#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pc_cache_cleaner.py — самостоятельная очистка кэша, временных файлов, корзины и ShellBag в Windows.

Каждый запуск пишет подробный отчёт: что именно удалено, сколько файлов и байт,
что пропущено и почему. Отчёты складываются в %LOCALAPPDATA%\\PC Cache Cleaner\\reports.

Примеры:
    py -3 pc_cache_cleaner.py --dry-run          # показать, что будет удалено (ничего не удаляет)
    py -3 pc_cache_cleaner.py                    # временные файлы + корзина + отчёт
    py -3 pc_cache_cleaner.py --quick            # temp + браузеры + корзина
    py -3 pc_cache_cleaner.py --shellbags        # очистить ShellBag (с бэкапом реестра)
    py -3 pc_cache_cleaner.py --all              # всё, включая системные кэши (нужны права админа)
    py -3 pc_cache_cleaner.py --report           # показать отчёт о последнем запуске
    py -3 pc_cache_cleaner.py --task install     # задача в Планировщике: чистка каждый день в 04:00
"""

from __future__ import annotations

import argparse
import ctypes
import glob
import logging
import os
import shutil
import stat
import subprocess
import sys
import time
import winreg
from dataclasses import dataclass, field
from datetime import datetime
from functools import lru_cache
from pathlib import Path

APP_NAME = "PC Cache Cleaner"
TASK_NAME = "PC Cache Cleaner"
LOG_DIR = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / APP_NAME
LOG_FILE = LOG_DIR / "clean.log"
REPORT_DIR = LOG_DIR / "reports"
BACKUP_DIR = LOG_DIR / "backups"
LAST_REPORT = REPORT_DIR / "last.txt"

# ─────────────────────────── консоль / утилиты ───────────────────────────

_COLORS = {
    "reset": "\033[0m", "dim": "\033[90m", "red": "\033[91m",
    "green": "\033[92m", "yellow": "\033[93m", "blue": "\033[94m", "cyan": "\033[96m",
}


def enable_unicode_console() -> None:
    """Чтобы русский текст не превращался в кракозябры в cmd/PowerShell."""
    try:
        if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    try:
        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass


def color(text: str, name: str, enabled: bool) -> str:
    return f"{_COLORS[name]}{text}{_COLORS['reset']}" if enabled else text


def human_size(num: float) -> str:
    num = float(num)
    if num <= 0:
        return "0 Б"
    for unit in ("Б", "КБ", "МБ", "ГБ", "ТБ"):
        if num < 1024 or unit == "ТБ":
            return f"{num:.0f} {unit}" if unit == "Б" else f"{num:.1f} {unit}"
        num /= 1024
    return f"{num:.1f} ТБ"


def human_count(num: int) -> str:
    return f"{num:,}".replace(",", " ")


def plural(num: int, one: str, few: str, many: str) -> str:
    """393 файла / 1 файл / 5 файлов — без ошибок склонения."""
    n = abs(num) % 100
    if 11 <= n <= 14:
        return many
    n %= 10
    if n == 1:
        return one
    if 2 <= n <= 4:
        return few
    return many


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def unlock(path: Path) -> None:
    """Снять read-only/системные/скрытые атрибуты (нужно для файлов Windows Update и т.п.)."""
    try:
        os.chmod(str(path), stat.S_IWRITE)
    except OSError:
        pass
    try:
        ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x80)  # FILE_ATTRIBUTE_NORMAL
    except Exception:
        pass


@lru_cache(maxsize=1)
def fixed_drives() -> tuple:
    """Список существующих локальных дисков (GetLogicalDrives + проверка типа)."""
    out = []
    try:
        mask = ctypes.windll.kernel32.GetLogicalDrives()
    except Exception:
        return (sysroot(),)
    for i in range(26):
        if not mask & (1 << i):
            continue
        letter = f"{chr(65 + i)}:\\"
        try:
            if ctypes.windll.kernel32.GetDriveTypeW(letter) == 3 and Path(letter).exists():
                out.append(Path(letter))
        except Exception:
            continue
    return tuple(out)


def sysroot() -> Path:
    """Корень системного диска: C:\\ (без слеша Path('C:') даёт путь относительно диска)."""
    return Path(os.environ.get("SystemDrive", "C:") + "\\")


def free_space(path: Path) -> int:
    try:
        return shutil.disk_usage(str(path)).free
    except OSError:
        return 0


def is_protected(path: Path) -> bool:
    """Стоп-лист: никогда не трогаем корень диска, профиль и пользовательские папки."""
    p = str(path.resolve() if path.exists() else path).lower().rstrip("\\/")
    home = str(Path.home()).lower().rstrip("\\/")
    system = str(sysroot()).lower().rstrip("\\/")
    drives = [str(d).lower().rstrip("\\/") for d in fixed_drives()]
    forbidden = {"downloads", "documents", "desktop", "pictures", "music", "videos",
                 "документы", "загрузки", "изображения", "музыка", "видео"}
    if p in drives or p in ("", "/") or p in (home, system):
        return True
    return any(part.lower() in forbidden for part in Path(p).parts[-2:])


# ──────────────────────────────── отчёт ────────────────────────────────

@dataclass
class Item:
    """Статистика по одной категории очистки."""
    title: str
    targets: int = 0          # сколько путей проверяли
    found: int = 0            # сколько из них существует
    bytes: int = 0            # освобождено / будет освобождено
    files: int = 0            # удалено файлов (или будет удалено в dry-run)
    errors: int = 0
    skipped: int = 0          # не существует / нет доступа
    seconds: float = 0.0
    note: str = ""
    _t0: float = field(default_factory=time.time, repr=False)

    def finish(self, note: str = "") -> None:
        self.seconds = time.time() - self._t0
        if note:
            self.note = note

    def summary(self, dry_run: bool, use_color: bool = True) -> str:
        parts = [f"{human_count(self.files)} {plural(self.files, 'файл', 'файла', 'файлов')}",
                 human_size(self.bytes)]
        if self.skipped:
            parts.append(f"пропущено путей: {self.skipped}")
        if self.errors:
            parts.append(f"ошибок: {self.errors}")
        line = f"  {self.title:<30} {' · '.join(parts)}"
        if self.note:
            line += f"\n      {self.note}"
        if self.seconds >= 0.5:
            line += color(f"  ({self.seconds:.1f} с)", "dim", use_color)
        return line


@dataclass
class Ctx:
    dry_run: bool = False
    cutoff: float | None = None          # файлы моложе этого времени не трогаем
    verbose: bool = False
    prefetch: bool = False               # чистить ли Prefetch
    resetbase: bool = False              # DISM /ResetBase
    restart_explorer: bool = False       # перезапускать проводник после чистки ShellBag
    gui_tool: bool = False               # запускать найденный ShellBag AnalyZer & Cleaner
    use_color: bool = True
    freed: int = 0
    files: int = 0
    errors: list = field(default_factory=list)
    items: list = field(default_factory=list)
    current: object = None
    big: list = field(default_factory=list)      # [(размер, путь)] крупнейшее удалённое
    info: dict = field(default_factory=dict)     # заметки для отчёта (shellbag и т.п.)

    def begin(self, title: str) -> Item:
        item = Item(title)
        self.items.append(item)
        self.current = item
        return item

    def err(self, path: Path | str, exc: Exception) -> None:
        if len(self.errors) < 300:          # больше не накапливаем — память не течёт
            self.errors.append(f"{path}: {exc}")
        if self.current is not None:
            self.current.errors += 1
        logging.warning("ошибка %s: %s", path, exc)

    def remember(self, size: int, path: Path) -> None:
        if size > 0 and len(self.big) < 5000:
            self.big.append((size, str(path)))


# ─────────────────────────────── удаление ───────────────────────────────

def path_size(path: Path, ctx: "Ctx | None" = None) -> int:
    """Размер папки/файла. Ошибки доступа не игнорируем, а пишем в отчёт."""
    total = 0

    def walk(folder: Path) -> None:
        nonlocal total
        try:
            with os.scandir(folder) as items:
                for entry in items:
                    try:
                        if entry.is_dir(follow_symlinks=False):
                            walk(Path(entry.path))
                        else:
                            total += entry.stat(follow_symlinks=False).st_size
                    except OSError as exc:
                        if ctx:
                            ctx.err(entry.path, exc)
        except OSError as exc:
            if ctx:
                ctx.err(folder, exc)

    try:
        if path.is_file() and not path.is_symlink():
            return path.stat().st_size
    except OSError as exc:
        if ctx:
            ctx.err(path, exc)
        return 0
    walk(path)
    return total


def _try_remove_file(path: Path) -> tuple:
    """Удалить файл, при отказе в доступе — снять атрибуты и повторить.
    Возвращает (True, size) либо (False, size, исключение)."""
    try:
        size = path.stat().st_size
    except OSError:
        size = 0
    try:
        os.remove(path)
        return True, size
    except PermissionError:
        pass
    except OSError:
        return False, size
    unlock(path)
    try:
        os.remove(path)
        return True, size
    except OSError as exc:
        return False, size, exc


def _rm(path: Path, ctx: Ctx) -> tuple:
    """Рекурсивное удаление с обходом read-only и занятых файлов.
    Возвращает (освобождённые байты, удалено файлов)."""
    freed = files = 0
    try:
        entries = list(os.scandir(path))
    except FileNotFoundError:
        return 0, 0
    except OSError:
        unlock(path)
        try:
            entries = list(os.scandir(path))
        except OSError as exc:
            ctx.err(path, exc)
            return 0, 0

    for entry in entries:
        child = Path(entry.path)
        try:
            is_dir = entry.is_dir(follow_symlinks=False)
        except OSError:
            continue

        if is_dir:
            sub_freed, sub_files = _rm(child, ctx)
            freed += sub_freed
            files += sub_files
            if ctx.dry_run:
                continue
            try:
                os.rmdir(child)
            except OSError:
                unlock(child)                     # read-only на папке
                try:
                    os.rmdir(child)
                except OSError as exc:
                    ctx.err(child, exc)
            continue

        try:
            info = entry.stat(follow_symlinks=False)
        except OSError:
            continue
        if ctx.cutoff is not None and info.st_mtime > ctx.cutoff:
            continue
        if ctx.dry_run:
            freed += info.st_size
            files += 1
            ctx.remember(info.st_size, child)
            continue

        result = _try_remove_file(child)
        if result[0]:
            freed += result[1]
            files += 1
            ctx.remember(result[1], child)
        else:
            ctx.err(child, result[2])
    return freed, files


def clean_path(path: Path | str, ctx: Ctx, label: str = "") -> int:
    """Удалить файл или содержимое папки. Пишет статистику в текущий Item отчёта."""
    p = Path(os.path.expandvars(str(path)))
    item = ctx.current
    if item is not None:
        item.targets += 1
    if not str(p):
        return 0
    if not p.exists():
        if item is not None:
            item.skipped += 1
        return 0
    if is_protected(p):
        logging.warning("защищённый путь, пропущен: %s", p)
        if item is not None:
            item.skipped += 1
        return 0
    if p.is_dir() and not p.is_symlink():
        try:
            next(os.scandir(p), None)              # сразу видно, есть ли доступ
        except PermissionError as exc:
            ctx.err(p, exc)
            if item is not None:
                item.skipped += 1
            return 0

    if item is not None:
        item.found += 1
    size = path_size(p, ctx)
    logging.info("%s: %s (%s)", "будет удалено" if ctx.dry_run else "удалено", p, human_size(size))
    if ctx.verbose:
        print(color(f"      {p}  [{human_size(size)}]", "dim", True))

    if ctx.dry_run:
        ctx.freed += size
        ctx.remember(size, p)
        planned = count_files(p)
        ctx.files += planned
        if item is not None:
            item.bytes += size
            item.files += planned
        return size

    errs_before = len(ctx.errors)
    try:
        if p.is_dir() and not p.is_symlink():
            freed, files = _rm(p, ctx)
            try:
                os.rmdir(p)                       # саму папку тоже, если освободилась
            except OSError:
                unlock(p)
                try:
                    os.rmdir(p)
                except OSError:
                    pass
            ctx.freed += freed
            ctx.files += files
            if item is not None:
                item.bytes += freed
                item.files += files
                item.errors += len(ctx.errors) - errs_before
            return freed
        if ctx.cutoff is not None and p.stat().st_mtime > ctx.cutoff:
            if item is not None:
                item.skipped += 1
            return 0
        ok, fsize, *rest = _try_remove_file(p)
        if not ok:
            raise rest[0]
        ctx.freed += fsize
        ctx.files += 1
        ctx.remember(fsize, p)
        if item is not None:
            item.bytes += fsize
            item.files += 1
        return fsize
    except OSError as exc:
        ctx.err(p, exc)
        if item is not None:
            item.errors += len(ctx.errors) - errs_before
        return 0


def count_files(path: Path, cap: int = 2_000_000) -> int:
    """Сколько файлов внутри (для dry-run, где удаления ещё не было)."""
    total = 0
    try:
        with os.scandir(path) as items:
            for entry in items:
                if total > cap:
                    break
                try:
                    if entry.is_dir(follow_symlinks=False):
                        total += count_files(Path(entry.path), cap - total)
                    else:
                        total += 1
                except OSError:
                    continue
    except OSError:
        return total
    return total


def find(pattern: str) -> list:
    """Развернуть шаблон с *, но без жадных рекурсивных обходов диска."""
    try:
        return [Path(p) for p in glob.glob(os.path.expandvars(pattern))]
    except Exception:
        return []


def stop_service(name: str, stop: bool) -> bool:
    """Остановить/запустить службу Windows. True — состояние изменилось."""
    try:
        out = subprocess.run(["sc", "query", name], capture_output=True, text=True, timeout=20)
        running = "RUNNING" in out.stdout
        if not running:
            return False
        subprocess.run(["net", "stop" if stop else "start", name, "/y"],
                       capture_output=True, timeout=90)
        logging.info("служба %s: %s", name, "stop" if stop else "start")
        return True
    except Exception as exc:
        logging.warning("служба %s: %s", name, exc)
        return False


def run_cmd(args: list, ctx: Ctx, timeout: int = 600, desc: str = "") -> bool:
    if ctx.dry_run:
        logging.info("пропущено (dry-run): %s", " ".join(args))
        return False
    try:
        res = subprocess.run(args, capture_output=True, timeout=timeout, text=True)
        logging.info("выполнено: %s (%s) rc=%s", " ".join(args), desc, res.returncode)
        return res.returncode == 0
    except Exception as exc:
        logging.warning("команда %s не выполнена: %s", " ".join(args), exc)
        return False


# ─────────────────────────────── чистильщики ───────────────────────────────
# Каждая функция: def name(ctx) -> None. Заполняет свой Item в отчёте.

def clean_temp(ctx: Ctx) -> None:
    """Временные файлы, логи, кэш шейдеров, краш-дампы"""
    local = Path(os.environ["LOCALAPPDATA"])
    sd = sysroot()
    progdata = Path(os.environ.get("ProgramData", r"C:\ProgramData"))
    groups = [
        ("TEMP пользователя", [
            Path(os.environ.get("TEMP", str(local / "Temp"))),
            Path(os.environ.get("TMP", "")),
            local / "Temp",
        ]),
        (r"Windows\Temp", [sd / r"Windows\Temp"]),
        ("Кэш шейдеров", [local / "D3DSCache"]
            + [local / gpu for gpu in ("NVIDIA/DXCache", "NVIDIA/GLCache", "NVIDIA/ComputeCache",
                                       "AMD/DxCache", "AMD/GLCache", "AMD/VkCache")]),
        ("Краш-дампы", [local / "CrashDumps", sd / r"Windows\Minidump", sd / r"Windows\MEMORY.DMP"]),
        ("Кэш интернета", [local / r"Microsoft\Windows\INetCache", local / r"Microsoft\Windows\WebCache"]),
        ("Системные логи", [
            sd / r"Windows\Logs\CBS", sd / r"Windows\Panther",
            *([progdata / r"Microsoft\Windows\WER\ReportQueue",
               progdata / r"Microsoft\Windows\WER\ReportArchive"] if is_admin() else []),
        ]),
        ("Кэш миниатюр и иконок",         # проводник перерисует их
         find(str(local / r"Microsoft\Windows\Explorer\thumbcache_*.db"))
         + find(str(local / r"Microsoft\Windows\Explorer\iconcache_*.db"))),
    ]

    seen: set = set()
    for title, paths in groups:
        item = ctx.begin(title)
        for path in paths:
            if path is None or not str(path):
                continue
            key = os.path.normcase(os.path.abspath(str(path)))
            if key in seen:
                continue
            seen.add(key)
            clean_path(path, ctx, title)
        item.finish()


CHROME_LIKE = [
    ("Google\\Chrome\\User Data", "Chrome"),
    ("Microsoft\\Edge\\User Data", "Edge"),
    ("BraveSoftware\\Brave-Browser\\User Data", "Brave"),
    ("Vivaldi\\User Data", "Vivaldi"),
    ("Opera Software\\Opera Stable", "Opera"),
    ("Yandex\\YandexBrowser\\User Data", "Яндекс.Браузер"),
]
CACHE_SUBDIRS = [
    "Cache\\Cache_Data", "Cache", "Code Cache", "GPUCache", "GrShaderCache",
    "ShaderCache", "DawnCache", "DawnGraphiteCache", "DawnWebGPUCache",
    "Service Worker\\CacheStorage", "Service Worker\\ScriptCache",
    "blob_storage", "Media Cache", "optimization_guide_model_store",
]


def drop_nested(paths: list) -> list:
    """Убирает пути, вложенные в другие пути из списка — иначе одни и те же
    байты (Cache и Cache\\Cache_Data) посчитались бы дважды."""
    norm = [os.path.normcase(os.path.abspath(str(p))) for p in paths]
    result = []
    for i, p in enumerate(paths):
        me = norm[i]
        covered = any(
            j != i and len(norm[j]) < len(me) and me.startswith(norm[j].rstrip("\\") + "\\")
            for j in range(len(paths))
        )
        if not covered:
            result.append(p)
    return result


def clean_browsers(ctx: Ctx) -> None:
    """Кэш браузеров"""
    local = Path(os.environ["LOCALAPPDATA"])
    for rel, name in CHROME_LIKE:
        base = local / rel
        if not base.exists():
            continue
        item = ctx.begin(name)
        candidates = []
        for prof in [base / "Default"] + sorted(base.glob("Profile *")) + [base]:
            candidates += [prof / sub for sub in CACHE_SUBDIRS]
        for path in drop_nested(candidates):
            clean_path(path, ctx)
        profiles = [base / "Default"] + sorted(base.glob("Profile *"))
        item.finish(f"профилей: {len(profiles)}, проверено папок: {item.targets}")

    ff = local / r"Mozilla\Firefox\Profiles"
    if ff.exists():
        item = ctx.begin("Firefox")
        for prof in ff.glob("*"):
            clean_path(prof / "cache2", ctx)
            clean_path(prof / "startupCache", ctx)
            storage = prof / "storage" / "default"
            for entry in storage.glob("*/cache2") if storage.exists() else []:
                clean_path(entry, ctx)
        item.finish()


def clean_system(ctx: Ctx) -> None:
    """Кэш Windows Update, Delivery Optimization, Prefetch (нужны права админа)"""
    item = ctx.begin("Windows Update и доставка")
    if not is_admin():
        item.finish("пропущено — нужны права администратора")
        return
    sd = sysroot()
    progdata = Path(os.environ.get("ProgramData", r"C:\ProgramData"))
    stop_service("wuauserv", True)
    stop_service("bits", True)
    stop_service("dosvc", True)
    clean_path(sd / r"Windows\SoftwareDistribution\Download", ctx)
    clean_path(sd / r"Windows\SoftwareDistribution\DataStore\Logs", ctx)
    clean_path(progdata / r"Microsoft\Windows\DeliveryOptimization\Cache", ctx)
    clean_path(sd / "$WINDOWS.~BT", ctx)
    clean_path(sd / "$WINDOWS.~WS", ctx)
    clean_path(sd / r"Windows\MEMORY.DMP", ctx)
    if ctx.prefetch:
        clean_path(sd / r"Windows\Prefetch", ctx)
    stop_service("dosvc", False)
    stop_service("bits", False)
    stop_service("wuauserv", False)
    stopped = "службы wuauserv/bits/dosvc были остановлены и запущены обратно"
    item.finish(stopped)


def recycle_size() -> tuple:
    """(размер корзины в байтах, число файлов) по всем дискам."""
    total = files = 0
    for drive in fixed_drives():
        root = drive / "$Recycle.Bin"
        try:
            if not root.exists():
                continue
            for sid_dir in root.iterdir():
                try:
                    for entry in sid_dir.iterdir():
                        if entry.name.startswith("$R"):
                            total += path_size(entry)
                            files += 1
                except OSError:
                    continue
        except OSError:
            continue
    return total, files


def clean_recycle(ctx: Ctx) -> None:
    """Очистить корзину на всех дисках"""
    item = ctx.begin("Корзина")
    size, files = recycle_size()
    item.found = 1 if files else 0
    item.bytes = size
    item.files = files
    if not files:
        item.finish("корзина уже пуста")
        return
    if ctx.dry_run:
        item.finish(f"будет очищена при следующем запуске без --dry-run")
        ctx.freed += size
        ctx.files += files
        return
    ok = ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x00000001 | 0x00000002 | 0x00000004)
    left_size, left_files = recycle_size()
    freed = max(0, size - left_size)
    ctx.freed += freed
    item.bytes = freed
    item.files = files - left_files
    item.note = "очищена (SHEmptyRecycleBin)" if ok is not None else "очищена"
    item.finish(f"осталось в корзине: {left_files} файлов ({human_size(left_size)})")


def clean_network(ctx: Ctx) -> None:
    """Сброс DNS-кэша и ARP-кэша"""
    item = ctx.begin("DNS / ARP кэш")
    ok_dns = run_cmd(["ipconfig", "/flushdns"], ctx, desc="сброс DNS")
    ok_arp = run_cmd(["arp", "-d", "*"], ctx, desc="сброс ARP")
    item.found = 1
    if ctx.dry_run:
        item.finish("будет сброшен")
    else:
        item.finish("DNS: " + ("ок" if ok_dns else "не требуется/нет прав") +
                    ", ARP: " + ("ок" if ok_arp else "не требуется/нет прав"))


def clean_component_store(ctx: Ctx) -> None:
    """DISM: очистка хранилища компонентов (медленно, нужно разрешение на перезагрузку)"""
    item = ctx.begin("Хранилище компонентов")
    if not is_admin():
        item.finish("пропущено — нужны права администратора")
        return
    if ctx.dry_run:
        item.finish("DISM пропущен в режиме проверки")
        return
    cmd = ["dism", "/Online", "/Cleanup-Image", "/StartComponentCleanup"]
    if ctx.resetbase:
        cmd.append("/ResetBase")
        item.note = "/ResetBase: откатить текущие обновления будет нельзя"
    res = subprocess.run(cmd, capture_output=True, text=True, timeout=3600)
    item.found = 1
    logging.info("DISM rc=%s", res.returncode)
    item.finish(f"DISM завершился с кодом {res.returncode}")


def clean_store_apps(ctx: Ctx) -> None:
    """LocalCache приложений Microsoft Store"""
    item = ctx.begin("Приложения Microsoft Store")
    pkgs = Path(os.environ["LOCALAPPDATA"]) / "Packages"
    if not pkgs.exists():
        item.finish("папка Packages не найдена")
        return
    for pkg in pkgs.glob("*/LocalCache"):
        clean_path(pkg, ctx)
    item.finish()


# ──────────────────────────────── ShellBag ────────────────────────────────

SHELLBAG_KEYS = [
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\BagMRU", "BagMRU — история проводника"),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\Bags", "Bags — метаданные папок"),
    (winreg.HKEY_CURRENT_USER, r"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags",
     "Bags (системные)"),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\RecentDocs", "RecentDocs — недавние файлы"),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\ComDlg32\OpenSavePidlMRU",
     "OpenSavePidlMRU — окна открыть/сохранить"),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\ComDlg32\LastVisitedPidlMRU",
     "LastVisitedPidlMRU"),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\ComDlg32\CIDSizeMRU", "CIDSizeMRU"),
    (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\WordWheelQuery", "WordWheelQuery — поиск в проводнике"),
]


def reg_count(root, path: str) -> int:
    """Сколько значений и подключей в ветке реестра (рекурсивно)."""
    try:
        key = winreg.OpenKey(root, path, 0, winreg.KEY_READ)
    except OSError:
        return 0
    count = 0
    i = 0
    while True:
        try:
            sub = winreg.EnumKey(key, i)
            count += 1 + reg_count(root, path + "\\" + sub)
        except OSError:
            break
        i += 1
    i = 0
    while True:
        try:
            winreg.EnumValue(key, i)
            count += 1
        except OSError:
            break
        i += 1
    winreg.CloseKey(key)
    return count


def reg_export(root, path: str, dest: Path) -> bool:
    """Бэкап ветки в .reg (reg export работает только с целым разделом)."""
    if dest.exists():
        return False
    prefix = {winreg.HKEY_CURRENT_USER: "HKCU", winreg.HKEY_LOCAL_MACHINE: "HKLM",
              winreg.HKEY_CLASSES_ROOT: "HKCR"}.get(root, "HKCU")
    try:
        res = subprocess.run(["reg", "export", f"{prefix}\\{path}", str(dest), "/y"],
                             capture_output=True, text=True, timeout=30)
        return res.returncode == 0
    except Exception as exc:
        logging.warning("reg export %s: %s", path, exc)
        return False


def reg_delete_tree(root, path: str) -> int:
    """Рекурсивно удалить ветку реестра. Возвращает число удалённых записей."""
    try:
        key = winreg.OpenKey(root, path, 0, winreg.KEY_ALL_ACCESS)
    except FileNotFoundError:
        return 0
    except OSError as exc:
        logging.warning("нет доступа к %s: %s", path, exc)
        return 0

    count = 0
    while True:                                  # всегда удаляем первый подключ
        try:
            sub = winreg.EnumKey(key, 0)
        except OSError:
            break
        count += reg_delete_tree(root, path + "\\" + sub)
    while True:
        try:
            name, _value, _type = winreg.EnumValue(key, 0)
            winreg.DeleteValue(key, name)
            count += 1
        except OSError:
            break
    winreg.CloseKey(key)
    try:
        winreg.DeleteKey(root, path)
    except OSError as exc:
        logging.warning("не удалось удалить ветку %s: %s", path, exc)
    return count


def find_shellbag_tool() -> Path | None:
    """Ищет GUI-утилиту ShellBag AnalyZer & Cleaner, если она скачана."""
    names = ("shellbag_analyzer_cleaner.exe", "ShellBagAnalyZer.exe",
             "shellbaganalyzer.exe", "ShellBagAnalyzer.exe")
    folders = [Path.home() / "Downloads", Path.home() / "Desktop",
               Path(__file__).resolve().parent, sysroot() / "Tools",
               Path(os.environ.get("ProgramFiles", r"C:\Program Files"))]
    for folder in folders:
        for name in names:
            cand = folder / name
            if cand.exists():
                return cand
    for folder in folders:
        if not folder.exists():
            continue
        try:
            for cand in folder.glob("*shellbag*.exe"):
                return cand
        except OSError:
            continue
    return None


def restart_explorer() -> bool:
    try:
        subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], capture_output=True, timeout=60)
        time.sleep(2)
        subprocess.Popen(["explorer.exe"], shell=False)
        logging.info("проводник перезапущен")
        return True
    except Exception as exc:
        logging.warning("не удалось перезапустить проводник: %s", exc)
        return False


def clean_shellbags(ctx: Ctx) -> None:
    """Чистка ShellBag: история проводника, недавние файлы, MRU окон (с бэкапом)"""
    item = ctx.begin("ShellBag (следы в реестре)")
    total_entries = 0
    per_key = []
    for root, path, title in SHELLBAG_KEYS:
        entries = reg_count(root, path)
        if not entries:
            per_key.append(f"{title}: не найдено")
            continue
        if ctx.dry_run:
            total_entries += entries
            per_key.append(f"{title}: записей {entries} (будет удалено)")
            continue
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = BACKUP_DIR / f"shellbags_{stamp}_{Path(path).name}.reg"
        backed_up = reg_export(root, path, backup)
        removed = reg_delete_tree(root, path)
        total_entries += removed
        per_key.append(f"{title}: удалено записей {removed}"
                       + (f", бэкап {backup.name}" if backed_up else ", бэкап не создан"))
        logging.info("ShellBag %s: удалено %s записей", path, removed)

    tool = find_shellbag_tool()
    notes = []
    if tool:
        notes.append(f"найдена утилита: {tool}")
        if ctx.gui_tool and not ctx.dry_run:
            try:
                subprocess.Popen([str(tool)])
                notes.append("утилита запущена вручную")
            except Exception as exc:
                notes.append(f"не удалось запустить утилиту: {exc}")
    else:
        notes.append("ShellBag AnalyZer & Cleaner не найден (чистка сделана напрямую через реестр)")

    if ctx.restart_explorer and not ctx.dry_run and total_entries:
        notes.append("проводник: " + ("перезапущен" if restart_explorer() else "не перезапущен"))
    elif total_entries:
        notes.append("проводник: перезапусти сам, иначе старые записи останутся в памяти "
                     "(или запусти с --restart-explorer)")

    item.found = 1
    ctx.info["Записей ShellBag в реестре"] = f"{total_entries} (бэкапы: {BACKUP_DIR})"
    item.finish("; ".join(per_key[:3]) + ("; …" if len(per_key) > 3 else "") + "\n      " + "\n      ".join(notes))


# ──────────────────────────── планировщик Windows ────────────────────────────

def task_ps(script_args: str, at_time: str) -> list:
    """Собирает команду PowerShell для регистрации задачи от имени SYSTEM."""
    esc = lambda s: s.replace("'", "''")  # noqa: E731
    return [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
        (
            f"$a = New-ScheduledTaskAction -Execute '{esc(sys.executable)}' -Argument '{esc(script_args)}' "
            f"-WorkingDirectory '{esc(str(Path(__file__).resolve().parent))}'; "
            f"$t = New-ScheduledTaskTrigger -Daily -At '{at_time}'; "
            "$p = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest; "
            "$s = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries "
            "-StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2); "
            f"Register-ScheduledTask -TaskName '{TASK_NAME}' -Action $a -Trigger $t -Principal $p "
            "-Settings $s -Description 'Ежедневная очистка кэша, корзины и временных файлов' -Force "
            f"| Out-Null; Write-Output 'OK'"
        ),
    ]


def task_install(args) -> int:
    if not is_admin():
        print(color("Для установки задачи нужны права администратора.", "red", True))
        print("Запусти PowerShell/cmd от имени администратора и повтори.")
        return 2
    flags = "--all --yes --shellbags --restart-explorer" if args.shellbags else "--all --yes"
    script_args = f'"{Path(__file__).resolve()}" {flags}'
    drive = os.path.splitdrive(str(Path(sys.executable)))[0] or "C:"
    try:
        drive_type = ctypes.windll.kernel32.GetDriveTypeW(f"{drive}\\")  # 3 = фиксированный
    except Exception:
        drive_type = 3
    if drive_type != 3:
        print(color(f"⚠ Диск {drive} не локальный — под SYSTEM задача может не запуститься.", "yellow", True))
    res = subprocess.run(task_ps(script_args, args.task_time), capture_output=True, text=True)
    if "OK" in (res.stdout or ""):
        print(color(f"Задача «{TASK_NAME}» установлена: ежедневно в {args.task_time} от имени SYSTEM.", "green", True))
        print(f"Команда задачи: {flags}")
        print(f"Лог: {LOG_FILE}\nОтчёты: {REPORT_DIR}")
        print(f'Проверить: py -3 "{Path(__file__).resolve()}" --report')
        return 0
    print(color("Не удалось создать задачу:", "red", True))
    print((res.stdout or "") + (res.stderr or ""))
    return 1


def task_remove(_args) -> int:
    res = subprocess.run(["schtasks", "/Delete", "/TN", TASK_NAME, "/F"],
                         capture_output=True, text=True)
    if res.returncode == 0:
        print(color(f"Задача «{TASK_NAME}» удалена.", "green", True))
        return 0
    print(color("Задача не найдена или удаление не удалось.", "yellow", True))
    return 1


def task_status(_args) -> int:
    res = subprocess.run(["schtasks", "/Query", "/TN", TASK_NAME, "/V", "/FO", "LIST"],
                         capture_output=True, text=True)
    if res.returncode != 0:
        print(color(f"Задача «{TASK_NAME}» не найдена. Создать: --task install", "yellow", True))
        return 1
    print(res.stdout)
    return 0


def task_run() -> int:
    res = subprocess.run(["schtasks", "/Run", "/TN", TASK_NAME], capture_output=True, text=True)
    if res.returncode == 0:
        print(color("Задача запущена.", "green", True))
    else:
        print(color("Не удалось запустить задачу.", "red", True), (res.stderr or "").strip())
    return res.returncode


# ──────────────────────────────── отчёт (вывод) ────────────────────────────────

def print_report(ctx: Ctx, started: float, free_before: int, free_after: int, use_color: bool) -> str:
    """Печатает сводку в консоль и возвращает тот же текст для файла отчёта."""
    lines: list[str] = []
    out = lines.append
    verb = "БУДЕТ ОЧИЩЕНО" if ctx.dry_run else "ОЧИЩЕНО"

    out("")
    out(color(f"ОТЧЁТ {datetime.now():%Y-%m-%d %H:%M:%S} — {verb}", "blue", use_color))
    out("─" * 72)
    for item in ctx.items:
        if item.targets == 0 and not item.note and not item.bytes:
            continue
        for line in item.summary(ctx.dry_run, ctx.use_color).splitlines():
            out(line)
    out("─" * 72)
    skipped = sum(i.skipped for i in ctx.items)
    errors = sum(i.errors for i in ctx.items)
    out(f"  ВСЕГО: {human_count(ctx.files)} {plural(ctx.files, 'файл', 'файла', 'файлов')}"
        f" · {human_size(ctx.freed)}"
        + (f" · пропущено путей: {skipped}" if skipped else "")
        + (f" · ошибок: {errors}" if errors else "")
        + f" · время: {time.time() - started:.1f} с")
    if free_before and not ctx.dry_run:
        delta = free_after - free_before
        out(f"  Свободно на {sysroot().drive}: было {human_size(free_before)} → "
            f"стало {human_size(free_after)} ({'+' if delta >= 0 else ''}{human_size(abs(delta))})")

    top = sorted(ctx.big, key=lambda t: -t[0])[:8]
    if top:
        out("")
        out(color("  Крупнейшее удалённое:", "cyan", use_color))
        for size, path in top:
            out(f"    {human_size(size):>10}  {path}")

    if ctx.info:
        out("")
        for key, value in ctx.info.items():
            out(f"  {key}: {value}")

    if ctx.errors:
        groups: dict = {}
        for err in ctx.errors:
            key = err.split(": ", 1)[0].rstrip("\\")     # ":" встречается в путях Windows
            p = Path(key)
            if "." in p.name:
                p = p.parent
            parts = p.parts
            short = str(Path(*parts[:7])) if len(parts) > 8 else str(p)
            groups[short] = groups.get(short, 0) + 1
        out("")
        out(color("  Пропущено (занято программами или нет прав):", "yellow", use_color))
        for path, count in sorted(groups.items(), key=lambda kv: -kv[1])[:12]:
            out(f"    · {path}" + (f"  ×{count}" if count > 1 else ""))
    out("")
    return "\n".join(lines)


def save_report(text: str, started: float) -> Path | None:
    try:
        REPORT_DIR.mkdir(parents=True, exist_ok=True)
        path = REPORT_DIR / f"{datetime.now():%Y-%m-%d_%H-%M}.txt"
        header = (f"{APP_NAME} — отчёт\n"
                  f"Компьютер: {os.environ.get('COMPUTERNAME', '?')}\n"
                  f"Пользователь: {os.environ.get('USERNAME', '?')}\n"
                  f"Права: {'администратор' if is_admin() else 'обычные'}\n"
                  f"Время запуска: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
                  f"Длительность: {time.time() - started:.1f} с\n"
                  f"{'─' * 72}\n")
        plain = text.replace("\033[90m", "").replace("\033[0m", "")
        for code in _COLORS.values():
            plain = plain.replace(code, "")
        path.write_text(header + plain + "\n", encoding="utf-8")
        LAST_REPORT.write_text(header + plain + "\n", encoding="utf-8")
        return path
    except Exception as exc:
        logging.warning("не удалось сохранить отчёт: %s", exc)
        return None


def show_last_report() -> int:
    if not LAST_REPORT.exists():
        print(f"Отчётов пока нет ({LAST_REPORT} не найден).")
        return 1
    print(LAST_REPORT.read_text(encoding="utf-8"))
    return 0


# ──────────────────────────────── основной код ────────────────────────────────

def setup_logging() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        if LOG_FILE.exists() and LOG_FILE.stat().st_size > 2 * 1024 * 1024:
            LOG_FILE.unlink(missing_ok=True)
    except OSError:
        pass
    logging.basicConfig(
        filename=str(LOG_FILE), encoding="utf-8", level=logging.DEBUG,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    logging.info("=== %s: запуск (python %s, pid %s) ===", APP_NAME,
                 sys.version.split()[0], os.getpid())


def confirm(question: str, auto_yes: bool) -> bool:
    if auto_yes:
        return True
    try:
        answer = input(color(question + " [y/N] ", "cyan", True)).strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    return answer in ("y", "yes", "д", "да")


def build_plan(args) -> list:
    plan = []
    if args.temp:
        plan.append(("Временные файлы и логи", clean_temp))
    if args.browsers:
        plan.append(("Кэш браузеров", clean_browsers))
    if args.store:
        plan.append(("Кэш приложений Microsoft Store", clean_store_apps))
    if args.recycle:
        plan.append(("Корзина", clean_recycle))
    if args.system:
        plan.append(("Системные кэши (нужен админ)", clean_system))
    if args.shellbags:
        plan.append(("ShellBag: следы в реестре", clean_shellbags))
    if args.network:
        plan.append(("Сетевые кэши", clean_network))
    if args.dism:
        plan.append(("Хранилище компонентов (DISM, долго)", clean_component_store))
    return plan


def main() -> int:
    enable_unicode_console()
    parser = argparse.ArgumentParser(
        description="Очистка кэша, корзины и ShellBag в Windows с подробным отчётом.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Пустой вызов (без флагов) = временные файлы + корзина + отчёт.\n"
               "Запуск от имени администратора открывает доступ к системным кэшам.",
    )
    parser.add_argument("--dry-run", action="store_true", help="только показать, ничего не удалять")
    parser.add_argument("--quick", action="store_true", help="temp + браузеры + корзина")
    parser.add_argument("--all", action="store_true", help="все категории, включая системные (нужен админ)")
    parser.add_argument("--temp", action="store_true", help="временные файлы, логи, кэш шейдеров")
    parser.add_argument("--browsers", action="store_true",
                        help="кэш Chrome/Edge/Firefox/Brave/Vivaldi/Opera/Яндекс")
    parser.add_argument("--store", action="store_true", help="LocalCache приложений Microsoft Store")
    parser.add_argument("--system", action="store_true",
                        help="Windows Update, Delivery Optimization, Prefetch")
    parser.add_argument("--network", action="store_true", help="сброс DNS и ARP кэша")
    parser.add_argument("--dism", action="store_true", help="очистка хранилища компонентов через DISM (медленно)")
    parser.add_argument("--resetbase", action="store_true", help="с --dism: запретить откат текущих обновлений")
    parser.add_argument("--prefetch", action="store_true", help="очистить Prefetch (с --system)")
    parser.add_argument("--recycle", action="store_true", help="очистить корзину на всех дисках")
    parser.add_argument("--shellbags", action="store_true",
                        help="очистить ShellBag: BagMRU/Bags/RecentDocs/MRU (с бэкапом в реестр)")
    parser.add_argument("--restart-explorer", action="store_true",
                        help="перезапустить проводник после очистки ShellBag")
    parser.add_argument("--with-gui-tool", action="store_true",
                        help="после очистки ShellBag запустить найденный ShellBag AnalyZer & Cleaner")
    parser.add_argument("--max-age-days", type=float, default=None,
                        help="удалять только файлы старше N дней (по умолчанию — все)")
    parser.add_argument("--no-report", action="store_true", help="не сохранять отчёт в файл")
    parser.add_argument("--report", action="store_true", help="показать отчёт о последнем запуске")
    parser.add_argument("-y", "--yes", action="store_true", help="без подтверждения")
    parser.add_argument("-v", "--verbose", action="store_true", help="показывать каждый путь")
    parser.add_argument("--no-color", action="store_true", help="без цветов в выводе")
    parser.add_argument("--task", choices=["install", "remove", "status", "run"],
                        help="управление задачей Планировщика: install | remove | status | run")
    parser.add_argument("--task-time", default="04:00", help="время ежедневного запуска (по умолчанию 04:00)")
    args = parser.parse_args()

    use_color = not args.no_color and sys.stdout.isatty() and not os.environ.get("NO_COLOR")
    setup_logging()

    if args.report:
        return show_last_report()
    if args.task:
        return {"install": task_install, "remove": task_remove,
                "status": task_status, "run": lambda _: task_run()}[args.task](args)

    # Набор категорий
    if args.all:
        args.temp = args.browsers = args.store = args.recycle = True
        args.network = args.dism = args.system = args.shellbags = True
    elif args.quick:
        args.temp = args.browsers = args.recycle = True
    if not any([args.temp, args.browsers, args.store, args.recycle, args.system,
                args.network, args.dism, args.shellbags]):
        args.temp = args.recycle = True          # по умолчанию: temp + корзина

    ctx = Ctx(dry_run=args.dry_run, verbose=args.verbose, prefetch=args.prefetch,
              resetbase=args.resetbase, restart_explorer=args.restart_explorer,
              gui_tool=args.with_gui_tool, use_color=use_color)
    if args.max_age_days:
        ctx.cutoff = time.time() - args.max_age_days * 86400

    admin = is_admin()
    started = time.time()
    free_before = free_space(sysroot())

    print(color(f"\n{APP_NAME}", "blue", use_color) + f"  ·  {datetime.now():%Y-%m-%d %H:%M:%S}")
    print(f"Компьютер: {os.environ.get('COMPUTERNAME', '?')}  ·  пользователь: {os.environ.get('USERNAME', '?')}")
    print(f"Режим: {'ПРОВЕРКА (ничего не удаляется)' if args.dry_run else 'ОЧИСТКА'}"
          f"  ·  Права: {'администратор' if admin else 'обычные'}"
          f"  ·  Возраст файлов: {'все' if not args.max_age_days else f'старше {args.max_age_days:g} дн.'}")

    if (args.system or args.dism) and not admin:
        print(color("⚠ Часть системных кэшей пропущена — запусти от имени администратора.", "yellow", use_color))

    plan = build_plan(args)
    print("\n" + color("План очистки", "cyan", use_color))
    for title, _ in plan:
        print(f"  • {title}")
    if args.shellbags:
        print(color("  ⚠ очистка ShellBag сбросит списки «недавних файлов» и папок в проводнике",
                    "yellow", use_color))
    if args.recycle and not args.dry_run:
        bin_size, bin_files = recycle_size()
        if bin_files:
            print(color(f"  ⚠ в корзине {human_count(bin_files)} "
                        f"{plural(bin_files, 'файл', 'файла', 'файлов')} ({human_size(bin_size)}) — "
                        "будет удалено безвозвратно", "yellow", use_color))
    if not args.dry_run and not confirm("\nНачать очистку?", args.yes):
        print("Отменено.")
        return 0

    for title, func in plan:
        try:
            func(ctx)
        except KeyboardInterrupt:
            print(color("\nПрервано пользователем.", "yellow", use_color))
            break
        except Exception as exc:
            ctx.err(title, exc)
            print(color(f"  ошибка в «{title}»: {exc}", "red", use_color))

    free_after = free_space(sysroot())
    text = print_report(ctx, started, free_before, free_after, use_color)
    print(text)

    if args.no_report:
        print(color(f"Отчёт не сохранён (--no-report). Лог: {LOG_FILE}\n", "dim", use_color))
    else:
        saved = save_report(text, started)
        print(color(f"Отчёт сохранён: {saved}", "green", use_color))
        print(color(f"Лог: {LOG_FILE}", "dim", use_color))
        print(color(f"Посмотреть позже: py -3 \"{Path(__file__).resolve()}\" --report", "dim", use_color))
        print()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        sys.exit(130)