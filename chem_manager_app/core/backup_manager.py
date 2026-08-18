"""Narrowly-scoped ChemicalList backup retention utilities."""

from __future__ import annotations

import calendar
import datetime as dt
import os
from pathlib import Path
import shutil

from core.config_manager import get_app_root


BACKUP_PREFIX = "ChemicalList_backup_"
BACKUP_SUFFIX = ".xlsx"


def get_backup_dir(app_root: str | None = None) -> Path:
    """Return the single program-local backup directory used by all features."""
    return Path(app_root or get_app_root()).resolve() / "backup"


def create_backup(source_file: str, now: dt.datetime | None = None, app_root: str | None = None) -> str:
    """Create a timestamped backup below the program directory."""
    timestamp = now or dt.datetime.now()
    backup_dir = get_backup_dir(app_root)
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_name = f"{BACKUP_PREFIX}{timestamp:%Y%m%d_%H%M%S_%f}{BACKUP_SUFFIX}"
    backup_path = backup_dir / backup_name
    shutil.copy2(source_file, backup_path)
    # copy2 carries the source workbook's mtime. Retention must use backup
    # creation time so a newly-created backup is never treated as expired.
    created_at = timestamp.timestamp()
    os.utime(backup_path, (created_at, created_at))
    return str(backup_path)


def subtract_calendar_months(value: dt.datetime, months: int) -> dt.datetime:
    total = value.year * 12 + (value.month - 1) - months
    year, month_zero = divmod(total, 12)
    month = month_zero + 1
    day = min(value.day, calendar.monthrange(year, month)[1])
    return value.replace(year=year, month=month, day=day)


def purge_expired_backups(
    months: int = 3,
    now: dt.datetime | None = None,
    app_root: str | None = None,
):
    """Delete only matching backup files older than the calendar-month cutoff."""
    if months < 1:
        raise ValueError("months must be at least 1")
    backup_dir = get_backup_dir(app_root)
    if not backup_dir.is_dir():
        return [], []

    cutoff = subtract_calendar_months(now or dt.datetime.now(), months)
    deleted, failures = [], []
    for path in backup_dir.iterdir():
        if (
            not path.is_file()
            or not path.name.startswith(BACKUP_PREFIX)
            or not path.name.lower().endswith(BACKUP_SUFFIX)
        ):
            continue
        try:
            if dt.datetime.fromtimestamp(path.stat().st_mtime) <= cutoff:
                os.remove(path)
                deleted.append(str(path))
        except OSError as error:
            failures.append((str(path), str(error)))
    return deleted, failures
