"""Time parsing for the Logs Insights window and Excel display."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone, tzinfo

from zoneinfo import ZoneInfo


_RELATIVE = re.compile(r"^-(\d+)([smhd])$")
_EPOCH_S = re.compile(r"^\d{10}$")
_EPOCH_MS = re.compile(r"^\d{13}$")
_OFFSET = re.compile(r"^([+-])(\d{2}):(\d{2})$")
_UNITS = {"s": 1, "m": 60, "h": 3600, "d": 86400}


def parse_zone(name: str) -> tzinfo:
    text = (name or "UTC").strip()
    if text.upper() in {"UTC", "Z", "+00:00"}:
        return timezone.utc
    offset = _OFFSET.fullmatch(text)
    if offset:
        sign = 1 if offset.group(1) == "+" else -1
        minutes = sign * (int(offset.group(2)) * 60 + int(offset.group(3)))
        return timezone(timedelta(minutes=minutes))
    try:
        return ZoneInfo(text)
    except Exception as exc:
        raise ValueError(
            f"未知のタイムゾーンです: {text}（RHEL の tzdata にある名前、UTC、または +09:00 形式）"
        ) from exc


def parse_time(value: str, now: datetime | None = None) -> int:
    """Return a UTC epoch second.

    Accepted forms are ``now``, a negative relative duration (``-15m``,
    ``-1h``, ``-2d``, ``-30s``), 10-digit epoch seconds, 13-digit epoch
    milliseconds, and ISO-8601. A timezone-naive ISO value is UTC.
    """
    text = (value or "").strip()
    if not text:
        raise ValueError("時刻が空です")
    current = now or datetime.now(timezone.utc)
    if current.tzinfo is None:
        current = current.replace(tzinfo=timezone.utc)
    if text == "now":
        return int(current.timestamp())
    relative = _RELATIVE.fullmatch(text)
    if relative:
        amount = int(relative.group(1))
        return int(current.timestamp()) - amount * _UNITS[relative.group(2)]
    if _EPOCH_S.fullmatch(text):
        return int(text)
    if _EPOCH_MS.fullmatch(text):
        return int(text) // 1000
    iso = text.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(iso)
    except ValueError as exc:
        raise ValueError(
            "時刻の形式が不正です。UNIX秒、ISO-8601、now、-15m、-1h、-1d を指定してください: "
            + text
        ) from exc
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return int(parsed.timestamp())


def format_epoch(epoch: int, zone: tzinfo) -> str:
    moment = datetime.fromtimestamp(int(epoch), timezone.utc).astimezone(zone)
    return moment.strftime("%Y-%m-%d %H:%M:%S %z")


def format_unix_nano(nanos: int | None, zone: tzinfo) -> str:
    if nanos is None:
        return ""
    seconds, remainder = divmod(int(nanos), 1_000_000_000)
    moment = datetime.fromtimestamp(seconds, timezone.utc).astimezone(zone)
    millis = remainder // 1_000_000
    return moment.strftime("%Y-%m-%d %H:%M:%S.") + f"{millis:03d} " + moment.strftime("%z")


def window_or_raise(start: int, end: int) -> None:
    if end <= start:
        raise ValueError(f"終了時刻は開始時刻より後である必要があります (start={start}, end={end})")
    # Logs Insights accepts wide windows. Callers warn above one day because
    # aws/spans is billed by scanned bytes.
    if end - start > 366 * 86400:
        raise ValueError("時間範囲が 366 日を超えています")


def human_bytes(size: int) -> str:
    value = float(size)
    for unit in ("B", "KiB", "MiB", "GiB", "TiB"):
        if value < 1024 or unit == "TiB":
            if unit == "B":
                return f"{int(value)} {unit}"
            return f"{value:.2f} {unit}"
        value /= 1024
    return f"{int(size)} B"


