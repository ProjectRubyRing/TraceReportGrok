"""Read CloudWatch Logs Insights API JSON without depending on jq."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def load_json(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"JSON を解釈できません: {path}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"JSON オブジェクトではありません: {path}")
    return data


def query_id(page: dict[str, Any]) -> str:
    value = str(page.get("queryId") or "").strip()
    if not value:
        raise ValueError("start-query の応答に queryId がありません")
    return value


def account_id(page: dict[str, Any]) -> str:
    return str(page.get("Account") or "").strip()


def log_group_present(page: dict[str, Any], name: str) -> int:
    """Return 0 when found, 1 when absent, 2 when the page is inconclusive."""
    groups = page.get("logGroups")
    if not isinstance(groups, list):
        return 2
    names = {str(item.get("logGroupName")) for item in groups if isinstance(item, dict)}
    if name in names:
        return 0
    if page.get("nextToken"):
        return 2
    return 1


def query_stats(page: dict[str, Any]) -> tuple[str, int, int, int, str]:
    status = str(page.get("status") or "").strip()
    stats = page.get("statistics") if isinstance(page.get("statistics"), dict) else {}
    token = str(page.get("nextToken") or "")
    if "\n" in token or "\r" in token:
        raise ValueError("nextToken に改行が含まれています")

    def number(key: str) -> int:
        value = stats.get(key)
        if value in (None, ""):
            return 0
        return int(float(value))

    return status, number("recordsMatched"), number("recordsScanned"), number("bytesScanned"), token


def load_result_rows(pages: list[Path]) -> tuple[list[dict[str, str]], dict[str, int]]:
    rows: list[dict[str, str]] = []
    stats = {"recordsMatched": 0, "recordsScanned": 0, "bytesScanned": 0}
    saw_stats = False
    for path in pages:
        page = load_json(path)
        if not saw_stats and isinstance(page.get("statistics"), dict):
            status, matched, scanned, size, _token = query_stats(page)
            del status
            stats = {
                "recordsMatched": matched,
                "recordsScanned": scanned,
                "bytesScanned": size,
            }
            saw_stats = True
        results = page.get("results")
        if not isinstance(results, list):
            continue
        for result in results:
            if not isinstance(result, list):
                continue
            row: dict[str, str] = {}
            for item in result:
                if not isinstance(item, dict):
                    continue
                field = item.get("field")
                if not isinstance(field, str) or not field:
                    continue
                value = item.get("value")
                row[field] = "" if value is None else str(value)
            if row:
                rows.append(row)
    return rows, stats


_PREFERRED = [
    "@timestamp",
    "@ingestionTime",
    "@logStream",
    "traceId",
    "spanId",
    "parentSpanId",
    "name",
    "kind",
    "durationNano",
    "startTimeUnixNano",
    "endTimeUnixNano",
    "statusCode",
    "awsLocalService",
    "awsRemoteService",
    "awsLocalOperation",
    "awsRemoteOperation",
    "httpMethod",
    "httpStatus",
    "httpRoute",
    "urlPath",
    "urlFull",
    "dbSystem",
    "dbStatement",
    "errorType",
    "resourceServiceName",
    "cloudRegion",
]


def result_columns(rows: list[dict[str, str]], *, include_ptr: bool) -> list[str]:
    seen: list[str] = []
    for row in rows:
        for key in row:
            if key not in seen:
                seen.append(key)
    if not include_ptr:
        seen = [key for key in seen if key != "@ptr"]
    preferred = [key for key in _PREFERRED if key in seen]
    rest = [key for key in seen if key not in preferred and key != "@message"]
    columns = preferred + rest
    if "@message" in seen:
        columns.append("@message")
    return columns
