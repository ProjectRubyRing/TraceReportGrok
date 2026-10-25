"""CLI used by export-span-traces.sh.

The shell script owns the AWS CLI calls. This module parses times, builds the
Logs Insights query, reads query-result JSON, and fills the Excel template.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Sequence

from spanreport import __version__
from spanreport.awsio import (
    account_id,
    load_json,
    load_result_rows,
    log_group_present,
    query_id,
    query_stats,
    result_columns,
)
from spanreport.model import (
    STATUS_ERROR,
    STATUS_FAULT,
    STATUS_OK,
    STATUS_THROTTLE,
    build_report,
    csv_text,
    detail_row,
    timeline_row,
    trace_row,
)
from spanreport.query import build_query
from spanreport.sample import BASE, sample_page
from spanreport.timeutil import format_epoch, human_bytes, parse_time, parse_zone
from spanreport.xlsxutil import (
    DETAIL_COLUMNS,
    SERVICE_COLUMNS,
    TIMELINE_COLUMNS,
    TRACE_COLUMNS,
    build_template,
    fill_template,
)


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            try:
                reconfigure(encoding="utf-8")
            except Exception:
                pass


def _write_csv(path: Path, columns: list[str], rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=columns,
            extrasaction="ignore",
            lineterminator="\n",
        )
        writer.writeheader()
        for row in rows:
            writer.writerow({key: _csv_cell(row.get(key)) for key in columns})


def _csv_cell(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        return f"{value:.3f}"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    return csv_text(str(value))


def _column_keys(columns: list[tuple[str, str, str]]) -> list[str]:
    return [key for key, _header, _kind in columns]


def render_report(
    *,
    rows: list[dict[str, str]],
    output_dir: Path,
    template: Path,
    timezone_name: str,
    log_group: str,
    region: str,
    account: str,
    start_epoch: int,
    end_epoch: int,
    query_text: str,
    query_status: str,
    records_matched: int,
    records_scanned: int,
    bytes_scanned: int,
    service: str,
    trace_id: str,
    filter_expr: str,
    limit: int,
    include_ptr: bool,
) -> dict[str, Any]:
    zone = parse_zone(timezone_name)
    report = build_report(
        rows,
        service=service,
        trace_id=trace_id,
        limit=limit,
        records_matched=records_matched,
    )
    raw_columns = result_columns(rows, include_ptr=include_ptr)
    raw_csv = output_dir / "spans-raw.csv"
    traces_csv = output_dir / "traces.csv"
    timeline_csv = output_dir / "timeline.csv"
    details_csv = output_dir / "span-details.csv"
    services_csv = output_dir / "services.csv"
    workbook = output_dir / "xray-trace-report.xlsx"
    _write_csv(raw_csv, raw_columns, rows)
    trace_records = [trace_row(trace, zone) for trace in report.traces]
    timeline_records = [
        timeline_row(trace, span, zone) for trace in report.traces for span in trace.spans
    ]
    detail_records = [
        detail_row(trace, span, zone) for trace in report.traces for span in trace.spans
    ]
    _write_csv(traces_csv, _column_keys(TRACE_COLUMNS), trace_records)
    _write_csv(timeline_csv, _column_keys(TIMELINE_COLUMNS), timeline_records)
    _write_csv(details_csv, _column_keys(DETAIL_COLUMNS), detail_records)
    _write_csv(services_csv, _column_keys(SERVICE_COLUMNS), report.services)
    counts = {
        STATUS_OK: 0,
        STATUS_ERROR: 0,
        STATUS_FAULT: 0,
        STATUS_THROTTLE: 0,
    }
    for trace in report.traces:
        counts[trace.trace_status] = counts.get(trace.trace_status, 0) + 1
    summary = {
        "QueryLogGroup": log_group,
        "QueryRegion": region,
        "QueryAccount": account,
        "QueryStart": format_epoch(start_epoch, zone),
        "QueryEnd": format_epoch(end_epoch, zone),
        "QueryTimeZone": timezone_name,
        "QueryStatus": query_status,
        "QueryLimit": limit,
        "QueryService": service,
        "QueryTraceId": trace_id,
        "QueryFilter": filter_expr,
        "QueryString": query_text,
        "RecordsMatched": records_matched,
        "RecordsScanned": records_scanned,
        "BytesScanned": f"{bytes_scanned} ({human_bytes(bytes_scanned)})",
        "SpanCount": len(report.spans),
        "TraceCount": len(report.traces),
        "FaultTraceCount": counts[STATUS_FAULT],
        "ThrottleTraceCount": counts[STATUS_THROTTLE],
        "ErrorTraceCount": counts[STATUS_ERROR],
        "OkTraceCount": counts[STATUS_OK],
        "GeneratedAt": datetime.now(timezone.utc).astimezone(zone).strftime("%Y-%m-%d %H:%M:%S %z"),
        "OutputCsv": str(raw_csv),
        "ParseNotes": "\n".join(report.notes),
    }
    fill_template(
        template,
        workbook,
        summary=summary,
        traces=trace_records,
        timeline=timeline_records,
        details=detail_records,
        services=report.services,
        raw_rows=rows,
        raw_columns=raw_columns or ["(no rows)"],
    )
    metadata = {
        "version": __version__,
        "log_group": log_group,
        "region": region,
        "account": account,
        "start_epoch": start_epoch,
        "end_epoch": end_epoch,
        "timezone": timezone_name,
        "query_status": query_status,
        "records_matched": records_matched,
        "records_scanned": records_scanned,
        "bytes_scanned": bytes_scanned,
        "span_count": len(report.spans),
        "trace_count": len(report.traces),
        "service": service,
        "trace_id": trace_id,
        "filter": filter_expr,
        "limit": limit,
        "notes": report.notes,
        "outputs": {
            "raw_csv": str(raw_csv),
            "traces_csv": str(traces_csv),
            "timeline_csv": str(timeline_csv),
            "details_csv": str(details_csv),
            "services_csv": str(services_csv),
            "workbook": str(workbook),
        },
    }
    (output_dir / "metadata.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return metadata


def _cmd_parse_time(args: argparse.Namespace) -> int:
    epoch = parse_time(args.value)
    sys.stdout.write(str(epoch) + "\n")
    return 0


def _cmd_build_query(args: argparse.Namespace) -> int:
    query_file_text = ""
    if args.query_file:
        query_file_text = Path(args.query_file).read_text(encoding="utf-8")
    text = build_query(
        limit=args.limit,
        include_message=not args.no_message,
        service=args.service,
        trace_id=args.trace_id,
        filter_expr=args.filter,
        query_file_text=query_file_text,
    )
    sys.stdout.write(text)
    return 0


def _cmd_init_template(args: argparse.Namespace) -> int:
    destination = Path(args.output)
    build_template(destination)
    sys.stdout.write(str(destination) + "\n")
    return 0


def _cmd_read_stats(args: argparse.Namespace) -> int:
    status, matched, scanned, size, token = query_stats(load_json(Path(args.page)))
    sys.stdout.write(f"{status}\n{matched}\n{scanned}\n{size}\n{token}\n")
    return 0


def _cmd_read_query_id(args: argparse.Namespace) -> int:
    sys.stdout.write(query_id(load_json(Path(args.page))) + "\n")
    return 0


def _cmd_read_account(args: argparse.Namespace) -> int:
    sys.stdout.write(account_id(load_json(Path(args.page))) + "\n")
    return 0


def _cmd_log_group_check(args: argparse.Namespace) -> int:
    return log_group_present(load_json(Path(args.page)), args.name)


def _render_from_args(args: argparse.Namespace, rows: list[dict[str, str]], matched: int, scanned: int, size: int) -> int:
    query_text = Path(args.query_file).read_text(encoding="utf-8") if args.query_file else ""
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    metadata = render_report(
        rows=rows,
        output_dir=output_dir,
        template=Path(args.template),
        timezone_name=args.timezone,
        log_group=args.log_group,
        region=args.region,
        account=args.account,
        start_epoch=args.start_epoch,
        end_epoch=args.end_epoch,
        query_text=query_text,
        query_status=args.query_status,
        records_matched=matched,
        records_scanned=scanned,
        bytes_scanned=size,
        service=args.service,
        trace_id=args.trace_id,
        filter_expr=args.filter,
        limit=args.limit,
        include_ptr=args.include_ptr,
    )
    sys.stdout.write(metadata["outputs"]["workbook"] + "\n")
    return 0


def _cmd_render(args: argparse.Namespace) -> int:
    pages = sorted(Path(args.pages).glob("page-*.json"))
    rows, stats = load_result_rows(pages)
    return _render_from_args(
        args,
        rows,
        stats["recordsMatched"],
        stats["recordsScanned"],
        stats["bytesScanned"],
    )


def _cmd_demo(args: argparse.Namespace) -> int:
    output_dir = Path(args.output_dir)
    pages = output_dir / "pages"
    pages.mkdir(parents=True, exist_ok=True)
    (pages / "page-0001.json").write_text(
        json.dumps(sample_page(), ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    query_path = output_dir / "query.cwli"
    query_text = build_query(limit=args.limit, include_message=True)
    query_path.write_text(query_text, encoding="utf-8")
    end_epoch = int(BASE.timestamp()) + 3600
    args.query_file = str(query_path)
    args.start_epoch = int(BASE.timestamp())
    args.end_epoch = end_epoch
    args.query_status = "Complete"
    args.log_group = args.log_group or "aws/spans"
    args.service = ""
    args.trace_id = ""
    args.filter = ""
    page = sample_page()
    _status, matched, scanned, size, _token = query_stats(page)
    rows, _stats = load_result_rows([pages / "page-0001.json"])
    return _render_from_args(args, rows, matched, scanned, size)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="spanreport")
    parser.add_argument("--version", action="version", version=__version__)
    sub = parser.add_subparsers(dest="command", required=True)

    parse_time = sub.add_parser("parse-time")
    parse_time.add_argument("--value", required=True)
    parse_time.set_defaults(func=_cmd_parse_time)

    query = sub.add_parser("build-query")
    query.add_argument("--limit", type=int, default=10000)
    query.add_argument("--no-message", action="store_true")
    query.add_argument("--service", default="")
    query.add_argument("--trace-id", default="")
    query.add_argument("--filter", default="")
    query.add_argument("--query-file", default="")
    query.set_defaults(func=_cmd_build_query)

    template = sub.add_parser("init-template")
    template.add_argument("--output", required=True)
    template.set_defaults(func=_cmd_init_template)

    stats = sub.add_parser("read-stats")
    stats.add_argument("--page", required=True)
    stats.set_defaults(func=_cmd_read_stats)

    qid = sub.add_parser("read-query-id")
    qid.add_argument("--page", required=True)
    qid.set_defaults(func=_cmd_read_query_id)

    account = sub.add_parser("read-account")
    account.add_argument("--page", required=True)
    account.set_defaults(func=_cmd_read_account)

    check = sub.add_parser("log-group-check")
    check.add_argument("--page", required=True)
    check.add_argument("--name", required=True)
    check.set_defaults(func=_cmd_log_group_check)

    def add_render_args(target: argparse.ArgumentParser, *, pages: bool) -> None:
        if pages:
            target.add_argument("--pages", required=True)
        target.add_argument("--output-dir", required=True)
        target.add_argument("--template", required=True)
        target.add_argument("--timezone", default="UTC")
        target.add_argument("--log-group", default="aws/spans")
        target.add_argument("--region", default="")
        target.add_argument("--account", default="")
        target.add_argument("--start-epoch", type=int, default=0)
        target.add_argument("--end-epoch", type=int, default=0)
        target.add_argument("--query-file", default="")
        target.add_argument("--query-status", default="Complete")
        target.add_argument("--service", default="")
        target.add_argument("--trace-id", default="")
        target.add_argument("--filter", default="")
        target.add_argument("--limit", type=int, default=10000)
        target.add_argument("--include-ptr", action="store_true")

    render = sub.add_parser("render")
    add_render_args(render, pages=True)
    render.set_defaults(func=_cmd_render)

    demo = sub.add_parser("demo")
    add_render_args(demo, pages=False)
    demo.set_defaults(func=_cmd_demo)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    _configure_stdio()
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        return int(args.func(args))
    except BrokenPipeError:
        return 0
    except (ValueError, OSError, json.JSONDecodeError) as exc:
        sys.stderr.write(f"spanreport: {exc}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
