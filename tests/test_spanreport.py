"""Unit tests for span parsing, the Logs Insights query, and the Excel template."""

from __future__ import annotations

import csv
import json
import sys
import tempfile
import unittest
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "lib"))

from spanreport.awsio import load_result_rows, query_stats, result_columns  # noqa: E402
from spanreport.model import (  # noqa: E402
    BAR_OFF,
    BAR_ON,
    STATUS_ERROR,
    STATUS_FAULT,
    STATUS_OK,
    STATUS_THROTTLE,
    WATERFALL_WIDTH,
    build_report,
    classify_status,
    csv_text,
    trace_row,
)
from spanreport.query import build_query, cwli_string, to_xray_trace_id, trace_id_variants  # noqa: E402
from spanreport.sample import BASE, TRACE_ERROR, TRACE_FAULT, TRACE_OK, sample_page, sample_spans  # noqa: E402
from spanreport.timeutil import format_unix_nano, parse_time, parse_zone  # noqa: E402
from spanreport.xlsxutil import (  # noqa: E402
    DATA_ROW,
    HEADER_ROW,
    TRACE_SHEET,
    build_template,
    fill_template,
    read_defined_values,
    read_sheet_matrix,
)
from spanreport.__main__ import render_report  # noqa: E402


NOW = datetime(2026, 10, 26, 12, 0, tzinfo=timezone.utc)


class TimeTests(unittest.TestCase):
    def test_parse_forms(self) -> None:
        self.assertEqual(parse_time("now", now=NOW), int(NOW.timestamp()))
        self.assertEqual(parse_time("-15m", now=NOW), int(NOW.timestamp()) - 900)
        self.assertEqual(parse_time("-2h", now=NOW), int(NOW.timestamp()) - 7200)
        self.assertEqual(parse_time("-1d", now=NOW), int(NOW.timestamp()) - 86400)
        self.assertEqual(parse_time("1700000000", now=NOW), 1700000000)
        self.assertEqual(parse_time("1700000000000", now=NOW), 1700000000)
        self.assertEqual(parse_time("2026-10-26T00:00:00Z", now=NOW), int(BASE.timestamp()))
        self.assertEqual(parse_time("2026-10-26T09:00:00+09:00", now=NOW), int(BASE.timestamp()))
        self.assertEqual(parse_time("2026-10-26T00:00:00", now=NOW), int(BASE.timestamp()))

    def test_reject_bad_time(self) -> None:
        with self.assertRaises(ValueError):
            parse_time("tomorrow", now=NOW)
        with self.assertRaises(ValueError):
            parse_time("", now=NOW)

    def test_fixed_offset_zone(self) -> None:
        zone = parse_zone("+09:00")
        text = format_unix_nano(int(BASE.timestamp()) * 1_000_000_000, zone)
        self.assertIn("2026-10-26 09:00:00.000 +0900", text)


class QueryTests(unittest.TestCase):
    def test_trace_id_variants(self) -> None:
        xray = "1-58406520-a006649127e371903a2de979"
        otel = "58406520a006649127e371903a2de979"
        self.assertEqual(trace_id_variants(xray), [xray, otel])
        self.assertEqual(to_xray_trace_id(otel), xray)
        self.assertEqual(to_xray_trace_id(xray), xray)

    def test_default_query_targets_span_fields(self) -> None:
        query = build_query(limit=25, service='pay"ment', trace_id="1-58406520-a006649127e371903a2de979")
        self.assertIn("durationNano", query)
        self.assertIn("`attributes.aws.local.service`", query)
        self.assertIn("`attributes.http.response.status_code`", query)
        self.assertIn("@message", query)
        self.assertIn("| limit 25", query)
        self.assertIn(cwli_string('pay"ment'), query)
        self.assertIn("`attributes.aws.remote.service`", query)
        self.assertIn("58406520a006649127e371903a2de979", query)
        self.assertNotIn("\r", query)

    def test_no_message_and_custom_query(self) -> None:
        query = build_query(limit=10, include_message=False, filter_expr='status.code = "ERROR"')
        self.assertNotIn("@message", query)
        self.assertIn('| filter status.code = "ERROR"', query)
        custom = build_query(query_file_text="fields @timestamp, @message\n| limit 3\n")
        self.assertTrue(custom.startswith("fields @timestamp"))
        with self.assertRaises(ValueError):
            build_query(query_file_text="fields @message", filter_expr="name = \"x\"")
        with self.assertRaises(ValueError):
            build_query(limit=0)
        with self.assertRaises(ValueError):
            build_query(filter_expr="status = 1\n| limit 1")


class ModelTests(unittest.TestCase):
    def test_status_mapping(self) -> None:
        self.assertEqual(classify_status(http_status=200, otel_status="OK", attributes={}, error_type="", error_message=""), STATUS_OK)
        self.assertEqual(classify_status(http_status=404, otel_status="ERROR", attributes={}, error_type="", error_message=""), STATUS_ERROR)
        self.assertEqual(classify_status(http_status=503, otel_status="OK", attributes={}, error_type="", error_message=""), STATUS_FAULT)
        self.assertEqual(classify_status(http_status=429, otel_status="ERROR", attributes={}, error_type="", error_message=""), STATUS_THROTTLE)
        self.assertEqual(
            classify_status(http_status=None, otel_status="UNSET", attributes={"aws.xray.fault": True}, error_type="", error_message=""),
            STATUS_FAULT,
        )

    def test_sample_report_matches_xray_screen(self) -> None:
        page = sample_page()
        rows, _stats = load_result_rows([_write_page(page)])
        report = build_report(rows, limit=10000, records_matched=len(rows))
        self.assertEqual(len(report.traces), 4)
        self.assertEqual(report.traces[0].trace_id, TRACE_ERROR)
        ok = next(trace for trace in report.traces if trace.trace_id == TRACE_OK)
        self.assertEqual(ok.trace_id_xray, "1-58406520-a006649127e371903a2de979")
        self.assertEqual(ok.http_method, "GET")
        self.assertEqual(ok.url, "/orders")
        self.assertEqual(ok.http_status, 200)
        self.assertEqual(ok.service, "order-api")
        self.assertEqual(ok.client_ip, "203.0.113.10")
        self.assertEqual(ok.user, "user-42")
        self.assertEqual(ok.root_status, STATUS_OK)
        self.assertEqual(ok.trace_status, STATUS_OK)
        self.assertIn("tenant_id=acme", ok.annotations)
        self.assertEqual([span.depth for span in ok.spans], [0, 1, 2])
        self.assertEqual(ok.spans[2].display_name, "    unmarshal")
        self.assertEqual(ok.spans[0].waterfall, BAR_ON * WATERFALL_WIDTH)
        child = ok.spans[1]
        self.assertTrue(child.waterfall.startswith(BAR_OFF * 5 + BAR_ON))
        self.assertEqual(child.remote_service, "DynamoDB")
        self.assertAlmostEqual((child.duration_ns or 0) / 1_000_000, 50.0)
        fault = next(trace for trace in report.traces if trace.trace_id == TRACE_FAULT)
        self.assertEqual(fault.root_status, STATUS_OK)
        self.assertEqual(fault.trace_status, STATUS_FAULT)
        self.assertEqual(fault.http_status, 200)
        self.assertEqual(fault.fault_count, 1)
        throttle = next(trace for trace in report.traces if "58406522" in trace.trace_id)
        self.assertEqual(throttle.trace_status, STATUS_THROTTLE)
        error = report.traces[0]
        self.assertEqual(error.trace_status, STATUS_ERROR)
        self.assertEqual(error.spans[0].error_type, "NotFoundError")
        self.assertIn("Catalog.get", error.spans[0].stacktrace)
        self.assertTrue(any("除外" in note for note in report.notes))
        order = next(row for row in report.services if row["service"] == "order-api")
        self.assertEqual(order["span_count"], 3)
        self.assertEqual(order["p50_ms"], 50.0)

    def test_service_filter_keeps_trace_siblings(self) -> None:
        rows, _stats = load_result_rows([_write_page(sample_page())])
        report = build_report(rows, service="charge-api")
        self.assertEqual(len(report.traces), 1)
        self.assertEqual(report.traces[0].trace_id, TRACE_FAULT)
        self.assertEqual(len(report.traces[0].spans), 2)

    def test_flat_row_and_cycle(self) -> None:
        flat = {
            "@timestamp": "2026-10-26 00:00:00.000",
            "@message": "not-json",
            "traceId": "ab" * 16,
            "spanId": "abcdabcdabcdabcd",
            "name": "flat-span",
            "durationNano": "5000000",
            "startTimeUnixNano": "1000",
            "endTimeUnixNano": "5001000",
            "awsLocalService": "flat-service",
            "httpMethod": "POST",
            "httpStatus": "201",
            "urlPath": "/flat",
        }
        report = build_report([flat])
        self.assertEqual(report.flattened_rows, 1)
        self.assertEqual(report.spans[0].service, "flat-service")
        self.assertEqual(report.spans[0].http_status, 201)
        self.assertEqual(report.spans[0].xray_status, STATUS_OK)
        cycle_a = {
            "traceId": "cd" * 16,
            "spanId": "aaaaaaaaaaaaaaaa",
            "parentSpanId": "bbbbbbbbbbbbbbbb",
            "name": "a",
            "kind": "INTERNAL",
            "startTimeUnixNano": 10,
            "endTimeUnixNano": 20,
            "durationNano": 10,
        }
        cycle_b = dict(cycle_a, spanId="bbbbbbbbbbbbbbbb", parentSpanId="aaaaaaaaaaaaaaaa", name="b")
        cycled = build_report([{"@message": json.dumps(cycle_a)}, {"@message": json.dumps(cycle_b)}])
        self.assertEqual(len(cycled.spans), 2)

    def test_csv_formula_guard_and_columns(self) -> None:
        self.assertTrue(csv_text("=cmd()").startswith("'"))
        self.assertNotIn("@ptr", result_columns([{"@ptr": "x", "traceId": "t", "@message": "{}"}], include_ptr=False))
        self.assertIn("@ptr", result_columns([{"@ptr": "x", "@message": "{}"}], include_ptr=True))

    def test_limit_note(self) -> None:
        rows, _stats = load_result_rows([_write_page(sample_page())])
        report = build_report(rows, limit=len(rows), records_matched=len(rows) + 5)
        self.assertTrue(any("上限" in note for note in report.notes))
        self.assertTrue(any("一致件数" in note for note in report.notes))

    def test_display_zone(self) -> None:
        rows, _stats = load_result_rows([_write_page(sample_page())])
        report = build_report(rows)
        ok = next(trace for trace in report.traces if trace.trace_id == TRACE_OK)
        text = trace_row(ok, parse_zone("+09:00"))["start_time"]
        self.assertEqual(text, "2026-10-26 09:00:00.000 +0900")


class WorkbookTests(unittest.TestCase):
    def test_template_contract_and_fill(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            directory = Path(tmp)
            template = directory / "template.xlsx"
            build_template(template)
            before = read_sheet_matrix(template, TRACE_SHEET)
            self.assertIn("trace_id_xray", before[2])
            self.assertIn("トレース ID", before[3])
            self.assertIn("EXAMPLE-TRACE", before[4])
            names = read_defined_values(template)
            self.assertIn("QueryLogGroup", names)
            page_path = _write_page(sample_page(), directory / "page-0001.json")
            rows, stats = load_result_rows([page_path])
            status, matched, scanned, size, _token = query_stats(json.loads(page_path.read_text(encoding="utf-8")))
            self.assertEqual(status, "Complete")
            self.assertEqual(matched, stats["recordsMatched"])
            metadata = render_report(
                rows=rows,
                output_dir=directory,
                template=template,
                timezone_name="+09:00",
                log_group="aws/spans",
                region="ap-northeast-1",
                account="123456789012",
                start_epoch=int(BASE.timestamp()),
                end_epoch=int(BASE.timestamp()) + 3600,
                query_text="fields @timestamp, @message\n| limit 10000\n",
                query_status="Complete",
                records_matched=matched,
                records_scanned=scanned,
                bytes_scanned=size,
                service="",
                trace_id="",
                filter_expr="",
                limit=10000,
                include_ptr=False,
            )
            workbook = Path(metadata["outputs"]["workbook"])
            values = read_defined_values(workbook)
            self.assertEqual(values["QueryLogGroup"], "aws/spans")
            self.assertEqual(values["QueryRegion"], "ap-northeast-1")
            self.assertEqual(values["TraceCount"], "4")
            self.assertEqual(values["FaultTraceCount"], "1")
            self.assertEqual(values["ThrottleTraceCount"], "1")
            self.assertEqual(values["ErrorTraceCount"], "1")
            self.assertIn("fields @timestamp", values["QueryString"])
            matrix = read_sheet_matrix(workbook, TRACE_SHEET)
            flat = "\n".join(",".join(row) for row in matrix)
            self.assertNotIn("EXAMPLE-TRACE", flat)
            self.assertIn("1-58406520-a006649127e371903a2de979", flat)
            self.assertIn("障害", flat)
            self.assertIn("スロットル", flat)
            self.assertIn("2026-10-26 09:00:00.000 +0900", flat)
            timeline = read_sheet_matrix(workbook, "タイムライン")
            timeline_text = "\n".join(",".join(row) for row in timeline)
            self.assertIn("unmarshal", timeline_text)
            self.assertIn(BAR_ON, timeline_text)
            raw_sheet = read_sheet_matrix(workbook, "クエリ結果")
            raw_header = next(row for row in raw_sheet if "@message" in row)
            self.assertIn("@timestamp", raw_header)
            self.assertNotIn("@ptr", raw_header)
            self.assertNotIn("（生成時に置き換わります）", " ".join(" ".join(row) for row in raw_sheet))
            raw_text = "\n".join(",".join(row) for row in raw_sheet)
            self.assertIn(TRACE_OK, raw_text)
            services = read_sheet_matrix(workbook, "サービス集計")
            service_text = "\n".join(",".join(row) for row in services)
            self.assertIn("order-api", service_text)
            self.assertIn("payment-api", service_text)
            _assert_meiryo_monotone(workbook)
            with (directory / "spans-raw.csv").open(encoding="utf-8-sig", newline="") as handle:
                header = next(csv.reader(handle))
            self.assertIn("@message", header)
            self.assertNotIn("@ptr", header)
            self.assertGreater(DATA_ROW, HEADER_ROW)


def _assert_meiryo_monotone(workbook: Path) -> None:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(workbook) as archive:
        styles = ET.fromstring(archive.read("xl/styles.xml"))
        names = [node.attrib.get("val") for node in styles.findall(".//m:font/m:name", ns)]
        assert names, "fonts missing"
        assert set(names) == {"Meiryo UI"}, names
        style_text = archive.read("xl/styles.xml").decode("utf-8").upper()
        for banned in ("FF0073BB", "FF1D8102", "FFEB5F07", "FFD13212", "FF6B4C9A", "CALIBRI", "CONSOLAS"):
            assert banned not in style_text
        for name in archive.namelist():
            if not name.startswith("xl/worksheets/sheet"):
                continue
            sheet = ET.fromstring(archive.read(name))
            tab = sheet.find("m:sheetPr/m:tabColor", ns)
            assert tab is not None
            rgb = tab.attrib["rgb"]
            assert rgb[2:4] == rgb[4:6] == rgb[6:8], rgb
            view = sheet.find("m:sheetViews/m:sheetView", ns)
            assert view is not None and view.attrib.get("showGridLines") == "0"


def _write_page(page: dict, path: Path | None = None) -> Path:
    if path is None:
        handle = tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8")
        path = Path(handle.name)
        handle.write(json.dumps(page))
        handle.close()
        return path
    path.write_text(json.dumps(page), encoding="utf-8")
    return path


if __name__ == "__main__":
    unittest.main()
