"""Excel template that mirrors the X-Ray trace list and timeline.

The workbook is Office Open XML written with the Python 3.9 standard library
so RHEL 9.8 does not need openpyxl. A fill copies the template and replaces
data rows. Column identity lives in a hidden key row, so display headers,
widths, and conditional formatting in the template survive a refill.
"""

from __future__ import annotations

import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from spanreport.model import BAR_ON, RESPONSE_BAR_WIDTH, WATERFALL_WIDTH, excel_text


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
NS_EXT = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

ET.register_namespace("", NS)
ET.register_namespace("r", NS_R)

SUMMARY_SHEET = "概要"
TRACE_SHEET = "トレース一覧"
TIMELINE_SHEET = "タイムライン"
DETAIL_SHEET = "スパン詳細"
SERVICE_SHEET = "サービス集計"
RAW_SHEET = "クエリ結果"

KEY_ROW = 3
HEADER_ROW = 4
DATA_ROW = 5

# Styles. Indexes are part of the template. The filler copies the style index
# stored on the sample row, so these numbers must stay aligned with styles.xml.
STYLE_NORMAL = 0
STYLE_TITLE = 1
STYLE_SUBTITLE = 2
STYLE_SECTION = 3
STYLE_LABEL = 4
STYLE_VALUE = 5
STYLE_HEADER = 6
STYLE_TEXT = 7
STYLE_MS = 8
STYLE_INT = 9
STYLE_BAR = 11
STYLE_WRAP = 12
STYLE_NOTE = 14
STYLE_TRACE = 20
STYLE_QUERY = 21
FONT_NAME = "Meiryo UI"
HEADER_HEIGHT = 32
# Grayscale only. Darker gray marks a worse X-Ray status.
INK = "FF1A1A1A"
INK_SOFT = "FF4D4D4D"
WHITE = "FFFFFFFF"
CHARCOAL = "FF2F2F2F"
FILL_BAND = "FFF3F3F3"
FILL_OK = "FFF7F7F7"
FILL_ERROR = "FFE4E4E4"
FILL_FAULT = "FFCFCFCF"
FILL_THROTTLE = "FFB0B0B0"
FILL_WHITE = "FFFFFFFF"
GRID = "FFD0D0D0"
DATA_BAR = "FF5A5A5A"
STYLE_LEGEND = {
    "OK": 15,
    "エラー": 16,
    "障害": 17,
    "スロットル": 18,
}

TRACE_COLUMNS: list[tuple[str, str, str]] = [
    ("trace_id_xray", "トレース ID", "trace"),
    ("response_time_ms", "応答時間 (ms)", "ms"),
    ("trace_time_ms", "トレース時間 (ms)", "ms"),
    ("response_bar", "応答時間バー", "bar"),
    ("start_time", "開始時刻", "text"),
    ("end_time", "終了時刻", "text"),
    ("http_method", "メソッド", "text"),
    ("url", "URL", "text"),
    ("root_status", "ルートステータス", "text"),
    ("trace_status", "トレースステータス", "text"),
    ("http_status", "レスポンス", "int"),
    ("service", "サービス", "text"),
    ("client_ip", "クライアント IP", "text"),
    ("user", "ユーザー", "text"),
    ("span_count", "スパン数", "int"),
    ("error_count", "エラー数", "int"),
    ("fault_count", "障害数", "int"),
    ("throttle_count", "スロットル数", "int"),
    ("annotations", "注釈", "wrap"),
    ("trace_id_raw", "traceId (raw)", "text"),
]
TIMELINE_COLUMNS: list[tuple[str, str, str]] = [
    ("trace_id_xray", "トレース ID", "trace"),
    ("depth", "深さ", "int"),
    ("waterfall", "タイムライン", "bar"),
    ("display_name", "名前", "text"),
    ("service", "サービス", "text"),
    ("offset_ms", "開始オフセット (ms)", "ms"),
    ("duration_ms", "所要時間 (ms)", "ms"),
    ("start_time", "開始時刻", "text"),
    ("end_time", "終了時刻", "text"),
    ("status", "ステータス", "text"),
    ("kind", "種別", "text"),
    ("http_method", "メソッド", "text"),
    ("operation", "URL / 操作", "text"),
    ("http_status", "レスポンス", "int"),
    ("remote_service", "リモートサービス", "text"),
    ("remote_operation", "リモート操作", "text"),
    ("span_id", "Span ID", "text"),
    ("parent_span_id", "Parent Span ID", "text"),
    ("origin", "オリジン", "text"),
    ("error_type", "エラー種別", "text"),
    ("error_message", "エラーメッセージ", "wrap"),
]
DETAIL_COLUMNS: list[tuple[str, str, str]] = [
    ("trace_id_xray", "トレース ID", "trace"),
    ("span_id", "Span ID", "text"),
    ("parent_span_id", "Parent Span ID", "text"),
    ("service", "サービス", "text"),
    ("name", "名前", "text"),
    ("kind", "種別", "text"),
    ("start_time", "開始時刻", "text"),
    ("end_time", "終了時刻", "text"),
    ("duration_ms", "所要時間 (ms)", "ms"),
    ("status", "ステータス", "text"),
    ("otel_status", "OTel ステータス", "text"),
    ("http_method", "メソッド", "text"),
    ("operation", "URL / 操作", "text"),
    ("http_status", "レスポンス", "int"),
    ("remote_service", "リモートサービス", "text"),
    ("remote_operation", "リモート操作", "text"),
    ("db_system", "DB システム", "text"),
    ("db_statement", "DB ステートメント", "wrap"),
    ("region", "リージョン", "text"),
    ("availability_zone", "AZ", "text"),
    ("environment", "環境", "text"),
    ("client_ip", "クライアント IP", "text"),
    ("user", "ユーザー", "text"),
    ("origin", "オリジン", "text"),
    ("annotations", "注釈", "wrap"),
    ("metadata", "メタデータ", "wrap"),
    ("error_type", "エラー種別", "text"),
    ("error_message", "エラーメッセージ", "wrap"),
    ("stacktrace", "スタックトレース", "wrap"),
    ("attributes_json", "属性 JSON", "wrap"),
    ("log_stream", "ログストリーム", "text"),
    ("ingestion_time", "取り込み時刻", "text"),
    ("trace_id_raw", "traceId (raw)", "text"),
]
SERVICE_COLUMNS: list[tuple[str, str, str]] = [
    ("service", "サービス", "text"),
    ("span_count", "スパン数", "int"),
    ("trace_count", "トレース数", "int"),
    ("ok_count", "OK", "int"),
    ("error_count", "エラー", "int"),
    ("fault_count", "障害", "int"),
    ("throttle_count", "スロットル", "int"),
    ("avg_ms", "平均 (ms)", "ms"),
    ("p50_ms", "p50 (ms)", "ms"),
    ("p95_ms", "p95 (ms)", "ms"),
    ("p99_ms", "p99 (ms)", "ms"),
    ("max_ms", "最大 (ms)", "ms"),
]

SUMMARY_NAMES: list[tuple[str, str, str]] = [
    ("QueryLogGroup", "B5", "ロググループ"),
    ("QueryRegion", "B6", "リージョン"),
    ("QueryAccount", "B7", "アカウント"),
    ("QueryStart", "B8", "開始"),
    ("QueryEnd", "B9", "終了"),
    ("QueryTimeZone", "B10", "表示タイムゾーン"),
    ("QueryStatus", "B11", "クエリステータス"),
    ("QueryLimit", "B12", "上限"),
    ("QueryService", "B13", "サービスフィルタ"),
    ("QueryTraceId", "B14", "トレース ID フィルタ"),
    ("QueryFilter", "B15", "追加フィルタ"),
    ("QueryString", "B16", "クエリ"),
    ("RecordsMatched", "B19", "一致レコード"),
    ("RecordsScanned", "B20", "スキャンレコード"),
    ("BytesScanned", "B21", "スキャンバイト"),
    ("SpanCount", "B22", "取得スパン数"),
    ("TraceCount", "B23", "トレース数"),
    ("FaultTraceCount", "B24", "障害のトレース"),
    ("ThrottleTraceCount", "B25", "スロットルのトレース"),
    ("ErrorTraceCount", "B26", "エラーのトレース"),
    ("OkTraceCount", "B27", "OK のトレース"),
    ("GeneratedAt", "B28", "生成日時"),
    ("OutputCsv", "B29", "生 CSV"),
    ("ParseNotes", "B30", "復元メモ"),
]


def col_letter(index: int) -> str:
    if index < 1:
        raise ValueError(f"column index must be positive: {index}")
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


def _style_for(kind: str) -> int:
    return {
        "trace": STYLE_TRACE,
        "text": STYLE_TEXT,
        "ms": STYLE_MS,
        "int": STYLE_INT,
        "bar": STYLE_BAR,
        "wrap": STYLE_WRAP,
    }[kind]


def _inline(ref: str, text: str, style: int | None) -> ET.Element:
    attrs = {"r": ref, "t": "inlineStr"}
    if style is not None:
        attrs["s"] = str(style)
    cell = ET.Element(_q("c"), attrs)
    inline = ET.SubElement(cell, _q("is"))
    node = ET.SubElement(inline, _q("t"))
    if text[:1].isspace() or text[-1:].isspace() or "\n" in text:
        node.set(XML_SPACE, "preserve")
    node.text = text
    return cell


def _number(ref: str, value: int | float, style: int | None, places: int) -> ET.Element:
    attrs = {"r": ref}
    if style is not None:
        attrs["s"] = str(style)
    cell = ET.Element(_q("c"), attrs)
    node = ET.SubElement(cell, _q("v"))
    if places == 0:
        node.text = str(int(round(float(value))))
    else:
        node.text = f"{float(value):.{places}f}"
    return cell


def _row(index: int, cells: list[ET.Element], *, hidden: bool = False, height: float | None = None) -> ET.Element:
    attrs: dict[str, str] = {"r": str(index)}
    if hidden:
        attrs["hidden"] = "1"
        attrs["ht"] = "0"
        attrs["customHeight"] = "1"
    elif height is not None:
        attrs["ht"] = str(height)
        attrs["customHeight"] = "1"
    row = ET.Element(_q("row"), attrs)
    for cell in cells:
        row.append(cell)
    return row


def _worksheet(
    *,
    rows: list[ET.Element],
    widths: list[float],
    freeze_cols: int,
    freeze_rows: int,
    tab_color: str,
    auto_filter: str | None,
    merges: list[str],
    conditional: list[ET.Element],
    last_ref: str,
) -> ET.Element:
    root = ET.Element(_q("worksheet"))
    sheet_pr = ET.SubElement(root, _q("sheetPr"))
    ET.SubElement(sheet_pr, _q("tabColor"), {"rgb": tab_color})
    ET.SubElement(sheet_pr, _q("pageSetUpPr"), {"fitToPage": "1"})
    ET.SubElement(root, _q("dimension"), {"ref": last_ref})
    views = ET.SubElement(root, _q("sheetViews"))
    view = ET.SubElement(
        views,
        _q("sheetView"),
        {"workbookViewId": "0", "showGridLines": "0", "zoomScale": "100"},
    )
    if freeze_rows or freeze_cols:
        top = f"{col_letter(freeze_cols + 1)}{freeze_rows + 1}"
        pane_attrs = {
            "topLeftCell": top,
            "activePane": "bottomRight" if freeze_cols and freeze_rows else "bottomLeft",
            "state": "frozen",
        }
        if freeze_cols:
            pane_attrs["xSplit"] = str(freeze_cols)
        if freeze_rows:
            pane_attrs["ySplit"] = str(freeze_rows)
        ET.SubElement(view, _q("pane"), pane_attrs)
        ET.SubElement(
            view,
            _q("selection"),
            {"pane": pane_attrs["activePane"], "activeCell": top, "sqref": top},
        )
    ET.SubElement(root, _q("sheetFormatPr"), {"defaultRowHeight": "21", "defaultColWidth": "12"})
    cols = ET.SubElement(root, _q("cols"))
    for index, width in enumerate(widths, start=1):
        ET.SubElement(
            cols,
            _q("col"),
            {
                "min": str(index),
                "max": str(index),
                "width": str(width),
                "customWidth": "1",
            },
        )
    sheet_data = ET.SubElement(root, _q("sheetData"))
    for row in rows:
        sheet_data.append(row)
    if auto_filter:
        ET.SubElement(root, _q("autoFilter"), {"ref": auto_filter})
    if merges:
        merge_parent = ET.SubElement(root, _q("mergeCells"), {"count": str(len(merges))})
        for ref in merges:
            ET.SubElement(merge_parent, _q("mergeCell"), {"ref": ref})
    for block in conditional:
        root.append(block)
    ET.SubElement(
        root,
        _q("pageMargins"),
        {"left": "0.5", "right": "0.5", "top": "0.6", "bottom": "0.6", "header": "0.3", "footer": "0.3"},
    )
    ET.SubElement(
        root,
        _q("pageSetup"),
        {"orientation": "landscape", "paperSize": "9", "fitToWidth": "1", "fitToHeight": "0"},
    )
    footer = ET.SubElement(root, _q("headerFooter"))
    ET.SubElement(footer, _q("oddHeader")).text = "&LTransaction Search / aws/spans&R&D"
    ET.SubElement(footer, _q("oddFooter")).text = "&LX-Ray trace report&R&P / &N"
    return root


def _status_cf(letter: str, priority_start: int) -> ET.Element:
    block = ET.Element(_q("conditionalFormatting"), {"sqref": f"{letter}{DATA_ROW}:{letter}1048576"})
    # dxf order: 0 OK, 1 エラー, 2 障害, 3 スロットル. Fault is listed first so a
    # cell cannot be repainted by a broader rule. Each rule is an exact match.
    rules = [(2, "障害"), (3, "スロットル"), (1, "エラー"), (0, "OK")]
    for offset, (dxf_id, label) in enumerate(rules):
        rule = ET.SubElement(
            block,
            _q("cfRule"),
            {
                "type": "cellIs",
                "dxfId": str(dxf_id),
                "operator": "equal",
                "priority": str(priority_start + offset),
            },
        )
        formula = ET.SubElement(rule, _q("formula"))
        formula.text = f'"{label}"'
    return block


def _data_bar(letter: str, priority: int) -> ET.Element:
    block = ET.Element(_q("conditionalFormatting"), {"sqref": f"{letter}{DATA_ROW}:{letter}1048576"})
    rule = ET.SubElement(block, _q("cfRule"), {"type": "dataBar", "priority": str(priority)})
    bar = ET.SubElement(rule, _q("dataBar"), {"showValue": "1"})
    ET.SubElement(bar, _q("cfvo"), {"type": "min"})
    ET.SubElement(bar, _q("cfvo"), {"type": "max"})
    ET.SubElement(bar, _q("color"), {"rgb": DATA_BAR})
    return block


def _example_value(key: str, kind: str) -> Any:
    samples: dict[str, Any] = {
        "trace_id_xray": "EXAMPLE-TRACE",
        "response_time_ms": 120.0,
        "trace_time_ms": 120.0,
        "response_bar": BAR_ON * RESPONSE_BAR_WIDTH,
        "start_time": "2026-10-26 09:00:00.000 +0900",
        "end_time": "2026-10-26 09:00:00.120 +0900",
        "http_method": "GET",
        "url": "/orders",
        "operation": "/orders",
        "root_status": "OK",
        "trace_status": "OK",
        "status": "OK",
        "http_status": 200,
        "service": "example-service",
        "client_ip": "203.0.113.10",
        "user": "example-user",
        "span_count": 3,
        "error_count": 0,
        "fault_count": 0,
        "throttle_count": 0,
        "annotations": "tenant_id=example",
        "trace_id_raw": "58406520a006649127e371903a2de979",
        "depth": 0,
        "waterfall": BAR_ON * WATERFALL_WIDTH,
        "display_name": "GET /orders",
        "name": "GET /orders",
        "offset_ms": 0.0,
        "duration_ms": 120.0,
        "kind": "SERVER",
        "remote_service": "DynamoDB",
        "remote_operation": "GetItem",
        "span_id": "53995c3f42cd8ad8",
        "parent_span_id": "",
        "origin": "aws_ecs",
        "error_type": "",
        "error_message": "",
        "otel_status": "OK",
        "db_system": "",
        "db_statement": "",
        "region": "ap-northeast-1",
        "availability_zone": "ap-northeast-1a",
        "environment": "prod",
        "metadata": "",
        "stacktrace": "",
        "attributes_json": "{}",
        "log_stream": "example",
        "ingestion_time": "2026-10-26 09:00:01.000 +0900",
        "ok_count": 3,
        "trace_count": 1,
        "avg_ms": 40.0,
        "p50_ms": 50.0,
        "p95_ms": 110.0,
        "p99_ms": 118.0,
        "max_ms": 120.0,
    }
    if key in samples:
        return samples[key]
    if kind in {"ms"}:
        return 0.0
    if kind == "int":
        return 0
    return ""


def _cells_for(row_index: int, values: list[tuple[Any, str]], styles: list[int | None]) -> list[ET.Element]:
    cells: list[ET.Element] = []
    for index, ((value, kind), style) in enumerate(zip(values, styles), start=1):
        ref = f"{col_letter(index)}{row_index}"
        if value is None or value == "":
            continue
        if kind in {"ms", "int"} and isinstance(value, (int, float)) and not isinstance(value, bool):
            places = 0 if kind == "int" else 3
            cells.append(_number(ref, value, style, places))
        else:
            cells.append(_inline(ref, excel_text(value), style))
    return cells


def _table_sheet(
    columns: list[tuple[str, str, str]],
    *,
    title: str,
    note: str,
    tab_color: str,
    freeze_cols: int,
    status_keys: list[str],
    bar_keys: list[str],
    priority_start: int,
) -> ET.Element:
    # Widths are in default-font character units. Meiryo UI is wider than
    # Calibri, and the timeline uses fullwidth squares.
    width_for = {"trace": 44, "text": 26, "ms": 20, "int": 14, "bar": 68, "wrap": 48}
    widths = [width_for[kind] for _, _, kind in columns]
    key_cells = [
        _inline(f"{col_letter(i)}{KEY_ROW}", key, STYLE_NOTE)
        for i, (key, _, _) in enumerate(columns, start=1)
    ]
    header_cells = [
        _inline(f"{col_letter(i)}{HEADER_ROW}", header, STYLE_HEADER)
        for i, (_, header, _) in enumerate(columns, start=1)
    ]
    sample_values = [( _example_value(key, kind), kind) for key, _, kind in columns]
    sample_styles = [_style_for(kind) for _, _, kind in columns]
    rows = [
        _row(1, [_inline("A1", title, STYLE_TITLE)], height=30),
        _row(2, [_inline("A2", note, STYLE_NOTE)], height=36),
        _row(KEY_ROW, key_cells, hidden=True),
        _row(HEADER_ROW, header_cells, height=HEADER_HEIGHT),
        _row(DATA_ROW, _cells_for(DATA_ROW, sample_values, sample_styles), height=18),
    ]
    last_col = col_letter(len(columns))
    conditional: list[ET.Element] = []
    priority = priority_start
    keys = [key for key, _, _ in columns]
    for key in bar_keys:
        conditional.append(_data_bar(col_letter(keys.index(key) + 1), priority))
        priority += 1
    for key in status_keys:
        conditional.append(_status_cf(col_letter(keys.index(key) + 1), priority))
        priority += 4
    return _worksheet(
        rows=rows,
        widths=widths,
        freeze_cols=freeze_cols,
        freeze_rows=HEADER_ROW,
        tab_color=tab_color,
        auto_filter=f"A{HEADER_ROW}:{last_col}{HEADER_ROW}",
        merges=[f"A1:{last_col}1", f"A2:{last_col}2"],
        conditional=conditional,
        last_ref=f"A1:{last_col}{DATA_ROW}",
    )


def _summary_sheet() -> ET.Element:
    rows: list[ET.Element] = [
        _row(1, [_inline("A1", "X-Ray トレースレポート", STYLE_TITLE)], height=32),
        _row(
            2,
            [
                _inline(
                    "A2",
                    "Transaction Search が aws/spans に書き込んだスパンを、CloudWatch Logs Insights で時間指定取得した結果です。"
                    "トレース一覧は X-Ray コンソールのトレース画面、タイムラインはトレース詳細のセグメントツリーに合わせています。",
                    STYLE_SUBTITLE,
                )
            ],
            height=36,
        ),
        _row(4, [_inline("A4", "クエリ", STYLE_SECTION)], height=24),
    ]
    for name, cell, label in SUMMARY_NAMES:
        row_index = int(cell[1:])
        if row_index == 19:
            rows.append(_row(18, [_inline("A18", "結果", STYLE_SECTION)], height=24))
        label_style = STYLE_LABEL
        value_style = STYLE_QUERY if name == "QueryString" else STYLE_VALUE
        height = 108 if name == "QueryString" else 48 if name == "ParseNotes" else 22
        rows.append(
            _row(
                row_index,
                [
                    _inline(f"A{row_index}", label, label_style),
                    _inline(cell, "", value_style),
                ],
                height=height,
            )
        )
    rows.extend(
        [
            _row(32, [_inline("A32", "ステータスの見方（X-Ray コンソールと同じ区分）", STYLE_SECTION)], height=24),
            _row(33, [_inline("A33", "OK", STYLE_LEGEND["OK"]), _inline("B33", "成功。HTTP 2xx / 3xx、または失敗を示すフラグがないスパン。", STYLE_TEXT)]),
            _row(34, [_inline("A34", "エラー", STYLE_LEGEND["エラー"]), _inline("B34", "クライアントエラー。HTTP 4xx（429 を除く）、OTel status=ERROR、記録された例外。", STYLE_TEXT)]),
            _row(35, [_inline("A35", "障害", STYLE_LEGEND["障害"]), _inline("B35", "サーバー障害。HTTP 5xx、または X-Ray の fault フラグ。", STYLE_TEXT)]),
            _row(36, [_inline("A36", "スロットル", STYLE_LEGEND["スロットル"]), _inline("B36", "HTTP 429、またはスロットリング。", STYLE_TEXT)]),
            _row(38, [_inline("A38", "シート", STYLE_SECTION)], height=24),
            _row(39, [_inline("A39", "トレース一覧", STYLE_LABEL), _inline("B39", "X-Ray のトレース一覧。応答時間はルートスパンの所要時間、トレース時間は最古の開始から最新の終了までです。ルートステータスは一覧のアイコン、トレースステータスはツリー内の最悪値です。応答時間バーは、このレポート内の最大応答時間との比です。", STYLE_WRAP)], height=48),
            _row(40, [_inline("A40", "タイムライン", STYLE_LABEL), _inline("B40", "トレース詳細のセグメントツリーです。名前は親子の深さだけ字下げしています。タイムライン列の ■ は所要時間、□ は待ち時間です。新しいトレースが上です。", STYLE_WRAP)], height=36),
            _row(41, [_inline("A41", "スパン詳細", STYLE_LABEL), _inline("B41", "セグメントを選んだときに X-Ray の右側へ出る情報です。注釈、メタデータ、例外、HTTP、AWS リモート呼び出し、属性 JSON を含みます。", STYLE_WRAP)], height=32),
            _row(42, [_inline("A42", "サービス集計", STYLE_LABEL), _inline("B42", "サービスごとの件数とレイテンシです。p50 / p95 / p99 は取得したスパンから算出しています。", STYLE_WRAP)], height=32),
            _row(43, [_inline("A43", "クエリ結果", STYLE_LABEL), _inline("B43", "Logs Insights の生結果です。長いセルは Excel の上限に合わせて切り詰めています。完全な値は生 CSV にあります。@ptr は既定では含みません。", STYLE_WRAP)], height=32),
            _row(45, [_inline("A45", "テンプレート", STYLE_SECTION)], height=24),
            _row(46, [_inline("A46", "契約", STYLE_LABEL), _inline("B46", "各表の 3 行目は機械可読のキーで、非表示です。定義名（QueryLogGroup など）もレポート投入先です。表示見出し、列幅、色、条件付き書式は変えて構いません。列を並べ替えるときは 3 行目のキーと条件付き書式の範囲も一緒に動かしてください。5 行目の見本は投入時に置き換わります。", STYLE_WRAP)], height=48),
        ]
    )
    return _worksheet(
        rows=rows,
        widths=[34, 118],
        freeze_cols=0,
        freeze_rows=0,
        tab_color="FF2F2F2F",
        auto_filter=None,
        merges=["A1:B1", "A2:B2", "A4:B4", "A18:B18", "A32:B32", "A38:B38", "A45:B45"],
        conditional=[],
        last_ref="A1:B46",
    )


def _raw_sheet() -> ET.Element:
    rows = [
        _row(1, [_inline("A1", "クエリ結果", STYLE_TITLE)], height=30),
        _row(
            2,
            [_inline("A2", "CloudWatch Logs Insights の生結果です。見出しと行は投入時に置き換わります。完全な CSV は出力ディレクトリの spans-raw.csv です。", STYLE_NOTE)],
            height=32,
        ),
        _row(HEADER_ROW, [_inline(f"A{HEADER_ROW}", "（生成時に置き換わります）", STYLE_HEADER)], height=HEADER_HEIGHT),
    ]
    return _worksheet(
        rows=rows,
        widths=[24] * 8,
        freeze_cols=0,
        freeze_rows=HEADER_ROW,
        tab_color="FF7F7F7F",
        auto_filter=f"A{HEADER_ROW}:A{HEADER_ROW}",
        merges=["A1:H1", "A2:H2"],
        conditional=[],
        last_ref=f"A1:H{HEADER_ROW}",
    )


def _font_xml(size: str, color: str, *, bold: bool = False, italic: bool = False) -> str:
    flags = ""
    if bold:
        flags += "<b/>"
    if italic:
        flags += "<i/>"
    return (
        f"<font>{flags}<sz val=\"{size}\"/><color rgb=\"{color}\"/>"
        f"<name val=\"{FONT_NAME}\"/><family val=\"3\"/><charset val=\"128\"/></font>"
    )


def _styles_xml() -> bytes:
    # Indexes are the cell style contract used by the filler. Do not reorder.
    fonts = [
        _font_xml("10", INK),  # 0 body
        _font_xml("16", INK, bold=True),  # 1 title
        _font_xml("9", INK_SOFT),  # 2 subtitle
        _font_xml("11", WHITE, bold=True),  # 3 section
        _font_xml("10", INK, bold=True),  # 4 label
        _font_xml("10", INK),  # 5 value
        _font_xml("10", WHITE, bold=True),  # 6 header
        _font_xml("10", INK),  # 7 text
        _font_xml("10", INK),  # 8 milliseconds
        _font_xml("10", INK),  # 9 integer
        _font_xml("10", INK),  # 10 reserved
        _font_xml("10", INK),  # 11 timeline
        _font_xml("10", INK),  # 12 wrap
        _font_xml("10", INK_SOFT, italic=True),  # 13 reserved italic
        _font_xml("9", INK_SOFT, italic=True),  # 14 note
        _font_xml("10", INK, bold=True),  # 15 OK
        _font_xml("10", INK, bold=True),  # 16 エラー
        _font_xml("10", INK, bold=True),  # 17 障害
        _font_xml("10", INK, bold=True),  # 18 スロットル
        _font_xml("10", INK),  # 19 reserved
        _font_xml("10", INK, bold=True),  # 20 trace id
        _font_xml("9", INK),  # 21 query
    ]
    fills = [
        "none",
        "gray",
        CHARCOAL,  # 2 section / header
        FILL_BAND,  # 3 label and query band
        FILL_OK,  # 4
        FILL_ERROR,  # 5
        FILL_FAULT,  # 6
        FILL_THROTTLE,  # 7
        FILL_WHITE,  # 8 body
    ]
    fill_body = [
        '<fill><patternFill patternType="none"/></fill>',
        '<fill><patternFill patternType="gray125"/></fill>',
    ]
    for rgb in fills[2:]:
        fill_body.append(
            "<fill><patternFill patternType=\"solid\">"
            f'<fgColor rgb="{rgb}"/><bgColor rgb="{rgb}"/></patternFill></fill>'
        )
    # font, fill, numFmt, border, alignment
    # border 1 is the hairline grid. Title, subtitle, and notes stay open.
    xfs = [
        (0, 0, 0, 0, "body"),
        (1, 0, 0, 0, "title"),
        (2, 0, 0, 0, "wrap"),
        (3, 2, 0, 0, "section"),
        (4, 3, 0, 1, "body"),
        (5, 8, 0, 1, "value"),
        (6, 2, 0, 1, "header"),
        (7, 8, 0, 1, "body"),
        (8, 8, 164, 1, "num"),
        (9, 8, 165, 1, "num"),
        (7, 8, 0, 1, "body"),
        (11, 8, 0, 1, "bar"),
        (12, 8, 0, 1, "wrap"),
        (13, 8, 0, 1, "body"),
        (14, 8, 0, 0, "wrap"),
        (15, 4, 0, 1, "body"),
        (16, 5, 0, 1, "body"),
        (17, 6, 0, 1, "body"),
        (18, 7, 0, 1, "body"),
        (9, 8, 165, 1, "num"),
        (20, 8, 0, 1, "body"),
        (21, 3, 0, 1, "wrap"),
    ]
    alignments = {
        "title": '<alignment vertical="center"/>',
        "section": '<alignment vertical="center" indent="1"/>',
        "body": '<alignment vertical="center" wrapText="1"/>',
        "value": '<alignment vertical="center" wrapText="1"/>',
        "header": '<alignment horizontal="center" vertical="center" wrapText="1"/>',
        "num": '<alignment horizontal="right" vertical="center"/>',
        "bar": '<alignment vertical="center" shrinkToFit="0"/>',
        "wrap": '<alignment vertical="top" wrapText="1"/>',
    }
    xf_body = []
    for font_id, fill_id, num_fmt, border_id, align in xfs:
        alignment = alignments[align]
        num = f' numFmtId="{num_fmt}"' if num_fmt else ' numFmtId="0"'
        apply_num = ' applyNumberFormat="1"' if num_fmt else ""
        apply_fill = ' applyFill="1"' if fill_id else ""
        apply_border = ' applyBorder="1"' if border_id else ""
        xf_body.append(
            f'<xf fontId="{font_id}" fillId="{fill_id}" borderId="{border_id}"{num}'
            f' applyFont="1"{apply_fill}{apply_border}{apply_num} applyAlignment="1" xfId="0">'
            f"{alignment}</xf>"
        )
    dxf_fills = [FILL_OK, FILL_ERROR, FILL_FAULT, FILL_THROTTLE]
    dxfs = []
    for fill_rgb in dxf_fills:
        dxfs.append(
            "<dxf>"
            f'<font><b/><sz val="10"/><color rgb="{INK}"/>'
            f'<name val="{FONT_NAME}"/><family val="3"/><charset val="128"/></font>'
            '<fill><patternFill patternType="solid">'
            f'<fgColor rgb="{fill_rgb}"/><bgColor rgb="{fill_rgb}"/></patternFill></fill>'
            '<border>'
            f'<left style="thin"><color rgb="{GRID}"/></left>'
            f'<right style="thin"><color rgb="{GRID}"/></right>'
            f'<top style="thin"><color rgb="{GRID}"/></top>'
            f'<bottom style="thin"><color rgb="{GRID}"/></bottom>'
            "</border>"
            "</dxf>"
        )
    grid = (
        f'<left style="thin"><color rgb="{GRID}"/></left>'
        f'<right style="thin"><color rgb="{GRID}"/></right>'
        f'<top style="thin"><color rgb="{GRID}"/></top>'
        f'<bottom style="thin"><color rgb="{GRID}"/></bottom>'
        "<diagonal/>"
    )
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{NS}">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="0.000"/>
    <numFmt numFmtId="165" formatCode="#,##0"/>
  </numFmts>
  <fonts count="{len(fonts)}">{"".join(fonts)}</fonts>
  <fills count="{len(fill_body)}">{"".join(fill_body)}</fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>{grid}</border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="{len(xf_body)}">{"".join(xf_body)}</cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="{len(dxfs)}">{"".join(dxfs)}</dxfs>
</styleSheet>
'''
    return xml.encode("utf-8")


def _xml_bytes(element: ET.Element) -> bytes:
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def _workbook(sheet_names: list[str]) -> bytes:
    root = ET.Element(_q("workbook"))
    ET.SubElement(root, _q("workbookPr"))
    ET.SubElement(root, _q("workbookView"), {"activeTab": "0"})
    sheets = ET.SubElement(root, _q("sheets"))
    for index, name in enumerate(sheet_names, start=1):
        ET.SubElement(
            sheets,
            _q("sheet"),
            {"name": name, "sheetId": str(index), f"{{{NS_R}}}id": f"rId{index}"},
        )
    defined = ET.SubElement(root, _q("definedNames"))
    for name, cell, _label in SUMMARY_NAMES:
        node = ET.SubElement(defined, _q("definedName"), {"name": name})
        node.text = f"'{SUMMARY_SHEET}'!${cell[0]}${cell[1:]}"
    for sheet_index, sheet_name in enumerate(sheet_names):
        if sheet_name == SUMMARY_SHEET:
            continue
        node = ET.SubElement(
            defined,
            _q("definedName"),
            {"name": "_xlnm.Print_Titles", "localSheetId": str(sheet_index)},
        )
        node.text = f"'{sheet_name}'!${HEADER_ROW}:${HEADER_ROW}"
    ET.SubElement(root, _q("calcPr"), {"calcMode": "auto"})
    return _xml_bytes(root)


def _content_types(sheet_count: int) -> bytes:
    overrides = [
        f'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        f'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        f'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        f'<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for index in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Types xmlns="{NS_CT}">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )
    return xml.encode("utf-8")


def _rels() -> bytes:
    xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )
    return xml.encode("utf-8")


def _workbook_rels(sheet_count: int) -> bytes:
    items = []
    for index in range(1, sheet_count + 1):
        items.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    items.append(
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    xml = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_REL}">'
        + "".join(items)
        + "</Relationships>"
    )
    return xml.encode("utf-8")


def _core() -> bytes:
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="{NS_CP}" xmlns:dc="{NS_DC}" xmlns:dcterms="{NS_DCTERMS}" xmlns:xsi="{NS_XSI}">
  <dc:title>X-Ray トレースレポート</dc:title>
  <dc:subject>Transaction Search aws/spans</dc:subject>
  <dc:creator>export-span-traces</dc:creator>
  <cp:lastModifiedBy>export-span-traces</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:modified>
</cp:coreProperties>
'''
    return xml.encode("utf-8")


def _app(sheet_names: list[str]) -> bytes:
    titles = "".join(f"<vt:lpstr>{_xml_escape(name)}</vt:lpstr>" for name in sheet_names)
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="{NS_EXT}" xmlns:vt="{NS_VT}">
  <Application>export-span-traces</Application>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="{len(sheet_names)}" baseType="lpstr">{titles}</vt:vector>
  </TitlesOfParts>
</Properties>
'''
    return xml.encode("utf-8")


def _xml_escape(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def build_template(path: Path) -> None:
    sheets = [
        (SUMMARY_SHEET, _summary_sheet()),
        (
            TRACE_SHEET,
            _table_sheet(
                TRACE_COLUMNS,
                title="トレース一覧",
                note="X-Ray コンソールのトレース一覧に相当します。5 行目の見本は投入時に置き換わります。ルートステータスはルートスパン、トレースステータスはツリー内の最悪値です。",
                tab_color="FF3F3F3F",
                freeze_cols=1,
                status_keys=["root_status", "trace_status"],
                bar_keys=["response_time_ms"],
                priority_start=1,
            ),
        ),
        (
            TIMELINE_SHEET,
            _table_sheet(
                TIMELINE_COLUMNS,
                title="タイムライン",
                note="X-Ray のトレース詳細タイムラインに相当します。■ の位置が開始オフセット、長さが所要時間です。□ はその前後です。新しいトレースが上で、スパンは親子の順です。",
                tab_color="FF4F4F4F",
                freeze_cols=4,
                status_keys=["status"],
                bar_keys=["duration_ms"],
                priority_start=20,
            ),
        ),
        (
            DETAIL_SHEET,
            _table_sheet(
                DETAIL_COLUMNS,
                title="スパン詳細",
                note="X-Ray でセグメントを選んだときの詳細ペインに相当します。注釈、メタデータ、例外、属性を含みます。",
                tab_color="FF5F5F5F",
                freeze_cols=2,
                status_keys=["status"],
                bar_keys=[],
                priority_start=40,
            ),
        ),
        (
            SERVICE_SHEET,
            _table_sheet(
                SERVICE_COLUMNS,
                title="サービス集計",
                note="取得したスパンをサービス名で集計しています。サービス名は aws.local.service、なければ resource の service.name です。",
                tab_color="FF6F6F6F",
                freeze_cols=1,
                status_keys=[],
                bar_keys=["avg_ms"],
                priority_start=60,
            ),
        ),
        (RAW_SHEET, _raw_sheet()),
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    names = [name for name, _element in sheets]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        _write_zip(archive, "[Content_Types].xml", _content_types(len(sheets)))
        _write_zip(archive, "_rels/.rels", _rels())
        _write_zip(archive, "xl/workbook.xml", _workbook(names))
        _write_zip(archive, "xl/_rels/workbook.xml.rels", _workbook_rels(len(sheets)))
        _write_zip(archive, "xl/styles.xml", _styles_xml())
        _write_zip(archive, "docProps/core.xml", _core())
        _write_zip(archive, "docProps/app.xml", _app(names))
        for index, (_name, element) in enumerate(sheets, start=1):
            _write_zip(archive, f"xl/worksheets/sheet{index}.xml", _xml_bytes(element))


def _write_zip(archive: zipfile.ZipFile, name: str, data: bytes) -> None:
    info = zipfile.ZipInfo(filename=name, date_time=(2026, 1, 1, 0, 0, 0))
    info.compress_type = zipfile.ZIP_DEFLATED
    archive.writestr(info, data)


def _parse(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def _sheet_targets(workbook: bytes, rels: bytes) -> dict[str, str]:
    wb = _parse(workbook)
    relationships = _parse(rels)
    targets = {}
    for rel in relationships:
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rel_id and target and "worksheets/" in target:
            targets[rel_id] = "xl/" + target if not target.startswith("xl/") else target
    mapping = {}
    sheets = wb.find(_q("sheets"))
    if sheets is None:
        raise ValueError("テンプレートにシートがありません")
    for sheet in sheets:
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{NS_R}}}id")
        if not name or not rel_id or rel_id not in targets:
            raise ValueError(f"シート {name} の参照を解決できません")
        mapping[name] = targets[rel_id]
    return mapping


def _defined_names(workbook: bytes) -> dict[str, tuple[str, str]]:
    wb = _parse(workbook)
    found: dict[str, tuple[str, str]] = {}
    container = wb.find(_q("definedNames"))
    if container is None:
        return found
    for node in container.findall(_q("definedName")):
        name = node.attrib.get("name") or ""
        if name.startswith("_xlnm."):
            continue
        ref = (node.text or "").strip()
        if "!" not in ref:
            continue
        sheet_part, cell_part = ref.split("!", 1)
        sheet = sheet_part.strip().strip("'").replace("''", "'")
        cell = cell_part.replace("$", "")
        if ":" in cell:
            continue
        found[name] = (sheet, cell)
    return found


def _column_styles(sheet_root: ET.Element) -> dict[int, str | None]:
    styles: dict[int, str | None] = {}
    for row in _rows(sheet_root):
        if int(row.attrib.get("r", "0")) != DATA_ROW:
            continue
        for cell in row.findall(_q("c")):
            ref = cell.attrib.get("r", "")
            column = _column_index(ref)
            styles[column] = cell.attrib.get("s")
    return styles


def _rows(sheet_root: ET.Element) -> list[ET.Element]:
    data = sheet_root.find(_q("sheetData"))
    if data is None:
        raise ValueError("sheetData がありません")
    return list(data.findall(_q("row")))


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch.upper()) - 64)
    return index


def _cell_text(cell: ET.Element) -> str:
    inline = cell.find(f"{_q('is')}/{_q('t')}")
    if inline is not None and inline.text:
        return inline.text
    value = cell.find(_q("v"))
    if value is not None and value.text:
        return value.text
    return ""


def _key_map(sheet_root: ET.Element) -> list[str]:
    for row in _rows(sheet_root):
        if int(row.attrib.get("r", "0")) != KEY_ROW:
            continue
        indexed: dict[int, str] = {}
        for cell in row.findall(_q("c")):
            indexed[_column_index(cell.attrib.get("r", ""))] = _cell_text(cell)
        if not indexed:
            return []
        last = max(indexed)
        return [indexed.get(index, "") for index in range(1, last + 1)]
    return []


def _set_dimension(sheet_root: ET.Element, ref: str) -> None:
    dimension = sheet_root.find(_q("dimension"))
    if dimension is None:
        dimension = ET.Element(_q("dimension"))
        sheet_root.insert(0, dimension)
    dimension.set("ref", ref)


def _set_autofilter(sheet_root: ET.Element, ref: str) -> None:
    auto_filter = sheet_root.find(_q("autoFilter"))
    if auto_filter is None:
        return
    auto_filter.set("ref", ref)


def _write_table_rows(
    sheet_root: ET.Element,
    records: list[dict[str, Any]],
    *,
    dynamic_headers: list[str] | None = None,
) -> None:
    data = sheet_root.find(_q("sheetData"))
    if data is None:
        raise ValueError("sheetData がありません")
    styles = _column_styles(sheet_root)
    keys = dynamic_headers if dynamic_headers is not None else _key_map(sheet_root)
    if dynamic_headers is not None:
        _replace_header(data, dynamic_headers)
    kept = []
    for row in list(data.findall(_q("row"))):
        row_index = int(row.attrib.get("r", "0"))
        if row_index >= DATA_ROW:
            data.remove(row)
        else:
            kept.append(row_index)
    numeric_places = {"ms": 3, "int": 0}
    column_kinds = _kinds_for(keys, dynamic_headers is not None)
    for offset, record in enumerate(records):
        row_index = DATA_ROW + offset
        values: list[tuple[Any, str]] = []
        style_list: list[int | None] = []
        for column, key in enumerate(keys, start=1):
            kind = column_kinds.get(key, "text")
            values.append((record.get(key), kind))
            raw_style = styles.get(column)
            style_list.append(int(raw_style) if raw_style is not None else _style_for(kind if kind in {"ms", "int", "bar", "wrap", "trace", "text"} else "text"))
        data.append(_row(row_index, _cells_for(row_index, values, style_list)))
    last_row = DATA_ROW + len(records) - 1 if records else HEADER_ROW
    last_col_index = max(1, len(keys))
    last_col = col_letter(last_col_index)
    _set_dimension(sheet_root, f"A1:{last_col}{max(last_row, 2)}")
    filter_end = last_row if records else HEADER_ROW
    _set_autofilter(sheet_root, f"A{HEADER_ROW}:{last_col}{filter_end}")


def _kinds_for(keys: list[str], dynamic: bool) -> dict[str, str]:
    if dynamic:
        return {key: "text" for key in keys}
    known: dict[str, str] = {}
    for columns in (TRACE_COLUMNS, TIMELINE_COLUMNS, DETAIL_COLUMNS, SERVICE_COLUMNS):
        for key, _header, kind in columns:
            known[key] = kind
    return {key: known.get(key, "text") for key in keys}


def _replace_header(sheet_data: ET.Element, headers: list[str]) -> None:
    for row in list(sheet_data.findall(_q("row"))):
        if int(row.attrib.get("r", "0")) == HEADER_ROW:
            sheet_data.remove(row)
    cells = [
        _inline(f"{col_letter(index)}{HEADER_ROW}", excel_text(header), STYLE_HEADER)
        for index, header in enumerate(headers, start=1)
        if header
    ]
    sheet_data.append(_row(HEADER_ROW, cells, height=HEADER_HEIGHT))
    # Keep header row in row-number order. Excel tolerates disorder, but sort
    # so later readers and tests see a stable document.
    rows = list(sheet_data.findall(_q("row")))
    for row in rows:
        sheet_data.remove(row)
    rows.sort(key=lambda item: int(item.attrib.get("r", "0")))
    for row in rows:
        sheet_data.append(row)


def _set_named_cell(sheet_root: ET.Element, cell_ref: str, value: Any, style: int) -> None:
    data = sheet_root.find(_q("sheetData"))
    if data is None:
        raise ValueError("sheetData がありません")
    row_index = int("".join(ch for ch in cell_ref if ch.isdigit()))
    row = None
    for candidate in data.findall(_q("row")):
        if int(candidate.attrib.get("r", "0")) == row_index:
            row = candidate
            break
    if row is None:
        row = _row(row_index, [])
        data.append(row)
    for cell in list(row.findall(_q("c"))):
        if cell.attrib.get("r") == cell_ref:
            row.remove(cell)
    text = excel_text(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        places = 0 if float(value).is_integer() else 3
        row.append(_number(cell_ref, value, style, places))
    else:
        row.append(_inline(cell_ref, text, style))
    cells = list(row.findall(_q("c")))
    for cell in cells:
        row.remove(cell)
    cells.sort(key=lambda item: _column_index(item.attrib.get("r", "A1")))
    for cell in cells:
        row.append(cell)


def fill_template(
    template: Path,
    output: Path,
    *,
    summary: dict[str, Any],
    traces: list[dict[str, Any]],
    timeline: list[dict[str, Any]],
    details: list[dict[str, Any]],
    services: list[dict[str, Any]],
    raw_rows: list[dict[str, str]],
    raw_columns: list[str],
) -> None:
    if not template.is_file():
        raise FileNotFoundError(f"Excel テンプレートがありません: {template}")
    with zipfile.ZipFile(template) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    targets = _sheet_targets(files["xl/workbook.xml"], files["xl/_rels/workbook.xml.rels"])
    names = _defined_names(files["xl/workbook.xml"])
    missing = [name for name, _cell, _label in SUMMARY_NAMES if name not in names]
    if missing:
        raise ValueError("テンプレートの定義名が不足しています: " + ", ".join(missing))
    summary_sheet_name, _summary_cell = names["QueryLogGroup"]
    summary_root = _parse(files[targets[summary_sheet_name]])
    for name, _cell, _label in SUMMARY_NAMES:
        sheet_name, cell = names[name]
        if sheet_name != summary_sheet_name:
            raise ValueError(f"定義名 {name} が概要シートにありません")
        style = STYLE_QUERY if name in {"QueryString", "ParseNotes"} else STYLE_VALUE
        _set_named_cell(summary_root, cell, summary.get(name, ""), style)
    files[targets[summary_sheet_name]] = _xml_bytes(summary_root)

    table_payload = {
        TRACE_SHEET: traces,
        TIMELINE_SHEET: timeline,
        DETAIL_SHEET: details,
        SERVICE_SHEET: services,
    }
    for sheet_name, records in table_payload.items():
        if sheet_name not in targets:
            raise ValueError(f"テンプレートにシートがありません: {sheet_name}")
        root = _parse(files[targets[sheet_name]])
        if not _key_map(root):
            raise ValueError(f"{sheet_name} の非表示キー行がありません")
        _write_table_rows(root, records)
        files[targets[sheet_name]] = _xml_bytes(root)
    if RAW_SHEET not in targets:
        raise ValueError("テンプレートにクエリ結果シートがありません")
    raw_root = _parse(files[targets[RAW_SHEET]])
    raw_records = [{column: row.get(column, "") for column in raw_columns} for row in raw_rows]
    _write_table_rows(raw_root, raw_records, dynamic_headers=raw_columns or ["(no rows)"])
    files[targets[RAW_SHEET]] = _xml_bytes(raw_root)

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            _write_zip(archive, name, data)


def read_sheet_matrix(path: Path, sheet_name: str) -> list[list[str]]:
    """Return each row as a dense list of displayed cell text. Test helper."""
    with zipfile.ZipFile(path) as archive:
        targets = _sheet_targets(archive.read("xl/workbook.xml"), archive.read("xl/_rels/workbook.xml.rels"))
        root = _parse(archive.read(targets[sheet_name]))
    matrix: list[list[str]] = []
    for row in _rows(root):
        indexed: dict[int, str] = {}
        for cell in row.findall(_q("c")):
            indexed[_column_index(cell.attrib.get("r", ""))] = _cell_text(cell)
        if not indexed:
            matrix.append([])
            continue
        last = max(indexed)
        matrix.append([indexed.get(index, "") for index in range(1, last + 1)])
    return matrix


def read_defined_values(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        names = _defined_names(archive.read("xl/workbook.xml"))
        targets = _sheet_targets(archive.read("xl/workbook.xml"), archive.read("xl/_rels/workbook.xml.rels"))
        cache: dict[str, ET.Element] = {}
        values: dict[str, str] = {}
        for name, (sheet_name, cell_ref) in names.items():
            if sheet_name not in cache:
                cache[sheet_name] = _parse(archive.read(targets[sheet_name]))
            root = cache[sheet_name]
            row_index = int("".join(ch for ch in cell_ref if ch.isdigit()))
            found = ""
            for row in _rows(root):
                if int(row.attrib.get("r", "0")) != row_index:
                    continue
                for cell in row.findall(_q("c")):
                    if cell.attrib.get("r") == cell_ref:
                        found = _cell_text(cell)
            values[name] = found
        return values


def utc_now_text() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S +0000")
