"""Parse aws/spans log events into an X-Ray-like trace model."""

from __future__ import annotations

import json
import math
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any

from spanreport.query import to_xray_trace_id
from spanreport.timeutil import format_unix_nano


WATERFALL_WIDTH = 32
RESPONSE_BAR_WIDTH = 20
# Meiryo UI is proportional. These two glyphs share the em square, so a
# timeline drawn with them stays aligned.
BAR_ON = "\u25a0"
BAR_OFF = "\u25a1"
EXCEL_TEXT_LIMIT = 32_000

STATUS_OK = "OK"
STATUS_ERROR = "エラー"
STATUS_FAULT = "障害"
STATUS_THROTTLE = "スロットル"
STATUS_RANK = {
    STATUS_OK: 0,
    STATUS_ERROR: 1,
    STATUS_THROTTLE: 2,
    STATUS_FAULT: 3,
}

_KIND = {
    "0": "UNSPECIFIED",
    "1": "INTERNAL",
    "2": "SERVER",
    "3": "CLIENT",
    "4": "PRODUCER",
    "5": "CONSUMER",
    "SPAN_KIND_UNSPECIFIED": "UNSPECIFIED",
    "SPAN_KIND_INTERNAL": "INTERNAL",
    "SPAN_KIND_SERVER": "SERVER",
    "SPAN_KIND_CLIENT": "CLIENT",
    "SPAN_KIND_PRODUCER": "PRODUCER",
    "SPAN_KIND_CONSUMER": "CONSUMER",
    "UNSPECIFIED": "UNSPECIFIED",
    "INTERNAL": "INTERNAL",
    "SERVER": "SERVER",
    "CLIENT": "CLIENT",
    "PRODUCER": "PRODUCER",
    "CONSUMER": "CONSUMER",
}
_STATUS = {
    "0": "UNSET",
    "1": "OK",
    "2": "ERROR",
    "STATUS_CODE_UNSET": "UNSET",
    "STATUS_CODE_OK": "OK",
    "STATUS_CODE_ERROR": "ERROR",
    "UNSET": "UNSET",
    "OK": "OK",
    "ERROR": "ERROR",
}
_ZERO_PARENT = re.compile(r"^0+$")
_CONTROL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# Shown in their own columns, so they are not repeated as annotations.
_SEMANTIC_PREFIXES = (
    "http.",
    "url.",
    "db.",
    "net.",
    "network.",
    "server.",
    "client.",
    "rpc.",
    "exception.",
    "error.",
    "otel.",
    "telemetry.",
    "thread.",
    "code.",
    "messaging.",
    "faas.",
    "gen_ai.",
    "aws.",
    "enduser.",
    "user_agent.",
    "peer.",
    "service.",
    "cloud.",
    "host.",
    "os.",
    "process.",
    "k8s.",
    "container.",
    "deployment.",
    "span.",
)

# Column aliases produced by the built-in Logs Insights query.
_FLAT = {
    "traceId": ("traceId",),
    "spanId": ("spanId",),
    "parentSpanId": ("parentSpanId",),
    "name": ("name",),
    "kind": ("kind",),
    "durationNano": ("durationNano",),
    "startTimeUnixNano": ("startTimeUnixNano",),
    "endTimeUnixNano": ("endTimeUnixNano",),
    "statusCode": ("status", "code"),
    "awsLocalService": ("attributes", "aws.local.service"),
    "awsRemoteService": ("attributes", "aws.remote.service"),
    "awsLocalOperation": ("attributes", "aws.local.operation"),
    "awsRemoteOperation": ("attributes", "aws.remote.operation"),
    "awsSpanKind": ("attributes", "aws.span.kind"),
    "httpMethod": ("attributes", "http.request.method"),
    "httpMethodLegacy": ("attributes", "http.method"),
    "httpStatus": ("attributes", "http.response.status_code"),
    "httpStatusLegacy": ("attributes", "http.status_code"),
    "httpRoute": ("attributes", "http.route"),
    "urlPath": ("attributes", "url.path"),
    "urlFull": ("attributes", "url.full"),
    "httpUrl": ("attributes", "http.url"),
    "httpTarget": ("attributes", "http.target"),
    "dbSystem": ("attributes", "db.system"),
    "dbStatement": ("attributes", "db.statement"),
    "errorType": ("attributes", "error.type"),
    "exceptionType": ("attributes", "exception.type"),
    "exceptionMessage": ("attributes", "exception.message"),
    "enduserId": ("attributes", "enduser.id"),
    "clientAddress": ("attributes", "client.address"),
    "clientIp": ("attributes", "http.client_ip"),
    "resourceServiceName": ("resource", "attributes", "service.name"),
    "cloudRegion": ("resource", "attributes", "cloud.region"),
    "cloudAz": ("resource", "attributes", "cloud.availability_zone"),
    "environmentName": ("resource", "attributes", "deployment.environment.name"),
}


@dataclass
class Span:
    trace_id: str
    span_id: str
    parent_span_id: str
    name: str
    service: str
    kind: str
    start_ns: int | None
    end_ns: int | None
    duration_ns: int | None
    otel_status: str
    xray_status: str
    http_method: str
    url: str
    http_status: int | None
    remote_service: str
    remote_operation: str
    origin: str
    client_ip: str
    user: str
    region: str
    availability_zone: str
    environment: str
    db_system: str
    db_statement: str
    annotations: str
    metadata: str
    error_type: str
    error_message: str
    stacktrace: str
    attributes_json: str
    log_stream: str
    ingestion_time: str
    depth: int = 0
    offset_ns: int = 0
    waterfall: str = ""
    display_name: str = ""
    raw: dict[str, Any] = field(default_factory=dict)


@dataclass
class Trace:
    trace_id: str
    trace_id_xray: str
    start_ns: int | None
    end_ns: int | None
    trace_time_ns: int | None
    response_time_ns: int | None
    root_status: str
    trace_status: str
    http_method: str
    url: str
    http_status: int | None
    service: str
    client_ip: str
    user: str
    annotations: str
    spans: list[Span]
    error_count: int
    fault_count: int
    throttle_count: int
    response_bar: str = ""


@dataclass
class ServiceStat:
    service: str
    span_count: int
    trace_ids: set[str]
    ok_count: int
    error_count: int
    fault_count: int
    throttle_count: int
    durations_ms: list[float]


@dataclass
class Report:
    traces: list[Trace]
    spans: list[Span]
    services: list[dict[str, Any]]
    notes: list[str]
    flattened_rows: int = 0
    skipped_rows: int = 0


def excel_text(value: Any, limit: int = EXCEL_TEXT_LIMIT) -> str:
    if value is None:
        return ""
    text = value if isinstance(value, str) else _stringify(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _CONTROL.sub("", text)
    if len(text) > limit:
        return text[: limit - 20] + "…(truncated)"
    return text


def csv_text(value: Any) -> str:
    text = excel_text(value, limit=1_000_000)
    if text.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def _stringify(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)
    try:
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"), default=str)
    except TypeError:
        return str(value)


def as_int(value: Any) -> int | None:
    if value is None or value is False or value is True or value == "":
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            return None
        return int(value)
    text = str(value).strip()
    if re.fullmatch(r"-?\d+", text):
        return int(text)
    if re.fullmatch(r"-?\d+\.0+", text):
        return int(text.split(".", 1)[0])
    return None


def _normalize_kind(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    return _KIND.get(text.upper(), _KIND.get(text, text.upper()))


def _normalize_otel_status(value: Any) -> str:
    text = str(value or "").strip()
    if not text:
        return "UNSET"
    return _STATUS.get(text.upper(), _STATUS.get(text, text.upper()))


def _ns_duration(start: int | None, end: int | None, duration: int | None) -> int | None:
    if duration is not None and duration >= 0:
        return duration
    if start is not None and end is not None and end >= start:
        return end - start
    return None


def _first(mapping: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        value = mapping.get(key)
        if value not in (None, ""):
            return value
    return None


def _attr(span_doc: dict[str, Any], *keys: str) -> Any:
    attributes = span_doc.get("attributes")
    if not isinstance(attributes, dict):
        attributes = {}
    resource = span_doc.get("resource")
    resource_attrs: dict[str, Any] = {}
    if isinstance(resource, dict):
        raw = resource.get("attributes")
        if isinstance(raw, dict):
            resource_attrs = raw
    for key in keys:
        if key in attributes and attributes[key] not in (None, ""):
            return attributes[key]
        if key in resource_attrs and resource_attrs[key] not in (None, ""):
            return resource_attrs[key]
    return None


def _dig(doc: dict[str, Any], path: tuple[str, ...]) -> Any:
    current: Any = doc
    for part in path:
        if not isinstance(current, dict) or part not in current:
            return None
        current = current[part]
    return current


def _put(doc: dict[str, Any], path: tuple[str, ...], value: Any) -> None:
    if value in (None, ""):
        return
    current = doc
    for part in path[:-1]:
        nxt = current.get(part)
        if not isinstance(nxt, dict):
            nxt = {}
            current[part] = nxt
        current = nxt
    leaf = path[-1]
    existing = current.get(leaf)
    if existing in (None, ""):
        current[leaf] = value


def loads_message(value: Any) -> dict[str, Any] | None:
    current = value
    for _ in range(2):
        if isinstance(current, dict):
            return current
        if not isinstance(current, str):
            return None
        text = current.strip()
        if not text or text[0] not in "{[":
            return None
        try:
            current = json.loads(text)
        except json.JSONDecodeError:
            return None
    return current if isinstance(current, dict) else None


def _looks_like_span(doc: dict[str, Any]) -> bool:
    if as_int(doc.get("startTimeUnixNano")) is not None or doc.get("spanId"):
        return True
    if doc.get("traceId") and (doc.get("name") or doc.get("durationNano") is not None):
        return True
    return False


def _is_event_record(doc: dict[str, Any]) -> bool:
    if _looks_like_span(doc):
        return False
    return "timeUnixNano" in doc or ("body" in doc and "spanId" not in doc)


def row_to_document(row: dict[str, str]) -> tuple[dict[str, Any] | None, str]:
    """Return ``(document, source)``.

    ``source`` is ``message``, ``flat``, or ``skip``.
    """
    message = loads_message(row.get("@message"))
    if message is not None and _is_event_record(message):
        return None, "skip"
    flat: dict[str, Any] = {}
    for column, path in _FLAT.items():
        _put(flat, path, row.get(column))
    if message is None:
        if _looks_like_span(flat):
            return flat, "flat"
        if row.get("traceId") or row.get("spanId") or row.get("name"):
            return flat, "flat"
        return None, "skip"
    if not _looks_like_span(message):
        if _looks_like_span(flat):
            return flat, "flat"
        return None, "skip"
    # Discovered columns fill gaps when the message omitted a field.
    for column, path in _FLAT.items():
        if _dig(message, path) in (None, ""):
            _put(message, path, row.get(column))
    return message, "message"


def _annotation_lists(attributes: dict[str, Any]) -> tuple[list[str], list[str]]:
    declared = attributes.get("aws.xray.annotations")
    declared_keys: list[str] = []
    if isinstance(declared, list):
        declared_keys = [str(item) for item in declared]
    elif isinstance(declared, str) and declared.strip():
        try:
            parsed = json.loads(declared)
        except json.JSONDecodeError:
            parsed = None
        if isinstance(parsed, list):
            declared_keys = [str(item) for item in parsed]
        else:
            declared_keys = [part.strip() for part in declared.split(",") if part.strip()]
    annotations: list[str] = []
    metadata: list[str] = []
    if declared_keys:
        for key in declared_keys:
            if key in attributes and key != "aws.xray.annotations":
                annotations.append(f"{key}={_stringify(attributes[key])}")
        for key, value in attributes.items():
            if key == "aws.xray.annotations" or key in declared_keys:
                continue
            if key.startswith(_SEMANTIC_PREFIXES):
                continue
            metadata.append(f"{key}={_stringify(value)}")
        return annotations, metadata
    for key, value in attributes.items():
        if key.startswith(_SEMANTIC_PREFIXES):
            continue
        rendered = _stringify(value)
        if len(rendered) <= 120 and not isinstance(value, (dict, list)):
            annotations.append(f"{key}={rendered}")
        else:
            metadata.append(f"{key}={rendered}")
    return annotations, metadata


def _events(doc: dict[str, Any]) -> tuple[str, str, str]:
    events = doc.get("events")
    if not isinstance(events, list):
        return "", "", ""
    error_type = ""
    error_message = ""
    stacks: list[str] = []
    for event in events:
        if not isinstance(event, dict):
            continue
        name = str(event.get("name") or "")
        attrs = event.get("attributes")
        if not isinstance(attrs, dict):
            attrs = {}
        interesting = name.lower() in {"exception", "error"} or "exception.message" in attrs
        if not interesting:
            continue
        error_type = error_type or str(
            _first(attrs, "exception.type", "error.type", "exception.type") or name
        )
        error_message = error_message or str(
            _first(attrs, "exception.message", "exception.escaped", "message") or ""
        )
        stack = _first(attrs, "exception.stacktrace", "exception.stack_trace")
        if stack:
            stacks.append(str(stack))
    return error_type, error_message, "\n".join(stacks)


def _service_name(doc: dict[str, Any]) -> str:
    value = _attr(
        doc,
        "aws.local.service",
        "service.name",
    )
    if value not in (None, ""):
        return str(value)
    return str(doc.get("name") or "(unknown)")


def _operation(doc: dict[str, Any]) -> str:
    value = _attr(
        doc,
        "url.full",
        "http.url",
        "http.target",
        "url.path",
        "http.route",
        "aws.local.operation",
    )
    if value not in (None, ""):
        return str(value)
    rpc_service = _attr(doc, "rpc.service")
    rpc_method = _attr(doc, "rpc.method")
    if rpc_service or rpc_method:
        return ".".join(part for part in (str(rpc_service or ""), str(rpc_method or "")) if part)
    db = _attr(doc, "db.statement")
    if db not in (None, ""):
        return str(db)
    remote = _attr(doc, "aws.remote.operation")
    if remote not in (None, ""):
        return str(remote)
    return ""


def _truthy(value: Any) -> bool:
    if value is True:
        return True
    if isinstance(value, (int, float)) and value == 1:
        return True
    if isinstance(value, str) and value.strip().lower() in {"1", "true", "yes"}:
        return True
    return False


def classify_status(
    *,
    http_status: int | None,
    otel_status: str,
    attributes: dict[str, Any],
    error_type: str,
    error_message: str,
) -> str:
    """Map a span onto the X-Ray status vocabulary.

    Throttle is HTTP 429. Fault is HTTP 5xx or an X-Ray fault flag. Error is
    HTTP 4xx or an OTel ERROR / recorded exception. Everything else is OK.
    """
    blob = " ".join(
        str(part or "")
        for part in (
            error_type,
            error_message,
            attributes.get("error.type"),
            attributes.get("aws.error"),
        )
    ).lower()
    throttle_flag = _truthy(attributes.get("aws.xray.throttle")) or _truthy(attributes.get("throttle"))
    fault_flag = _truthy(attributes.get("aws.xray.fault")) or _truthy(attributes.get("fault"))
    error_flag = _truthy(attributes.get("aws.xray.error")) or _truthy(attributes.get("error"))
    if http_status == 429 or throttle_flag or "throttl" in blob:
        return STATUS_THROTTLE
    if (http_status is not None and http_status >= 500) or fault_flag:
        return STATUS_FAULT
    if (http_status is not None and http_status >= 400) or error_flag or otel_status == "ERROR":
        return STATUS_ERROR
    if error_type or error_message:
        return STATUS_ERROR
    return STATUS_OK


def worst_status(statuses: list[str]) -> str:
    if not statuses:
        return STATUS_OK
    return max(statuses, key=lambda item: STATUS_RANK.get(item, 0))


def span_from_document(doc: dict[str, Any], row: dict[str, str]) -> Span | None:
    trace_id = str(doc.get("traceId") or row.get("traceId") or "").strip()
    if not trace_id:
        return None
    attributes = doc.get("attributes")
    if not isinstance(attributes, dict):
        attributes = {}
    status_obj = doc.get("status")
    otel_code: Any = ""
    if isinstance(status_obj, dict):
        otel_code = status_obj.get("code")
    otel_status = _normalize_otel_status(otel_code)
    http_status = as_int(
        _attr(doc, "http.response.status_code", "http.status_code")
    )
    start_ns = as_int(doc.get("startTimeUnixNano"))
    end_ns = as_int(doc.get("endTimeUnixNano"))
    duration_ns = _ns_duration(start_ns, end_ns, as_int(doc.get("durationNano")))
    event_type, event_message, stack = _events(doc)
    error_type = str(_attr(doc, "error.type", "exception.type") or event_type or "")
    error_message = str(
        _attr(doc, "exception.message", "error.message")
        or (status_obj.get("message") if isinstance(status_obj, dict) else "")
        or event_message
        or ""
    )
    stacktrace = str(_attr(doc, "exception.stacktrace") or stack or "")
    annotations, metadata = _annotation_lists(attributes)
    parent = str(doc.get("parentSpanId") or "").strip()
    if _ZERO_PARENT.fullmatch(parent or ""):
        parent = ""
    origin = str(
        _attr(doc, "cloud.platform", "faas.name", "rpc.system", "aws.origin") or ""
    )
    return Span(
        trace_id=trace_id,
        span_id=str(doc.get("spanId") or "").strip(),
        parent_span_id=parent,
        name=str(doc.get("name") or "(unnamed)"),
        service=_service_name(doc),
        kind=_normalize_kind(doc.get("kind") or _attr(doc, "aws.span.kind")),
        start_ns=start_ns,
        end_ns=end_ns,
        duration_ns=duration_ns,
        otel_status=otel_status,
        xray_status=classify_status(
            http_status=http_status,
            otel_status=otel_status,
            attributes=attributes,
            error_type=error_type,
            error_message=error_message,
        ),
        http_method=str(_attr(doc, "http.request.method", "http.method") or ""),
        url=_operation(doc),
        http_status=http_status,
        remote_service=str(_attr(doc, "aws.remote.service", "peer.service") or ""),
        remote_operation=str(_attr(doc, "aws.remote.operation", "rpc.method") or ""),
        origin=origin,
        client_ip=str(_attr(doc, "client.address", "http.client_ip", "net.peer.ip", "network.peer.address") or ""),
        user=str(_attr(doc, "enduser.id", "user.id", "aws.xray.user") or ""),
        region=str(_attr(doc, "cloud.region", "aws.region") or ""),
        availability_zone=str(_attr(doc, "cloud.availability_zone") or ""),
        environment=str(_attr(doc, "deployment.environment.name", "deployment.environment") or ""),
        db_system=str(_attr(doc, "db.system") or ""),
        db_statement=str(_attr(doc, "db.statement") or ""),
        annotations="; ".join(annotations),
        metadata="; ".join(metadata),
        error_type=error_type,
        error_message=error_message,
        stacktrace=stacktrace,
        attributes_json=_stringify(attributes) if attributes else "",
        log_stream=str(row.get("@logStream") or ""),
        ingestion_time=str(row.get("@ingestionTime") or row.get("@timestamp") or ""),
        raw=doc,
    )


def _waterfall(offset_ns: int, duration_ns: int, trace_ns: int, width: int = WATERFALL_WIDTH) -> str:
    if trace_ns <= 0:
        return BAR_ON
    start = int(round(offset_ns / trace_ns * width))
    end = int(round((offset_ns + max(duration_ns, 0)) / trace_ns * width))
    if end <= start:
        end = start + 1
    start = max(0, min(width - 1, start))
    end = max(start + 1, min(width, end))
    return BAR_OFF * start + BAR_ON * (end - start) + BAR_OFF * (width - end)


def _magnitude_bar(value_ns: int | None, max_ns: int, width: int = RESPONSE_BAR_WIDTH) -> str:
    if value_ns is None or value_ns <= 0 or max_ns <= 0:
        return ""
    cells = int(round(value_ns / max_ns * width))
    cells = max(1, min(width, cells))
    return BAR_ON * cells


def _primary_root(roots: list[Span]) -> Span:
    def rank(span: Span) -> tuple[int, int, str]:
        kind_rank = 0 if span.kind == "SERVER" else 1 if span.kind == "LOCAL_ROOT" else 2
        return (kind_rank, span.start_ns if span.start_ns is not None else 2**62, span.span_id)

    return sorted(roots, key=rank)[0]


def _order_tree(spans: list[Span]) -> list[Span]:
    canonical: dict[str, Span] = {}
    for span in spans:
        if span.span_id and span.span_id not in canonical:
            canonical[span.span_id] = span
    children: dict[int, list[Span]] = defaultdict(list)
    roots: list[Span] = []
    for span in spans:
        parent = canonical.get(span.parent_span_id) if span.parent_span_id else None
        if parent is not None and parent is not span:
            children[id(parent)].append(span)
        else:
            roots.append(span)

    def sort_key(span: Span) -> tuple[int, str]:
        return (span.start_ns if span.start_ns is not None else 2**62, span.span_id)

    roots.sort(key=sort_key)
    for kids in children.values():
        kids.sort(key=sort_key)
    ordered: list[Span] = []
    seen: set[int] = set()

    def walk(span: Span, depth: int, stack: set[int]) -> None:
        if id(span) in seen or id(span) in stack or depth > 200:
            return
        stack.add(id(span))
        seen.add(id(span))
        span.depth = depth
        span.display_name = ("  " * depth) + span.name
        ordered.append(span)
        for child in children.get(id(span), []):
            walk(child, depth + 1, stack)
        stack.remove(id(span))

    for root in roots:
        walk(root, 0, set())
    # A parent cycle has no root. Emit the remaining spans once instead of
    # dropping them.
    for span in spans:
        if id(span) not in seen:
            walk(span, 0, set())
    return ordered


def _build_trace(trace_id: str, spans: list[Span]) -> Trace:
    ordered = _order_tree(spans)
    starts = [span.start_ns for span in ordered if span.start_ns is not None]
    ends = [span.end_ns for span in ordered if span.end_ns is not None]
    start_ns = min(starts) if starts else None
    end_ns = max(ends) if ends else None
    trace_time = None
    if start_ns is not None and end_ns is not None and end_ns >= start_ns:
        trace_time = end_ns - start_ns
    elif any(span.duration_ns is not None for span in ordered):
        trace_time = max(span.duration_ns or 0 for span in ordered)
    roots = [span for span in ordered if span.depth == 0]
    root = _primary_root(roots) if roots else ordered[0]
    for span in ordered:
        offset = 0
        if start_ns is not None and span.start_ns is not None:
            offset = max(0, span.start_ns - start_ns)
        span.offset_ns = offset
        duration = span.duration_ns or 0
        span.waterfall = _waterfall(offset, duration, trace_time or duration or 1)
    counts = {STATUS_ERROR: 0, STATUS_FAULT: 0, STATUS_THROTTLE: 0}
    for span in ordered:
        if span.xray_status in counts:
            counts[span.xray_status] += 1
    annotations = root.annotations
    return Trace(
        trace_id=trace_id,
        trace_id_xray=to_xray_trace_id(trace_id),
        start_ns=start_ns,
        end_ns=end_ns,
        trace_time_ns=trace_time,
        response_time_ns=root.duration_ns if root.duration_ns is not None else trace_time,
        root_status=root.xray_status,
        trace_status=worst_status([span.xray_status for span in ordered]),
        http_method=root.http_method,
        url=root.url,
        http_status=root.http_status,
        service=root.service,
        client_ip=root.client_ip,
        user=root.user,
        annotations=annotations,
        spans=ordered,
        error_count=counts[STATUS_ERROR],
        fault_count=counts[STATUS_FAULT],
        throttle_count=counts[STATUS_THROTTLE],
    )


def _percentile(sorted_values: list[float], percentile: float) -> float | None:
    if not sorted_values:
        return None
    if len(sorted_values) == 1:
        return sorted_values[0]
    position = (len(sorted_values) - 1) * (percentile / 100.0)
    low = math.floor(position)
    high = math.ceil(position)
    if low == high:
        return sorted_values[low]
    weight = position - low
    return sorted_values[low] * (1 - weight) + sorted_values[high] * weight


def _service_rows(spans: list[Span]) -> list[dict[str, Any]]:
    grouped: dict[str, ServiceStat] = {}
    for span in spans:
        stat = grouped.get(span.service)
        if stat is None:
            stat = ServiceStat(span.service, 0, set(), 0, 0, 0, 0, [])
            grouped[span.service] = stat
        stat.span_count += 1
        stat.trace_ids.add(span.trace_id)
        if span.xray_status == STATUS_OK:
            stat.ok_count += 1
        elif span.xray_status == STATUS_ERROR:
            stat.error_count += 1
        elif span.xray_status == STATUS_FAULT:
            stat.fault_count += 1
        elif span.xray_status == STATUS_THROTTLE:
            stat.throttle_count += 1
        if span.duration_ns is not None:
            stat.durations_ms.append(span.duration_ns / 1_000_000)
    rows: list[dict[str, Any]] = []
    for stat in grouped.values():
        durations = sorted(stat.durations_ms)
        average = sum(durations) / len(durations) if durations else None
        rows.append(
            {
                "service": stat.service,
                "span_count": stat.span_count,
                "trace_count": len(stat.trace_ids),
                "ok_count": stat.ok_count,
                "error_count": stat.error_count,
                "fault_count": stat.fault_count,
                "throttle_count": stat.throttle_count,
                "avg_ms": average,
                "p50_ms": _percentile(durations, 50),
                "p95_ms": _percentile(durations, 95),
                "p99_ms": _percentile(durations, 99),
                "max_ms": durations[-1] if durations else None,
            }
        )
    rows.sort(key=lambda item: (-int(item["span_count"]), str(item["service"])))
    return rows


def _matches_trace_filter(span: Span, trace_id: str) -> bool:
    if not trace_id:
        return True
    from spanreport.query import trace_id_variants

    wanted = {item.lower() for item in trace_id_variants(trace_id)}
    return span.trace_id.lower() in wanted or to_xray_trace_id(span.trace_id).lower() in wanted


def _matches_service_filter(span: Span, service: str) -> bool:
    if not service:
        return True
    target = service.casefold()
    values = [span.service, span.name, span.remote_service]
    return any(target == str(value or "").casefold() for value in values)


def build_report(
    rows: list[dict[str, str]],
    *,
    service: str = "",
    trace_id: str = "",
    limit: int | None = None,
    records_matched: int | None = None,
) -> Report:
    notes: list[str] = []
    spans: list[Span] = []
    flattened = 0
    skipped = 0
    for row in rows:
        document, source = row_to_document(row)
        if document is None:
            skipped += 1
            continue
        if source == "flat":
            flattened += 1
        span = span_from_document(document, row)
        if span is None:
            skipped += 1
            continue
        spans.append(span)
    # Keep every returned span of a matching trace. A service filter then still
    # shows the siblings that the query managed to return, which is what the
    # X-Ray trace screen does once you open a trace.
    if trace_id:
        matched = {span.trace_id for span in spans if _matches_trace_filter(span, trace_id)}
        spans = [span for span in spans if span.trace_id in matched]
    if service:
        matched = {span.trace_id for span in spans if _matches_service_filter(span, service)}
        spans = [span for span in spans if span.trace_id in matched]
    grouped: dict[str, list[Span]] = defaultdict(list)
    for span in spans:
        grouped[span.trace_id].append(span)
    traces = [_build_trace(trace_id_key, items) for trace_id_key, items in grouped.items()]
    traces.sort(key=lambda item: (-(item.start_ns or 0), item.trace_id_xray))
    max_response = max((item.response_time_ns or 0 for item in traces), default=0)
    for trace in traces:
        trace.response_bar = _magnitude_bar(trace.response_time_ns, max_response)
    ordered_spans = [span for trace in traces for span in trace.spans]
    if flattened:
        notes.append(f"@message を解釈できない、または欠落した {flattened} 件は平坦化フィールドから復元しました。")
    if skipped:
        notes.append(f"スパンではない、または traceId がない {skipped} 件を除外しました。")
    returned = len(rows)
    if limit is not None and returned >= limit:
        notes.append(
            f"取得件数が上限 {limit} に達しています。時間範囲を狭めるか、--limit を上げてください（最大 100000）。"
        )
    if records_matched is not None and records_matched > returned:
        notes.append(
            f"クエリの一致件数は {records_matched} 件で、出力したログイベントは {returned} 件です。"
        )
    return Report(
        traces=traces,
        spans=ordered_spans,
        services=_service_rows(ordered_spans),
        notes=notes,
        flattened_rows=flattened,
        skipped_rows=skipped,
    )


def ns_to_ms(value: int | None) -> float | None:
    if value is None:
        return None
    return value / 1_000_000


def trace_row(trace: Trace, zone: Any) -> dict[str, Any]:
    return {
        "trace_id_xray": trace.trace_id_xray,
        "response_time_ms": ns_to_ms(trace.response_time_ns),
        "trace_time_ms": ns_to_ms(trace.trace_time_ns),
        "response_bar": trace.response_bar,
        "start_time": format_unix_nano(trace.start_ns, zone),
        "end_time": format_unix_nano(trace.end_ns, zone),
        "http_method": trace.http_method,
        "url": trace.url,
        "root_status": trace.root_status,
        "trace_status": trace.trace_status,
        "http_status": trace.http_status,
        "service": trace.service,
        "client_ip": trace.client_ip,
        "user": trace.user,
        "span_count": len(trace.spans),
        "error_count": trace.error_count,
        "fault_count": trace.fault_count,
        "throttle_count": trace.throttle_count,
        "annotations": trace.annotations,
        "trace_id_raw": trace.trace_id,
    }


def timeline_row(trace: Trace, span: Span, zone: Any) -> dict[str, Any]:
    return {
        "trace_id_xray": trace.trace_id_xray,
        "depth": span.depth,
        "waterfall": span.waterfall,
        "display_name": span.display_name,
        "service": span.service,
        "offset_ms": ns_to_ms(span.offset_ns),
        "duration_ms": ns_to_ms(span.duration_ns),
        "start_time": format_unix_nano(span.start_ns, zone),
        "end_time": format_unix_nano(span.end_ns, zone),
        "status": span.xray_status,
        "kind": span.kind,
        "http_method": span.http_method,
        "operation": span.url,
        "http_status": span.http_status,
        "remote_service": span.remote_service,
        "remote_operation": span.remote_operation,
        "span_id": span.span_id,
        "parent_span_id": span.parent_span_id,
        "origin": span.origin,
        "error_type": span.error_type,
        "error_message": span.error_message,
    }


def detail_row(trace: Trace, span: Span, zone: Any) -> dict[str, Any]:
    row = timeline_row(trace, span, zone)
    row.update(
        {
            "name": span.name,
            "db_system": span.db_system,
            "db_statement": span.db_statement,
            "region": span.region,
            "availability_zone": span.availability_zone,
            "environment": span.environment,
            "client_ip": span.client_ip,
            "user": span.user,
            "annotations": span.annotations,
            "metadata": span.metadata,
            "stacktrace": span.stacktrace,
            "attributes_json": span.attributes_json,
            "log_stream": span.log_stream,
            "ingestion_time": span.ingestion_time,
            "otel_status": span.otel_status,
            "trace_id_raw": trace.trace_id,
        }
    )
    return row
