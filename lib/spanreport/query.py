"""CloudWatch Logs Insights query construction for the aws/spans log group."""

from __future__ import annotations

import re


MAX_LIMIT = 100_000
DEFAULT_LIMIT = 10_000

# Discovered field names documented for Transaction Search spans. Dotted
# attribute keys are a single Logs Insights field, so they are backticked.
_FIELDS = """\
@timestamp, @ingestionTime, @logStream,
traceId, spanId, parentSpanId, name, kind,
durationNano, startTimeUnixNano, endTimeUnixNano,
status.code as statusCode,
`attributes.aws.local.service` as awsLocalService,
`attributes.aws.remote.service` as awsRemoteService,
`attributes.aws.local.operation` as awsLocalOperation,
`attributes.aws.remote.operation` as awsRemoteOperation,
`attributes.aws.span.kind` as awsSpanKind,
`attributes.http.request.method` as httpMethod,
`attributes.http.method` as httpMethodLegacy,
`attributes.http.response.status_code` as httpStatus,
`attributes.http.status_code` as httpStatusLegacy,
`attributes.http.route` as httpRoute,
`attributes.url.path` as urlPath,
`attributes.url.full` as urlFull,
`attributes.http.url` as httpUrl,
`attributes.http.target` as httpTarget,
`attributes.db.system` as dbSystem,
`attributes.db.statement` as dbStatement,
`attributes.error.type` as errorType,
`attributes.exception.type` as exceptionType,
`attributes.exception.message` as exceptionMessage,
`attributes.enduser.id` as enduserId,
`attributes.client.address` as clientAddress,
`attributes.http.client_ip` as clientIp,
`resource.attributes.service.name` as resourceServiceName,
`resource.attributes.cloud.region` as cloudRegion,
`resource.attributes.cloud.availability_zone` as cloudAz,
`resource.attributes.deployment.environment.name` as environmentName"""

_XRAY_TRACE = re.compile(r"^1-([0-9a-fA-F]{8})-([0-9a-fA-F]{24})$")
_OTEL_TRACE = re.compile(r"^[0-9a-fA-F]{32}$")
_SAFE_LIKE = re.compile(r"[A-Za-z0-9_.:/=+-]{1,200}")


def clamp_limit(value: int) -> int:
    if value < 1 or value > MAX_LIMIT:
        raise ValueError(f"limit は 1 から {MAX_LIMIT} の整数です: {value}")
    return value


def cwli_string(value: str) -> str:
    """Quote a value for a CloudWatch Logs Insights double-quoted string."""
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return '"' + escaped + '"'


def trace_id_variants(trace_id: str) -> list[str]:
    """Return the stored OTel form and the X-Ray console form.

    aws/spans keeps the 32-hex OpenTelemetry trace id. The X-Ray console shows
    ``1-{first8}-{last24}``. Equality filters include both.
    """
    text = trace_id.strip()
    if not text or any(ch in text for ch in "\r\n"):
        raise ValueError("trace id が空、または改行を含んでいます")
    found: list[str] = []

    def add(item: str) -> None:
        if item not in found:
            found.append(item)

    add(text)
    xray = _XRAY_TRACE.fullmatch(text)
    if xray:
        add((xray.group(1) + xray.group(2)).lower())
        add(f"1-{xray.group(1).lower()}-{xray.group(2).lower()}")
        return found
    otel = _OTEL_TRACE.fullmatch(text)
    if otel:
        lowered = text.lower()
        add(lowered)
        add(f"1-{lowered[:8]}-{lowered[8:]}")
    return found


def to_xray_trace_id(trace_id: str) -> str:
    text = (trace_id or "").strip()
    if _XRAY_TRACE.fullmatch(text):
        return text.lower()
    if _OTEL_TRACE.fullmatch(text):
        lowered = text.lower()
        return f"1-{lowered[:8]}-{lowered[8:]}"
    return text


def _like_clause(literal: str) -> str | None:
    if _SAFE_LIKE.fullmatch(literal) is None:
        return None
    pattern = re.escape(literal).replace("/", r"\/")
    return f"@message like /{pattern}/"


def _trace_filter(trace_id: str) -> str:
    parts = [f"traceId = {cwli_string(item)}" for item in trace_id_variants(trace_id)]
    for item in trace_id_variants(trace_id):
        like = _like_clause(item)
        if like:
            parts.append(like)
    return "| filter (" + " or ".join(parts) + ")"


def _service_filter(service: str) -> str:
    if not service or any(ch in service for ch in "\r\n"):
        raise ValueError("service が空、または改行を含んでいます")
    quoted = cwli_string(service)
    # Field equality only. A substring match on @message would admit unrelated
    # spans and could crowd out the real service once the result limit is hit.
    # Parsed spans are filtered again locally on the service name.
    parts = [
        f"`attributes.aws.local.service` = {quoted}",
        f"`attributes.aws.remote.service` = {quoted}",
        f"`attributes.service.name` = {quoted}",
        f"`resource.attributes.service.name` = {quoted}",
        f"name = {quoted}",
    ]
    return "| filter (" + " or ".join(parts) + ")"


def build_query(
    *,
    limit: int = DEFAULT_LIMIT,
    include_message: bool = True,
    service: str = "",
    trace_id: str = "",
    filter_expr: str = "",
    query_file_text: str = "",
) -> str:
    """Build the Logs Insights query.

    ``query_file_text`` replaces the built-in query. Service and trace filters
    are then applied again locally, because appending a filter after ``stats``
    would change the caller's query. ``filter_expr`` is only valid for the
    built-in query.
    """
    if query_file_text.strip():
        if filter_expr.strip():
            raise ValueError(
                "--filter は既定クエリ専用です。--query-file の中に filter を書いてください"
            )
        return query_file_text.strip() + "\n"

    if filter_expr and any(ch in filter_expr for ch in "\r\n"):
        raise ValueError("--filter に改行は使えません")
    limit = clamp_limit(limit)
    lines = ["fields " + " ".join(_FIELDS.split())]
    if include_message:
        lines[0] += ", @message"
    # Do not require ispresent(traceId). Field discovery for aws/spans can lag,
    # and a missing discovered field would drop every row before @message is
    # parsed locally.
    if trace_id.strip():
        lines.append(_trace_filter(trace_id.strip()))
    if service.strip():
        lines.append(_service_filter(service.strip()))
    if filter_expr.strip():
        lines.append("| filter " + filter_expr.strip())
    lines.append("| sort @timestamp asc")
    lines.append(f"| limit {limit}")
    return "\n".join(lines) + "\n"
