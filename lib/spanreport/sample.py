"""Fixed spans used by --demo and the tests.

The shape matches a Transaction Search record in aws/spans: an OpenTelemetry
JSON document with dotted attribute keys, nanosecond timestamps, and the
trace id stored as 32 hex digits.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any


BASE = datetime(2026, 10, 26, tzinfo=timezone.utc)
BASE_NS = int(BASE.timestamp()) * 1_000_000_000

TRACE_OK = "58406520a006649127e371903a2de979"
TRACE_FAULT = "58406521bbbbbbbbbbbbbbbbbbbbbbbb"
TRACE_THROTTLE = "58406522cccccccccccccccccccccccc"
TRACE_ERROR = "58406523dddddddddddddddddddddddd"


def _span(
    *,
    trace_id: str,
    span_id: str,
    name: str,
    kind: str,
    start_ns: int,
    duration_ns: int,
    service: str,
    parent: str = "",
    status: str = "OK",
    attributes: dict[str, Any] | None = None,
    resource: dict[str, Any] | None = None,
    events: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    document: dict[str, Any] = {
        "traceId": trace_id,
        "spanId": span_id,
        "name": name,
        "kind": kind,
        "startTimeUnixNano": start_ns,
        "endTimeUnixNano": start_ns + duration_ns,
        "durationNano": duration_ns,
        "status": {"code": status},
        "attributes": {
            "aws.local.service": service,
            **(attributes or {}),
        },
        "resource": {
            "attributes": {
                "service.name": service,
                "cloud.region": "ap-northeast-1",
                "cloud.availability_zone": "ap-northeast-1a",
                "deployment.environment.name": "prod",
                **(resource or {}),
            }
        },
    }
    if parent:
        document["parentSpanId"] = parent
    if events:
        document["events"] = events
    return document


def sample_spans() -> list[dict[str, Any]]:
    root = BASE_NS
    return [
        _span(
            trace_id=TRACE_OK,
            span_id="53995c3f42cd8ad8",
            name="GET /orders",
            kind="SPAN_KIND_SERVER",
            start_ns=root,
            duration_ns=120_000_000,
            service="order-api",
            status="STATUS_CODE_OK",
            attributes={
                "aws.local.operation": "GET /orders",
                "http.request.method": "GET",
                "http.response.status_code": 200,
                "url.path": "/orders",
                "http.route": "/orders",
                "client.address": "203.0.113.10",
                "enduser.id": "user-42",
                "tenant_id": "acme",
                "aws.xray.annotations": ["tenant_id"],
            },
            resource={"cloud.platform": "aws_ecs"},
        ),
        _span(
            trace_id=TRACE_OK,
            span_id="1111111111111111",
            parent="53995c3f42cd8ad8",
            name="DynamoDB GetItem",
            kind="CLIENT",
            start_ns=root + 20_000_000,
            duration_ns=50_000_000,
            service="order-api",
            attributes={
                "aws.remote.service": "DynamoDB",
                "aws.remote.operation": "GetItem",
                "db.system": "dynamodb",
                "db.statement": "GetItem orders",
                "rpc.system": "aws-api",
            },
        ),
        _span(
            trace_id=TRACE_OK,
            span_id="2222222222222222",
            parent="1111111111111111",
            name="unmarshal",
            kind="INTERNAL",
            start_ns=root + 30_000_000,
            duration_ns=10_000_000,
            service="order-api",
        ),
        _span(
            trace_id=TRACE_FAULT,
            span_id="aaaaaaaaaaaaaaaa",
            name="POST /pay",
            kind="SERVER",
            start_ns=root + 5_000_000_000,
            duration_ns=80_000_000,
            service="payment-api",
            attributes={
                "http.request.method": "POST",
                "http.response.status_code": 200,
                "url.path": "/pay",
                "client.address": "203.0.113.20",
                "enduser.id": "user-7",
            },
        ),
        _span(
            trace_id=TRACE_FAULT,
            span_id="bbbbbbbbbbbbbbbb",
            parent="aaaaaaaaaaaaaaaa",
            name="POST /charge",
            kind="CLIENT",
            start_ns=root + 5_010_000_000,
            duration_ns=40_000_000,
            service="payment-api",
            status="ERROR",
            attributes={
                "http.request.method": "POST",
                "http.response.status_code": 500,
                "url.full": "https://charge.internal/charge",
                "aws.remote.service": "charge-api",
                "aws.remote.operation": "Charge",
                "aws.xray.fault": True,
                "error.type": "HTTPError",
                "exception.message": "upstream failed",
            },
        ),
        _span(
            trace_id=TRACE_THROTTLE,
            span_id="cccccccccccccccc",
            name="PUT /stock",
            kind="SERVER",
            start_ns=root + 9_000_000_000,
            duration_ns=40_000_000,
            service="inventory-api",
            attributes={
                "http.request.method": "PUT",
                "http.response.status_code": 429,
                "url.path": "/stock",
                "aws.xray.throttle": True,
                "error.type": "ThrottlingException",
            },
        ),
        _span(
            trace_id=TRACE_ERROR,
            span_id="dddddddddddddddd",
            name="GET /items/1",
            kind="SERVER",
            start_ns=root + 12_000_000_000,
            duration_ns=30_000_000,
            service="catalog-api",
            status="ERROR",
            attributes={
                "http.request.method": "GET",
                "http.response.status_code": 404,
                "http.route": "/items/{id}",
                "url.path": "/items/1",
            },
            events=[
                {
                    "name": "exception",
                    "attributes": {
                        "exception.type": "NotFoundError",
                        "exception.message": "item 1 was not found",
                        "exception.stacktrace": "NotFoundError: item 1 was not found\n  at Catalog.get",
                    },
                }
            ],
        ),
    ]


def sample_page() -> dict[str, Any]:
    results = []
    for span in sample_spans():
        start_ns = int(span["startTimeUnixNano"])
        seconds, remainder = divmod(start_ns, 1_000_000_000)
        stamp = datetime.fromtimestamp(seconds, timezone.utc).strftime("%Y-%m-%d %H:%M:%S.")
        stamp += f"{remainder // 1_000_000:03d}"
        results.append(
            [
                {"field": "@timestamp", "value": stamp},
                {"field": "@ingestionTime", "value": stamp},
                {"field": "@logStream", "value": "spans-demo"},
                {"field": "@ptr", "value": "demo-pointer-should-not-be-exported"},
                {"field": "traceId", "value": span["traceId"]},
                {"field": "spanId", "value": span["spanId"]},
                {"field": "@message", "value": json.dumps(span, ensure_ascii=False)},
            ]
        )
    results.append(
        [
            {"field": "@timestamp", "value": "2026-10-26 00:00:20.000"},
            {"field": "@message", "value": json.dumps({"body": "span event only", "timeUnixNano": 1})},
            {"field": "@ptr", "value": "event-pointer"},
        ]
    )
    return {
        "status": "Complete",
        "statistics": {
            "recordsMatched": float(len(results)),
            "recordsScanned": 40.0,
            "bytesScanned": 8192.0,
        },
        "results": results,
    }
