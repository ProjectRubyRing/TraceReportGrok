#!/usr/bin/env bash
# End-to-end test of export-span-traces.sh with a fake aws command.
# Runs on bash (RHEL 9.8 / WSL). It does not call AWS.

set -euo pipefail
IFS=$'\n\t'

ROOT=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)
export PATH="${ROOT}/tests/bin:${PATH}"
chmod +x "${ROOT}/tests/bin/aws" "${ROOT}/export-span-traces.sh" || true

fail() {
  echo "FAIL: $*" >&2
  exit 1
}

assert_eq() {
  local actual=$1
  local expected=$2
  local label=$3
  if [[ "$actual" != "$expected" ]]; then
    fail "${label}: expected [${expected}] got [${actual}]"
  fi
}

new_state() {
  local dir
  dir=$(mktemp -d)
  python3 - "$dir" <<'PY'
import json
import sys
from pathlib import Path

dest = Path(sys.argv[1])

def span(trace_id: str, span_id: str, service: str):
    start = 1761436800000000000
    return {
        "traceId": trace_id,
        "spanId": span_id,
        "name": "GET /" + service,
        "kind": "SERVER",
        "startTimeUnixNano": start,
        "endTimeUnixNano": start + 15000000,
        "durationNano": 15000000,
        "status": {"code": "OK"},
        "attributes": {
            "aws.local.service": service,
            "http.request.method": "GET",
            "http.response.status_code": 200,
            "url.path": "/" + service,
        },
        "resource": {"attributes": {"service.name": service, "cloud.region": "ap-northeast-1"}},
    }

def page(document: dict, matched: int):
    return {
        "status": "Complete",
        "statistics": {"recordsMatched": matched, "recordsScanned": 10, "bytesScanned": 2048},
        "results": [[
            {"field": "@timestamp", "value": "2026-10-26 00:00:00.000"},
            {"field": "@logStream", "value": "spans"},
            {"field": "@ptr", "value": "pointer-secret"},
            {"field": "traceId", "value": document["traceId"]},
            {"field": "@message", "value": json.dumps(document)},
        ]],
    }

alpha = span("a" * 32, "1" * 16, "alpha")
beta = span("b" * 32, "2" * 16, "beta")
(dest / "page1.json").write_text(json.dumps(page(alpha, 2)), encoding="utf-8")
(dest / "page2.json").write_text(json.dumps(page(beta, 2)), encoding="utf-8")
PY
  echo "$dir"
}

run_export() {
  local state=$1
  shift
  FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" "$@"
}

echo "1. pagination exports both pages and omits @ptr"
state=$(new_state)
out="${state}/out"
run_export "$state" \
  --start 1700000000 \
  --end 1700003600 \
  --region ap-northeast-1 \
  --output-dir "$out" \
  --limit 100 \
  --timezone UTC >/dev/null
PYTHONPATH="${ROOT}/lib${PYTHONPATH:+:$PYTHONPATH}" python3 - "$out" "$state" "${ROOT}/lib" <<'PY'
import csv, json, sys
from pathlib import Path
sys.path.insert(0, sys.argv[3])
from spanreport.xlsxutil import read_sheet_matrix
out = Path(sys.argv[1])
state = Path(sys.argv[2])
invocations = [json.loads(line) for line in (state / "invocations.jsonl").read_text(encoding="utf-8").splitlines()]
starts = [item for item in invocations if item["operation"] == "start-query"]
polls = [item for item in invocations if item["operation"] == "get-query-results"]
assert len(starts) == 1, starts
assert starts[0]["args"]["start-time"] == "1700000000"
assert starts[0]["args"]["end-time"] == "1700003600"
assert starts[0]["args"]["log-group-name"] == "aws/spans"
assert len(polls) == 3, polls
assert polls[2]["args"].get("next-token") == "tok-2"
with (out / "spans-raw.csv").open(encoding="utf-8-sig", newline="") as handle:
    rows = list(csv.DictReader(handle))
assert len(rows) == 2, rows
assert "@ptr" not in rows[0]
assert {row["traceId"] for row in rows} == {"a" * 32, "b" * 32}
matrix = read_sheet_matrix(out / "xray-trace-report.xlsx", "トレース一覧")
text = "\n".join(",".join(row) for row in matrix)
assert "1-" + "a" * 8 in text, text
assert "1-" + "b" * 8 in text, text
meta = json.loads((out / "metadata.json").read_text(encoding="utf-8"))
assert meta["trace_count"] == 2
assert meta["records_matched"] == 2
assert meta["account"] == "123456789012"
PY

echo "2. dry-run writes both trace-id forms and does not call aws"
state=$(new_state)
out="${state}/dry"
bash "${ROOT}/export-span-traces.sh" \
  --last 15m \
  --service 'a"b' \
  --trace-id '1-58406520-a006649127e371903a2de979' \
  --filter 'status.code = "ERROR"' \
  --region ap-northeast-1 \
  --output-dir "$out" \
  --dry-run >/dev/null
[[ ! -f "${state}/invocations.jsonl" ]] || fail "dry-run called aws"
query=$(cat "${out}/query.cwli")
[[ "$query" == *'a\"b'* ]] || fail "service quote was not escaped: $query"
[[ "$query" == *'58406520a006649127e371903a2de979'* ]] || fail "otel trace id missing"
[[ "$query" == *'1-58406520-a006649127e371903a2de979'* ]] || fail "xray trace id missing"
[[ "$query" == *'status.code = "ERROR"'* ]] || fail "filter missing"
python3 - "$out" <<'PY'
import sys
from pathlib import Path
lines = Path(sys.argv[1], "dry-run.txt").read_text(encoding="utf-8")
values = dict(line.split("=", 1) for line in lines.splitlines())
delta = int(values["end_epoch"]) - int(values["start_epoch"])
if delta not in {899, 900, 901, 902}:
    raise SystemExit(f"unexpected window {delta}")
PY

echo "3. failed query exits non-zero"
state=$(new_state)
set +e
FAKE_AWS_MODE=fail FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --output-dir "${state}/fail" --skip-preflight >/dev/null
rc=$?
set -e
[[ "$rc" -ne 0 ]] || fail "failed query returned success"
[[ ! -f "${state}/fail/xray-trace-report.xlsx" ]] || fail "workbook written after failure"

echo "4. missing log group stops before start-query"
state=$(new_state)
set +e
FAKE_AWS_MODE=missing-group FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --output-dir "${state}/missing" >/dev/null
rc=$?
set -e
[[ "$rc" -eq 1 ]] || fail "missing group exit ${rc}"
python3 - "$state" <<'PY'
import json, sys
from pathlib import Path
lines = Path(sys.argv[1], "invocations.jsonl").read_text(encoding="utf-8").splitlines()
ops = [json.loads(line)["operation"] for line in lines]
assert "start-query" not in ops, ops
assert "describe-log-groups" in ops
PY

echo "5. permission denial warns and exits without querying"
state=$(new_state)
set +e
FAKE_AWS_MODE=describe-denied FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --switchback-script "${ROOT}/tests/fixtures/switchback.sh" \
  --output-dir "${state}/denied" >"${state}/denied.out" 2>"${state}/denied.err"
rc=$?
set -e
[[ "$rc" -ne 0 ]] || fail "access denied returned success"
[[ ! -f "${state}/denied/xray-trace-report.xlsx" ]] || fail "workbook written after access denied"
[[ ! -f "${state}/switchback.log" ]] || fail "warn mode sourced the switchback script"
grep -q 'CodeCommit' "${state}/denied.err" || fail "warning did not mention CodeCommit"
grep -q 'スイッチバック' "${state}/denied.err" || fail "warning did not mention switchback"
python3 - "$state" <<'PY'
import json, sys
from pathlib import Path
lines = Path(sys.argv[1], "invocations.jsonl").read_text(encoding="utf-8").splitlines()
ops = [json.loads(line)["operation"] for line in lines]
assert "start-query" not in ops, ops
assert ops.count("describe-log-groups") == 1
PY

echo "5b. switchback sources the script and continues"
state=$(new_state)
FAKE_AWS_MODE=describe-denied FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --on-access-denied switchback \
  --switchback-script "${ROOT}/tests/fixtures/switchback.sh" \
  --output-dir "${state}/switched" >/dev/null
[[ -f "${state}/switched/xray-trace-report.xlsx" ]] || fail "switchback did not continue"
[[ "$(cat "${state}/switchback.log")" == "sourced" ]] || fail "switchback script was not sourced"
python3 - "$state" <<'PY'
import json, sys
from pathlib import Path
lines = [json.loads(line) for line in Path(sys.argv[1], "invocations.jsonl").read_text(encoding="utf-8").splitlines()]
describes = [item for item in lines if item["operation"] == "describe-log-groups"]
starts = [item for item in lines if item["operation"] == "start-query"]
assert len(describes) == 2, describes
assert len(starts) == 1, starts
PY

echo "5c. start-query denial switchback retries that call"
state=$(new_state)
FAKE_AWS_MODE=start-denied FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --on-access-denied switchback \
  --switchback-script "${ROOT}/tests/fixtures/switchback.sh" \
  --output-dir "${state}/start-switched" >/dev/null
[[ -f "${state}/start-switched/xray-trace-report.xlsx" ]] || fail "start-query switchback did not continue"
python3 - "$state" <<'PY'
import json, sys
from pathlib import Path
lines = [json.loads(line) for line in Path(sys.argv[1], "invocations.jsonl").read_text(encoding="utf-8").splitlines()]
starts = [item for item in lines if item["operation"] == "start-query"]
assert len(starts) == 2, starts
PY

echo "5d. automatic switchback requires a readable script path"
state=$(new_state)
set +e
FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --on-access-denied switchback \
  --output-dir "${state}/noscript" >"${state}/noscript.out" 2>"${state}/noscript.err"
rc=$?
set -e
[[ "$rc" -eq 1 ]] || fail "missing switchback script exit ${rc}"
[[ ! -f "${state}/invocations.jsonl" ]] || fail "AWS was called without a switchback script"
grep -q -- '--switchback-script' "${state}/noscript.err" || fail "missing path was not explained"
set +e
FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --on-access-denied switchback \
  --switchback-script "${state}/missing.sh" \
  --output-dir "${state}/badpath" >"${state}/badpath.out" 2>"${state}/badpath.err"
rc=$?
set -e
[[ "$rc" -eq 1 ]] || fail "unreadable switchback script exit ${rc}"
grep -q '読み取れません' "${state}/badpath.err" || fail "unreadable script was not explained"

echo "6. CLI without start-query --limit is retried"
state=$(new_state)
FAKE_AWS_MODE=no-limit FAKE_AWS_DIR="$state" bash "${ROOT}/export-span-traces.sh" \
  --start 1700000000 --end 1700000300 --region ap-northeast-1 \
  --output-dir "${state}/nolimit" --limit 50 >/dev/null
python3 - "$state" <<'PY'
import json, sys
from pathlib import Path
lines = [json.loads(line) for line in Path(sys.argv[1], "invocations.jsonl").read_text(encoding="utf-8").splitlines()]
starts = [item for item in lines if item["operation"] == "start-query"]
assert len(starts) == 2, starts
assert "limit" in starts[0]["args"]
assert "limit" not in starts[1]["args"]
assert Path(sys.argv[1], "nolimit", "xray-trace-report.xlsx").is_file()
PY

echo "7. demo fills the template in JST offset"
state=$(new_state)
bash "${ROOT}/export-span-traces.sh" --demo --timezone +09:00 --output-dir "${state}/demo" >/dev/null
PYTHONPATH="${ROOT}/lib${PYTHONPATH:+:$PYTHONPATH}" python3 - "$state/demo" "${ROOT}/lib" <<'PY'
import sys
from pathlib import Path
sys.path.insert(0, sys.argv[2])
from spanreport.xlsxutil import read_defined_values, read_sheet_matrix
demo = Path(sys.argv[1])
values = read_defined_values(demo / "xray-trace-report.xlsx")
assert values["QueryLogGroup"] == "aws/spans", values["QueryLogGroup"]
assert values["TraceCount"] == "4"
assert values["QueryTimeZone"] == "+09:00"
matrix = read_sheet_matrix(demo / "xray-trace-report.xlsx", "トレース一覧")
text = "\n".join(",".join(row) for row in matrix)
assert "EXAMPLE-TRACE" not in text
assert "1-58406520-a006649127e371903a2de979" in text
assert "障害" in text
assert "スロットル" in text
assert "2026-10-26 09:00:00.000 +0900" in text
assert "tenant_id=acme" in text
timeline = "\n".join(",".join(row) for row in read_sheet_matrix(demo / "xray-trace-report.xlsx", "タイムライン"))
assert "unmarshal" in timeline
assert "DynamoDB" in timeline
PY

echo "8. usage errors"
set +e
bash "${ROOT}/export-span-traces.sh" --output-dir "${state}/bad" >/tmp/span-usage.err
rc=$?
set -e
[[ "$rc" -eq 1 ]] || fail "missing window exit ${rc}"
bash "${ROOT}/export-span-traces.sh" --help >/dev/null
query="${state}/custom.cwli"
echo 'fields @timestamp, @message | limit 5' >"$query"
set +e
bash "${ROOT}/export-span-traces.sh" --start 1700000000 --end 1700000300 --query-file "$query" --filter 'name = "x"' --dry-run --output-dir "${state}/both" >/tmp/span-both.err 2>&1
rc=$?
set -e
[[ "$rc" -eq 1 ]] || fail "query-file plus filter should fail"
grep -q -- '--filter' /tmp/span-both.err || fail "error did not mention --filter"

echo "shell tests passed"
