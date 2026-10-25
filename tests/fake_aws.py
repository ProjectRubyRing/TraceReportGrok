#!/usr/bin/env python3
"""Stand-in for the AWS CLI. Used only by tests/test_export_script.sh."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path


def state_dir() -> Path:
    raw = os.environ.get("FAKE_AWS_DIR")
    if not raw:
        sys.stderr.write("FAKE_AWS_DIR is not set\n")
        raise SystemExit(2)
    path = Path(raw)
    path.mkdir(parents=True, exist_ok=True)
    return path


def record(entry: dict) -> None:
    path = state_dir() / "invocations.jsonl"
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(entry, ensure_ascii=False) + "\n")


def split_global(argv: list[str]) -> tuple[dict[str, str], list[str]]:
    options: dict[str, str] = {}
    index = 0
    services = {"logs", "sts", "configure", "help"}
    flags = {"--no-cli-pager", "--no-paginate", "--debug"}
    valued = {"--region", "--profile", "--output", "--endpoint-url"}
    while index < len(argv):
        token = argv[index]
        if token in services:
            break
        if token in flags:
            index += 1
            continue
        if token in valued:
            options[token[2:]] = argv[index + 1]
            index += 2
            continue
        sys.stderr.write(f"unrecognized global option: {token}\n")
        raise SystemExit(2)
    return options, argv[index:]


def parse_operation(argv: list[str]) -> dict[str, str]:
    parsed: dict[str, str] = {}
    index = 0
    while index < len(argv):
        token = argv[index]
        if token.startswith("--"):
            key = token[2:]
            if index + 1 < len(argv) and not argv[index + 1].startswith("--"):
                parsed[key] = argv[index + 1]
                index += 2
            else:
                parsed[key] = "true"
                index += 1
            continue
        parsed.setdefault("positional", "")
        parsed["positional"] = (parsed["positional"] + " " + token).strip()
        index += 1
    return parsed


def read_query(value: str) -> str:
    prefix = "file://"
    if value.startswith(prefix):
        return Path(value[len(prefix):]).read_text(encoding="utf-8")
    return value


def main(argv: list[str]) -> int:
    mode = os.environ.get("FAKE_AWS_MODE", "ok")
    directory = state_dir()
    options, rest = split_global(argv)
    if not rest:
        sys.stderr.write("missing service\n")
        return 2
    service = rest[0]
    operation_argv = rest[1:]
    operation = operation_argv[0] if operation_argv else ""
    op_args = parse_operation(operation_argv[1:])
    record(
        {
            "service": service,
            "operation": operation,
            "global": options,
            "args": op_args,
            "mode": mode,
        }
    )
    if service == "configure" and operation == "get":
        sys.stdout.write(os.environ.get("FAKE_AWS_REGION", "ap-northeast-1") + "\n")
        return 0
    if service == "sts" and operation == "get-caller-identity":
        sys.stdout.write(json.dumps({"Account": "123456789012", "UserId": "TEST", "Arn": "arn:aws:iam::123456789012:user/test"}) + "\n")
        return 0
    if service == "logs" and operation == "describe-log-groups":
        if mode == "describe-denied":
            sys.stderr.write("AccessDeniedException: describe denied\n")
            return 254
        name = "aws/spans"
        if mode == "missing-group":
            payload = {"logGroups": []}
        else:
            payload = {"logGroups": [{"logGroupName": name}]}
        sys.stdout.write(json.dumps(payload) + "\n")
        return 0
    if service == "logs" and operation == "start-query":
        if mode == "start-denied":
            sys.stderr.write("AccessDeniedException: not authorized to perform logs:StartQuery\n")
            return 254
        if "limit" in op_args and mode == "no-limit" and not (directory / "limit-rejected").exists():
            (directory / "limit-rejected").write_text("1", encoding="utf-8")
            sys.stderr.write("Unknown options: --limit\n")
            return 252
        query = read_query(op_args.get("query-string", ""))
        (directory / "last-query.cwli").write_text(query, encoding="utf-8")
        (directory / "start-query.json").write_text(json.dumps(op_args), encoding="utf-8")
        sys.stdout.write(json.dumps({"queryId": "qid-test"}) + "\n")
        return 0
    if service == "logs" and operation == "stop-query":
        sys.stdout.write("{}\n")
        return 0
    if service == "logs" and operation == "get-query-results":
        counter_path = directory / "polls"
        count = int(counter_path.read_text(encoding="utf-8")) if counter_path.exists() else 0
        count += 1
        counter_path.write_text(str(count), encoding="utf-8")
        if mode == "fail":
            sys.stdout.write(json.dumps({"status": "Failed", "results": [], "statistics": {}}) + "\n")
            return 0
        if count == 1:
            sys.stdout.write(json.dumps({
                "status": "Running",
                "results": [],
                "statistics": {"recordsMatched": 0, "recordsScanned": 0, "bytesScanned": 0},
            }) + "\n")
            return 0
        page_name = "page1.json" if count == 2 else "page2.json"
        payload = json.loads((directory / page_name).read_text(encoding="utf-8"))
        if count == 2:
            payload["nextToken"] = "tok-2"
        else:
            payload.pop("nextToken", None)
        sys.stdout.write(json.dumps(payload) + "\n")
        return 0
    sys.stderr.write(f"unhandled aws command: {service} {operation}\n")
    return 2


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except BrokenPipeError:
        raise SystemExit(0)
