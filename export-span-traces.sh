#!/usr/bin/env bash
# Query the Transaction Search span log group (aws/spans) with CloudWatch Logs
# Insights and pour the CSV into an X-Ray-style Excel template.
#
# Target host: RHEL 9.8
#   bash (system), Python 3.9+ standard library, AWS CLI v2
# Python builds the query, parses span JSON, and fills the workbook.
# This script owns the AWS CLI calls so the scanned time range stays visible.

set -euo pipefail
IFS=$'\n\t'

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
export PYTHONPATH="${SCRIPT_DIR}/lib${PYTHONPATH:+:$PYTHONPATH}"

case "${LANG:-C}" in
  C|POSIX) export LANG=C.UTF-8 ;;
esac
export PYTHONUTF8=1
export AWS_PAGER=""
umask 077

PYTHON=""
query_id=""
profile=""
region=""
log_group="aws/spans"
log_group_identifier=""
start_raw=""
end_raw=""
last=""
service=""
trace_id=""
filter_expr=""
query_file_user=""
limit=10000
timezone_name="UTC"
outdir=""
template=""
poll_interval=2
timeout_sec=600
include_message=1
include_ptr=0
skip_preflight=0
dry_run=0
demo=0
init_only=0
template_out=""
self_test=0
use_limit_flag=1
# warn: 権限不足ならスイッチバックを促して終了する。
# switchback: 指定シェルを source し、終了せずに AWS 操作を再試行する。
on_access_denied=${SPAN_REPORT_ON_ACCESS_DENIED:-warn}
switchback_script=${SPAN_REPORT_SWITCHBACK_SCRIPT:-}
switchback_done=0

usage() {
  cat <<'EOF'
使い方:
  export-span-traces.sh (--start TIME --end TIME | --last DURATION) [オプション]
  export-span-traces.sh --demo [オプション]
  export-span-traces.sh --init-template [出力パス]
  export-span-traces.sh --self-test

CloudWatch Logs Insights で Transaction Search のスパンロググループ aws/spans を
時間指定クエリし、結果を CSV にしたうえで X-Ray トレース画面に近い Excel
テンプレートへ流し込みます。

TIME:
  UNIX 秒、13 桁ミリ秒、ISO-8601（2026-10-26T00:00:00Z または +09:00）、
  now、-15m、-1h、-2d、-30s。タイムゾーンなしの ISO 値は UTC です。

DURATION:
  15m、1h、2d、30s（末尾の単位は s/m/h/d）。終了は現在時刻です。

オプション:
  --log-group NAME              既定: aws/spans（先頭の / は付きません）
  --log-group-identifier ARN    クロスアカウントのロググループ ARN
  --region REGION               未指定時は AWS_REGION、なければ AWS CLI の既定
  --profile NAME
  --service NAME                aws.local.service / service.name / スパン名
  --trace-id ID                 X-Ray 形式 1-8桁-24桁、または 32 桁の traceId
  --filter EXPR                 既定クエリへ追加する CWLI の filter 式
  --query-file PATH             クエリ全文。--filter とは併用できません
  --limit N                     既定 10000、最大 100000
  --timezone TZ                 Excel の表示時刻。既定 UTC。+09:00 も可
  --output-dir DIR              既定: ./out/span-report-YYYYmmddTHHMMSSZ
  --template PATH               既定: templates/xray-trace-report.xlsx
  --poll-interval SEC           既定 2
  --timeout SEC                 既定 600。超過時は stop-query します
  --no-message                  既定クエリから @message を外す
  --include-ptr                 生 CSV に @ptr を含める
  --skip-preflight              describe-log-groups による存在確認を省略
  --on-access-denied MODE       warn（既定）または switchback
                                warn: 現在の権限で AWS を操作できないとき、
                                スイッチバックを促すメッセージを出して終了する
                                switchback: 終了せず、--switchback-script を
                                source してから同じ AWS 操作を再試行する
  --switchback-script PATH      source するスイッチバック用シェル。
                                環境変数 SPAN_REPORT_SWITCHBACK_SCRIPT でも指定できる
  --dry-run                     AWS を呼ばず、クエリと時間範囲だけを出力
  --demo                        サンプルスパンでテンプレート投入を確認
  --init-template [PATH]        Excel テンプレートを書き出す
  --self-test                   単体テストとシェルテストを実行
  -h, --help

必要な IAM:
  logs:StartQuery, logs:GetQueryResults, logs:StopQuery, logs:DescribeLogGroups
  （リソースは log-group:aws/spans）
  sts:GetCallerIdentity（アカウント表示のみ。失敗しても続行します）

この処理に CodeCommit の操作は不要です。CloudWatch Logs を拒否された場合は、
別チームのスイッチバック用シェルを source して権限を戻します。既定では警告して
終了します。自動で切り替えるときは --on-access-denied switchback と
--switchback-script を指定してください。

課金:
  Logs Insights はスキャンした非圧縮データ量で課金されます。時間範囲を絞ってください。

RHEL 9.8 の前提:
  python3（3.9 以上、追加の pip パッケージは不要）、aws（AWS CLI v2）、bash、coreutils。
EOF
}

die() {
  echo "export-span-traces: $*" >&2
  exit 1
}

require_value() {
  local flag=$1
  local value=${2-}
  if [[ -z "$value" || "$value" == --* ]]; then
    die "$flag には値が必要です"
  fi
}

spanreport() {
  "$PYTHON" -m spanreport "$@"
}

aws_raw() {
  local args=()
  if [[ -n "$profile" ]]; then
    args+=(--profile "$profile")
  fi
  if [[ -n "$region" ]]; then
    args+=(--region "$region")
  fi
  aws --no-cli-pager --no-paginate "${args[@]}" "$@"
}

access_denied_text() {
  local file=$1
  [[ -f "$file" ]] || return 1
  grep -Eiq 'AccessDenied|UnauthorizedOperation|not authorized to perform|is not authorized|ExplicitDeny|explicit deny' "$file"
}

require_switchback_script() {
  if [[ -z "$switchback_script" ]]; then
    die "スイッチバックを自動実行するには --switchback-script または SPAN_REPORT_SWITCHBACK_SCRIPT でシェルのパスを指定してください"
  fi
  if [[ ! -f "$switchback_script" || ! -r "$switchback_script" ]]; then
    die "スイッチバック用シェルを読み取れません: ${switchback_script}"
  fi
}

warn_switchback_and_exit() {
  local err_file=$1
  echo "export-span-traces: 現在の権限では AWS を操作できません。" >&2
  echo "この処理に CodeCommit の操作は不要です。CloudWatch Logs（logs:DescribeLogGroups, logs:StartQuery, logs:GetQueryResults, logs:StopQuery）の権限が必要です。" >&2
  if [[ "$switchback_done" -eq 1 ]]; then
    echo "スイッチバック後も操作できませんでした。処理を終了します。" >&2
  else
    echo "スイッチバックしてから、同じコマンドを再実行してください。" >&2
  fi
  if [[ -n "$switchback_script" ]]; then
    echo "スイッチバック: source ${switchback_script}" >&2
  else
    echo "スイッチバック用シェルのパスは --switchback-script または SPAN_REPORT_SWITCHBACK_SCRIPT で指定できます。" >&2
    echo "終了せずに自動で切り替える場合は --on-access-denied switchback を指定してください。" >&2
  fi
  if [[ -s "$err_file" ]]; then
    echo "AWS の応答:" >&2
    sed -n '1,20p' "$err_file" >&2
  fi
  exit 1
}

# Returns 0 when the caller should retry the AWS call. Warn mode does not return.
perform_switchback() {
  local err_file=$1
  if [[ "$on_access_denied" != "switchback" || "$switchback_done" -eq 1 ]]; then
    warn_switchback_and_exit "$err_file"
  fi
  require_switchback_script
  switchback_done=1
  echo "現在の権限では AWS を操作できません。この処理に CodeCommit の操作は不要なため、スイッチバックします。" >&2
  if [[ -s "$err_file" ]]; then
    echo "AWS の応答: $(sed -n '1p' "$err_file")" >&2
  fi
  echo "source ${switchback_script}" >&2
  local flags=$-
  set +e
  # The other team's script is sourced on purpose so credential changes stay
  # in this shell. It must be a regular readable file chosen by the operator.
  # shellcheck disable=SC1090
  source "$switchback_script"
  local rc=$?
  case "$flags" in
    *e*) set -e ;;
  esac
  if [[ "$rc" -ne 0 ]]; then
    echo "スイッチバック用シェルが失敗しました (exit ${rc}): ${switchback_script}" >&2
    exit "$rc"
  fi
  echo "スイッチバック用シェルを読み込みました。AWS 操作を再試行します。" >&2
  return 0
}

aws_cmd() {
  local out err rc
  out=$(mktemp)
  err=$(mktemp)
  # Do not toggle set -e here. A non-zero return must reach callers that
  # temporarily disabled errexit, such as the start-query --limit retry.
  if aws_raw "$@" >"$out" 2>"$err"; then
    rc=0
  else
    rc=$?
  fi
  if [[ "$rc" -ne 0 && "${1-}" == "logs" ]] && access_denied_text "$err"; then
    perform_switchback "$err"
    rm -f "$out" "$err"
    out=$(mktemp)
    err=$(mktemp)
    if aws_raw "$@" >"$out" 2>"$err"; then
      rc=0
    else
      rc=$?
    fi
    if [[ "$rc" -ne 0 ]] && access_denied_text "$err"; then
      warn_switchback_and_exit "$err"
    fi
  fi
  if [[ -s "$err" ]]; then
    cat "$err" >&2
  fi
  if [[ -n "${AWS_CMD_ERRCOPY:-}" ]]; then
    cp "$err" "$AWS_CMD_ERRCOPY"
  fi
  cat "$out"
  rm -f "$out" "$err"
  return "$rc"
}

on_exit() {
  local code=$?
  if [[ -n "${query_id}" ]]; then
    echo "未完了のクエリを停止します: ${query_id}" >&2
    aws_raw logs stop-query --query-id "$query_id" >/dev/null 2>&1 || true
    query_id=""
  fi
  exit "$code"
}

parse_args() {
  while [[ $# -gt 0 ]]; do
    case "$1" in
      -h|--help)
        usage
        exit 0
        ;;
      --start)
        require_value "$1" "${2-}"
        start_raw=$2
        shift 2
        ;;
      --end)
        require_value "$1" "${2-}"
        end_raw=$2
        shift 2
        ;;
      --last)
        require_value "$1" "${2-}"
        last=$2
        shift 2
        ;;
      --log-group)
        require_value "$1" "${2-}"
        log_group=$2
        shift 2
        ;;
      --log-group-identifier)
        require_value "$1" "${2-}"
        log_group_identifier=$2
        shift 2
        ;;
      --region)
        require_value "$1" "${2-}"
        region=$2
        shift 2
        ;;
      --profile)
        require_value "$1" "${2-}"
        profile=$2
        shift 2
        ;;
      --service)
        require_value "$1" "${2-}"
        service=$2
        shift 2
        ;;
      --trace-id)
        require_value "$1" "${2-}"
        trace_id=$2
        shift 2
        ;;
      --filter)
        require_value "$1" "${2-}"
        filter_expr=$2
        shift 2
        ;;
      --query-file)
        require_value "$1" "${2-}"
        query_file_user=$2
        shift 2
        ;;
      --limit)
        require_value "$1" "${2-}"
        limit=$2
        shift 2
        ;;
      --timezone)
        require_value "$1" "${2-}"
        timezone_name=$2
        shift 2
        ;;
      --output-dir)
        require_value "$1" "${2-}"
        outdir=$2
        shift 2
        ;;
      --template)
        require_value "$1" "${2-}"
        template=$2
        shift 2
        ;;
      --poll-interval)
        require_value "$1" "${2-}"
        poll_interval=$2
        shift 2
        ;;
      --timeout)
        require_value "$1" "${2-}"
        timeout_sec=$2
        shift 2
        ;;
      --no-message)
        include_message=0
        shift
        ;;
      --include-ptr)
        include_ptr=1
        shift
        ;;
      --skip-preflight)
        skip_preflight=1
        shift
        ;;
      --on-access-denied)
        require_value "$1" "${2-}"
        on_access_denied=$2
        shift 2
        ;;
      --switchback-script)
        require_value "$1" "${2-}"
        switchback_script=$2
        shift 2
        ;;
      --dry-run)
        dry_run=1
        shift
        ;;
      --demo)
        demo=1
        shift
        ;;
      --self-test)
        self_test=1
        shift
        ;;
      --init-template)
        init_only=1
        if [[ $# -ge 2 && "${2-}" != --* ]]; then
          template_out=$2
          shift 2
        else
          shift
        fi
        ;;
      *)
        die "不明な引数です: $1 （--help を参照）"
        ;;
    esac
  done
}

require_python() {
  if command -v python3 >/dev/null 2>&1; then
    PYTHON=python3
  elif command -v python >/dev/null 2>&1; then
    PYTHON=python
  else
    echo "python3 が必要です。RHEL 9.8 では python3 パッケージを導入してください。" >&2
    exit 127
  fi
  if ! "$PYTHON" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)'; then
    die "Python 3.9 以上が必要です（RHEL 9.8 の python3 は 3.9 です）"
  fi
}

require_aws() {
  if ! command -v aws >/dev/null 2>&1; then
    die "aws コマンドがありません。AWS CLI v2 を導入してください"
  fi
}

validate_numbers() {
  [[ "$limit" =~ ^[0-9]+$ ]] || die "--limit は正の整数です"
  [[ "$poll_interval" =~ ^[0-9]+$ ]] || die "--poll-interval は秒の整数です"
  [[ "$timeout_sec" =~ ^[0-9]+$ ]] || die "--timeout は秒の整数です"
  if [[ "$limit" -lt 1 || "$limit" -gt 100000 ]]; then
    die "--limit は 1 から 100000 です"
  fi
  if [[ "$poll_interval" -lt 1 ]]; then
    die "--poll-interval は 1 秒以上です"
  fi
  if [[ "$timeout_sec" -lt 1 ]]; then
    die "--timeout は 1 秒以上です"
  fi
  if [[ -n "$last" && ! "$last" =~ ^[0-9]+[smhd]$ ]]; then
    die "--last は 15m や 1h の形式です"
  fi
}

prepare_out() {
  if [[ -z "$outdir" ]]; then
    outdir="$PWD/out/span-report-$(date -u +%Y%m%dT%H%M%SZ)"
  fi
  mkdir -p "$outdir"
  outdir=$(CDPATH= cd -- "$outdir" && pwd)
}

ensure_template() {
  if [[ -z "$template" ]]; then
    template="${SCRIPT_DIR}/templates/xray-trace-report.xlsx"
  fi
  if [[ ! -f "$template" ]]; then
    echo "テンプレートがないため生成します: $template" >&2
    mkdir -p "$(dirname -- "$template")"
    spanreport init-template --output "$template"
  fi
}

resolve_region() {
  if [[ -n "$region" ]]; then
    return
  fi
  region="${AWS_REGION:-${AWS_DEFAULT_REGION:-}}"
  if [[ -z "$region" ]]; then
    region=$(aws_cmd configure get region 2>/dev/null || true)
  fi
}

run_self_test() {
  require_python
  echo "Python 単体テストを実行します" >&2
  (
    cd "$SCRIPT_DIR"
    "$PYTHON" -m unittest discover -s tests -p 'test_*.py' -v
  )
  bash -n "$SCRIPT_DIR/export-span-traces.sh"
  echo "シェル結合テストを実行します" >&2
  bash "$SCRIPT_DIR/tests/test_export_script.sh"
}

run_init_template() {
  require_python
  local destination=${template_out:-${SCRIPT_DIR}/templates/xray-trace-report.xlsx}
  mkdir -p "$(dirname -- "$destination")"
  spanreport init-template --output "$destination"
}

run_demo() {
  if [[ -n "$start_raw" || -n "$end_raw" || -n "$last" || -n "$query_file_user" || -n "$filter_expr" ]]; then
    die "--demo は --start / --end / --last / --query-file / --filter と併用できません"
  fi
  prepare_out
  ensure_template
  local render_args=(
    demo
    --output-dir "$outdir"
    --template "$template"
    --timezone "$timezone_name"
    --log-group "$log_group"
    --limit "$limit"
  )
  if [[ -n "$region" ]]; then
    render_args+=(--region "$region")
  fi
  if [[ "$include_ptr" -eq 1 ]]; then
    render_args+=(--include-ptr)
  fi
  spanreport "${render_args[@]}" >/dev/null
  echo "生CSV: ${outdir}/spans-raw.csv"
  echo "トレースCSV: ${outdir}/traces.csv"
  echo "タイムラインCSV: ${outdir}/timeline.csv"
  echo "スパン詳細CSV: ${outdir}/span-details.csv"
  echo "サービスCSV: ${outdir}/services.csv"
  echo "Excel: ${outdir}/xray-trace-report.xlsx"
}

build_query_file() {
  local args=(
    build-query
    "--limit=${limit}"
    "--service=${service}"
    "--trace-id=${trace_id}"
    "--filter=${filter_expr}"
  )
  if [[ "$include_message" -eq 0 ]]; then
    args+=(--no-message)
  fi
  if [[ -n "$query_file_user" ]]; then
    args+=(--query-file "$query_file_user")
  fi
  local destination=$1
  spanreport "${args[@]}" >"$destination"
}

resolve_window() {
  if [[ -n "$last" && ( -n "$start_raw" || -n "$end_raw" ) ]]; then
    die "--last は --start / --end と併用できません"
  fi
  if [[ -n "$last" ]]; then
    end_epoch=$(spanreport parse-time --value=now)
    # The leading dash must stay attached to --value. argparse treats a bare
    # -15m as another option.
    start_epoch=$(spanreport parse-time --value="-${last}")
  else
    if [[ -z "$start_raw" || -z "$end_raw" ]]; then
      die "--start と --end、または --last が必要です"
    fi
    start_epoch=$(spanreport parse-time --value="${start_raw}")
    end_epoch=$(spanreport parse-time --value="${end_raw}")
  fi
  if [[ "$end_epoch" -le "$start_epoch" ]]; then
    die "終了時刻は開始時刻より後である必要があります"
  fi
  local span=$((end_epoch - start_epoch))
  if [[ "$span" -gt 86400 ]]; then
    echo "注意: 時間範囲が 24 時間を超えています。aws/spans のスキャン量に応じて Logs Insights の料金が発生します。" >&2
  fi
}

write_dry_run() {
  local query_path=$1
  cat >"${outdir}/dry-run.txt" <<EOF
start_epoch=${start_epoch}
end_epoch=${end_epoch}
log_group=${log_group}
log_group_identifier=${log_group_identifier}
region=${region}
query_file=${query_path}
limit=${limit}
EOF
  echo "dry-run: start=${start_epoch} end=${end_epoch} query=${query_path}"
  if [[ -n "$log_group_identifier" ]]; then
    echo "aws logs start-query --log-group-identifiers ${log_group_identifier} --start-time ${start_epoch} --end-time ${end_epoch} --limit ${limit} --query-string file://${query_path}"
  else
    echo "aws logs start-query --log-group-name ${log_group} --start-time ${start_epoch} --end-time ${end_epoch} --limit ${limit} --query-string file://${query_path}"
  fi
}

preflight_log_group() {
  if [[ "$skip_preflight" -eq 1 || -n "$log_group_identifier" ]]; then
    return
  fi
  local desc rc
  desc=$(mktemp)
  set +e
  aws_cmd logs describe-log-groups --log-group-name-prefix "$log_group" --output json >"$desc"
  rc=$?
  set -e
  if [[ "$rc" -ne 0 ]]; then
    echo "ロググループの事前確認をスキップします（describe-log-groups が失敗しました）。" >&2
    rm -f "$desc"
    return
  fi
  set +e
  spanreport log-group-check --page "$desc" --name "$log_group"
  rc=$?
  set -e
  rm -f "$desc"
  if [[ "$rc" -eq 1 ]]; then
    die "ロググループ ${log_group} が見つかりません。リージョンと Transaction Search の有効化を確認してください"
  fi
}

start_query() {
  local query_path=$1
  local start_out err rc
  start_out=$(mktemp)
  err=$(mktemp)
  local limit_args=()
  if [[ "$use_limit_flag" -eq 1 ]]; then
    limit_args=(--limit "$limit")
  fi
  local group_args=()
  if [[ -n "$log_group_identifier" ]]; then
    group_args=(--log-group-identifiers "$log_group_identifier")
  else
    group_args=(--log-group-name "$log_group")
  fi
  AWS_CMD_ERRCOPY=$err
  set +e
  aws_cmd logs start-query \
    "${group_args[@]}" \
    --start-time "$start_epoch" \
    --end-time "$end_epoch" \
    --query-string "file://${query_path}" \
    "${limit_args[@]}" \
    --output json >"$start_out"
  rc=$?
  set -e
  unset AWS_CMD_ERRCOPY
  if [[ "$rc" -ne 0 ]]; then
    if [[ "$use_limit_flag" -eq 1 ]] && grep -q -- '--limit' "$err"; then
      echo "この AWS CLI は start-query --limit に未対応です。クエリ内の limit だけを使います。" >&2
      use_limit_flag=0
      rm -f "$start_out" "$err"
      start_query "$query_path"
      return
    fi
    rm -f "$start_out" "$err"
    exit "$rc"
  fi
  query_id=$(spanreport read-query-id --page "$start_out")
  cp "$start_out" "${outdir}/start-query.json"
  rm -f "$start_out" "$err"
  echo "query id: ${query_id}" >&2
}

poll_query() {
  local pages=$1
  local deadline=$(( $(date +%s) + timeout_sec ))
  local tmp stats status matched scanned bytes next_token page_index last_status
  tmp=$(mktemp)
  stats=$(mktemp)
  page_index=0
  last_status=""
  next_token=""
  while true; do
    if [[ "$(date +%s)" -gt "$deadline" ]]; then
      echo "タイムアウトしました。クエリを停止します: ${query_id}" >&2
      aws_raw logs stop-query --query-id "$query_id" >/dev/null 2>&1 || true
      rm -f "$tmp" "$stats"
      exit 124
    fi
    local token_args=()
    if [[ -n "$next_token" ]]; then
      token_args=(--next-token "$next_token")
    fi
    aws_cmd logs get-query-results \
      --query-id "$query_id" \
      "${token_args[@]}" \
      --output json >"$tmp"
    spanreport read-stats --page "$tmp" >"$stats"
    status=$(sed -n '1p' "$stats")
    matched=$(sed -n '2p' "$stats")
    scanned=$(sed -n '3p' "$stats")
    bytes=$(sed -n '4p' "$stats")
    next_token=$(sed -n '5p' "$stats")
    if [[ "$status" != "$last_status" ]]; then
      echo "クエリ状態: ${status}（一致 ${matched} / スキャン ${scanned} / ${bytes} bytes）" >&2
      last_status=$status
    fi
    case "$status" in
      Scheduled|Running|Unknown)
        next_token=""
        sleep "$poll_interval"
        ;;
      Complete)
        page_index=$((page_index + 1))
        printf -v dest '%s/page-%04d.json' "$pages" "$page_index"
        mv "$tmp" "$dest"
        tmp=$(mktemp)
        echo "結果ページ ${page_index} を保存しました" >&2
        if [[ -z "$next_token" ]]; then
          break
        fi
        ;;
      Failed|Cancelled|Timeout)
        echo "クエリが終了しました: ${status}" >&2
        cat "$tmp" >&2
        rm -f "$tmp" "$stats"
        exit 1
        ;;
      *)
        echo "未知のクエリ状態です: ${status}" >&2
        cat "$tmp" >&2
        rm -f "$tmp" "$stats"
        exit 1
        ;;
    esac
  done
  rm -f "$tmp" "$stats"
}

render_workbook() {
  local pages=$1
  local query_path=$2
  local render_args=(
    render
    "--pages=${pages}"
    "--output-dir=${outdir}"
    "--template=${template}"
    "--timezone=${timezone_name}"
    "--log-group=${log_group}"
    "--region=${region}"
    "--account=${account}"
    "--start-epoch=${start_epoch}"
    "--end-epoch=${end_epoch}"
    "--query-file=${query_path}"
    --query-status Complete
    "--service=${service}"
    "--trace-id=${trace_id}"
    "--filter=${filter_expr}"
    "--limit=${limit}"
  )
  if [[ "$include_ptr" -eq 1 ]]; then
    render_args+=(--include-ptr)
  fi
  spanreport "${render_args[@]}" >/dev/null
}

lookup_account() {
  account=""
  local identity rc
  identity=$(mktemp)
  set +e
  aws_cmd sts get-caller-identity --output json >"$identity"
  rc=$?
  set -e
  if [[ "$rc" -eq 0 ]]; then
    account=$(spanreport read-account --page "$identity")
  else
    echo "アカウント ID は取得できませんでした。レポートのアカウントは空欄になります。" >&2
  fi
  rm -f "$identity"
}

run_export() {
  validate_numbers
  resolve_window
  prepare_out
  ensure_template
  local query_path="${outdir}/query.cwli"
  build_query_file "$query_path"
  echo "注意: aws/spans への Logs Insights はスキャン量に応じて課金されます。範囲 ${start_epoch} から ${end_epoch} です。" >&2
  if [[ "$dry_run" -eq 1 ]]; then
    write_dry_run "$query_path"
    return
  fi
  require_aws
  if [[ "$on_access_denied" == "switchback" ]]; then
    require_switchback_script
  fi
  resolve_region
  trap on_exit EXIT
  trap 'exit 130' INT TERM
  preflight_log_group
  lookup_account
  start_query "$query_path"
  local pages="${outdir}/pages"
  mkdir -p "$pages"
  poll_query "$pages"
  query_id=""
  render_workbook "$pages" "$query_path"
  echo "生CSV: ${outdir}/spans-raw.csv"
  echo "トレースCSV: ${outdir}/traces.csv"
  echo "タイムラインCSV: ${outdir}/timeline.csv"
  echo "スパン詳細CSV: ${outdir}/span-details.csv"
  echo "サービスCSV: ${outdir}/services.csv"
  echo "Excel: ${outdir}/xray-trace-report.xlsx"
}

main() {
  parse_args "$@"
  case "$on_access_denied" in
    warn|switchback) ;;
    *)
      die "--on-access-denied は warn または switchback です"
      ;;
  esac
  if [[ "$self_test" -eq 1 ]]; then
    run_self_test
    return
  fi
  if [[ "$init_only" -eq 1 ]]; then
    run_init_template
    return
  fi
  require_python
  validate_numbers
  if [[ "$demo" -eq 1 ]]; then
    run_demo
    return
  fi
  run_export
}

main "$@"
