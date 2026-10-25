# aws/spans トレースレポート

Transaction Search が書き込むスパンロググループ `aws/spans` を、CloudWatch Logs Insights で時間指定クエリします。結果は CSV になり、X-Ray コンソールのトレース一覧とトレース詳細タイムラインに近い Excel テンプレートへ流し込まれます。

RHEL 9.8 を対象にしています。追加の Python パッケージは使いません。

## 前提

- bash
- `python3` 3.9 以上（RHEL 9.8 の標準）
- AWS CLI v2
- ロググループ名は `aws/spans` です。先頭に `/` は付きません

IAM の例:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Action": [
        "logs:StartQuery",
        "logs:GetQueryResults",
        "logs:StopQuery",
        "logs:DescribeLogGroups"
      ],
      "Resource": "arn:aws:logs:*:*:log-group:aws/spans:*"
    },
    {
      "Effect": "Allow",
      "Action": "sts:GetCallerIdentity",
      "Resource": "*"
    }
  ]
}
```

Logs Insights はスキャンした非圧縮データ量で課金されます。時間範囲を必要な幅にしてください。

この処理に CodeCommit の操作は不要です。現在の認証情報で CloudWatch Logs を操作できない場合、既定ではスイッチバックを促して終了します。別チームのシェルを `source` して続行するには、次のように指定します。

```bash
./export-span-traces.sh --last 15m --region ap-northeast-1 \
  --on-access-denied switchback \
  --switchback-script /opt/team/switchback.sh
```

パスは環境変数 `SPAN_REPORT_SWITCHBACK_SCRIPT` でも指定できます。動作の切り替えは `SPAN_REPORT_ON_ACCESS_DENIED`（`warn` または `switchback`）です。

## 実行例

```bash
chmod +x export-span-traces.sh
./export-span-traces.sh --last 15m --region ap-northeast-1 --timezone Asia/Tokyo
./export-span-traces.sh \
  --start 2026-10-26T00:00:00Z \
  --end 2026-10-26T01:00:00Z \
  --trace-id 1-58406520-a006649127e371903a2de979 \
  --region ap-northeast-1
./export-span-traces.sh --last 1h --service order-api --filter 'status.code = "ERROR"'
./export-span-traces.sh --demo --timezone +09:00
./export-span-traces.sh --self-test
```

時刻は UNIX 秒、ISO-8601、`now`、`-15m`、`-1h`、`-1d` を受け付けます。タイムゾーンのない ISO 値は UTC です。Excel の表示タイムゾーンは `--timezone`（`Asia/Tokyo` または `+09:00`）です。

`--query-file` を渡すと、その Logs Insights クエリをそのまま実行します。`--filter` は既定クエリ専用です。`--service` と `--trace-id` は、カスタムクエリの結果に対してもレポート側で絞り込まれます。

## 出力

| ファイル | 内容 |
| --- | --- |
| `query.cwli` | 実行したクエリ |
| `spans-raw.csv` | Logs Insights の生結果。Excel 向けに UTF-8 BOM |
| `traces.csv` | トレース一覧 |
| `timeline.csv` | セグメントツリー |
| `span-details.csv` | セグメント詳細 |
| `services.csv` | サービス集計 |
| `xray-trace-report.xlsx` | テンプレートへ流し込んだレポート |
| `metadata.json` | 時間範囲、件数、スキャン量 |

Excel のシート:

- **トレース一覧**: トレース ID、応答時間、メソッド、URL、レスポンス、ルートステータス、トレースステータス、注釈
- **タイムライン**: 親子の深さと、トレース全体に対する位置を表す `■` / `□` のバー
- **スパン詳細**: HTTP、リモート呼び出し、注釈、メタデータ、例外、属性 JSON
- **サービス集計**: 件数と p50 / p95 / p99
- **クエリ結果**: 生の Logs Insights 行。長いセルは Excel の上限で切り詰め、完全な値は CSV に残します

ルートステータスはルートスパンの区分で、X-Ray 一覧のアイコンに相当します。トレースステータスはツリー内の最悪値です（障害 > スロットル > エラー > OK）。応答時間はルートスパン、トレース時間は最古の開始から最新の終了までです。トレース ID はコンソールと同じ `1-8桁-24桁` で、`aws/spans` 上の 32 桁 ID は raw 列に残します。

## テンプレート

`templates/xray-trace-report.xlsx` が投入先です。フォントは Meiryo UI、配色はグレーの濃淡、セルには細い罫線と縦位置を設定しています。ステータスは色相ではなく塗りつぶしの濃さで分けます。各表の 3 行目は非表示のキー、概要シートの値は定義名（`QueryLogGroup` など）です。表示見出し、列幅、色、条件付き書式は変更できます。列を並べ替えるときは非表示キーと条件付き書式の範囲も一緒に動かしてください。5 行目の見本は投入時に置き換わります。

再生成:

```bash
./export-span-traces.sh --init-template templates/xray-trace-report.xlsx
```

## 制限

- 1 回の取得は最大 100000 件です。API は 10000 件ずつページングします。上限に達した場合は概要シートにその旨を書きます。
- `@ptr` は既定では捨てます。残す場合は `--include-ptr` です。
- 古い AWS CLI が `start-query --limit` を拒否した場合は、クエリ内の `limit` だけで再実行します。
- タイムアウト時と中断時は `stop-query` します。
