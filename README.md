# CloudWatch Logs Insights to X-Ray Trace Excel Report

AWS CloudWatch Logs Insights を利用して、Transaction Search により生成されるスパン専用ロググループ（`aws/spans`）に対して時間指定クエリを実行し、結果を CSV エクスポートした上で、**AWS X-Ray のトレース画面とほぼ同様の表示を行う Excel レポート（.xlsx）** へ自動流し込みを行う RHEL 9.8 向けシェルスクリプト実装です。

---

## 1. 概要とアーキテクチャ

AWS CloudWatch Transaction Search（Application Signals / OpenTelemetry）を有効化すると、分散トレースのスパンデータが `aws/spans` ロググループに構造化 JSON ログとして蓄積されます。

本ツールは、RHEL 9.8 上で動作するシェルスクリプトを起点とし、以下のパイプラインを自動実行します。

```text
[ CloudWatch Logs Insights ]
  Log Group: aws/spans
       │  (時間指定・フィルタ指定クエリ実行: aws logs start-query)
       ▼
[ Query Results Polling ]
  ステータス Complete までポーリング (aws logs get-query-results)
       │
       ▼
[ 1. Raw JSON 出力 ]
  output/query_results_YYYYMMDD_HHMMSS.json
       │
       ▼
[ 2. CSV エクスポート ]
  output/trace_spans_YYYYMMDD_HHMMSS.csv
  ・スパン階層、時間、レイテンシ(ms)、HTTPステータス等を正規化
       │
       ▼
[ 3. Excel テンプレート流し込み ]
  templates/trace_report_template.xlsx  (X-Ray画面を忠実に再現)
       │
       ▼
[ 4. 最終 Excel レポート生成 ]
  output/trace_report_YYYYMMDD_HHMMSS.xlsx
  ・Summary シート (サマリーKPIカード、トレース一覧インデックス)
  ・Trace_Waterfall シート (ツリー階層、インセルGanttバー、積み上げ横棒グラフ)
  ・Raw_Data シート (全スパンの生データ一覧)
```

---

## 2. ディレクトリ構成

```text
TraceReportAgy/
├── bin/
│   └── export_trace_report.sh     # RHEL 9.8 向けメイン実行シェルスクリプト
├── scripts/
│   ├── generate_template.py       # Excelテンプレート生成スクリプト
│   ├── inject_to_excel.py         # CSVデータ流し込み・ツリー構築・グラフ生成
│   └── csv_transformer.py        # Logs Insights JSON -> CSV 変換スクリプト
├── templates/
│   └── trace_report_template.xlsx # X-Ray 画面再現 Excel テンプレート
├── sample_data/
│   ├── spans_query_result.json    # モック用クエリ結果 JSON
│   └── spans_sample.csv           # モック用スパン CSV
├── tests/
│   ├── test_csv_transformer.py   # CSV変換単体テスト
│   ├── test_inject_to_excel.py    # Excel生成・ツリー構築単体テスト
│   ├── test_export_trace_report.sh# シェルスクリプト統合テスト
│   └── run_all_tests.sh           # 全テスト一括実行スクリプト
├── README.md                      # 本ドキュメント
└── INSTALL_RHEL9.md               # RHEL 9.8 環境構築・IAM 設定手順書
```

---

## 3. Excel レポートの画面構成（X-Ray トレース画面の再現）

生成される Excel ワークブックは、AWS Management Console の **X-Ray Trace Details** 画面を忠実に再現しています。

### ① `Summary` シート（エグゼクティブサマリー ＆ トレース一覧）
- **ヘッダーバナー**: ロググループ名、クエリ期間（UTC）、ステータスフィルタ、レポート生成時刻。
- **KPI カード**:
  - `TOTAL TRACES`（総トレース数）
  - `TOTAL SPANS`（総スパン数）
  - `OK (2xx)`（正常トランザクション数 - グリーンバッジ）
  - `ERROR / FAULT`（エラー・障害トランザクション数 - レッドバッジ）
  - `MAX DURATION`（最大レイテンシ）
- **トレース一覧テーブル**:
  - トレースID、開始時刻、ルートサービス、ルート操作/URL、実行時間(ms)、スパン数、参加サービス数、ステータス、エラー数。
  - 各行から該当トレースの Waterfall シートへ直接ジャンプできるハイパーリンクを配置。

### ② `Trace_Waterfall` シート（X-Ray ウォーターフォール・タイムライン）
- **トレース個別サマリーカード**: トレースID、ステータス（OK/ERROR/FAULT）、ルートサービス:操作、総所要時間、スパン数、開始タイムスタンプ。
- **階層化スパンテーブル**:
  - **Status**: OK (緑), ERROR (黄), FAULT (赤) のカラーバッジ。
  - **Hierarchy & Operation**: 親子関係（`parentSpanId`）を解析し、ツリー構造（`├── `, `└── `）で視覚的にインデント表示。
  - **Service**: サービス名（太字）。
  - **Kind**: SERVER, CLIENT, INTERNAL など。
  - **HTTP**: レスポンスコード（200, 404, 500 など）。
  - **Start Offset**: トレース開始からの相対オフセット時間（`+0.00 ms`, `+12.50 ms` 等）。
  - **Duration**: スパン単体の所要時間（ms）。
  - **% of Trace**: トレース全体に占める割合（%）。
- **インセル タイムライン Gantt バー (20 分割グリッド)**:
  - 列 J 〜 AC に 0% 〜 100% の時間軸を配置。
  - スパンの実行区間に応じてセルがステータスカラー（OK=緑、Client/DB=青、Error=黄/赤）で自動塗り分けされ、**X-Ray のタイムラインバーと全く同じ感覚**で遅延箇所を直感把握可能。
- **ネイティブ Excel 積み上げ横棒グラフ (Gantt Chart)**:
  - Excel のネイティブチャート機能による Gantt グラフをテーブル下部に自動埋め込み。
  - オフセット系列（透明）＋所要時間系列（着色）で視覚的なウォーターフォールを描画。

### ③ `Raw_Data` シート（全スパン生データ）
- `timestamp`, `trace_id`, `span_id`, `parent_span_id`, `service_name`, `operation_name`, `duration_ms`, `duration_nano`, `http_status`, `http_method`, `http_url`, `status_code`, `status_message`, `kind` の全レコードを格納。
- オートフィルタと交互行カラーリングが適用済み。

---

## 4. 使い方

### 基本実行コマンド
```bash
# 過去 1 時間のスパンを取得して Excel レポートを作成
./bin/export_trace_report.sh --start-time 1h

# 時間範囲（ISO 8601 または 日時形式）を指定して実行
./bin/export_trace_report.sh \
  --start-time "2026-10-26 00:00:00" \
  --end-time "2026-10-26 01:00:00"

# 特定の Trace ID に絞り込んで詳細レポートを生成
./bin/export_trace_report.sh \
  --trace-id "1-671c6d1f-8181283921839182391283" \
  --start-time 24h

# サービス名フィルタ、プロファイル、リージョンを指定
./bin/export_trace_report.sh \
  --start-time 2h \
  --filter "attributes.aws.local.service = 'order-frontend'" \
  --profile prod \
  --region ap-northeast-1
```

### オフライン・モックテスト実行（AWS 認証情報不要）
```bash
./bin/export_trace_report.sh --mock
```
同梱されているサンプルデータ（17スパン、正常系＋504タイムアウト障害＋404エラー）を用いて、全パイプラインと Excel 生成を即座に検証できます。

### ドライラン実行（クエリ・コマンドの事前確認）
```bash
./bin/export_trace_report.sh --dry-run -s 1h -t "trace-id-example"
```
AWS API を呼び出さずに、構築された CloudWatch Logs Insights クエリ構文と AWS CLI コマンドプレビューを表示します。

---

## 5. CLI オプション一覧

| オプション | 短縮 | デフォルト値 | 説明 |
|---|---|---|---|
| `--start-time <TIME>` | `-s` | `1h` | クエリ開始時刻（`1h`, `30m`, `24h`, `2d`, ISO8601, Epoch秒） |
| `--end-time <TIME>` | `-e` | `now` | クエリ終了時刻（同上） |
| `--log-group <NAME>` | `-g` | `aws/spans` | 対象 CloudWatch ロググループ名 |
| `--trace-id <ID>` | `-t` | なし | 特定のトレースIDで絞り込み |
| `--filter <COND>` | `-f` | なし | 追加の Logs Insights フィルタ式 |
| `--query <QUERY>` | `-q` | なし | 独自のカスタムクエリ文字列（デフォルトクエリを完全上書き） |
| `--output-dir <DIR>` | `-o` | `./output` | CSV および Excel レポートの出力先ディレクトリ |
| `--profile <NAME>` | `-p` | なし | 使用する AWS CLI プロファイル名 |
| `--region <REGION>` | `-r` | なし | AWS リージョン（例: `ap-northeast-1`） |
| `--template <PATH>` | - | `templates/trace_report_template.xlsx` | 使用する Excel テンプレートのパス |
| `--limit <NUM>` | - | `10000` | 取得する最大ログレコード件数 |
| `--mock` | - | - | モックデータを使用してオフライン実行 |
| `--dry-run` | - | - | 実行計画・クエリを表示して終了 |
| `--help` | `-h` | - | ヘルプメッセージを表示 |

---

## 6. テストの実行

本プロジェクトには包括的なテストスイートが付属しています。

```bash
# 全テスト（Python単体テスト + Bash結合テスト）を一括実行
bash tests/run_all_tests.sh
```

実行されるテスト項目:
1. `tests/test_csv_transformer.py`: Logs Insights JSON から CSV への正規化変換ロジックの検証。
2. `tests/test_inject_to_excel.py`:
   - トレースツリー構築（親子関係、インデント、オフセット計算）
   - トレース全体のステータス自動判定（OK / ERROR / FAULT）
   - Excel テンプレート流し込み、シート構成、セル書式、グラフ生成
   - 0 件クエリ時の空データ耐性
3. `tests/test_export_trace_report.sh`:
   - Bash 構文検証（`bash -n`）
   - ヘルプおよびドライラン表示
   - 不正日付・開始終了逆転のバリデーション
   - モック実行と成果物（JSON, CSV, XLSX）の整合性確認
