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
├── examples/
│   └── switchback_sample.sh       # スイッチバック用シェルスクリプト実装サンプル
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

## 3. Excel レポートの画面構成（Meiryo UI ＆ モノトーンベースデザイン）

生成される Excel ワークブックは、AWS Management Console の **X-Ray Trace Details** 画面をモデルとしつつ、**フォントに Meiryo UI** を全面採用し、上質で洗練された **モノトーンベース（グレースケール・チャコール・スレート）のコーポレートデザイン** でレイアウト・セル設定を最適化しています。

### デザイン・セル設定の特徴
- **フォント**: 全シート・全セルに `Meiryo UI` を適用。日本語・英数記号ともに高い視認性と端正なUI表示を実現。
- **モノトーンベース配色**:
  - メインタイトル・バナー: シックな漆黒チャコール（`#212529`）
  - テーブルヘッダー: 深みのあるダークスレート（`#343A40`）
  - タイムライン時間軸スケール: グラファイトグレー（`#495057`）
  - カードパネル: 清潔感のあるライトグレー（`#F8F9FA`）
  - ステータス表現:
    - `OK`: 上品なシルバーグレー（`#E9ECEF` / テキスト `#212529`）
    - `ERROR (4xx)`: ミディアムチャコール（`#6C757D` / テキスト `#FFFFFF`）
    - `FAULT (5xx)`: 最も注意を引く高コントラストブラック（`#212529` / テキスト `#FFFFFF`）
- **セル書式・配置の最適化**:
  - **表示形式 (Number Format)**: ミリ秒レイテンシは `#,##0.00 "ms"`、件数は `#,##0`、割合は `0.0%`、IDやタイムスタンプは `@`（文字列）。
  - **配置 (Alignment)**: 数値は右揃え、ステータスやコード・日時は中央揃え、名称やIDは左揃え、全セル垂直中央揃え（`vertical="center"`）。
  - **行高・列幅**: Meiryo UI の行送りに合わせ、窮屈感のないゆとりある行の高さ（データ行 20〜22pt、ヘッダー 24〜26pt、バナー 26〜28pt）と十分な列幅を確保。
  - **枠線・グリッド線**: `showGridLines = True` による方眼視認性の確保と、薄いグレー罫線（`#E9ECEF`, `#CED4DA`）による整然とした表組。
  - **ウィンドウ枠の固定 (Freeze Panes)**: 大量トレース・スパンのスクロール時にもヘッダーが常に画面上部に固定表示。
  - **オートフィルタ (AutoFilter)**: トレース一覧および生データ一覧に自動適用され、Excel 上での即座な並べ替え・絞り込みが可能。

### ① `Summary` シート（エグゼクティブサマリー ＆ トレース一覧）
- **ヘッダーバナー**: ロググループ名、クエリ期間（UTC）、ステータスフィルタ、レポート生成時刻。
- **モノトーン KPI カード**:
  - `TOTAL TRACES`（総トレース数）
  - `TOTAL SPANS`（総スパン数）
  - `OK (2xx)`（正常トランザクション数）
  - `ERROR / FAULT`（エラー・障害トランザクション数 - 障害検知時は反転強調）
  - `MAX DURATION`（最大レイテンシ）
- **トレース一覧テーブル**:
  - トレースID、開始時刻、ルートサービス、ルート操作/URL、実行時間(ms)、スパン数、参加サービス数、ステータス、エラー数。
  - オートフィルタとウィンドウ枠固定（13行目以降固定）を完備。
  - 各行から該当トレースの Waterfall シートへ直接ジャンプできるリンク（`View Timeline →`）を配置。

### ② `Trace_Waterfall` シート（X-Ray ウォーターフォール・タイムライン）
- **トレース個別サマリーカード**: トレースID、ステータス（OK/ERROR/FAULT）、ルートサービス:操作、総所要時間、スパン数、開始タイムスタンプ。
- **階層化スパンテーブル**:
  - **Status**: モノトーン階調のステータスバッジ（OK / ERROR / FAULT）。
  - **Hierarchy & Operation**: 親子関係（`parentSpanId`）を解析し、Meiryo UI によるツリー構造（`├── `, `└── `）で視覚的にインデント表示。
  - **Service**: サービス名（ルート・直下は太字強調）。
  - **Kind**: SERVER, CLIENT, INTERNAL など。
  - **HTTP**: レスポンスコード（200, 404, 500 など）。
  - **Start Offset**: トレース開始からの相対オフセット時間（`+0.00 ms`, `+12.50 ms` 等）。
  - **Duration**: スパン単体の所要時間（ms）。
  - **% of Trace**: トレース全体に占める割合（%）。
- **インセル タイムライン Gantt バー (20 分割グリッド)**:
  - 列 J 〜 AC に 0% 〜 100% の時間軸を配置。
  - スパンの実行区間に応じてセルがモノトーンの明淡バー（FAULT=深黒、ERROR=ダークグレー、CLIENT=ミディアムグレー、OK=シルバーグレー）で自動塗り分けされ、遅延箇所を一目で特定可能。
- **ネイティブ Excel 積み上げ横棒グラフ (Gantt Chart)**:
  - Excel のネイティブチャート機能による Gantt グラフをテーブル下部に自動埋め込み。
  - オフセット系列（透明）＋所要時間系列（モノトーン着色）で視覚的なウォーターフォールを描画。
  - ウィンドウ枠固定（11行目固定）により、多数のスパンがあってもヘッダーが常に可視。

### ③ `Raw_Data` シート（全スパン生データ）
- `timestamp`, `trace_id`, `span_id`, `parent_span_id`, `service_name`, `operation_name`, `duration_ms`, `duration_nano`, `http_status`, `http_method`, `http_url`, `status_code`, `status_message`, `kind` の全レコードを格納。
- オートフィルタ、ウィンドウ枠固定（5行目固定）、および交互行カラーリングが適用済み。

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

# スイッチバック用シェルを指定し、自動スイッチバックを有効化して実行
./bin/export_trace_report.sh \
  --start-time 1h \
  --auto-switchback \
  --switchback-script /opt/team/aws_switchback.sh
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
| `-a, --auto-switchback` | `-a` | 無効 (false) | AWS 操作権限不足時に終了せず、専用シェルを自動で `source` してスイッチバックを実行し処理を継続 |
| `--no-auto-switchback` | - | 有効 (true) | AWS 操作権限不足時に警告メッセージを出力して終了する（デフォルト動作） |
| `-S, --switchback-script <PATH>` | `-S` | `$AWS_SWITCHBACK_SCRIPT` 又は `./switchback.sh` | スイッチバック用シェルのファイル配置場所 |
| `--template <PATH>` | - | `templates/trace_report_template.xlsx` | 使用する Excel テンプレートのパス |
| `--limit <NUM>` | - | `10000` | 取得する最大ログレコード件数 |
| `--mock` | - | - | モックデータを使用してオフライン実行 |
| `--dry-run` | - | - | 実行計画・クエリを表示して終了 |
| `--help` | `-h` | - | ヘルプメッセージを表示 |

---

## 6. AWS 権限チェックとスイッチバック機能

本ツールは CodeCommit の操作は不要であり、CloudWatch Logs (`aws/spans`) へのクエリおよびデータ取得のための各種 AWS 操作権限が必要です。

CodeCommit 専用ロールや限定権限のアカウントが適用されている場合、スクリプト実行時に自動で権限チェックが行われ、以下の2つの挙動をオプションで切り替えることができます。

### ① 警告メッセージを出力して終了する機能（デフォルト動作）
AWS 操作権限（CloudWatch Logs へのアクセス）が不足している場合、警告メッセージを表示し、スイッチバック用シェルの手動実行方法を案内した上で終了コード `2` で安全に停止します。

```text
[WARN] 現在の操作権限では AWS (CloudWatch Logs) への操作が不可能です (AccessDenied)。
[WARN] 現在の Identity ARN: arn:aws:sts::123456789012:assumed-role/CodeCommit-Dev-Role/session
[WARN] ※ CodeCommit 専用の操作権限が適用されている可能性があります。

[スイッチバック警告]
本ツールでは CodeCommit の操作は不要であり、各種 AWS (CloudWatch Logs) への操作権限が必要です。
別のチームが用意している専用のシェルを source してスイッチバックを実行してください。

  スイッチバック用シェル配置場所: /path/to/switchback.sh
  手動実行例: source "/path/to/switchback.sh"

※ '--auto-switchback' (-a) オプションを指定すると、
   終了せずに上記シェルを自動的に source してスイッチバックを実行し処理を継続できます。
※ シェルの配置場所は '--switchback-script <PATH>' (-S) または環境変数 AWS_SWITCHBACK_SCRIPT で指定可能です。
```

### ② 終了せずにスイッチバックを自動実行する機能 (`--auto-switchback` / `-a`)
`--auto-switchback` オプションを指定した場合、権限不足を検知しても終了せず、指定された専用シェルを自動的に `source` で呼び出してスイッチバックを実行します。
スイッチバック完了後に権限が正常に復帰したことを再検証した上で、クエリ処理をそのままシームレスに継続します。

### ③ スイッチバック用シェルの配置場所指定
- **CLI オプション**: `-S <PATH>` または `--switchback-script <PATH>`
- **環境変数**: `AWS_SWITCHBACK_SCRIPT="/path/to/switchback.sh"`
- **デフォルトパス**: プロジェクト直下の `switchback.sh`
- **参考サンプル**: `examples/switchback_sample.sh` にスイッチバック用シェルの実装サンプルを同梱しています。

### 終了コード一覧
| 終了コード | 状態 | 内容 |
|---|---|---|
| `0` | 正常完了 | クエリ実行、CSV変換、Excelレポート生成まで正常完了 |
| `1` | 一般エラー / 認証未完了 | 引数不正、日時パース失敗、または `aws sts get-caller-identity` 未ログイン |
| `2` | 権限不足（スイッチバック警告） | 操作権限不足（CodeCommitロール等）。スイッチバック案内を出力して終了 |
| `3` | スイッチバックシェル未存在 | 自動スイッチバック有効時に指定されたシェルファイルが存在しない |
| `4` | スイッチバック失敗 / 権限不足継続 | スイッチバックシェルの実行失敗、または実行後も依然として権限が不足 |

---

## 7. テストの実行

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
   - 未ログイン（未認証トークン）検出とエラーハンドリング（Exit 1）
   - CodeCommit 専用権限 / 操作権限不足時のスイッチバック警告と終了（Exit 2）
   - `--auto-switchback` による専用シェルの自動 `source` 呼び出し・復帰完了検証（Exit 0）
   - スイッチバックシェルのファイル未存在エラー検出（Exit 3）
   - スイッチバック後も権限不足が継続する場合のエラー検出（Exit 4）
   - 環境変数 `AWS_SWITCHBACK_SCRIPT` によるパス解決検証
