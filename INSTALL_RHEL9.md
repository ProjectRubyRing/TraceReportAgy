# RHEL 9.8 環境構築 ＆ 導入ガイド

本ドキュメントでは、Red Hat Enterprise Linux 9.8 (RHEL 9.8) 環境で本ツール（`export_trace_report.sh`）をセットアップし、CloudWatch Logs Insights と連携して実行する手順を説明します。

---

## 1. 前提条件とシステム要件

- **OS**: Red Hat Enterprise Linux 9.8 (x86_64 / aarch64)
- **シェル**: GNU bash 5.1 以上
- **必須ツール**:
  - `aws-cli` (v2)
  - `python3` (3.9 以上 - RHEL 9.8 標準パッケージ)
  - `openpyxl` (Python Excel 操作ライブラリ)

---

## 2. RHEL 9.8 パッケージのインストール

RHEL 9.8 では、`dnf` パッケージマネージャーを使用して必要なソフトウェアを導入します。

### (1) Python3 および openpyxl のインストール
```bash
# Python 3 と openpyxl パッケージのインストール
sudo dnf install -y python3 python3-pip python3-openpyxl
```

※ `python3-openpyxl` は EPEL または CodeReady Linux Builder (CRB) リポジトリから提供されています。もし `dnf` で見つからない場合は、以下のいずれかで導入可能です：
```bash
# オプション A: pip を利用してシステムまたはユーザー環境にインストール
pip3 install --user openpyxl

# オプション B: 仮想環境（venv）を利用
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl
```
※ 本スクリプト `bin/export_trace_report.sh` は、`openpyxl` が見つからない場合でも自動的にプロジェクト直下に `.venv` を検出し、必要に応じて自動セットアップを試行するフォールバック機能を内包しています。

### (2) AWS CLI v2 のインストール
RHEL 9.8 に AWS CLI v2 がまだインストールされていない場合は、公式バイナリをインストールします：

```bash
# AWS CLI v2 インストーラのダウンロードと解凍
curl "https://awscli.amazonaws.com/awscli-exe-linux-x86_64.zip" -o "awscliv2.zip"
unzip awscliv2.zip

# インストール実行
sudo ./aws/install

# バージョン確認
aws --version
# 出力例: aws-cli/2.x.x Python/3.x.x Linux/5.14.0-xxx.el9.x86_64 ...
```

---

## 3. AWS IAM 権限の設定

本ツールを実行する IAM ユーザーまたは EC2 / ECS インスタンスロールには、以下の CloudWatch Logs Insights 権限が必要です。

### 必要とされる IAM ポリシー例
```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Sid": "CloudWatchLogsInsightsSpansQuery",
      "Effect": "Allow",
      "Action": [
        "logs:StartQuery",
        "logs:GetQueryResults",
        "logs:StopQuery",
        "logs:DescribeLogGroups",
        "logs:GetLogEvents"
      ],
      "Resource": [
        "arn:aws:logs:*:*:log-group:aws/spans",
        "arn:aws:logs:*:*:log-group:aws/spans:*"
      ]
    }
  ]
}
```

### AWS CLI 認証情報の設定
```bash
aws configure
# AWS Access Key ID [None]: <アクセスキー>
# AWS Secret Access Key [None]: <シークレットキー>
# Default region name [None]: ap-northeast-1
# Default output format [None]: json
```
※ EC2 上で実行する場合は、IAM インスタンスプロファイル（ロール）をアタッチすることを推奨します。

---

## 4. スクリプトの実行権限付与と動作確認

```bash
# 実行権限を付与
chmod +x bin/export_trace_report.sh
chmod +x tests/*.sh

# (1) オフラインモックテストを実行して環境動作を確認
./bin/export_trace_report.sh --mock

# (2) 全テストスイートの実行
bash tests/run_all_tests.sh
```

---

## 5. 本番実行例

### 過去 1 時間の全スパンをレポート化
```bash
./bin/export_trace_report.sh --start-time 1h
```

### 日時範囲を明示指定して実行
```bash
./bin/export_trace_report.sh \
  --start-time "2026-10-26 09:00:00" \
  --end-time "2026-10-26 10:00:00" \
  --region ap-northeast-1
```

### 特定の Trace ID にフォーカスしたウォーターフォールレポート
```bash
./bin/export_trace_report.sh \
  --trace-id "1-671c6d1f-8181283921839182391283" \
  --start-time 24h
```

### 出力ファイル
実行が完了すると、`output/` ディレクトリ配下に以下の 3 つのファイルが生成されます：
- `output/query_results_<TIMESTAMP>.json`: CloudWatch から取得した生のクエリ応答
- `output/trace_spans_<TIMESTAMP>.csv`: スパン情報を整形・正規化した CSV ファイル
- `output/trace_report_<TIMESTAMP>.xlsx`: X-Ray 画面を再現した Excel レポート
