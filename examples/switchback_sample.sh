#!/usr/bin/env bash
# ==============================================================================
# Script: switchback_sample.sh
# Purpose: Sample switchback script provided for reference to switch back from
#          a restricted CodeCommit role or scoped credentials to operational
#          credentials with CloudWatch Logs permissions.
# Target OS: RHEL 9.8 / Bash 5.1+
#
# Usage:
#   source ./examples/switchback_sample.sh
# ==============================================================================

echo "[SWITCHBACK] スイッチバック処理を開始します..."

# ------------------------------------------------------------------------------
# パターン 1: AWS CLI のプロファイルを運用用プロファイルに切り替える場合
# ------------------------------------------------------------------------------
# export AWS_PROFILE="operations"
# export AWS_DEFAULT_REGION="ap-northeast-1"

# ------------------------------------------------------------------------------
# パターン 2: 一時認証情報 (AssumeRole等) を環境変数に設定する場合
# ------------------------------------------------------------------------------
# export AWS_ACCESS_KEY_ID="ASIAXXXXXXXXXXXXXXXX"
# export AWS_SECRET_ACCESS_KEY="YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY"
# export AWS_SESSION_TOKEN="ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ..."
# export AWS_DEFAULT_REGION="ap-northeast-1"
# unset AWS_PROFILE

# ------------------------------------------------------------------------------
# パターン 3: 外部の認証コマンドやAssumeRoleツールを実行する場合
# ------------------------------------------------------------------------------
# eval "$(assume-role OperationsRole)"

echo "[SWITCHBACK] スイッチバック処理が完了しました。現在の AWS 認証情報を確認します:"
aws sts get-caller-identity --output table 2>/dev/null || echo "[SWITCHBACK] 警告: get-caller-identity に失敗しました。"
