#!/usr/bin/env bash
# 毎朝6:00に fortune.py を実行するための cron 登録スクリプト。
# 使い方:  bash schedule_6am.sh        # 登録
#         bash schedule_6am.sh remove  # 解除
set -euo pipefail

APP_DIR="$(cd "$(dirname "$0")" && pwd)"
PY="$(command -v python3)"
CMD="0 6 * * * cd ${APP_DIR} && ${PY} fortune.py >> ${APP_DIR}/logs/cron.log 2>&1"
TAG="# fortune-divination-app"

mkdir -p "${APP_DIR}/logs"

if [[ "${1:-install}" == "remove" ]]; then
    crontab -l 2>/dev/null | grep -v "${TAG}" | crontab -
    echo "removed 6am fortune cron."
    exit 0
fi

# 既存の同タグ行を消してから追記（重複防止）
( crontab -l 2>/dev/null | grep -v "${TAG}"; echo "${CMD} ${TAG}" ) | crontab -
echo "installed: 毎朝06:00 に ${APP_DIR}/fortune.py を実行します。"
echo "確認: crontab -l"
