# report_maker/core/settings.py
from datetime import timezone, timedelta

JST = timezone(timedelta(hours=9))

APP_TITLE = "故障報告書自動生成"

SHEET_NAME = "緊急出動報告書（リンク付き）"
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]

REQUIRED_KEYS = [
    "通報者", "受信内容", "現着状況", "原因", "処置内容", "処理修理後", "所属",
]
