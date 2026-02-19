"""ثابت‌ها و مسیرهای پروژه — یک جا تغییر بده، همه‌جا اعمال بشه."""

# ── مسیر فایل ورودی ──
MONITORING_PATH = "1 الی 23 بهمن مونیتورینگ.xlsx"

# ── مسیرهای خروجی ──
MATCHED_REPORT_PATH = "matched_report.xlsx"
ENRICHED_PATH = "enriched_monitoring.xlsx"
ANALYTICS_PATH = "analytics_report.xlsx"

# ── آستانه‌های مچینگ ──
SUPPORT_MATCH_THRESHOLD = 75
RATE_MATCH_THRESHOLD = 80

# ── بازه‌های زمانی (دقیقه) ──
SUPPORT_WINDOW_BEFORE = 45  # دقیقه قبل
SUPPORT_WINDOW_AFTER = 5  # دقیقه بعد
RATE_WINDOW_BEFORE = 5  # دقیقه قبل از connect_time
RATE_WINDOW_AFTER = 45  # دقیقه بعد از connect_time
