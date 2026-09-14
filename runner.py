import datetime
import os
import requests

# جلب توكن التليجرام والآيدي من GitHub Secrets
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# قائمة بجميع ملفات التكست واسم كل استراتيجية
TXT_FILES = [
    ("BROADING BOTTOMS.txt", "Broadening Bottoms"),
    ("FLAGS.txt", "Flags"),
    ("adam and adam.txt", "Adam & Adam"),
    ("adam and eva.txt", "Adam & Eve"),
    ("asending trainangle.txt", "Ascending Triangle"),
    ("pipe bottom.txt", "Pipe Bottom"),
    ("tripple bottom.txt", "Triple Bottom"),
]


def send_telegram_message(message: str):
    """إرسال التنبيهات عبر التليجرام"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️ Telegram Secrets missing/not configured.")
        return

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML",
    }
    try:
        res = requests.post(url, json=payload, timeout=15)
        if res.status_code == 200:
            print("✅ Telegram notification sent successfully!")
        else:
            print(f"❌ Telegram API Error: {res.text}")
    except Exception as e:
        print(f"⚠️ Exception sending Telegram message: {e}")


def run_txt_script(filename: str, strategy_name: str):
    """قراءة وتحديث تشغيل ملف الـ txt واستخراج الصفقات المفتوحة"""
    if not os.path.exists(filename):
        print(f"⚠️ File not found: {filename}")
        return []

    print(f"🔍 Processing: {filename} ({strategy_name})...")

    # بيئة معزولة لتشغيل الكود بداخلها
    local_scope = {}

    try:
        with open(filename, "r", encoding="utf-8") as f:
            code = f.read()

        # تنفيذ كود الملف داخل النطاق المعزول
        exec(code, local_scope)

        trades = []
        # البحث عن دالة الفحص المتاحة بداخل ملف الـ txt
        if "run_full_backtest" in local_scope:
            trades = local_scope["run_full_backtest"]()
        elif "run_full_scan" in local_scope:
            trades = local_scope["run_full_scan"]()

        open_trades = []
        if trades:
            for t in trades:
                # تصفية الصفقات المفتوحة فقط
                if isinstance(t, dict) and t.get("Status") == "Open":
                    t["Strategy"] = strategy_name
                    open_trades.append(t)

        return open_trades

    except Exception as e:
        print(f"⚠️ Error executing {filename}: {e}")
        return []


def main():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    print(f"🚀 Starting EGX Multi-Strategy TXT Scan ({today_str})...\n")

    all_open_signals = []

    # المرور على الملفات السبعة واحدًا تلو الآخر
    for filename, strat_name in TXT_FILES:
        signals = run_txt_script(filename, strat_name)
        all_open_signals.extend(signals)

    # إذا لم تكن هناك أي صفقة مفتوحة
    if not all_open_signals:
        msg = f"📊 <b>EGX Market Scan ({today_str})</b>\n\nNo active OPEN signals found across all strategies today."
        print(msg)
        send_telegram_message(msg)
        return

    # صياغة رسالة التليجرام المجمعة
    msg_lines = [
        f"🚨 <b>EGX ALL STRATEGIES - ACTIVE SIGNALS</b> 🚨",
        f"📅 <i>Date: {today_str}</i>",
        f"🌐 Total Active Signals: <b>{len(all_open_signals)}</b>\n",
        "========================================",
    ]

    for sig in all_open_signals:
        stock = str(sig.get("Stock Name", "N/A")).replace(".CA", "")
        strat = sig.get("Strategy", "N/A")
        entry_d = sig.get("Entry Date", "N/A")
        entry_p = sig.get("Entry Price", 0.0)
        curr_p = sig.get("Current Price", 0.0)
        target = sig.get("Target Price", 0.0)
        stop = sig.get("Stop Loss", 0.0)
        pnl = sig.get("Win Rate", 0.0) * 100

        pnl_emoji = "🟢" if pnl >= 0 else "🔴"

        card = (
            f"🎯 <b>Strategy: {strat}</b>\n"
            f"📈 <b>Stock: #{stock}</b>\n"
            f"📅 Entry Date: {entry_d}\n"
            f"💵 Entry Price: <b>{entry_p} EGP</b>\n"
            f"📊 Current Price: {curr_p} EGP ({pnl_emoji} {pnl:+.2f}%)\n"
            f"🎯 Target: <b>{target} EGP</b>\n"
            f"🛑 Stop Loss: <b>{stop} EGP</b>\n"
            "----------------------------------------"
        )
        msg_lines.append(card)

    final_msg = "\n".join(msg_lines)
    print("\n" + final_msg)

    # إرسال الرسالة إلى تليجرام
    send_telegram_message(final_msg)


if __name__ == "__main__":
    main()
