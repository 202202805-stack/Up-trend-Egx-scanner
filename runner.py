import datetime
import importlib
import os
import requests

# جلب توكن التليجرام والآيدي من GitHub Secrets
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# قائمة الملفات بعد تحويلها لـ .py وأسمائها البرمجية
PYTHON_STRATEGIES = [
    ("broadening_bottoms", "Broadening Bottoms"),
    ("FLAGS", "Flags"),
    ("adam_and_adam", "Adam & Adam"),
    ("adam_and_eva", "Adam & Eve"),
    ("ascending_triangle", "Ascending Triangle"),
    ("pipe_bottom", "Pipe Bottom"),
    ("tripple_bottom", "Triple Bottom"),
]


def send_telegram_message(message: str):
    """إرسال التنبيهات عبر التليجرام مع التعامل مع حدود طول الرسالة"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️ Telegram Secrets missing/not configured.")
        return

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    max_len = 4000
    chunks = [message[i : i + max_len] for i in range(0, len(message), max_len)]

    for chunk in chunks:
        payload = {
            "chat_id": TELEGRAM_CHAT_ID,
            "text": chunk,
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


def run_py_strategy(module_name: str, strategy_name: str) -> list:
    """استدعاء ملف البايثون مباشرة وتشغيل دالة الفحص فيه"""
    print(f"🔍 Processing: {module_name}.py ({strategy_name})...")
    
    try:
        # استدعاء الملف كـ Module
        strat_module = importlib.import_module(module_name)
        
        # التأكد من وجود صفقات مخرجة
        trades = []
        if hasattr(strat_module, "all_trades") and strat_module.all_trades:
            trades = strat_module.all_trades
        elif hasattr(strat_module, "main"):
            # في حال كان الكود يعمل داخل دالة main
            trades = getattr(strat_module, "main")() or []

        open_signals = []
        for trade in trades:
            # فلترة الصفقات المفتوحة فقط
            status = str(trade.get("Status", "")).strip().upper()
            exit_date = trade.get("Exit Date")
            
            is_open = status in ["OPEN", "ACTIVE", "مفتوحة"] or exit_date is None or str(exit_date).strip() in ["", "None", "nan", "NaT"]

            if is_open:
                entry_p = float(trade.get("Entry Price", 0.0))
                curr_p = float(trade.get("Current Price", entry_p))
                
                pnl = round(((curr_p - entry_p) / entry_p) * 100, 2) if entry_p > 0 else 0.0

                open_signals.append({
                    "Strategy": strategy_name,
                    "Stock": str(trade.get("Stock Name", trade.get("Ticker", "N/A"))).replace(".CA", ""),
                    "Entry Date": str(trade.get("Entry Date", "N/A"))[:10],
                    "Entry Price": round(entry_p, 3),
                    "Current Price": round(curr_p, 3),
                    "Target": round(float(trade.get("Target", 0.0)), 3),
                    "Stop Loss": round(float(trade.get("Stop Loss", 0.0)), 3),
                    "PnL": pnl,
                })

        print(f"  └─ 🟢 Found {len(open_signals)} OPEN position(s).")
        return open_signals

    except ModuleNotFoundError:
        print(f"⚠️ Module not found: {module_name}.py. Skipping.")
        return []
    except Exception as e:
        print(f"⚠️ Error executing {module_name}.py: {e}")
        return []


def main():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    print(f"🚀 Starting EGX Multi-Strategy Python Scan ({today_str})...\n")

    all_open_signals = []

    for module_name, strat_name in PYTHON_STRATEGIES:
        signals = run_py_strategy(module_name, strat_name)
        all_open_signals.extend(signals)

    if not all_open_signals:
        msg = f"📊 <b>EGX Market Scan ({today_str})</b>\n\nNo active OPEN signals found across all strategies today."
        print("\n" + msg)
        send_telegram_message(msg)
        return

    msg_lines = [
        "🚨 <b>EGX ALL STRATEGIES - ACTIVE SIGNALS</b> 🚨",
        f"📅 <i>Date: {today_str}</i>",
        f"🌐 Total Active Signals: <b>{len(all_open_signals)}</b>\n",
        "========================================",
    ]

    for sig in all_open_signals:
        pnl_val = sig["PnL"]
        pnl_emoji = "🟢" if pnl_val >= 0 else "🔴"

        card = (
            f"🎯 <b>Strategy: {sig['Strategy']}</b>\n"
            f"📈 <b>Stock: #{sig['Stock']}</b>\n"
            f"📅 Entry Date: {sig['Entry Date']}\n"
            f"💵 Entry Price: <b>{sig['Entry Price']} EGP</b>\n"
            f"📊 Current Price: {sig['Current Price']} EGP ({pnl_emoji} {pnl_val:+.2f}%)\n"
            f"🎯 Target: <b>{sig['Target']} EGP</b>\n"
            f"🛑 Stop Loss: <b>{sig['Stop Loss']} EGP</b>\n"
            f"----------------------------------------"
        )
        msg_lines.append(card)

    final_msg = "\n".join(msg_lines)
    print("\n" + final_msg)
    send_telegram_message(final_msg)


if __name__ == "__main__":
    main()
