import datetime
import glob
import os
import subprocess
import pandas as pd
import requests

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# قائمة الاستراتيجيات السبع (اسم الملف، اسم الاستراتيجية للتليجرام)
PYTHON_STRATEGIES = [
    ("broadening_bottoms.py", "Broadening Bottoms"),
    ("FLAGS.py", "Flags"),
    ("adam_and_adam.py", "Adam & Adam"),
    ("adam_and_eva.py", "Adam & Eve"),
    ("ascending_triangle.py", "Ascending Triangle"),
    ("pipe_bottom.py", "Pipe Bottom"),
    ("tripple_bottom.py", "Triple Bottom"),
]


def send_telegram_message(message: str):
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
            if res.status_code != 200:
                print(f"❌ Telegram API Error: {res.text}")
        except Exception as e:
            print(f"⚠️ Exception sending Telegram message: {e}")


def parse_trades_from_excel(excel_file: str, strategy_name: str) -> list:
    open_signals = []
    if not os.path.exists(excel_file):
        return open_signals

    try:
        df = pd.read_excel(excel_file)
        if df.empty:
            return open_signals

        df.columns = [str(c).strip().title() for c in df.columns]

        status_col = next((c for c in df.columns if any(k in c.lower() for k in ["state", "status", "trade status"])), None)
        exit_col = next((c for c in df.columns if any(k in c.lower() for k in ["exit date", "exit_date", "close date", "exit"])), None)

        for _, row in df.iterrows():
            is_open = False

            if status_col:
                val = str(row.get(status_col, "")).strip().upper()
                if any(s in val for s in ["OPEN", "ACTIVE", "مفتوحة", "مستمرة"]):
                    is_open = True

            if not is_open and exit_col:
                val = str(row.get(exit_col, "")).strip().lower()
                if val in ["", "none", "nan", "nat", "0", "null"]:
                    is_open = True

            if not status_col and not exit_col:
                is_open = True

            if is_open:
                entry_p = float(row.get("Entry Price", row.get("Buy Price", row.get("Entry_Price", 0.0))) or 0.0)
                curr_p = float(row.get("Current Price", row.get("Last Price", row.get("Close", entry_p))) or entry_p)

                # حساب نسبة الربح مباشرة وبدقة
                if entry_p > 0 and curr_p > 0:
                    pnl = round(((curr_p - entry_p) / entry_p) * 100, 2)
                else:
                    pnl = 0.0

                stock_name = row.get("Stock Name", row.get("Ticker", row.get("Stock", row.get("Symbol", "N/A"))))

                open_signals.append({
                    "Strategy": strategy_name,
                    "Stock": str(stock_name).replace(".CA", ""),
                    "Entry Date": str(row.get("Entry Date", row.get("Date", "N/A")))[:10],
                    "Entry Price": round(entry_p, 3),
                    "Current Price": round(curr_p, 3),
                    "Target": round(float(row.get("Target", row.get("Target Price", 0.0)) or 0.0), 3),
                    "Stop Loss": round(float(row.get("Stop Loss", row.get("Stop", 0.0)) or 0.0), 3),
                    "PnL": pnl,
                })
    except Exception as e:
        print(f"⚠️ Error reading excel file {excel_file}: {e}")

    return open_signals


def main():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    print(f"🚀 Starting EGX Multi-Strategy Full Scan ({today_str})...\n")

    all_open_signals = []

    for file_name, strat_name in PYTHON_STRATEGIES:
        if not os.path.exists(file_name):
            print(f"⚠️ File {file_name} not found. Skipping...")
            continue

        print(f"🔍 Running {file_name} ({strat_name})...")
        excel_before = set(glob.glob("*.xlsx"))

        # تشغيل السكربت كـ Subprocess مستقل
        try:
            subprocess.run(["python", file_name], check=True)
        except Exception as e:
            print(f"⚠️ Error executing {file_name}: {e}")
            continue

        # التقاط ملف الإكسيل الجديد الذي أنشأته الاستراتيجية
        excel_after = set(glob.glob("*.xlsx"))
        new_files = list(excel_after - excel_before)

        if new_files:
            latest_excel = max(new_files, key=os.path.getmtime)
            signals = parse_trades_from_excel(latest_excel, strat_name)
            all_open_signals.extend(signals)
            print(f"  └─ 🟢 Found {len(signals)} OPEN position(s).")
        else:
            print("  └─ ⚪ No new Excel output generated.")

    print(f"\n🌐 Total Active Signals across all strategies: {len(all_open_signals)}")

    if not all_open_signals:
        msg = f"📊 <b>EGX Market Scan ({today_str})</b>\n\nNo active OPEN signals found across all 7 strategies today."
        print("\n" + msg)
        send_telegram_message(msg)
        return

    msg_lines = [
        "🚨 <b>EGX ALL STRATEGIES - ACTIVE SIGNALS</b> 🚨",
        f"📅 <i>Date: {today_str}</i>",
        f"🌐 Total Signals: <b>{len(all_open_signals)}</b>\n",
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
