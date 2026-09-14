import datetime
import glob
import os
import subprocess
import pandas as pd
import requests

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")


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
    """استخراج الصفقات المفتوحة مباشرة من ملف الإكسيل المولد"""
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

                pnl_val = row.get("Pnp_Ratio", row.get("Pnl %", row.get("Return %", row.get("PnL", 0.0))))
                try:
                    pnl = float(pnl_val)
                    if abs(pnl) <= 1.0 and pnl != 0.0:
                        pnl = pnl * 100
                except Exception:
                    pnl = round(((curr_p - entry_p) / entry_p) * 100, 2) if entry_p > 0 else 0.0

                stock_name = row.get("Stock Name", row.get("Ticker", row.get("Stock", row.get("Symbol", "N/A"))))

                open_signals.append({
                    "Strategy": strategy_name,
                    "Stock": str(stock_name).replace(".CA", ""),
                    "Entry Date": str(row.get("Entry Date", row.get("Date", "N/A")))[:10],
                    "Entry Price": round(entry_p, 3),
                    "Current Price": round(curr_p, 3),
                    "Target": round(float(row.get("Target", row.get("Target Price", 0.0)) or 0.0), 3),
                    "Stop Loss": round(float(row.get("Stop Loss", row.get("Stop", 0.0)) or 0.0), 3),
                    "PnL": round(pnl, 2),
                })
    except Exception as e:
        print(f"⚠️ Error reading excel file: {e}")

    return open_signals


def main():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    print(f"🚀 Running broadening_bottoms.py directly as Subprocess ({today_str})...\n")

    excel_before = set(glob.glob("*.xlsx"))

    # تشغيل ملف الاستراتيجية بالكامل كأنك شغلته يدويًا بالضبط
    try:
        subprocess.run(["python", "broadening_bottoms.py"], check=True)
    except Exception as e:
        print(f"⚠️ Error running script via subprocess: {e}")

    # البحث عن ملف Excel الذي نتج عن العملية
    excel_after = set(glob.glob("*.xlsx"))
    new_files = list(excel_after - excel_before)

    all_open_signals = []
    if new_files:
        latest_excel = max(new_files, key=os.path.getmtime)
        print(f"📁 Reading output file: {latest_excel}")
        all_open_signals = parse_trades_from_excel(latest_excel, "Broadening Bottoms")
    else:
        # لو لم يتولد ملف جديد نتحقق من الملفات الموجودة حالياً
        all_excels = glob.glob("*.xlsx")
        if all_excels:
            latest_excel = max(all_excels, key=os.path.getmtime)
            all_open_signals = parse_trades_from_excel(latest_excel, "Broadening Bottoms")

    print(f"\n  └─ 🟢 Found {len(all_open_signals)} OPEN position(s).")

    if not all_open_signals:
        msg = f"📊 <b>EGX Market Scan ({today_str})</b>\n\nNo active OPEN signals found for Broadening Bottoms."
        print("\n" + msg)
        send_telegram_message(msg)
        return

    msg_lines = [
        "🚨 <b>EGX SCAN - ACTIVE BROADENING BOTTOM SIGNALS</b> 🚨",
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
