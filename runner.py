import datetime
import os
import requests
import pandas as pd

# جلب توكن التليجرام والآيدي من GitHub Secrets
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

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

def extract_trades_from_scope(local_scope, strategy_name):
    """البحث الشامل والدقيق عن الصفقات المفتوحة بداخل كل ملف"""
    extracted_open_trades = []

    raw_trades = None

    # 1. البحث في المتغيرات المخزنة مباشرة في الذاكرة بعد التنفيذ
    for var_name in ["all_trades", "trades", "trades_df", "results", "open_positions", "df_results", "open_trades"]:
        if var_name in local_scope and local_scope[var_name] is not None:
            raw_trades = local_scope[var_name]
            break

    # 2. تشغيل دالة الفحص إن لم تكن النتائج مخزنة في متغير جاهز
    if raw_trades is None:
        for func_name in ["run_full_backtest", "run_full_scan", "main", "scan_market"]:
            if func_name in local_scope and callable(local_scope[func_name]):
                try:
                    res = local_scope[func_name]()
                    if res is not None:
                        raw_trades = res
                        break
                except Exception:
                    pass

    # تحويل النتائج إلى قائمة
    trades_list = []
    if hasattr(raw_trades, "to_dict"):  # إذا كانت DataFrame
        trades_list = raw_trades.to_dict(orient="records")
    elif isinstance(raw_trades, list):
        trades_list = raw_trades

    # 3. تصفية واستخراج الصفقات المفتوحة
    for t in trades_list:
        if not isinstance(t, dict):
            continue

        status = str(
            t.get("Status") or 
            t.get("Trade Status") or 
            t.get("Position Status") or 
            t.get("status") or ""
        ).strip().upper()

        if status in ["OPEN", "ACTIVE", "مفتوحة", "مستمرة"]:
            stock = t.get("Stock Name") or t.get("Ticker") or t.get("Stock") or t.get("Symbol") or "N/A"
            entry_d = t.get("Entry Date") or t.get("Date") or "N/A"
            entry_p = t.get("Entry Price") or t.get("Buy Price") or 0.0
            curr_p = t.get("Current Price") or t.get("Last Price") or entry_p
            target = t.get("Target Price") or t.get("Target") or 0.0
            stop = t.get("Stop Loss") or t.get("Stop") or 0.0
            pnl = t.get("PnL %") or t.get("Unrealized PnL %") or t.get("Return %") or t.get("Win Rate") or 0.0

            if isinstance(pnl, (float, int)) and abs(pnl) <= 1.0:
                pnl = pnl * 100

            extracted_open_trades.append({
                "Strategy": strategy_name,
                "Stock": str(stock).replace(".CA", ""),
                "Entry Date": str(entry_d)[:10],
                "Entry Price": round(float(entry_p), 2) if entry_p else 0.0,
                "Current Price": round(float(curr_p), 2) if curr_p else 0.0,
                "Target": round(float(target), 2) if target else 0.0,
                "Stop Loss": round(float(stop), 2) if stop else 0.0,
                "PnL": round(float(pnl), 2) if pnl else 0.0,
            })

    return extracted_open_trades

def run_txt_script(filename: str, strategy_name: str):
    if not os.path.exists(filename):
        print(f"⚠️ File not found: {filename}")
        return []

    print(f"🔍 Processing: {filename} ({strategy_name})...")
    local_scope = {}

    try:
        with open(filename, "r", encoding="utf-8") as f:
            code = f.read()

        # تنفيذ كود الملف
        exec(code, local_scope)

        # استخراج الصفقات المفتوحة
        open_trades = extract_trades_from_scope(local_scope, strategy_name)
        print(f"  └─ 🟢 Found {len(open_trades)} OPEN position(s).")
        return open_trades

    except Exception as e:
        print(f"⚠️ Error executing {filename}: {e}")
        return []

def main():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    print(f"🚀 Starting EGX Multi-Strategy TXT Scan ({today_str})...\n")

    all_open_signals = []

    for filename, strat_name in TXT_FILES:
        signals = run_txt_script(filename, strat_name)
        all_open_signals.extend(signals)

    if not all_open_signals:
        msg = f"📊 <b>EGX Market Scan ({today_str})</b>\n\nNo active OPEN signals found across all 7 strategies today."
        print("\n" + msg)
        send_telegram_message(msg)
        return

    msg_lines = [
        f"🚨 <b>EGX ALL STRATEGIES - ACTIVE SIGNALS</b> 🚨",
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
            "----------------------------------------"
        )
        msg_lines.append(card)

    final_msg = "\n".join(msg_lines)
    print("\n" + final_msg)

    send_telegram_message(final_msg)

if __name__ == "__main__":
    main()
