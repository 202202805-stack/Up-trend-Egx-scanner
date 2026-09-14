import datetime
import glob
import os
import pandas as pd
import requests

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
    """إرسال التنبيهات عبر التليجرام مع التعامل مع حدود طول الرسالة (4096 حرف)"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️ Telegram Secrets missing/not configured.")
        return

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"

    # تقطيع الرسالة إلى أجزاء إذا تجاوزت 4000 حرف لتجنب رفض Telegram API
    max_len = 4000
    chunks = [
        message[i : i + max_len] for i in range(0, len(message), max_len)
    ]

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


def extract_trades(local_scope, strategy_name, excel_files_before):
    """استخراج الصفقات المفتوحة بناءً على شرط State/Status == Open أو عدم وجود تاريخ خروج"""
    extracted_open_trades = []
    df_result = None

    # 1. البحث عن ملف Excel الجديد الذي تم إنشاؤه أثناء تشغيل هذا الملف
    excel_files_after = set(glob.glob("*.xlsx"))
    new_excel_files = list(excel_files_after - excel_files_before)

    if new_excel_files:
        latest_file = max(new_excel_files, key=os.path.getmtime)
        try:
            df_result = pd.read_excel(latest_file)
            print(f"  └─ 📁 Read results from generated Excel: {latest_file}")
        except Exception as e:
            print(f"  └─ ⚠️ Could not read Excel {latest_file}: {e}")

    # 2. إذا لم يجد ملف Excel جديد، يبحث في متغيرات الذاكرة
    if df_result is None or df_result.empty:
        for var_name in [
            "all_trades",
            "trades",
            "trades_df",
            "results",
            "open_positions",
            "df_results",
            "open_trades",
        ]:
            if (
                var_name in local_scope
                and local_scope[var_name] is not None
            ):
                val = local_scope[var_name]
                if isinstance(val, pd.DataFrame):
                    df_result = val.copy()
                elif isinstance(val, list):
                    df_result = pd.DataFrame(val)
                break

    if df_result is None or df_result.empty:
        return []

    # توحيد أسماء الأعمدة وإزالة المسافات الزائدة
    df_result.columns = [str(c).strip().title() for c in df_result.columns]

    # 3. تحديد الصفقات المفتوحة بمرونة عالية (دعم State و Status)
    status_col = None
    for col in ["State", "Status", "Trade Status", "Position Status"]:
        if col in df_result.columns:
            status_col = col
            break

    exit_date_col = None
    for col in ["Exit Date", "Exit_Date", "Exitdate", "Close Date"]:
        if col in df_result.columns:
            exit_date_col = col
            break

    # الأولوية لعمود State/Status إذا وجد، ثم الاعتماد على تاريخ الخروج
    if status_col:
        open_df = df_result[
            df_result[status_col]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.contains("OPEN|ACTIVE|مفتوحة|مستمرة")
        ]
    elif exit_date_col:
        open_df = df_result[
            df_result[exit_date_col].isna()
            | (df_result[exit_date_col].astype(str).str.strip() == "")
            | (df_result[exit_date_col].astype(str).str.lower() == "nan")
            | (df_result[exit_date_col].astype(str).str.lower() == "nat")
        ]
    else:
        open_df = pd.DataFrame()

    if open_df.empty:
        return []

    # 4. قراءة تفاصيل الصفقات المفتوحة
    for _, row in open_df.iterrows():

        def get_val(keys, default=0.0):
            for k in keys:
                for col in df_result.columns:
                    if k.lower() == col.lower():
                        val = row.get(col)
                        if pd.notnull(val):
                            return val
            return default

        stock = get_val(["Stock Name", "Ticker", "Stock", "Symbol"], "N/A")
        entry_d = get_val(["Entry Date", "Date", "Entry_Date"], "N/A")
        entry_p = get_val(
            ["Entry Price", "Buy Price", "Entry_Price", "Price"], 0.0
        )
        curr_p = get_val(
            ["Current Price", "Last Price", "Close", "Current_Price"], entry_p
        )
        target = get_val(["Target", "Target Price", "Target_Price"], 0.0)
        stop = get_val(["Stop Loss", "Stop", "Stop_Loss"], 0.0)
        pnl = get_val(
            ["Pnp_Ratio", "PnL %", "Unrealized PnL %", "Return %", "Pnl"],
            0.0,
        )

        try:
            entry_float = float(entry_p)
            curr_float = float(curr_p)
            pnl_float = float(pnl)

            if pnl_float == 0.0 and entry_float > 0 and curr_float > 0:
                pnl = ((curr_float - entry_float) / entry_float) * 100
            elif abs(pnl_float) <= 1.0 and pnl_float != 0.0:
                pnl = pnl_float * 100
            else:
                pnl = pnl_float
        except Exception:
            pnl = 0.0

        extracted_open_trades.append({
            "Strategy": strategy_name,
            "Stock": str(stock).replace(".CA", ""),
            "Entry Date": str(entry_d)[:10],
            "Entry Price": round(float(entry_p), 3) if entry_p else 0.0,
            "Current Price": round(float(curr_p), 3) if curr_p else 0.0,
            "Target": round(float(target), 3) if target else 0.0,
            "Stop Loss": round(float(stop), 3) if stop else 0.0,
            "PnL": round(float(pnl), 2),
        })

    return extracted_open_trades


def run_txt_script(filename: str, strategy_name: str):
    if not os.path.exists(filename):
        print(f"⚠️ File not found: {filename}")
        return []

    print(f"🔍 Processing: {filename} ({strategy_name})...")
    local_scope = {}
    excel_files_before = set(glob.glob("*.xlsx"))

    try:
        with open(filename, "r", encoding="utf-8") as f:
            code = f.read()

        exec(code, local_scope)

        open_trades = extract_trades(
            local_scope, strategy_name, excel_files_before
        )
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
