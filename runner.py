import datetime
import glob
import os
import pandas as pd
import requests

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
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️ Telegram Secrets missing/not configured.")
        return

    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
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


def find_actual_file(target_filename: str) -> str:
    if os.path.exists(target_filename):
        return target_filename
    for f in os.listdir("."):
        if f.lower() == target_filename.lower():
            return f
    return target_filename


def extract_trades(local_scope, strategy_name, excel_files_before):
    extracted_open_trades = []
    df_result = None

    # 1. البحث عن ملف Excel جديد تم إنشاؤه
    excel_files_after = set(glob.glob("*.xlsx"))
    new_excel_files = list(excel_files_after - excel_files_before)

    if new_excel_files:
        latest_file = max(new_excel_files, key=os.path.getmtime)
        try:
            df_result = pd.read_excel(latest_file)
            print(f"  └─ 📁 Read from Excel: {latest_file}")
        except Exception as e:
            print(f"  └─ ⚠️ Could not read Excel: {e}")

    # 2. فحص شامل للذاكرة: البحث عن أي DataFrame موجود داخل local_scope بغض النظر عن اسمه
    if df_result is None or df_result.empty:
        for var_name, var_value in local_scope.items():
            if var_name.startswith("__"):
                continue
            if isinstance(var_value, pd.DataFrame) and not var_value.empty:
                df_result = var_value.copy()
                print(f"  └─ 🧠 Found DataFrame in memory: '{var_name}'")
                break
            elif isinstance(var_value, list) and len(var_value) > 0:
                if isinstance(var_value[0], dict):
                    df_result = pd.DataFrame(var_value)
                    print(
                        f"  └─ 🧠 Converted list of dicts from memory: '{var_name}'"
                    )
                    break

    if df_result is None or df_result.empty:
        print("  └─ ⚠️ No DataFrame or list of trades found in memory.")
        return []

    # توحيد أسماء الأعمدة
    df_result.columns = [str(c).strip().title() for c in df_result.columns]

    # 3. تحديد عمود الحالة بمرونة كاملة
    status_col = None
    for col in [
        "State",
        "Status",
        "Trade Status",
        "Position Status",
        "Trade_State",
    ]:
        if col in df_result.columns:
            status_col = col
            break

    exit_date_col = None
    for col in [
        "Exit Date",
        "Exit_Date",
        "Exitdate",
        "Close Date",
        "Close_Date",
    ]:
        if col in df_result.columns:
            exit_date_col = col
            break

    # 4. تصفية الصفقات المفتوحة
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
            | (df_result[exit_date_col].astype(str).str.lower() == "none")
        ]
    else:
        # إذا لم يوجد عمود حالة ولا تاريخ خروج، نعتبر جميع الأسطر صفقات متاحة
        open_df = df_result.copy()

    if open_df.empty:
        print("  └─ ℹ️ Trades found, but 0 positions matched 'OPEN' state.")
        return []

    # 5. استخراج البيانات
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
    actual_filename = find_actual_file(filename)

    if not os.path.exists(actual_filename):
        print(f"⚠️ File not found: {filename}")
        return []

    print(f"🔍 Processing: {actual_filename} ({strategy_name})...")

    local_scope = {}
    excel_files_before = set(glob.glob("*.xlsx"))

    try:
        with open(actual_filename, "r", encoding="utf-8") as f:
            code = f.read()

        exec(code, local_scope)

        open_trades = extract_trades(
            local_scope, strategy_name, excel_files_before
        )
        print(f"  └─ 🟢 Result: {len(open_trades)} OPEN position(s).")

        # حذف ملفات Excel الناتجة لتجنب قراءتها في الاستراتيجية التالية
        excel_files_after = set(glob.glob("*.xlsx"))
        for nf in excel_files_after - excel_files_before:
            try:
                os.remove(nf)
            except Exception:
                pass

        return open_trades

    except Exception as e:
        print(f"⚠️ Error executing {actual_filename}: {e}")
        return []


def main():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    print(f"🚀 Starting EGX Scan ({today_str})...\n")

    all_open_signals = []

    for filename, strat_name in TXT_FILES:
        signals = run_txt_script(filename, strat_name)
        all_open_signals.extend(signals)

    if not all_open_signals:
        msg = f"📊 <b>EGX Market Scan ({today_str})</b>\n\nNo active OPEN signals found today."
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
