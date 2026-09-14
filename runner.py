import datetime
import glob
import importlib
import os
import pandas as pd
import requests

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

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


def parse_trades_from_dataframe(df: pd.DataFrame, strategy_name: str) -> list:
    """تحديد واستخراج الصفقات المفتوحة بجميع الصياغات الممكنة"""
    open_signals = []
    if df.empty:
        return open_signals

    # توحيد أسماء الأعمدة
    df.columns = [str(c).strip().title() for c in df.columns]

    # البحث عن أعمدة الحالة وتاريخ الخروج
    status_col = next((c for c in df.columns if any(k in c.lower() for k in ["state", "status", "trade status"])), None)
    exit_col = next((c for c in df.columns if any(k in c.lower() for k in ["exit date", "exit_date", "close date", "exit"])), None)

    for _, row in df.iterrows():
        is_open = False
        
        # 1. التحقق من عمود الحالة
        if status_col:
            val = str(row.get(status_col, "")).strip().upper()
            if any(s in val for s in ["OPEN", "ACTIVE", "مفتوحة", "مستمرة"]):
                is_open = True

        # 2. التحقق من عمود تاريخ الخروج إذا لم تقتنع بقيمة الحالة
        if not is_open and exit_col:
            val = str(row.get(exit_col, "")).strip().lower()
            if val in ["", "none", "nan", "nat", "0", "null"]:
                is_open = True

        # 3. إذا لم يوجد العمودان، نعتبر الصفقات التي ليس لها خروج مفتوحة
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

    return open_signals


def run_py_strategy(module_name: str, strategy_name: str) -> list:
    print(f"🔍 Processing: {module_name}.py ({strategy_name})...")
    excel_before = set(glob.glob("*.xlsx"))

    try:
        strat_module = importlib.import_module(module_name)
        open_signals = []

        # 1. البحث عن الـ DataFrames المتولدة في ملف البايثون
        for var_name in ["all_trades", "trades", "trades_df", "results", "df_results", "open_trades"]:
            trades_raw = getattr(strat_module, var_name, None)
            if trades_raw is not None:
                df = trades_raw if isinstance(trades_raw, pd.DataFrame) else pd.DataFrame(trades_raw)
                open_signals = parse_trades_from_dataframe(df, strategy_name)
                if open_signals:
                    break

        # 2. القراءة مباشرة من ملف الـ Excel الجديد إن لم يجد الصفقات في الذاكرة
        if not open_signals:
            excel_after = set(glob.glob("*.xlsx"))
            new_files = list(excel_after - excel_before)
            if new_files:
                latest_excel = max(new_files, key=os.path.getmtime)
                df = pd.read_excel(latest_excel)
                open_signals = parse_trades_from_dataframe(df, strategy_name)

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
    print(f"🚀 Starting EGX Multi-Strategy Scan ({today_str})...\n")

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
