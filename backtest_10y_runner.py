import datetime
import glob
import importlib
import os
import sys
import pandas as pd
import requests
import yfinance as yf

# ==========================================
# 1. اعتراض جلب البيانات لإجبار فترة 10 سنوات (Monkey Patching)
# ==========================================
_original_download = yf.download
_original_ticker_history = yf.Ticker.history

def custom_download(*args, **kwargs):
    kwargs["period"] = "10y"
    kwargs.pop("start", None)
    kwargs.pop("end", None)
    return _original_download(*args, **kwargs)

def custom_history(self, *args, **kwargs):
    kwargs["period"] = "10y"
    kwargs.pop("start", None)
    kwargs.pop("end", None)
    return _original_ticker_history(self, *args, **kwargs)

yf.download = custom_download
yf.Ticker.history = custom_history

# ==========================================
# 2. إعدادات تليجرام
# ==========================================
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
            res = requests.post(url, json=payload, timeout=20)
            if res.status_code != 200:
                print(f"❌ Telegram API Error: {res.text}")
        except Exception as e:
            print(f"⚠️ Exception sending Telegram message: {e}")

# ==========================================
# 3. اكتشاف وقراءة الملفات
# ==========================================
def get_all_strategy_files() -> list:
    all_py_files = glob.glob("*.py")
    strategies = []

    for file_path in sorted(all_py_files):
        file_name = os.path.basename(file_path)
        # استثناء جميع ملفات הـ runner والملفات المساعدة
        if file_name.lower().endswith("runner.py") or file_name.startswith("_"):
            continue

        clean_name = os.path.splitext(file_name)[0]
        pretty_strategy_name = clean_name.replace("_", " ").title()
        strategies.append((file_name, clean_name, pretty_strategy_name))

    return strategies

def parse_all_historical_trades(excel_file: str, strategy_name: str) -> dict:
    """قراءة وتحليل كافة صفقات 10 سنوات (الرابحة والخاسرة والمفتوحة)"""
    stats = {
        "Strategy": strategy_name,
        "Total Trades": 0,
        "Winning Trades": 0,
        "Losing Trades": 0,
        "Open Trades": 0,
        "Win Rate": 0.0,
        "Total PnL %": 0.0,
    }

    if not os.path.exists(excel_file):
        return stats

    try:
        df = pd.read_excel(excel_file)
        if df.empty:
            return stats

        df.columns = [str(c).strip().title() for c in df.columns]
        stats["Total Trades"] = len(df)

        # استخراج الأعمدة المتاحة
        pnl_col = next((c for c in df.columns if any(k in c.lower() for k in ["pnl", "profit", "return", "gain", "net"])), None)
        status_col = next((c for c in df.columns if any(k in c.lower() for k in ["state", "status", "trade status"])), None)
        
        total_pnl = 0.0
        wins = 0
        losses = 0
        open_cnt = 0

        for _, row in df.iterrows():
            # حساب الأرباح والخسائر
            pnl_val = 0.0
            if pnl_col:
                try:
                    pnl_val = float(str(row.get(pnl_col, 0)).replace("%", "").strip() or 0.0)
                except ValueError:
                    pnl_val = 0.0
            else:
                entry_p = float(row.get("Entry Price", row.get("Buy Price", 0.0)) or 0.0)
                exit_p = float(row.get("Exit Price", row.get("Current Price", row.get("Close", 0.0))) or 0.0)
                if entry_p > 0 and exit_p > 0:
                    pnl_val = ((exit_p - entry_p) / entry_p) * 100

            total_pnl += pnl_val

            # تحديد حالة الصفقة
            is_open = False
            if status_col:
                val = str(row.get(status_col, "")).strip().upper()
                if any(s in val for s in ["OPEN", "ACTIVE", "مفتوحة", "مستمرة"]):
                    is_open = True
            
            if is_open:
                open_cnt += 1
            else:
                if pnl_val > 0:
                    wins += 1
                elif pnl_val < 0:
                    losses += 1

        stats["Winning Trades"] = wins
        stats["Losing Trades"] = losses
        stats["Open Trades"] = open_cnt
        stats["Total PnL %"] = round(total_pnl, 2)
        
        closed_trades = wins + losses
        if closed_trades > 0:
            stats["Win Rate"] = round((wins / closed_trades) * 100, 1)

    except Exception as e:
        print(f"⚠️ Error reading excel file {excel_file}: {e}")

    return stats

# ==========================================
# 4. الدالة الرئيسية
# ==========================================
def main():
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    python_strategies = get_all_strategy_files()

    print(f"🚀 Starting 10-Year Backtest Scan ({today_str})...")
    print(f"💡 Detected {len(python_strategies)} strategy file(s).\n")

    summary_reports = []
    total_market_trades = 0

    for file_name, module_name, strat_name in python_strategies:
        print(f"🔍 Executing 10Y Backtest for: {file_name} ({strat_name})...")
        excel_before = set(glob.glob("*.xlsx"))

        try:
            # تشغيل الملف في نفس الذاكرة لتطبيق الـ Patching
            if module_name in sys.modules:
                imported_module = importlib.reload(sys.modules[module_name])
            else:
                imported_module = importlib.import_module(module_name)

            if hasattr(imported_module, "main"):
                imported_module.main()

        except Exception as e:
            print(f"⚠️ Error executing {file_name}: {e}")
            continue

        excel_after = set(glob.glob("*.xlsx"))
        new_files = list(excel_after - excel_before)

        if new_files:
            latest_excel = max(new_files, key=os.path.getmtime)
            stats = parse_all_historical_trades(latest_excel, strat_name)
            summary_reports.append(stats)
            total_market_trades += stats["Total Trades"]
            print(f"   └─ 📊 Total 10Y Trades Found: {stats['Total Trades']} (Wins: {stats['Winning Trades']}, Losses: {stats['Losing Trades']})")
        else:
            print("   └─ ⚪ No Excel file generated.")

    # ==========================================
    # 5. صياغة تقرير تليجرام الشامل
    # ==========================================
    if not summary_reports:
        msg = f"📊 <b>10-Year Backtest Scan ({today_str})</b>\n\nNo trade data generated across strategies."
        print("\n" + msg)
        send_telegram_message(msg)
        return

    msg_lines = [
        "📈 <b>10-YEAR HISTORICAL BACKTEST REPORT</b> 📈",
        f"📅 <i>Execution Date: {today_str}</i>",
        f"🌐 Total Strategies Analyzed: <b>{len(summary_reports)}</b>",
        f"🔢 Cumulative Total Trades: <b>{total_market_trades}</b>\n",
        "========================================",
    ]

    for rep in summary_reports:
        card = (
            f"🎯 <b>Strategy: {rep['Strategy']}</b>\n"
            f"📦 <b>Total Trades (10Y): {rep['Total Trades']}</b>\n"
            f"🟢 Winning Trades: {rep['Winning Trades']}\n"
            f"🔴 Losing Trades: {rep['Losing Trades']}\n"
            f"⏳ Open Positions: {rep['Open Trades']}\n"
            f"🎯 Win Rate: <b>{rep['Win Rate']}%</b>\n"
            f"💰 Cumulative PnL: <b>{rep['Total PnL %']:+.2f}%</b>\n"
            f"----------------------------------------"
        )
        msg_lines.append(card)

    final_msg = "\n".join(msg_lines)
    print("\n" + final_msg)
    send_telegram_message(final_msg)

if __name__ == "__main__":
    main()
