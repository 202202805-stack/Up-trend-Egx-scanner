import datetime
import json
import os
import re
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
    """إرسال الرسالة إلى التليجرام"""
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        print("⚠️ Telegram Secrets missing.")
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
            print(f"⚠️ Telegram Exception: {e}")


def find_actual_file(target_filename: str) -> str:
    """معالجة اختلاف حالة الأحرف في نظام Linux"""
    if os.path.exists(target_filename):
        return target_filename
    for f in os.listdir("."):
        if f.lower() == target_filename.lower():
            return f
    return target_filename


def process_and_save_strategy(filename: str, strategy_name: str):
    """تشغيل كل استراتيجية بشكل منفصل واستخراج صفقاتها في ملف JSON محلي"""
    actual_filename = find_actual_file(filename)
    json_output_name = (
        f"output_{re.sub(r'[^a-zA-Z0-9]', '_', strategy_name).lower()}.json"
    )

    if not os.path.exists(actual_filename):
        print(f"⚠️ [Missing File]: {filename}")
        return

    print(f"\n🔍 Executing: {actual_filename} ({strategy_name})...")

    scope = {}
    try:
        with open(actual_filename, "r", encoding="utf-8") as f:
            code = f.read()

        # تنفيذ الكود في بيئة مستقلة
        exec(code, scope)

        found_dfs = []

        # 1. البحث في الذاكرة عن أي DataFrame
        for var_name, var_val in scope.items():
            if not var_name.startswith("__") and isinstance(
                var_val, pd.DataFrame
            ):
                if not var_val.empty:
                    found_dfs.append(var_val.copy())

        # 2. قراءة أي ملف Excel جُدَّدَ على القرص
        for file in os.listdir("."):
            if file.endswith(".xlsx") and not file.startswith("~$"):
                try:
                    df = pd.read_excel(file)
                    if not df.empty:
                        found_dfs.append(df)
                except Exception:
                    pass

        if not found_dfs:
            print(f"  └─ ⚠️ No data extracted for {strategy_name}.")
            return

        # دمج البيانات المكتشفة
        final_df = pd.concat(found_dfs, ignore_index=True).drop_duplicates()

        # تنظيف أسماء الأعمدة (حذف المسافات وتحويلها لـ Title Case)
        final_df.columns = [
            str(c).strip().replace("_", " ").title() for c in final_df.columns
        ]

        # البحث عن عمود حالة الصفقة (State / Status)
        state_col = None
        for col in ["State", "Status", "Trade State", "Trade Status"]:
            if col in final_df.columns:
                state_col = col
                break

        open_trades = []

        # الفلترة الدقيقة للصفقات التي حالتها OPEN فقط
        for _, row in final_df.iterrows():
            is_open = False

            if state_col:
                val = str(row.get(state_col, "")).strip().lower()
                # الاعتماد المباشر على مطابقة open وتجاهل win / loss
                if "open" in val or val == "active":
                    is_open = True
            else:
                # إذا لم يوجد عمود state نعتبر الصفقة مفتوحة بشكل افتراضي
                is_open = True

            if is_open:
                # استخراج القيم الأساسية
                def get_val(keys, default="N/A"):
                    for k in keys:
                        for c in final_df.columns:
                            if k.lower() == c.lower():
                                return row[c]
                    return default

                stock = get_val(
                    ["Stock Name", "Ticker", "Stock", "Symbol"], "N/A"
                )
                entry_d = get_val(["Entry Date", "Date", "Entry_Date"], "N/A")
                entry_p = get_val(
                    ["Entry Price", "Buy Price", "Price", "Entry_Price"], 0.0
                )
                curr_p = get_val(
                    ["Current Price", "Last Price", "Close"], entry_p
                )
                target = get_val(["Target", "Target Price"], 0.0)
                stop = get_val(["Stop Loss", "Stop"], 0.0)

                # حساب نسبة الربح/الخسارة PnL
                try:
                    ep, cp = float(entry_p), float(curr_p)
                    pnl = ((cp - ep) / ep) * 100 if ep > 0 else 0.0
                except Exception:
                    pnl = 0.0

                open_trades.append({
                    "Strategy": strategy_name,
                    "Stock": str(stock).replace(".CA", ""),
                    "Entry Date": str(entry_d)[:10],
                    "Entry Price": (
                        round(float(entry_p), 3) if entry_p != "N/A" else 0.0
                    ),
                    "Current Price": (
                        round(float(curr_p), 3) if curr_p != "N/A" else 0.0
                    ),
                    "Target": round(float(target), 3) if target != "N/A" else 0.0,
                    "Stop Loss": (
                        round(float(stop), 3) if stop != "N/A" else 0.0
                    ),
                    "PnL": round(pnl, 2),
                })

        # حفظ النتيجة في ملف JSON مستقل لكل استراتيجية
        with open(json_output_name, "w", encoding="utf-8") as jf:
            json.dump(open_trades, jf, ensure_ascii=False, indent=2)

        print(
            f"  └─ 🟢 Successfully saved {len(open_trades)} OPEN trade(s) to"
            f" {json_output_name}"
        )

    except Exception as e:
        print(f"  └─ ❌ Error processing {filename}: {e}")


def aggregate_and_notify():
    """تجميع كافة الصفقات من ملفات الـ JSON وإرسال إشعار موحد للتليجرام"""
    all_signals = []

    # قراءة كل ملفات الـ JSON الناتجة
    for file in os.listdir("."):
        if file.startswith("output_") and file.endswith(".json"):
            try:
                with open(file, "r", encoding="utf-8") as jf:
                    data = json.load(jf)
                    all_signals.extend(data)
                os.remove(file)  # تنظيف الملفات المؤقتة بعد القراءة
            except Exception as e:
                print(f"⚠️ Error reading {file}: {e}")

    today_str = datetime.date.today().strftime("%Y-%m-%d")

    if not all_signals:
        msg = f"📊 <b>EGX Market Scan ({today_str})</b>\n\nNo active OPEN signals found today."
        print("\n" + msg)
        send_telegram_message(msg)
        return

    # صياغة رسالة التليجرام المجمعة
    msg_lines = [
        "🚨 <b>EGX ALL STRATEGIES - ACTIVE SIGNALS</b> 🚨",
        f"📅 <i>Date: {today_str}</i>",
        f"🌐 Total Active Signals: <b>{len(all_signals)}</b>\n",
        "========================================",
    ]

    for sig in all_signals:
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


def main():
    print("🚀 Starting EGX Multi-Strategy Execution & JSON Logging...\n")

    # 1. تشغيل الفحص واستخراج الصفقات لكل استراتيجية على حدة
    for filename, strat_name in TXT_FILES:
        process_and_save_strategy(filename, strat_name)

    # 2. تجميع كل النتائج وإرسال التنبيه
    print("\n📦 Aggregating all results and sending notification...")
    aggregate_and_notify()


if __name__ == "__main__":
    main()
