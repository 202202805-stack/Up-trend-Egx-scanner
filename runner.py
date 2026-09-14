import os
import requests

# 1. بيانات تليجرام (استبدل القيم بالتالي الخاصة بك)
TELEGRAM_TOKEN = "YOUR_BOT_TOKEN_HERE"
CHAT_ID = "YOUR_CHAT_ID_HERE"

# 2. قائمة أسماء الملفات الـ 7 بالظبط كما في الصورة
strategy_files = [
    "BROADING BOTTOMS.txt",
    "FLAGS.txt",
    "adam and adam.txt",
    "adam and eva.txt",
    "asending trainangle.txt",
    "pipe bottom.txt",
    "tripple bottom.txt"
]

all_open_trades = []

# 3. تشغيل كل ملف بالتتابع واستخراج الصفقات
for file_name in strategy_files:
    if os.path.exists(file_name):
        print(f"Running strategy: {file_name}...")
        
        # تجهيز نطاق للمتغيرات لاستلام النتيجة من الملف
        local_scope = {}
        
        with open(file_name, "r", encoding="utf-8") as f:
            code = f.read()
            # تنفيذ كود الملف
            exec(code, globals(), local_scope)
        
        # نفترض أن كل ملف بيخزن الصفقات المفتوحة في متغير اسمه open_trades
        # أو يمكنك تعديل اسم المتغير حسب المكتوب داخل ملفاتك
        trades = local_scope.get("open_trades", [])
        
        if trades:
            all_open_trades.append(f"📌 **{file_name.replace('.txt', '')}**:")
            for trade in trades:
                all_open_trades.append(f" - {trade}")
    else:
        print(f"Warning: File {file_name} not found!")

# 4. تجميع الرسالة وإرسالها لتليجرام
if all_open_trades:
    message = "🚀 **تقرير الصفقات المفتوحة الجديد:**\n\n" + "\n".join(all_open_trades)
else:
    message = "ℹ️ لا توجد صفقات مفتوحة حالياً في جميع النماذج."

def send_telegram_message(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    payload = {
        "chat_id": CHAT_ID,
        "text": text,
        "parse_mode": "Markdown"
    }
    requests.post(url, json=payload)

send_telegram_message(message)
print("Done! Report sent to Telegram.")
