# سامانه مدیریت نامه‌های شرکت گام آبی فردا (هنزا)

نسخه اولیه یک سامانه فارسی برای تولید نامه‌های Word بر اساس قالب‌های `.docx`.

## اجرا در سیستم شخصی
```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
uvicorn app:app --reload
```
سپس `http://127.0.0.1:8000` را باز کنید.

## Deploy روی Render
Repository را به Render متصل کنید. `render.yaml` یا Start Command زیر را استفاده کنید:
```bash
uvicorn app:app --host 0.0.0.0 --port $PORT
```

## قالب‌ها
فایل قالب اصلی شرکت در `templates/قالب-اصلی.docx` قرار گرفته است. قالب‌های جدید از داخل پنل قابل افزودن هستند.

> نسخه فعلی MVP است. برای محیط واقعی، در مرحله بعد PostgreSQL، احراز هویت، ذخیره‌سازی پایدار فایل‌ها و موتور Template حرفه‌ای‌تر اضافه می‌شود.
