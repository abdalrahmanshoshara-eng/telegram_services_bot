# بوت خدمات تيليغرام

بوت خدمات احترافي يدعم الخدمات المرتبطة بمنصات خارجية ومعالجة الملفات محليًا. يختار المستخدم الخدمة بالزر أو بإرسال رقمها، وتُحذف ملفات العمل المؤقتة بعد إرسال النتيجة.

## الخدمات

1. منصة إعداد العقود.
2. منصة حفظ الكروت الشخصية.
3. تنسيق جداول Word ومنع انقسام الصفوف.
4. إزالة خلفية الصور باستخدام النموذج الخفيف `u2netp`.
5. دمج عدة ملفات PDF حسب ترتيب الإرسال.
6. استخراج صفحات ونطاقات من PDF، مثل `1-3,5,8`، مع دعم الأرقام العربية.
7. تحويل عدة صور إلى PDF.
8. ضغط الصور وتحويلها إلى JPG أو PNG أو WEBP وتصغيرها.
9. إضافة علامة مائية نصية عربية أو صورة شعار إلى PDF.
10. إنشاء QR Code من رابط أو نص.
11. استخراج النص العربي والإنجليزي من الصور إلى TXT.

الأوامر: `/start` و`/services` و`/done` و`/cancel`.

## التشغيل باستخدام Docker - الطريقة الموصى بها

المتطلبات:

- Docker Engine مع Docker Compose.
- خادوم موصى به: 2 vCPU و4GB RAM و20GB مساحة.
- اتصال بالإنترنت في أول تشغيل لتنزيل نموذج `u2netp` الصغير.

انسخ ملف الإعداد:

```bash
cp .env.example .env
nano .env
```

ضع التوكن الجديد في:

```dotenv
BOT_TOKEN=ضع_التوكن_الجديد_هنا
```

لا ترفع ملف `.env` إلى GitHub. شغّل البوت:

```bash
docker compose up -d --build
docker compose ps
docker compose logs -f
```

في أول تشغيل يجهّز الـentrypoint نموذج إزالة الخلفية داخل volume دائم، ولن يعيد تنزيله بعد إعادة إنشاء الحاوية.

للتحديث من GitHub:

```bash
git pull --ff-only
docker compose up -d --build
docker image prune -f
```

لإيقاف البوت دون حذف النموذج:

```bash
docker compose down
```

لا تستخدم `docker compose down -v` إلا إذا أردت حذف volume النموذج وإعادة تنزيله.

## متغيرات البيئة

| المتغير | الافتراضي | الاستخدام |
|---|---:|---|
| `BOT_TOKEN` | مطلوب | توكن BotFather |
| `MAX_UPLOAD_MB` | 20 | الحد الأقصى لكل ملف |
| `MAX_OUTPUT_MB` | 45 | الحد الأقصى للملف الناتج |
| `MAX_MULTI_FILES` | 10 | عدد الملفات في الدمج أو صور إلى PDF |
| `MAX_TOTAL_UPLOAD_MB` | 100 | مجموع أحجام عملية متعددة الملفات |
| `MAX_IMAGE_PIXELS` | 40000000 | أقصى مجموع بكسلات للصورة لمنع استنزاف الذاكرة |
| `MAX_PDF_PAGES` | 500 | أقصى عدد صفحات لعملية PDF واحدة |
| `MAX_CONCURRENT_JOBS` | 2 | عدد المعالجات المتزامنة |
| `PROCESSING_TIMEOUT_SECONDS` | 300 | مهلة العملية الواحدة |
| `OCR_LANG` | `ara+eng` | لغات Tesseract |
| `REMBG_MODEL` | `u2netp` | نموذج إزالة الخلفية |
| `NUMBA_CACHE_DIR` | `/tmp/numba` | كاش Numba داخل المسار القابل للكتابة في الحاوية |
| `XDG_CACHE_HOME` | `/tmp/cache` | كاش مكتبات ONNX والنظام |
| `BOT_MEMORY_LIMIT` | `4g` | حد ذاكرة الحاوية |
| `BOT_CPU_LIMIT` | `2.0` | حد أنوية CPU للحاوية |

## التشغيل دون Docker

على Ubuntu/Debian:

```bash
sudo apt update
sudo apt install -y \
  python3 python3-venv fonts-dejavu-core libgl1 libglib2.0-0 \
  libmagic1 tesseract-ocr tesseract-ocr-ara tesseract-ocr-eng

python3 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

export BOT_TOKEN="ضع_التوكن_الجديد_هنا"
export REMBG_MODEL="u2netp"
export REMBG_HOME="$PWD/.rembg"
.venv/bin/python -m services.background_remover --prepare
.venv/bin/python bot.py
```

يوجد مثال خدمة systemd في `deploy/telegram-services-bot.service`، لكن Docker Compose هو المسار الأسهل لتثبيت Tesseract وبقية الاعتماديات بنفس الإصدارات في كل خادوم.

## الاختبارات

```bash
python3 -m venv .test-venv
.test-venv/bin/pip install -r requirements-dev.txt
PYTHONPATH=. .test-venv/bin/pytest -q
```

تشغيل OCR الفعلي للصور يحتاج Tesseract واللغات المذكورة أعلاه. الاختبارات لا تنزّل نموذج إزالة الخلفية؛ يتم استبداله بمحاكاة محلية أثناء الاختبار.

## إضافة خدمة جديدة

1. أضف تعريف الخدمة إلى `SERVICES` في `config.py`.
2. أضف المعالج إلى مجلد `services/`.
3. اربط الخدمة بحالة المحادثة المناسبة في `bot.py`.
4. أضف اختبارًا يغطي ملفًا صحيحًا ومدخلًا غير صالح.

## الأمان والخصوصية

- لا يوجد توكن داخل الكود، وملف `.env` مستبعد من Git وDocker build context.
- الحاوية تعمل بمستخدم غير root وبنظام ملفات للقراءة فقط، ولا يُسمح لها باكتساب صلاحيات جديدة.
- ملفات المستخدم تُنزّل إلى `/tmp` وتُحذف تلقائيًا بعد المعالجة.
- في العمليات متعددة الخطوات يحتفظ البوت بمعرفات Telegram فقط إلى أن يضغط المستخدم «انتهيت».
- توجد حدود لكل ملف، وإجمالي العملية، وعدد الملفات، وحجم النتيجة، والوقت، والتزامن.
- ملفات PDF المشفرة تُرفض بدل محاولة تجاوز الحماية.

> التوكن الذي كان موجودًا في النسخة القديمة مكشوف؛ يجب إلغاؤه من BotFather واستخدام توكن جديد.

## حل خطأ Numba داخل Docker

إذا ظهر الخطأ `cannot cache function ... no locator available` فهذا يعني أن Numba حاولت الكتابة داخل نظام الملفات المقروء فقط. النسخة الحالية توجه الكاش إلى `/tmp/numba` وتسمح للبوت بالبدء حتى لو تعذر تجهيز نموذج إزالة الخلفية.

بعد تحديث الملفات نفذ:

```bash
docker compose down
docker compose build --no-cache
docker compose up -d
docker compose ps
docker compose logs -f --tail=100
```

لا حاجة إلى حذف volume `rembg-models`.
# telegram_services_bot
