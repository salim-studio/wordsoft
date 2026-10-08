# WordSoft 1.0 — معالج نصوص عربي بأسلوب Word

مستوحى من فكرة [wordcraft](https://github.com/storytold/wordcraft) (إعادة بناء Word مفتوحة المصدر بلغة Rust)،
لكن مكتوب بـ **Python + Tkinter ليعمل فوراً** بدون تجميع، وبواجهة عربية داكنة عملية.

## التشغيل (Windows)

```bat
cd Wordsoft
python app.py
```

أو:

```bat
python -m Wordsoft
```

أو نقر مزدوج على `run.bat`.

المتطلبات: `pip install -r requirements.txt`

## المزايا

- **الشريط**: الرئيسية / إدراج / تخطيط / مراجعة / عرض — مثل Word
- **كتابة**: تراجع/إعادة، استمرار القوائم عند Enter، تصحيح تلقائي، علامات ذكية، نسخ تنسيق، بحث واستبدال
- **تنسيق**: خطوط وأحجام، عريض/مائل/تحته خط/مشطوب، ألوان وتمييز، محاذاة، RTL، مسافة بادئة
- **أنماط**: عادي، عنوان، عنوان 1/2/3، اقتباس، نقطية، مرقمة + جزء تنقل للعناوين
- **جداول**: إدراج n×m يُحفظ في DOCX/PDF/HTML
- **صور**: إدراج وتقليص تلقائي
- **مراجعة**: تعداد كلمات، تدقيق F7، تعقب تغييرات، تعليقات
- **ملفات**: فتح/حفظ DOCX يعمل مع Word + تصدير PDF (نص حقيقي قابل للتحديد) و HTML/MD/TXT
- **قوالب**: فارغ، عينة الاستوديو، رسالة، تقرير، سيرة ذاتية
- **عرض**: زوم، داكن/فاتح، تركيز، تنقل

## CLI (مثل wordcraft-cli)

```bat
python cli.py convert report.docx report.pdf
python cli.py text report.docx
python cli.py stats report.docx
python cli.py inspect report.docx
python cli.py new --template sample --save out.docx
```

## البنية

```
Wordsoft/
  engine.py   # نموذج المستند: فقرات + Runs + جداول + عينة وقوالب (طبقة L1)
  formats.py  # فتح/حفظ docx/txt/md/html/pdf (طبقة L2-L3)
  app.py      # الواجهة بأسلوب Word (طبقة L6)
  cli.py      # أداة سطر الأوامر والأتمتة (مثل wordcraft-cli)
  api/index.py# نسخة الويب (Flask)
```

## مقارنة مع wordcraft

| | wordcraft | WordSoft |
|---|---|---|
| اللغة | Rust + egui GPU | Python + Tkinter |
| التثبيت | cargo / مثبتات | `python app.py` فوراً |
| DOCX/PDF | كامل OOXML + krilla | أساسي عملي (python-docx + reportlab) |
| المزايا | ~389 أمراً + MCP | ~40 أمراً أساسياً + CLI |
| الهدف | بديل Word كامل | أداة عربية خفيفة وسريعة |

## نسخة الويب

```
cd Wordsoft
python api/index.py
```

ثم افتح http://127.0.0.1:5000 — محرر كامل: تنسيق، عناوين، قوائم، جداول، رفع/تنزيل DOCX، تصدير PDF/HTML.
