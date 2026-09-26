# bch

## مجموعه آموزش مهارت‌های کامپیوتری

خروجی‌های طرح تفصیلی سه‌پایه‌ای:

- `output/computer-skills-curriculum-outline.docx` — نسخه قابل ویرایش
- `output/computer-skills-curriculum-outline.pdf` — نسخه آماده چاپ
- `docs/computer-skills-curriculum-outline.md` — متن منبع

بازسازی خروجی‌ها:

```bash
python3 -m pip install -r requirements-docs.txt
python3 scripts/generate_curriculum_booklet.py
```

## جزوه اجرایی Microsoft Word پایه هفتم

- `output/word-grade7/word-grade7-booklet.docx` — جزوه قابل ویرایش
- `output/word-grade7/word-grade7-booklet.pdf` — جزوه آماده چاپ
- `output/word-grade7/files/word-grade7-practice-starter.docx` — فایل تمرین خام
- `output/word-grade7/files/word-grade7-sample-project.docx` — نمونه پروژه
- `output/word-grade7/files/word-grade7-sample-project.pdf` — نسخه PDF نمونه پروژه

بازسازی جزوه:

```bash
python3 scripts/generate_word_grade7_booklet.py
```
