# لوحة تسجيل الطلاب – School Registration Dashboard

## تشغيل التطبيق (للمستخدمين)

### الخطوات:

1. **تثبيت Python** من https://www.python.org/downloads/ (إذا لم يكن مثبتاً)

2. **فتح Terminal أو Command Prompt** في مجلد التطبيق

3. **تثبيت المتطلبات** (مرة واحدة فقط):
   ```
   pip install -r requirements.txt
   ```

4. **تشغيل التطبيق**:
   ```
   streamlit run app.py
   ```

5. سيفتح المتصفح تلقائياً على العنوان: http://localhost:8501

6. **ارفع ملف Excel** من الشريط الجانبي وابدأ الاستخدام.

---

## How to Run (English)

1. Install Python from https://www.python.org/downloads/
2. Open Terminal/Command Prompt in this folder
3. Run once: `pip install -r requirements.txt`
4. Run the app: `streamlit run app.py`
5. Browser opens at http://localhost:8501
6. Upload your Excel file from the sidebar.

---

## ملاحظات

- يمكن رفع نسخة جديدة من الملف في أي وقت دون إعادة تشغيل التطبيق.
- لا حاجة لأي معرفة برمجية لاستخدام التطبيق.
- إذا تغيرت أسماء الأعمدة في Excel، يمكن تعديلها في قسم `COL` في بداية ملف `app.py`.
