import os

import pandas as pd
import win32com.client as win32

# ===================== إعداد أسماء الملفات =====================
TEMPLATE_FILE = "EUS Open & Closed Tickets _12-18.02.2026.xlsx"  # ملف القالب الأصلي
OPEN_FILE = "open+call+report_test.xlsx"                         # البيانات الخام - التذاكر المفتوحة
CLOSED_FILE = "closed+call+report_test.xlsx"                     # البيانات الخام - التذاكر المغلقة
OUTPUT_FILE = "Final_Automated_Report.xlsx"                      # اسم الملف الناتج المطلوب

# اسم شيت البيانات التفصيلية داخل الـ Template التي تعتمد عليها الـ Pivots والـ Charts
# لو تعرف الاسم (مثلاً "Data") اكتبه هنا، وإلا خليه None عشان السكربت يختار أنسب شيت تلقائياً
DATA_SHEET_NAME = None

# صف الهيدر في ملف القالب (غالباً 1)
HEADER_ROW = 1
# أول صف يبدأ فيه الداتا تحت الهيدر
DATA_START_ROW = HEADER_ROW + 1

# ===================== قراءة البيانات الخام =====================
open_df = pd.read_excel(OPEN_FILE, sheet_name=0)
closed_df = pd.read_excel(CLOSED_FILE, sheet_name=0)

# دمج التذاكر المفتوحة والمغلقة
raw_combined = pd.concat([open_df, closed_df], ignore_index=True)

# ===================== إعداد الأعمدة وترتيبها =====================
"""
لو أسماء الأعمدة في الملفات الخام = نفس أسماء الأعمدة في الملف المنسق،
فالسكربت رح يلتقط الترتيب أوتوماتيكياً من القالب.

لو في اختلاف، عدّل القاموس COLUMN_MAPPING بحيث:
    المفتاح   = اسم العمود في القالب (الهيدر في الملف المنسق)
    القيمة     = اسم العمود المقابل في بيانات الـ DataFrame المدمجة (raw_combined)
"""
COLUMN_MAPPING = {
    # مثال للاستخدام (احذف المثال وعدّل حسب حاجتك):
    # "Ticket ID": "Ticket_ID",
    # "Status": "Ticket_Status",
}

# ===================== التفاعل مع ملف الـ Template عبر Excel (COM) =====================

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    # فتح ملف الـ Template الأصلي
    wb = excel.Workbooks.Open(os.path.abspath(TEMPLATE_FILE))

    # اختيار شيت البيانات الرئيسي
    if DATA_SHEET_NAME:
        ws = wb.Worksheets(DATA_SHEET_NAME)
    else:
        candidate_ws = None
        candidate_rows = -1
        for ws_iter in wb.Worksheets:
            title = ws_iter.Name.lower()
            if "pivot" in title or "chart" in title:
                continue
            rows_count = ws_iter.UsedRange.Rows.Count
            if rows_count > candidate_rows:
                candidate_rows = rows_count
                candidate_ws = ws_iter

        ws = candidate_ws if candidate_ws is not None else wb.Worksheets(1)

    # جلب عناوين الأعمدة من صف الهيدر في القالب
    template_headers = []
    col_idx = 1
    while True:
        value = ws.Cells(HEADER_ROW, col_idx).Value
        if value is None or str(value).strip() == "":
            break
        template_headers.append(str(value).strip())
        col_idx += 1

    # بناء DataFrame بنفس ترتيب أعمدة القالب
    final_data = pd.DataFrame()

    for header in template_headers:
        if header in COLUMN_MAPPING:
            source_col = COLUMN_MAPPING[header]
        else:
            source_col = header  # نفترض تطابق الاسم

        if source_col not in raw_combined.columns:
            final_data[header] = ""  # عمود غير موجود في الخام
        else:
            final_data[header] = raw_combined[source_col]

    # تحويل كل NaN إلى فراغ "" قبل الكتابة في الإكسل
    final_data = final_data.fillna("")

    # 1) مسح البيانات القديمة تحت صف الهيدر (مع الحفاظ على الهيدر، المعادلات، الرسوم، الـ pivots)
    start_col_idx = 1
    end_col_idx = len(template_headers)

    last_row = ws.Cells(ws.Rows.Count, start_col_idx).End(-4162).Row  # xlUp
    if last_row >= DATA_START_ROW:
        clear_range = ws.Range(
            ws.Cells(DATA_START_ROW, start_col_idx),
            ws.Cells(last_row, end_col_idx),
        )
        clear_range.ClearContents()

    # 2) كتابة البيانات الجديدة في نفس النطاق (نفس الأعمدة، نفس الترتيب)
    num_rows = len(final_data)
    if num_rows > 0:
        end_row = DATA_START_ROW + num_rows - 1
        write_range = ws.Range(
            ws.Cells(DATA_START_ROW, start_col_idx),
            ws.Cells(end_row, end_col_idx),
        )

        values = [
            tuple(row)
            for row in final_data[template_headers].itertuples(index=False, name=None)
        ]

        write_range.Value = values

    # 3) RefreshAll لتحديث الـ PivotTables والرسوم البيانية
    try:
        wb.RefreshAll()
    except Exception:
        pass

    # 4) حفظ كملف جديد مع الحفاظ على ملف الـ Template الأصلي
    wb.SaveAs(os.path.abspath(OUTPUT_FILE))
    wb.Close(SaveChanges=False)

finally:
    excel.Quit()

print(f"تم إنشاء الملف: {OUTPUT_FILE}")