import os
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import pandas as pd


def clean_numbers(val):
    """تنسيق الأرقام لإزالة .00 للأرقام الصحيحة مع الإبقاء على الأرقام العشرية إن وجدت"""
    if pd.isna(val):
        return ""
    if isinstance(val, (int, float)):
        # إذا كان الرقم صحيحة تماماً (مثل 4.0 أو 4.00) يحوله إلى 4
        if val == int(val):
            return str(int(val))
    return str(val)


def select_files():
    files = filedialog.askopenfilenames(
        title="اختر ملفات Excel",
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
    )
    if files:
        selected_files_list.clear()
        selected_files_list.extend(files)
        lbl_status.config(
            text=f"تم تحديد {len(files)} ملف/ملفات", fg="#000000"
        )


def convert_files():
    if not selected_files_list:
        messagebox.showwarning("تنبيه", "الرجاء اختيار ملفات Excel أولاً!")
        return

    success_count = 0
    fail_count = 0

    for file_path in selected_files_list:
        try:
            base_dir = os.path.dirname(file_path)
            file_name = os.path.splitext(os.path.basename(file_path))[0]
            output_txt_path = os.path.join(base_dir, f"{file_name}.txt")

            # قراءة ملف الإكسل
            df = pd.read_excel(file_path)

            # 1. تنظيف الأرقام لمنع ظهور .00 بالأرقام الصحيحة
            df = df.applymap(clean_numbers)

            # 2. الحفظ بترميز UTF-16 (Unicode التقليدي المطابق لإكسل والأمين)
            df.to_csv(
                output_txt_path, sep="\t", index=False, encoding="utf-16"
            )

            success_count += 1
        except Exception as e:
            fail_count += 1

    messagebox.showinfo(
        "تمت العملية",
        f"تم تحويل {success_count} ملف بنجاح!\nفشل: {fail_count}",
    )
    lbl_status.config(text="اكتملت عملية التحويل", fg="green")


# إعداد النافذة الرئيسية
root = tk.Tk()
root.title("محول ملفات Excel إلى TXT (Unicode)")
root.geometry("450x300")
root.resizable(False, False)

selected_files_list = []

lbl_title = tk.Label(
    root,
    text="محول ملفات Excel إلى TXT (Unicode)",
    font=("Arial", 14, "bold"),
    bg="#2B579A",
    fg="white",
    pady=10,
)
lbl_title.pack(fill=tk.X)

btn_select = tk.Button(
    root,
    text="اختر ملفات Excel (واحدة أو أكثر)",
    font=("Arial", 11, "bold"),
    bg="#28a745",
    fg="white",
    command=select_files,
    padx=10,
    pady=5,
)
btn_select.pack(pady=20)

lbl_status = tk.Label(
    root, text="لم يتم اختيار أي ملف", font=("Arial", 10), fg="gray"
)
lbl_status.pack(pady=5)

btn_convert = tk.Button(
    root,
    text="تحويل الملفات إلى TXT (Unicode)",
    font=("Arial", 11, "bold"),
    bg="#007bff",
    fg="white",
    command=convert_files,
    padx=10,
    pady=5,
)
btn_convert.pack(pady=15)

root.mainloop()
