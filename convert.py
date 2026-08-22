import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl

class ExcelToTextConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("محول ملفات إكسل إلى نص Unicode (.txt)")
        self.root.geometry("500x320")
        self.root.resizable(False, False)
        self.setup_ui()

    def setup_ui(self):
        header_frame = tk.Frame(self.root, bg="#2b579a", height=60)
        header_frame.pack(fill=tk.X)
        lbl_title = tk.Label(header_frame, text="محول ملفات Excel إلى TXT (Unicode)", fg="white", bg="#2b579a", font=("Segoe UI", 14, "bold"))
        lbl_title.pack(pady=15)

        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        self.lbl_file = tk.Label(main_frame, text="لم يتم اختيار أي ملف إكسل بعد", font=("Segoe UI", 10), fg="#555555", wraplength=440, justify="center")
        self.lbl_file.pack(pady=(10, 15))

        btn_select = tk.Button(main_frame, text="اختيار ملف Excel (.xlsx / .xls)", command=self.select_file, font=("Segoe UI", 10, "bold"), bg="#107c41", fg="white", padx=15, pady=8, relief=tk.FLAT, cursor="hand2")
        btn_select.pack()

        self.btn_convert = tk.Button(main_frame, text="تحويل إلى ملف نصي TXT واحد (Unicode)", command=self.convert_file, font=("Segoe UI", 11, "bold"), bg="#2b579a", fg="white", padx=15, pady=8, relief=tk.FLAT, state=tk.DISABLED, cursor="hand2")
        self.btn_convert.pack(pady=(20, 0))

        self.selected_file_path = None

    def select_file(self):
        file_path = filedialog.askopenfilename(title="اختر ملف إكسل", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            self.selected_file_path = file_path
            self.lbl_file.config(text=f"الملف المحدد: {os.path.basename(file_path)}", fg="#000000")
            self.btn_convert.config(state=tk.NORMAL, bg="#2b579a")

    def convert_file(self):
        if not self.selected_file_path:
            return
        try:
            # فتح ملف الإكسل وقراءة ورقة العمل الأولى النشطة فقط بنفس طريقة إكسل اليدوية
            wb = openpyxl.load_workbook(self.selected_file_path, data_only=True)
            sheet = wb.active

            # اسم الملف الناتج سيكون بالضبط بنفس اسم ملف الإكسل الاصلي (مثال: البرامكة.txt)
            base_name = os.path.splitext(self.selected_file_path)[0]
            output_path = f"{base_name}.txt"

            with open(output_path, "w", encoding="utf-16", newline="") as f:
                for row in sheet.iter_rows(values_only=True):
                    # تجاهل الصفوف الفارغة بالكامل
                    if all(cell is None for cell in row):
                        continue
                    # تحويل القيم إلى نصوص ومفصولة بـ Tab (	) بنفس تنسيق Excel بالضبط
                    row_str = "	".join([str(cell) if cell is not None else "" for cell in row])
                    f.write(row_str + "
")

            messagebox.showinfo("تم بنجاح", f"تم إنشاء ملف نصي واحد بنجاح:
{os.path.basename(output_path)}")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء التحويل:
{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToTextConverterApp(root)
    root.mainloop()
