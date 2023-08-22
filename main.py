import os
import re
from docx import Document
import tkinter as tk
from tkinter import filedialog
from tkcalendar import Calendar  # 導入 Calendar 模組
from datetime import datetime, timedelta


class DocxToMdConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Docx to Md Converter")

        self.input_folder_label = tk.Label(root, text="Input Folder:")  # 輸入資料夾標籤
        self.input_folder_label.pack()
        self.input_folder_entry = tk.Entry(root, width=50)  # 輸入資料夾輸入框
        self.input_folder_entry.pack()
        # 瀏覽按鈕，連結到 browse_input_folder 函數
        self.input_folder_button = tk.Button(root, text="Browse", command=self.browse_input_folder)
        self.input_folder_button.pack()

        self.output_folder_label = tk.Label(root, text="Output Folder:")  # 輸出資料夾標籤
        self.output_folder_label.pack()
        self.output_folder_entry = tk.Entry(root, width=50)  # 輸出資料夾輸入框
        self.output_folder_entry.pack()
        # 瀏覽按鈕，連結到 browse_output_folder 函數
        self.output_folder_button = tk.Button(root, text="Browse", command=self.browse_output_folder)
        self.output_folder_button.pack()

        self.start_date_label = tk.Label(root, text="Start Date:")  # 開始日期標籤
        self.start_date_label.pack()

        self.start_date_calendar = Calendar(root, selectmode="day")  # 選擇日期的日曆部件
        self.start_date_calendar.pack()

        self.categories_label = tk.Label(root, text="Categories:")  # 類別標籤
        self.categories_label.pack()

        self.categories_listbox = tk.Listbox(root, selectmode=tk.MULTIPLE)  # 使用多選模式的列表框
        categories = ["JAVA基本語法", "Java物件導向", "Java其他概念", "Java執行原理", "Java常用API", "JDK版本特性"]
        for category in categories:
            self.categories_listbox.insert(tk.END, category)
        self.categories_listbox.pack()

        # 轉換按鈕，連結到 start_conversion 函數
        self.convert_button = tk.Button(root, text="Convert", command=self.start_conversion)
        self.convert_button.pack()

    def browse_input_folder(self):
        folder_path = filedialog.askdirectory()
        self.input_folder_entry.delete(0, tk.END)
        self.input_folder_entry.insert(0, folder_path)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory()
        self.output_folder_entry.delete(0, tk.END)
        self.output_folder_entry.insert(0, folder_path)

    def convert_docx_to_md(self, docx_filename, start_date):
        doc = Document(docx_filename)
        text = ""

        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"

        text = re.sub(r'\n', '  \n', text)

        if not text.strip():
            text = "(This document is empty)"

        md_filename = os.path.splitext(os.path.basename(docx_filename))[0] + ".md"
        md_title = os.path.splitext(os.path.basename(docx_filename))[0]
        md_dir = os.path.dirname(docx_filename)
        md_path = os.path.join(md_dir, md_filename)

        selected_categories = self.categories_listbox.curselection()  # 獲取使用者選擇的類別
        selected_categories = [self.categories_listbox.get(index) for index in selected_categories]

        parent_folder_names = os.path.basename(md_dir)
        tag_name = re.sub(r'^\d+\s+', '', parent_folder_names)

        with open(md_path, 'w', encoding='utf-8') as md_file:
            md_file.write("---\n")
            md_file.write(f"title: {md_title}\n")
            md_file.write(f"date: {start_date.strftime('%Y-%m-%d %H:%M:%S')}\n")
            md_file.write("categories:\n")
            for category in selected_categories:  # 寫入使用者選擇的類別
                md_file.write(f"- {category}\n")
            md_file.write("tags:\n")
            md_file.write(f"- {tag_name}\n")  # 添加父文件夾名稱作為 tag
            md_file.write("---\n\n")
            md_file.write(text)

    def batch_convert(self, folder_path):
        selected_date = self.start_date_calendar.get_date()  # 獲取使用者選擇的日期
        selected_date = datetime.strptime(selected_date, "%m/%d/%y").date()  # 將字串日期轉換為 date 類型

        docx_files = []

        for root_folder, _, filenames in os.walk(folder_path):
            for filename in filenames:
                if filename.endswith('.docx'):
                    docx_files.append(os.path.join(root_folder, filename))

        start_date = datetime.combine(selected_date, datetime.min.time())
        file_count = 0

        for docx_path in docx_files:
            self.convert_docx_to_md(docx_path, start_date)
            print(f"Converted '{docx_path}'")

            file_count += 1
            if file_count % 5 == 0:
                start_date += timedelta(days=1)

    def start_conversion(self):
        input_folder = self.input_folder_entry.get()
        output_folder = self.output_folder_entry.get()

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        self.batch_convert(input_folder)


if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToMdConverterApp(root)
    root.mainloop()
