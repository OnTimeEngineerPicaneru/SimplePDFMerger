import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from PyPDF2 import PdfMerger
import os
import threading

MAX_FILES = 15


class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF結合アプリ（ドラッグ＆ドロップ対応）")
        self.pdf_files = []

        self.setup_ui()

    def setup_ui(self):
        self.instructions = tk.Label(
            self.root,
            text=f"ここにPDFファイルをドラッグ＆ドロップ（最大{MAX_FILES}枚）",
            bg="lightgray",
            height=3,
        )
        self.instructions.pack(fill="x", padx=10, pady=10)

        self.file_listbox = tk.Listbox(self.root, height=8, width=60)
        self.file_listbox.pack(padx=10)

        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.drop_files)

        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)

        tk.Button(button_frame, text="↑ 上へ", command=self.move_up).pack(
            side="left", padx=5
        )
        tk.Button(button_frame, text="↓ 下へ", command=self.move_down).pack(
            side="left", padx=5
        )
        tk.Button(button_frame, text="削除", command=self.delete_selected).pack(
            side="left", padx=5
        )
        tk.Button(
            button_frame, text="PDFを結合", bg="lightblue", command=self.merge_pdfs
        ).pack(side="left", padx=5)

        tk.Button(
            self.root, text="ファイルを参照して追加", command=self.browse_files
        ).pack(pady=5)

    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            if f.lower().endswith(".pdf"):
                self.add_pdf(f)

    def browse_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDFファイル", "*.pdf")])
        for path in paths:
            self.add_pdf(path)

    def add_pdf(self, path):
        if path not in self.pdf_files:
            if len(self.pdf_files) >= MAX_FILES:
                messagebox.showwarning(
                    "制限", f"最大{MAX_FILES}ファイルまでしか追加できません。"
                )
                return
            self.pdf_files.append(path)
            self.update_listbox()

    def update_listbox(self):
        self.file_listbox.delete(0, tk.END)
        for f in self.pdf_files:
            self.file_listbox.insert(tk.END, os.path.basename(f))

    def move_up(self):
        index = self.file_listbox.curselection()
        if index and index[0] > 0:
            i = index[0]
            self.pdf_files[i - 1], self.pdf_files[i] = (
                self.pdf_files[i],
                self.pdf_files[i - 1],
            )
            self.update_listbox()
            self.file_listbox.select_set(i - 1)

    def move_down(self):
        index = self.file_listbox.curselection()
        if index and index[0] < len(self.pdf_files) - 1:
            i = index[0]
            self.pdf_files[i + 1], self.pdf_files[i] = (
                self.pdf_files[i],
                self.pdf_files[i + 1],
            )
            self.update_listbox()
            self.file_listbox.select_set(i + 1)

    def delete_selected(self):
        index = self.file_listbox.curselection()
        if index:
            del self.pdf_files[index[0]]
            self.update_listbox()

    def merge_pdfs(self):
        if not self.pdf_files:
            messagebox.showwarning("警告", "PDFファイルを追加してください。")
            return

        try:

            def task():
                merger = PdfMerger()
                for pdf in self.pdf_files:
                    merger.append(pdf)

                save_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf", filetypes=[("PDFファイル", "*.pdf")]
                )
                if save_path:
                    merger.write(save_path)
                    merger.close()
                    messagebox.showinfo("成功", f"PDFを保存しました：\n{save_path}")

            pdf_task = threading.Thread(target=task, daemon=True)
            pdf_task.start()
        except Exception as e:
            messagebox.showerror("エラー", f"結合中にエラーが発生しました:\n{str(e)}")


# 実行部：TkinterDnDのインスタンスで初期化
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
