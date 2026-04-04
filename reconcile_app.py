import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from reconcile_core import load_xlsx, load_csv, reconcile, get_summary_text, export_report


class ReconcileApp:
    def __init__(self, root):
        self.root = root
        self.root.title('每日發票對帳工具')
        self.root.geometry('620x520')
        self.root.resizable(False, False)
        self.result = None

        # 檔案選擇區
        frame_files = tk.LabelFrame(root, text='檔案選擇', padx=10, pady=10)
        frame_files.pack(fill='x', padx=10, pady=(10, 5))

        tk.Label(frame_files, text='發票檔案 (XLSX)：').grid(row=0, column=0, sticky='w')
        self.xlsx_var = tk.StringVar()
        tk.Entry(frame_files, textvariable=self.xlsx_var, width=45).grid(row=0, column=1, padx=5)
        tk.Button(frame_files, text='瀏覽', command=self.browse_xlsx).grid(row=0, column=2)

        tk.Label(frame_files, text='交易檔案 (CSV)：').grid(row=1, column=0, sticky='w', pady=(5, 0))
        self.csv_var = tk.StringVar()
        tk.Entry(frame_files, textvariable=self.csv_var, width=45).grid(row=1, column=1, padx=5, pady=(5, 0))
        tk.Button(frame_files, text='瀏覽', command=self.browse_csv).grid(row=1, column=2, pady=(5, 0))

        # 按鈕區
        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=10)
        tk.Button(frame_btns, text='開始對帳', command=self.run_reconcile,
                  bg='#4472C4', fg='white', font=('Arial', 12, 'bold'),
                  padx=20, pady=5).pack(side='left', padx=10)
        self.export_btn = tk.Button(frame_btns, text='匯出報表', command=self.export,
                                    state='disabled', font=('Arial', 12),
                                    padx=20, pady=5)
        self.export_btn.pack(side='left', padx=10)

        # 結果顯示區
        frame_result = tk.LabelFrame(root, text='對帳結果', padx=10, pady=10)
        frame_result.pack(fill='both', expand=True, padx=10, pady=(5, 10))
        self.result_text = scrolledtext.ScrolledText(frame_result, height=15, font=('Courier', 11))
        self.result_text.pack(fill='both', expand=True)

    def browse_xlsx(self):
        path = filedialog.askopenfilename(
            title='選擇發票檔案',
            filetypes=[('Excel 檔案', '*.xlsx'), ('所有檔案', '*.*')]
        )
        if path:
            self.xlsx_var.set(path)

    def browse_csv(self):
        path = filedialog.askopenfilename(
            title='選擇交易檔案',
            filetypes=[('CSV 檔案', '*.csv'), ('所有檔案', '*.*')]
        )
        if path:
            self.csv_var.set(path)

    def run_reconcile(self):
        xlsx_path = self.xlsx_var.get().strip()
        csv_path = self.csv_var.get().strip()

        if not xlsx_path or not csv_path:
            messagebox.showwarning('提示', '請先選擇兩個檔案')
            return

        if not os.path.exists(xlsx_path):
            messagebox.showerror('錯誤', f'找不到發票檔案：\n{xlsx_path}')
            return
        if not os.path.exists(csv_path):
            messagebox.showerror('錯誤', f'找不到交易檔案：\n{csv_path}')
            return

        try:
            xlsx_df = load_xlsx(xlsx_path)
            csv_df = load_csv(csv_path)
            self.result = reconcile(xlsx_df, csv_df)
            summary = get_summary_text(self.result)

            self.result_text.delete('1.0', tk.END)
            self.result_text.insert(tk.END, summary)
            self.export_btn.config(state='normal')
        except Exception as e:
            messagebox.showerror('錯誤', f'對帳過程發生錯誤：\n{str(e)}')

    def export(self):
        if not self.result:
            messagebox.showwarning('提示', '請先執行對帳')
            return

        path = filedialog.asksaveasfilename(
            title='儲存對帳報表',
            defaultextension='.xlsx',
            filetypes=[('Excel 檔案', '*.xlsx')],
            initialfile='對帳報表.xlsx'
        )
        if path:
            try:
                export_report(self.result, path)
                messagebox.showinfo('完成', f'報表已儲存至：\n{path}')
            except Exception as e:
                messagebox.showerror('錯誤', f'匯出失敗：\n{str(e)}')


if __name__ == '__main__':
    root = tk.Tk()
    app = ReconcileApp(root)
    root.mainloop()
