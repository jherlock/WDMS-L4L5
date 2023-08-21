import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog, messagebox
import pandas as pd
import os
from datetime import datetime

# 文件名
EXCEL_FILE = "Train_Wheel.xlsx"

# 列名
COLUMNS = ['地铁线路', '车号', '车轮直径', '轮径值修改日期', '修改人', '修改原因']
MODIFICATION_REASONS = ['镟轮', '三个月减少 2mm', '其他']

# 检查文件是否存在，如果不存在则创建
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=COLUMNS)
    df.to_excel(EXCEL_FILE, index=False)


# 主应用类
class WheelManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("轮径值管理程序")

        # 地铁线路选择
        self.line_var = tk.StringVar(value="4 号线")
        ttk.Label(root, text="选择地铁线路:").grid(row=0, column=0, padx=10, pady=10)
        ttk.Combobox(root, textvariable=self.line_var, values=["4 号线", "5 号线"]).grid(row=0, column=1, padx=10,
                                                                                         pady=10)

        # 表格
        self.tree = ttk.Treeview(root, columns=COLUMNS, show="headings")
        for col in COLUMNS:
            self.tree.heading(col, text=col, anchor='center')
            self.tree.column(col, anchor='center')
        self.tree.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        # 按钮
        ttk.Button(root, text="添加", command=self.add_entry).grid(row=2, column=0, padx=10, pady=10)
        ttk.Button(root, text="编辑", command=self.edit_entry).grid(row=2, column=1, padx=10, pady=10)
        ttk.Button(root, text="删除", command=self.delete_entry).grid(row=3, column=0, padx=10, pady=10)
        ttk.Button(root, text="刷新", command=self.load_data).grid(row=3, column=1, padx=10, pady=10)
        ttk.Button(root, text="查询需要减少2mm的列车", command=self.query_reduction).grid(row=4, column=0, columnspan=2,
                                                                                          padx=10, pady=10)

        self.load_data()

    def load_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        df = pd.read_excel(EXCEL_FILE)
        for _, row in df.iterrows():
            if row['地铁线路'] == self.line_var.get():
                self.tree.insert("", "end", values=[row[col] for col in COLUMNS])

    def query_reduction(self):
        three_months_ago = (datetime.now() - pd.DateOffset(months=3)).strftime('%Y%m%d')
        df = pd.read_excel(EXCEL_FILE)

        # 获取本月已经有修改原因是“三个月减少 2mm”的车号
        current_month_start = datetime.now().replace(day=1).strftime('%Y%m%d')
        excluded_cars = df[(df['轮径值修改日期'].astype(str) >= current_month_start) &
                           (df['修改原因'] == '三个月减少 2mm')]['车号'].tolist()

        # 过滤查询结果
        results = df[(df['轮径值修改日期'].astype(str) <= three_months_ago) &
                     (df['修改原因'] == '镟轮') &
                     (df['地铁线路'] == self.line_var.get()) &
                     (~df['车号'].isin(excluded_cars))]

        # 新窗口来显示查询结果
        result_window = tk.Toplevel(self.root)
        result_window.title("需要减少2mm的列车")
        tree = ttk.Treeview(result_window, columns=COLUMNS, show="headings")
        for col in COLUMNS:
            tree.heading(col, text=col, anchor='center')
            tree.column(col, anchor='center')
        tree.pack(padx=10, pady=10)

        for _, row in results.iterrows():
            tree.insert("", "end", values=[row[col] for col in COLUMNS])

    def add_entry(self):
        dialog = EntryDialog(self.root, line=self.line_var.get(), title="添加条目")
        self.root.wait_window(dialog.top)
        self.load_data()

    def edit_entry(self):
        item = self.tree.selection()
        if not item:
            messagebox.showinfo("提示", "请选择一个条目来编辑")
            return
        old_data = self.tree.item(item, "values")
        dialog = EntryDialog(self.root, old_data, self.line_var.get(), title="编辑条目")
        self.root.wait_window(dialog.top)
        self.load_data()

    def delete_entry(self):
        item = self.tree.selection()
        if not item:
            messagebox.showinfo("提示", "请选择一个条目来删除")
            return
        if messagebox.askyesno("确认", "您确定要删除这个条目吗？"):
            car_number = self.tree.item(item, "values")[1]
            df = pd.read_excel(EXCEL_FILE)
            df = df[(df['地铁线路'] != self.line_var.get()) | (df['车号'] != car_number)]
            df.to_excel(EXCEL_FILE, index=False)
            self.load_data()


# 输入和编辑条目的对话框
class EntryDialog:
    def __init__(self, parent, old_data=None, line=None, title=""):
        self.line = line
        self.original_car_number = old_data[1] if old_data else None
        self.top = tk.Toplevel(parent)
        self.top.title(title)

        self.values = [tk.StringVar(value=old_data[i] if old_data else "") for i in range(len(COLUMNS))]
        row_index = 0
        for i, (var, col) in enumerate(zip(self.values, COLUMNS)):
            if col == "地铁线路":
                continue
            ttk.Label(self.top, text=col).grid(row=row_index, column=0, padx=10, pady=10)

            if col == "修改原因":
                reason_combobox = ttk.Combobox(self.top, textvariable=var, values=MODIFICATION_REASONS)
                reason_combobox.grid(row=row_index, column=1, padx=10, pady=10)
                if old_data:
                    reason_combobox.set(old_data[i])
            else:
                if col == "轮径值修改日期":
                    self.values[i].set(datetime.now().strftime('%Y%m%d'))
                ttk.Entry(self.top, textvariable=var).grid(row=row_index, column=1, padx=10, pady=10)
            row_index += 1

        ttk.Button(self.top, text="确定", command=self.save_entry).grid(row=row_index, column=0, padx=10, pady=10)
        ttk.Button(self.top, text="取消", command=self.top.destroy).grid(row=row_index, column=1, padx=10, pady=10)

    def save_entry(self):
        df = pd.read_excel(EXCEL_FILE)
        new_data = {col: var.get() for col, var in zip(COLUMNS, self.values)}
        new_data['地铁线路'] = self.line

        if self.original_car_number:
            idx = df[(df['地铁线路'] == self.line) & (df['车号'] == self.original_car_number)].index
            df.loc[idx[0]] = new_data
        else:
            df = df.append(new_data, ignore_index=True)

        df.to_excel(EXCEL_FILE, index=False)
        self.top.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = WheelManagerApp(root)
    root.mainloop()
