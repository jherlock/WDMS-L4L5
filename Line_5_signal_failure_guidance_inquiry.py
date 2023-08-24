import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog, messagebox
import pandas as pd
import os
from datetime import datetime, timedelta

# 文件名
EXCEL_FILE = "Train_Wheel.xlsx"

# 列名
COLUMNS = ['地铁线路', '车号', '车轮直径', '轮径值修改日期', '修改人', '修改原因']
MODIFICATION_REASONS = ['镟轮', '三个月减少 2mm', '其他']

if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=COLUMNS)
    df.to_excel(EXCEL_FILE, index=False)


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
        self.tree = ttk.Treeview(root, columns=['地铁线路', '车号'], show="headings")
        self.tree.heading('地铁线路', text='地铁线路', anchor='center')
        self.tree.heading('车号', text='车号', anchor='center')
        self.tree.column('地铁线路', anchor='center')
        self.tree.column('车号', anchor='center')
        self.tree.bind("<Double-1>", self.view_car_records)
        self.tree.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        # 按钮
        ttk.Button(root, text="添加", command=self.add_entry).grid(row=2, column=0, padx=10, pady=10)
        ttk.Button(root, text="刷新", command=self.load_data_unique).grid(row=2, column=1, padx=10, pady=10)
        ttk.Button(root, text="查询需要减少2mm的列车", command=self.query_reduction).grid(row=3, column=0, columnspan=2,
                                                                                          padx=10, pady=10)

        self.load_data_unique()

    def load_data_unique(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        df = pd.read_excel(EXCEL_FILE)
        unique_cars = df[df['地铁线路'] == self.line_var.get()]['车号'].unique()
        for car in unique_cars:
            self.tree.insert("", "end", values=[self.line_var.get(), car])

    def view_car_records(self, event):
        item = self.tree.selection()[0]
        car_number = self.tree.item(item, 'values')[1]
        df = pd.read_excel(EXCEL_FILE)
        records = df[(df['地铁线路'] == self.line_var.get()) & (df['车号'] == car_number)]

        records_window = tk.Toplevel(self.root)
        records_window.title(f"车号 {car_number} 的记录")

        tree = ttk.Treeview(records_window, columns=COLUMNS, show="headings")
        for col in COLUMNS:
            tree.heading(col, text=col, anchor='center')
            tree.column(col, anchor='center')
        tree.pack(padx=10, pady=10)
        tree.bind("<Double-1>",
                  lambda event, tree=tree, car_number=car_number: self.edit_record(event, tree, car_number,
                                                                                   records_window))

        for _, row in records.iterrows():
            tree.insert("", "end", values=[row[col] for col in COLUMNS])

        ttk.Button(records_window, text="删除", command=lambda tree=tree: self.delete_record(tree)).pack(padx=10,
                                                                                                         pady=10)

    def delete_record(self, tree):
        item = tree.selection()[0]
        data = {col: tree.item(item, 'values')[COLUMNS.index(col)] for col in COLUMNS}

        # 密码验证窗口
        password = simpledialog.askstring("密码验证", "请输入密码：", show='*')
        if password == "admin":
            df = pd.read_excel(EXCEL_FILE)

            # 转换轮径值修改日期为字符串进行比较
            matching_rows = df[(df['车号'] == data['车号']) &
                               (df['地铁线路'] == data['地铁线路']) &
                               (df['轮径值修改日期'].astype(str) == data['轮径值修改日期'])]

            if not matching_rows.empty:
                idx = matching_rows.index[0]
                df.drop(idx, inplace=True)
                df.to_excel(EXCEL_FILE, index=False)
                tree.delete(item)
            else:
                messagebox.showwarning("警告", "所选记录未找到!")
        else:
            messagebox.showerror("错误", "密码错误!")

    def add_entry(self):
        dialog = EntryDialog(self.root, line=self.line_var.get(), title="添加条目")
        self.root.wait_window(dialog.top)
        self.load_data_unique()

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

    def edit_record(self, event, tree, car_number, records_window):
        item = tree.selection()[0]
        old_data = {col: tree.item(item, 'values')[COLUMNS.index(col)] for col in COLUMNS}

        # 密码验证窗口
        password = simpledialog.askstring("密码验证", "请输入密码：", show='*')
        if password == "admin":
            dialog = EntryDialog(self.root, old_data=old_data, line=self.line_var.get(), title="编辑条目",
                                 callback=lambda: self.refresh_records(tree, car_number))
            self.root.wait_window(dialog.top)
            if dialog.top:
                self.refresh_records(tree, car_number)  # 在此处刷新车辆记录
        else:
            messagebox.showerror("错误", "密码错误!")

    def refresh_list(self):
        df = pd.read_excel(EXCEL_FILE)
        unique_cars = sorted(df['车号'].unique())  # 对车号进行排序

        self.car_listbox.delete(0, tk.END)
        for car in unique_cars:
            self.car_listbox.insert(tk.END, car)


class EntryDialog:
    def __init__(self, parent, old_data=None, line=None, title="", callback=None):
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.line = line
        self.callback = callback
        self.old_data = old_data

        self.car_number_var = tk.StringVar(value=old_data.get("车号") if old_data else "")
        self.wheel_diameter_var = tk.StringVar(value=old_data.get("车轮直径") if old_data else "")
        self.modification_date_var = tk.StringVar(
            value=old_data.get("轮径值修改日期") if old_data else datetime.now().strftime('%Y%m%d'))
        self.modifier_var = tk.StringVar(value=old_data.get("修改人") if old_data else "")
        self.modification_reason_var = tk.StringVar(value=old_data.get("修改原因") if old_data else "镟轮")

        ttk.Label(self.top, text="车号:").grid(row=0, column=0, padx=10, pady=10)
        ttk.Entry(self.top, textvariable=self.car_number_var).grid(row=0, column=1, padx=10, pady=10)

        ttk.Label(self.top, text="车轮直径:").grid(row=1, column=0, padx=10, pady=10)
        ttk.Entry(self.top, textvariable=self.wheel_diameter_var).grid(row=1, column=1, padx=10, pady=10)

        ttk.Label(self.top, text="轮径值修改日期:").grid(row=2, column=0, padx=10, pady=10)
        ttk.Entry(self.top, textvariable=self.modification_date_var).grid(row=2, column=1, padx=10, pady=10)

        ttk.Label(self.top, text="修改人:").grid(row=3, column=0, padx=10, pady=10)
        ttk.Entry(self.top, textvariable=self.modifier_var).grid(row=3, column=1, padx=10, pady=10)

        ttk.Label(self.top, text="修改原因:").grid(row=4, column=0, padx=10, pady=10)
        ttk.Combobox(self.top, textvariable=self.modification_reason_var, values=MODIFICATION_REASONS).grid(row=4,
                                                                                                            column=1,
                                                                                                            padx=10,
                                                                                                            pady=10)

        ttk.Button(self.top, text="提交", command=self.save_entry).grid(row=5, column=0, padx=10, pady=10)
        ttk.Button(self.top, text="取消", command=self.top.destroy).grid(row=5, column=1, padx=10, pady=10)

    def save_entry(self):
        new_data = {
            '地铁线路': self.line,
            '车号': self.car_number_var.get(),
            '车轮直径': self.wheel_diameter_var.get(),
            '轮径值修改日期': self.modification_date_var.get(),
            '修改人': self.modifier_var.get(),
            '修改原因': self.modification_reason_var.get()
        }
        df = pd.read_excel(EXCEL_FILE)

        if self.old_data:
            # 仅基于车号和地铁线路进行索引
            idx = df[(df['车号'] == self.old_data['车号']) &
                     (df['地铁线路'] == self.line)].index[0]

            # 更新该索引上的数据
            for col in COLUMNS:
                df.at[idx, col] = new_data[col]
        else:
            df = df.append(new_data, ignore_index=True)

        df.to_excel(EXCEL_FILE, index=False)
        self.top.destroy()

        if self.callback:
            self.callback()


if __name__ == "__main__":
    root = tk.Tk()
    app = WheelManagerApp(root)
    root.mainloop()
