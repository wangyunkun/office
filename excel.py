import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm


df = None

def load_excel_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        df = pd.read_excel(file_path)
        update_dropdowns(df)


def update_dropdowns(df):
    row_var.set("")
    col_var.set("")
    row_dropdown["values"] = df.iloc[:, 0].tolist()
    col_dropdown["values"] = list(df.columns)


def display_selected_row():
    global df
    if df is not None:
        selected_row_label = row_var.get()
        selected_row = df[df.iloc[:, 0] == selected_row_label].index
        if len(selected_row) > 0:
            selected_row = selected_row[0]
            selected_columns = col_dropdown.get()

            if isinstance(selected_columns, str):
                selected_columns = [selected_columns]

            selected_data = df.loc[selected_row, selected_columns]
            text.delete(1.0, tk.END)
            text.insert(tk.END, selected_data.to_string(index=False))

            # 绘制柱状图
            # 指定中文字体
            font_path = "C:/Windows/Fonts/simsun.ttc"  # 宋体
            font_prop = fm.FontProperties(fname=font_path)

            # 设置字体
            plt.rcParams['font.family'] = font_prop.get_name()

            # 检查是否已经开了一个窗口，如果是，则关闭
            if plt.fignum_exists(1):
                plt.close()
            # 获取选择列的全部数据
            selected_columns = col_dropdown.get()
            y_data = df.loc[:, selected_columns].astype(float)


            # 设置绘图窗口大小
            plt.figure(figsize=(15, 8))


            # 获取横坐标数据
            x_data = df.iloc[:, 0].tolist()

            # 设置纵坐标范围
            plt.ylim(bottom=0, top=max(y_data) * 2)

            # 绘制柱状图
            bars = plt.bar(x_data, y_data)
            plt.title('普洱版纳农村供水保障项目')
            plt.xlabel('县')
            plt.ylabel(selected_columns)

            # 在每个柱子上方添加数值
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width() / 2, height, str(height), ha='center', va='bottom')

            plt.show()
root = tk.Tk()
root.title("Excel快速查询器")

load_button = ttk.Button(root, text="选择EXCEL文件", command=load_excel_file)
load_button.grid(column=0, row=0, padx=10, pady=10)

row_var = tk.StringVar()
col_var = tk.StringVar()

row_label = ttk.Label(root, text="选择行名称")
row_label.grid(column=0, row=1, padx=10, pady=10, sticky=tk.W)
row_dropdown = ttk.Combobox(root, textvariable=row_var, state="readonly")
row_dropdown.grid(column=1, row=1, padx=10, pady=10, sticky=tk.W)

col_label = ttk.Label(root, text="选择所需要列名称:")
col_label.grid(column=0, row=2, padx=10, pady=10, sticky=tk.W)
col_dropdown = ttk.Combobox(root, textvariable=col_var, state="readonly", values=[])
col_dropdown.grid(column=1, row=2, padx=10, pady=10, sticky=tk.W)

display_button = ttk.Button(root, text="点击查询数据", command=display_selected_row)
display_button.grid(column=0, row=3, padx=10, pady=10)

text = tk.Text(root, wrap=tk.WORD, width=50, height=20)
text.grid(column=0, row=4, columnspan=2, padx=10, pady=10)

root.mainloop()
