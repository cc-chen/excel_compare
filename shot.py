# -*- coding: utf-8 -*-
"""渲染真实窗口并截图，用于人工核对双栏界面。"""
import os
import sys
import time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from openpyxl import Workbook
import main

HERE = os.path.dirname(os.path.abspath(__file__))


def make(path, rows, headers):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)


headers = ["编号", "项目", "负责人", "状态", "预算(万元)", "备注"]
rows_a = [
    ["GC1008", "主体结构", "张三", "进行中", 1200, ""],
    ["GC1009", "外立面", "李四", "未开始", 800, ""],
    ["GC1010", "机电", "王五", "进行中", 1500, "需复核"],
    ["GC1011", "消防", "赵六", "已完成", 600, ""],
    ["GC1012", "精装修", "钱七", "进行中", 950, ""],
]
rows_b = [
    ["GC1008", "主体结构", "张三", "进行中", 1250, ""],       # 预算变
    ["GC1009", "外立面", "李四", "进行中", 800, ""],          # 状态变
    ["GC1010", "机电", "王五", "进行中", 1500, "需复核"],     # 相同
    ["GC1011", "消防", "赵六", "已完成", 600, ""],            # 相同
    ["GC1013", "景观工程", "孙八", "未开始", 400, ""],        # 新增
]
a = os.path.join(HERE, "shot_a.xlsx")
b = os.path.join(HERE, "shot_b.xlsx")
make(a, rows_a, headers)
make(b, rows_b, headers)

root = tk.Tk()
root.geometry("1180x720+40+40")
app = main.App(root)
app.file_a.set(a)
app.file_b.set(b)
app.sheet_a.set("Sheet")
app.sheet_b.set("Sheet")
app.mode.set("key")
app.key_col.set("编号")

ha, da = main.load_excel(a)
hb, db = main.load_excel(b)
res = main.compare(ha, da, hb, db, "key", "编号", False, True)
app._show_result(a, b, res)   # 内部会 set_data + 设置状态栏
# 注意：set_data 会清空选择，因此必须在 _show_result 之后再设选中
app.compare_view.sel = {1, 5}
app.compare_view.sel_pane = "L"
app.compare_view._redraw_sel()
root.update_idletasks()
root.update()
try:
    root.lift()
    root.attributes("-topmost", True)
except Exception:
    pass
time.sleep(0.8)
root.update()

from PIL import ImageGrab
x = root.winfo_rootx()
y = root.winfo_rooty()
w = root.winfo_width()
h = root.winfo_height()
img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
out = os.path.join(HERE, "preview.png")
img.save(out)
print("saved", out, img.size, "stats", res["stats"])
root.destroy()
