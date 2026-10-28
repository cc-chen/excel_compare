# -*- coding: utf-8 -*-
"""验证双栏视图的选中与复制逻辑。"""
import os
import sys
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


def sample():
    a = os.path.join(HERE, "s_a.xlsx")
    b = os.path.join(HERE, "s_b.xlsx")
    make(a, [["K1", "苹果", 5], ["K2", "香蕉", 3], ["K3", "橙子", 8]],
         ["编号", "名称", "数量"])
    make(b, [["K1", "苹果", 5], ["K2", "香蕉", 9], ["K4", "葡萄", 2]],
         ["编号", "名称", "数量"])
    ha, da = main.load_excel(a)
    hb, db = main.load_excel(b)
    return main.compare(ha, da, hb, db, "key", "编号", False, True)


class FakeEvent:
    def __init__(self, y, x=10):
        self.y = y
        self.x = x
        self.x_root = 0
        self.y_root = 0


def run():
    root = tk.Tk()
    root.withdraw()
    app = main.App(root)
    res = sample()
    cv = app.compare_view
    cv.set_data(res["aligned"], res["unified"])
    cv.update()
    root.update()

    captured = {}
    cv._to_clipboard = lambda text, n: captured.update(text=text, n=n)

    cols = cv.cols
    n = len(cv.aligned)
    # K1 same, K2 changed, K3 removed, K4 added = 4 行
    assert n == 4, n
    print("aligned rows:", n, "cols:", cols)

    # 1) 行点击定位：y = hdrh + row*rowh + 5
    i = cv._row_at(cv.cvL, FakeEvent(cv.hdrh + 2 * cv.rowh + 5))
    assert i == 2, i

    # 2) 单击选中
    cv._on_press(FakeEvent(cv.hdrh + 0 * cv.rowh + 5), cv.cvL, "L")
    assert cv.sel == {0}, cv.sel
    # 3) Shift 范围选择 0..2
    cv._on_shift(FakeEvent(cv.hdrh + 2 * cv.rowh + 5), cv.cvL, "L")
    assert cv.sel == {0, 1, 2}, cv.sel
    # 4) Ctrl 取消 1
    cv._on_ctrl(FakeEvent(cv.hdrh + 1 * cv.rowh + 5), cv.cvL, "L")
    assert cv.sel == {0, 2}, cv.sel
    # 5) 全选
    cv.select_all()
    assert cv.sel == set(range(n)), cv.sel

    # 6) 复制左栏：行数 == n，每行 tab 数 == len(cols)-1
    cv.sel_pane = "L"
    cv.copy()
    txt = captured["text"]
    lines = txt.split("\n")
    assert len(lines) == n, (len(lines), n)
    assert all(line.count("\t") == len(cols) - 1 for line in lines), lines
    print("copy L first line:", lines[0])

    # 7) 含表头
    cv.copy(header=True)
    lines = captured["text"].split("\n")
    assert lines[0] == "\t".join(cols), lines[0]

    # 8) 只选一行（K2 修改），复制右栏应含 9
    cv.sel = {1}
    cv.sel_pane = "R"
    cv.copy()
    assert "9" in captured["text"], captured["text"]

    # 9) 差异文本
    cv.select_all()
    cv.copy_diff_text()
    dt = captured["text"]
    assert "修改" in dt and "新增" in dt and "删除" in dt, dt
    print("---- diff text ----")
    print(dt)

    # 10) 选中覆盖层图元存在
    cv.sel = {0}
    cv._redraw_sel()
    assert len(cv.cvL.find_withtag("sel")) >= 2, len(cv.cvL.find_withtag("sel"))

    # 11) 清除
    cv.clear_selection()
    assert cv.sel == set()
    assert len(cv.cvL.find_withtag("sel")) == 0

    root.update()
    root.destroy()
    print("SELECT/COPY TESTS PASSED")


if __name__ == "__main__":
    try:
        run()
    finally:
        for f in ("s_a.xlsx", "s_b.xlsx"):
            p = os.path.join(HERE, f)
            if os.path.exists(p):
                os.remove(p)
