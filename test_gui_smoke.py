# -*- coding: utf-8 -*-
"""GUI 冒烟测试：真正创建 Tk 窗口，把样例数据喂给双栏视图并强制重绘，检查不抛异常。"""
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


def build_aligned():
    a = os.path.join(HERE, "g_a.xlsx")
    b = os.path.join(HERE, "g_b.xlsx")
    make(a, [
        ["K1", "苹果", 5], ["K2", "香蕉", 3], ["K3", "橙子", 8],
    ], ["编号", "名称", "数量"])
    make(b, [
        ["K1", "苹果", 5], ["K2", "香蕉", 9], ["K4", "葡萄", 2],
    ], ["编号", "名称", "数量"])
    ha, da = main.load_excel(a)
    hb, db = main.load_excel(b)
    res = main.compare(ha, da, hb, db, "key", "编号", False, True)
    return res


def run():
    root = tk.Tk()
    root.withdraw()  # 不真正显示，但仍可布局与绘制
    app = main.App(root)
    res = build_aligned()
    app.compare_view.set_data(res["aligned"], res["unified"])
    app.compare_view.update()
    root.update()
    # 校验画布已生成图元
    n_l = len(app.compare_view.cvL.find_all())
    n_r = len(app.compare_view.cvR.find_all())
    assert n_l > 0 and n_r > 0, (n_l, n_r)
    # 校验滚动同步函数不报错
    app.compare_view.cvL.yview_scroll(1, "units")
    app.compare_view._sync_v_from(app.compare_view.cvL)
    app.compare_view._hscrollL("moveto", 0.1)
    app.compare_view.clear()
    root.update()
    root.destroy()
    print("GUI smoke OK: cvL items=%d, cvR items=%d, stats=%s"
          % (n_l, n_r, res["stats"]))


if __name__ == "__main__":
    try:
        run()
        print("GUI SMOKE PASSED")
    finally:
        for f in ("g_a.xlsx", "g_b.xlsx"):
            p = os.path.join(HERE, f)
            if os.path.exists(p):
                os.remove(p)
