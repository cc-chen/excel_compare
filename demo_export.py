# -*- coding: utf-8 -*-
"""生成一份较宽的示例差异（多列多行），导出 Excel + HTML，并渲染 xlsx 预览。"""
import os
import main
import preview_xlsx
from openpyxl import Workbook

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_t")
os.makedirs(BASE, exist_ok=True)

HEADER = ["编号", "名称", "专业", "计划开始", "进度%", "状态", "预算(万)", "备注"]

A = [
    ["GC1008", "桩基工程", "土建", "2026-03-01", 100, "已完成", 1200, "已验收"],
    ["GC1009", "主体结构", "土建", "2026-04-10", 65, "进行中", 3600, "正常"],
    ["GC1010", "机电安装", "MEP", "2026-05-01", 40, "进行中", 2400, "管综待确认"],
    ["GC1011", "幕墙工程", "Façade", "2026-06-15", 20, "进行中", 1800, "样板未定"],
    ["GC1012", "消防工程", "Fire", "2026-06-20", 15, "进行中", 900, "图审滞后"],
    ["GC1013", "精装修", "Fit out", "2026-07-01", 5, "未开始", 2600, "待移交"],
    ["GC1014", "景观绿化", "Landscape", "2026-08-01", 0, "未开始", 700, ""],
]
B = [
    ["GC1008", "桩基工程", "土建", "2026-03-01", 100, "已完成", 1200, "已验收"],   # 相同
    ["GC1009", "主体结构", "土建", "2026-04-10", 72, "进行中", 3600, "正常"],     # 进度变化
    ["GC1010", "机电安装", "MEP", "2026-05-06", 45, "进行中", 2500, "管综已确认"],  # 开始日/进度/预算/备注
    ["GC1011", "幕墙工程", "Façade", "2026-06-15", 20, "暂停", 1800, "样板未定"],   # 状态变化
    # GC1012 删除
    ["GC1013", "精装修", "Fit out", "2026-07-05", 8, "进行中", 2600, "已移交"],     # 4 处变化
    ["GC1014", "景观绿化", "Landscape", "2026-08-01", 0, "未开始", 700, ""],        # 相同
    ["GC1015", "电梯工程", "MEP", "2026-09-01", 0, "未开始", 480, "新签"],          # 新增
]


def mk(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "进度"
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    wb.save(path)


pa = os.path.join(BASE, "A.xlsx")
pb = os.path.join(BASE, "B.xlsx")
mk(pa, A)
mk(pb, B)

ha, da = main.load_excel(pa)
hb, db = main.load_excel(pb)
res = main.compare(ha, da, hb, db, "key", "编号", False, True)
print("stats:", res["stats"], "aligned:", len(res["aligned"]))

xp = os.path.join(BASE, "demo_diff.xlsx")
hp = os.path.join(BASE, "demo_diff.html")
info = {"file_a": "D:/项目/罗氏/进度计划_V1.xlsx", "sheet_a": "进度",
        "file_b": "D:/项目/罗氏/进度计划_V2.xlsx", "sheet_b": "进度",
        "opts": "按关键字列「编号」（表头行=1）"}
main.export_excel(res, xp, info)
main.build_html(res, info["file_a"], info["sheet_a"], info["file_b"], info["sheet_b"], info["opts"])
with open(hp, "w", encoding="utf-8") as f:
    f.write(main.build_html(res, info["file_a"], info["sheet_a"],
                            info["file_b"], info["sheet_b"], info["opts"]))
preview_xlsx.dump(xp, os.path.join(BASE, "xlsx_preview.html"))
print("ok:", xp, hp)
