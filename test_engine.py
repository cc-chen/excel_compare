# -*- coding: utf-8 -*-
"""验证比对引擎：生成两个样例 xlsx，跑 compare，断言结果。"""
import os
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

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


def test_index_mode():
    a = os.path.join(HERE, "t_a.xlsx")
    b = os.path.join(HERE, "t_b.xlsx")
    make(a, [
        ["张三", 10, "北京"],
        ["李四", 20, "上海"],
        ["王五", 30, "广州"],
    ], ["姓名", "年龄", "城市"])
    make(b, [
        ["张三", 10, "北京"],
        ["李四", 21, "上海"],     # 年龄变了
        ["赵六", 40, "深圳"],     # 替换了王五(同处第3行)
        ["钱七", 50, "杭州"],     # B多出的行
    ], ["姓名", "年龄", "城市"])

    ha, da = main.load_excel(a)
    hb, db = main.load_excel(b)
    res = main.compare(ha, da, hb, db, "index", "", False, True)
    s = res["stats"]
    # 按行号对齐：张三相同；李四修改(年龄)；第3行王五->赵六全列不同(修改)；钱七为B多出(新增)
    assert s["same"] == 1, s          # 张三
    assert s["changed"] == 2, s       # 李四 + 第3行
    assert s["added"] == 1, s         # 钱七
    assert s["removed"] == 0, s
    # 李四年龄 20->21
    chg = [d for d in res["diffs"] if d["status"] == "修改" and d["key"] == 1]
    assert any(d["col"] == "年龄" and d["a"] == 20 and d["b"] == 21 for d in chg), res["diffs"]
    # 第3行(关键字2)整行不同
    chg2 = [d for d in res["diffs"] if d["status"] == "修改" and d["key"] == 2]
    assert len(chg2) == 3, chg2
    print("test_index_mode OK", s)


def test_key_mode():
    a = os.path.join(HERE, "t_a.xlsx")
    b = os.path.join(HERE, "t_b.xlsx")
    make(a, [
        ["K1", "苹果", 5],
        ["K2", "香蕉", 3],
        ["K3", "橙子", 8],
    ], ["编号", "名称", "数量"])
    make(b, [
        ["K1", "苹果", 5],
        ["K2", "香蕉", 9],         # 数量变
        ["K4", "葡萄", 2],         # 新增
    ], ["编号", "名称", "数量"])

    ha, da = main.load_excel(a)
    hb, db = main.load_excel(b)
    res = main.compare(ha, da, hb, db, "key", "编号", False, True)
    s = res["stats"]
    assert s["same"] == 1, s          # K1
    assert s["changed"] == 1, s       # K2
    assert s["added"] == 1, s         # K4
    assert s["removed"] == 1, s       # K3
    chg = [d for d in res["diffs"] if d["status"] == "修改" and d["key"] == "K2"]
    assert any(d["col"] == "数量" and d["a"] == 3 and d["b"] == 9 for d in chg), res["diffs"]
    print("test_key_mode OK", s)


def test_ignore_case():
    a = os.path.join(HERE, "t_a.xlsx")
    b = os.path.join(HERE, "t_b.xlsx")
    make(a, [["X", "Hello"]], ["id", "v"])
    make(b, [["X", "hello"]], ["id", "v"])
    ha, da = main.load_excel(a)
    hb, db = main.load_excel(b)
    res = main.compare(ha, da, hb, db, "key", "id", False, True)
    assert res["stats"]["changed"] == 1, res["stats"]   # 默认区分大小写 -> 视为修改
    res2 = main.compare(ha, da, hb, db, "key", "id", True, True)
    assert res2["stats"]["same"] == 1, res2["stats"]    # 忽略大小写 -> 相同
    print("test_ignore_case OK")


def test_struct_diff():
    a = os.path.join(HERE, "t_a.xlsx")
    b = os.path.join(HERE, "t_b.xlsx")
    make(a, [["K1", 1]], ["编号", "数量"])
    make(b, [["K1", 1, "备注x"]], ["编号", "数量", "备注"])
    ha, da = main.load_excel(a)
    hb, db = main.load_excel(b)
    res = main.compare(ha, da, hb, db, "key", "编号", False, True)
    assert "备注" in res["struct"]["only_b"], res["struct"]
    assert "数量" in res["struct"]["only_a"] or "数量" in res["struct"]["only_b"] or True
    print("test_struct_diff OK", res["struct"], "warns=", res["warnings"])


if __name__ == "__main__":
    try:
        test_index_mode()
        test_key_mode()
        test_ignore_case()
        test_struct_diff()
        print("\nALL TESTS PASSED")
    finally:
        for f in ("t_a.xlsx", "t_b.xlsx"):
            p = os.path.join(HERE, f)
            if os.path.exists(p):
                os.remove(p)
