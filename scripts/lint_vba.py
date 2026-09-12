# -*- coding: utf-8 -*-
"""对生成的 .bas 做结构自检（导入 PowerPoint 前把能查的错都查掉）。

用法: python lint_vba.py <输出目录>
"""
import io
import os
import re
import sys

GEOM = ["CANVAS_W_PX", "CANVAS_H_PX", "PX_TO_IN", "OFFSET_X_IN", "OFFSET_Y_IN",
        "SLIDE_W_IN", "SLIDE_H_IN", "CUSTOM_SIZE", "FONT_NAME"]
ENGINE_API = ["PX2X", "PX2Y", "PX2L", "ShapeTypeFor", "AddNode", "AddPath",
              "DrawPolyline", "SetLabel", "ClearPrevious", "TagShape", "HasTag",
              "BuildFlowchart", "RemoveFlowchart"]
ENGINE_CONST = ["TAG_NAME", "ARROWHEAD_TRIANGLE", "LINE_SOLID", "LINE_DASH",
                "LINE_DOT", "LINE_DASHDOT"]
KINDS = {"rect", "rectangle", "box", "round_rect", "rounded", "rounded_rect", "oval",
         "ellipse", "circle", "diamond", "decision", "parallelogram", "trapezoid",
         "pentagon", "hexagon", "stadium", "pill", "terminator", "cylinder",
         "document", "right_arrow", "arrow_right", "left_arrow", "arrow_left",
         "up_arrow", "arrow_up", "down_arrow", "arrow_down", "chevron"}


def split_args(s):
    out, cur, depth, q = [], [], 0, False
    for ch in s:
        if ch == '"':
            q = not q
        if ch == "(" and not q:
            depth += 1
        elif ch == ")" and not q:
            depth -= 1
        if ch == "," and depth == 0 and not q:
            out.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    cur = "".join(cur).strip()
    if cur:
        out.append(cur)
    return out


def main(d):
    errs, warns, info = [], [], []
    eng = io.open(os.path.join(d, "modFlow_Engine.bas"), "rb").read()
    txt = eng.decode("latin-1")
    mods = {}
    for fn in sorted(os.listdir(d)):
        if not fn.endswith(".bas"):
            continue
        raw = open(os.path.join(d, fn), "rb").read()
        try:
            raw.decode("ascii")
        except UnicodeDecodeError as e:
            errs.append("%s: 含非 ASCII 字节 (%s)" % (fn, e))
        t = raw.decode("latin-1")
        if not t.startswith("Attribute VB_Name = "):
            errs.append("%s: 首行不是 Attribute VB_Name" % fn)
        if b"\r\n" not in raw:
            warns.append("%s: 行尾不是 CRLF" % fn)
        mods[fn] = t

    consts, subs, calls = set(), {}, []
    for fn, t in mods.items():
        for m in re.finditer(r"Public Const (\w+)\s+As\s+\w+\s*=", t):
            consts.add(m.group(1))
        for m in re.finditer(r"^(?:Public |Private )?Sub (\w+)\(", t, re.M):
            subs[m.group(1)] = fn
        for m in re.finditer(r"^\s*(AddNode|AddPath|DrawPolyline|SetLabel|PX2X|PX2Y|PX2L)\b", t, re.M):
            calls.append(m.group(1))

    if "DrawAll" not in subs:
        errs.append("缺少 Public Sub DrawAll")
    missing = [g for g in GEOM if g not in consts] if "modFlow_Engine.bas" not in mods else \
              [g for g in GEOM if g not in consts]
    if missing:
        errs.append("缺几何常量: %s" % missing)

    # 内容模块里的 Public Sub Draw* 必须都存在（跨模块调用）
    for fn, t in mods.items():
        if fn == "modFlow_Engine.bas":
            continue
        for m in re.finditer(r"^\s{4}([A-Z]\w+) sld$", t, re.M):
            if m.group(1) not in subs:
                errs.append("%s: 调用了未定义的 %s" % (fn, m.group(1)))

    # 颜色常量引用是否都有定义（引擎常量 + 本套调色板）
    known = consts | set(ENGINE_CONST)
    for fn, t in mods.items():
        if fn == "modFlow_Engine.bas":
            continue
        for m in re.finditer(r"\b(C_\w+)\b", t):
            if m.group(1) not in consts:
                errs.append("%s: 未定义的颜色常量 %s" % (fn, m.group(1)))

    # 逐行检查 AddNode / AddPath
    n_nodes = n_paths = 0
    for fn, t in mods.items():
        for ln, line in enumerate(t.splitlines(), 1):
            s = line.strip()
            if s.startswith("AddNode "):
                n_nodes += 1
                a = split_args(s[8:])
                if len(a) < 11:
                    errs.append("%s:%d AddNode 实参不足 (%d)" % (fn, ln, len(a)))
                if len(a) > 20:
                    errs.append("%s:%d AddNode 实参过多 (%d)" % (fn, ln, len(a)))
                if any(x == "" for x in a):
                    errs.append("%s:%d AddNode 存在空实参" % (fn, ln))
                kind = a[2].strip('"')
                if kind not in KINDS:
                    errs.append("%s:%d 未知形状 kind=%s" % (fn, ln, kind))
                dsh = a[10]
                if dsh not in ENGINE_CONST and not dsh.isdigit():
                    errs.append("%s:%d 非法 dash=%s" % (fn, ln, dsh))
                if not os.path.exists(os.path.join(d, fn)):
                    pass
            elif s.startswith("AddPath "):
                n_paths += 1
                a = split_args(s[8:])
                if len(a) != 7:
                    errs.append("%s:%d AddPath 实参应为 7，实为 %d" % (fn, ln, len(a)))
                pts = a[2].strip('"').split(";")
                for p in pts:
                    if len(p.split(",")) != 2:
                        errs.append("%s:%d 点串非法: %s" % (fn, ln, p))
                if a[5] not in ENGINE_CONST and not a[5].isdigit():
                    errs.append("%s:%d 非法 dash=%s" % (fn, ln, a[5]))
                if a[6] not in ("True", "False"):
                    errs.append("%s:%d arrowEnd 非布尔: %s" % (fn, ln, a[6]))
            if len(line) > 900:
                warns.append("%s:%d 行过长 (%d 字符)" % (fn, ln, len(line)))

    # 中文检测（只在字符串字面量里）
    for fn, t in mods.items():
        for m in re.finditer(r'"([^"]*)"', t):
            if any(ord(c) > 127 for c in m.group(1)):
                errs.append("%s: 字符串含非 ASCII: %r" % (fn, m.group(1)))

    info.append("模块: %s" % ", ".join(sorted(mods)))
    info.append("Sub: %d 个, AddNode %d 条, AddPath %d 条" % (len(subs), n_nodes, n_paths))
    print("== INFO ==")
    for i in info:
        print("  " + i)
    print("== WARN (%d) ==" % len(warns))
    for w in warns:
        print("  " + w)
    print("== ERROR (%d) ==" % len(errs))
    for e in errs:
        print("  " + e)
    return 1 if errs else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "."))
