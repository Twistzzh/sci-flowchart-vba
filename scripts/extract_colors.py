# -*- coding: utf-8 -*-
"""稳健提取：容器填充 bbox（行/列计数阈值）+ 关键描边色 + 文字色。"""
import json
import sys

import numpy as np
from PIL import Image


def robust_bbox(a, color, tol=20, min_run=25):
    m = np.all(np.abs(a - np.array(color, np.int16)) <= tol, axis=2)
    rows = m.sum(axis=1)
    cols = m.sum(axis=0)
    ry = np.where(rows >= min_run)[0]
    rx = np.where(cols >= min_run)[0]
    if len(ry) == 0 or len(rx) == 0:
        return None
    return [int(rx[0]), int(ry[0]), int(rx[-1]), int(ry[-1])]


def saturated_extreme(a, box, want="dark"):
    x1, y1, x2, y2 = box
    sub = a[y1:y2, x1:x2].reshape(-1, 3).astype(np.int32)
    lum = sub.sum(axis=1)
    order = np.argsort(lum)
    sel = order[: max(1, len(order) // 400)] if want == "dark" else order[-max(1, len(order) // 400):]
    return [int(v) for v in sub[sel].mean(axis=0).round()]


def main():
    a = np.asarray(Image.open(sys.argv[1]).convert("RGB")).astype(np.int16)
    res = {}
    fills = {
        "blue": [156, 195, 231],
        "peach": [248, 202, 169],
        "green": [196, 224, 179],
        "yellow": [254, 230, 152],
        "gray_outer": [242, 240, 231],
    }
    for k, c in fills.items():
        res.setdefault("containers", {})[k] = robust_bbox(a, c)
    # 描边色：在各容器边框附近取最饱和像素
    probe = {
        "blue_line": [66, 300, 74, 420],
        "orange_line": [560, 300, 572, 420],
        "green_line": [70, 700, 80, 780],
        "gold_line": [336, 700, 346, 780],
        "gray_line": [30, 500, 42, 600],
    }
    for k, box in probe.items():
        x1, y1, x2, y2 = box
        sub = a[y1:y2, x1:x2].reshape(-1, 3).astype(np.int32)
        lum = sub.sum(axis=1)
        sel = sub[lum <= np.percentile(lum, 4)]
        res.setdefault("strokes", {})[k] = [int(v) for v in sel.mean(axis=0).round()]
    # 文字色
    texts = {
        "title": [240, 40, 1040, 95],
        "phase1_head": [150, 165, 420, 200],
        "phase1_sub": [90, 205, 470, 245],
        "phase2_head": [600, 185, 1050, 225],
        "phase2_phase": [592, 235, 700, 275],
        "phase3_head": [370, 645, 620, 685],
        "phase3_phase": [368, 692, 470, 730],
        "message_lbl": [150, 890, 320, 940],
        "filt_lbl": [480, 320, 570, 360],
        "hier_lbl": [600, 1000, 1000, 1050],
        "node_text": [130, 275, 370, 310],
    }
    for k, box in texts.items():
        res.setdefault("text_colors", {})[k] = saturated_extreme(a, box)
    print(json.dumps(res, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
