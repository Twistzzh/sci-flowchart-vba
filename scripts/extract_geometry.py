# -*- coding: utf-8 -*-
"""从流程图 JPG 中提取容器填充色 bbox、白色节点框 bbox、主色板。
用法: python extract_geometry.py <image> [out.json]
"""
import json
import sys
from collections import deque

import numpy as np
from PIL import Image


def load(path):
    return np.asarray(Image.open(path).convert("RGB")).astype(np.int16)


def flat_palette(a, tol=18, min_count=3000):
    q = (a // 16) * 16
    flat = q.reshape(-1, 3)
    uniq, cnt = np.unique(flat, axis=0, return_counts=True)
    out = []
    for c, n in zip(uniq, cnt):
        if n < min_count:
            continue
        m = np.all(np.abs(a - c[None, None, :]) <= tol, axis=2)
        real = a[m].mean(axis=0)
        out.append((int(n), [int(round(v)) for v in real]))
    out.sort(reverse=True)
    return out


def mask_bbox(m):
    ys, xs = np.where(m)
    if len(xs) == 0:
        return None
    return [int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())], int(m.sum())


def components(m, min_area=1500):
    """连通域(4邻域)标注, 返回 [(bbox, area)]"""
    h, w = m.shape
    lab = np.zeros((h, w), np.int32)
    res = []
    cur = 0
    idx = np.argwhere(m)
    for y0, x0 in idx:
        if lab[y0, x0]:
            continue
        cur += 1
        q = deque([(y0, x0)])
        lab[y0, x0] = cur
        x1 = x2 = x0
        y1 = y2 = y0
        area = 0
        while q:
            y, x = q.popleft()
            area += 1
            if x < x1:
                x1 = x
            if x > x2:
                x2 = x
            if y < y1:
                y1 = y
            if y > y2:
                y2 = y
            for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                ny, nx = y + dy, x + dx
                if 0 <= ny < h and 0 <= nx < w and m[ny, nx] and not lab[ny, nx]:
                    lab[ny, nx] = cur
                    q.append((ny, nx))
        if area >= min_area:
            res.append(([int(x1), int(y1), int(x2), int(y2)], int(area)))
    res.sort(key=lambda t: -t[1])
    return res


def main():
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    a = load(src)
    H, W, _ = a.shape
    rep = {"size": [W, H], "palette": [], "containers": {}, "white_boxes": []}
    pal = flat_palette(a)
    rep["palette"] = [{"rgb": c, "count": n} for n, c in pal[:14]]
    for n, c in pal[:10]:
        m = np.all(np.abs(a - np.array(c)[None, None, :]) <= 18, axis=2)
        bb = mask_bbox(m)
        if bb:
            rep["containers"]["rgb(%d,%d,%d)" % tuple(c)] = {
                "count": n,
                "bbox": bb[0],
                "xywh": [bb[0][0], bb[0][1], bb[0][2] - bb[0][0] + 1, bb[0][3] - bb[0][1] + 1],
            }
    white = np.all(a >= 248, axis=2)
    for bb, area in components(white, 1200):
        rep["white_boxes"].append(
            {
                "bbox": bb,
                "xywh": [bb[0], bb[1], bb[2] - bb[0] + 1, bb[3] - bb[1] + 1],
                "area": area,
                "fill_ratio": round(area / float((bb[2] - bb[0] + 1) * (bb[3] - bb[1] + 1)), 3),
            }
        )
    txt = json.dumps(rep, ensure_ascii=False, indent=1)
    if out:
        open(out, "w", encoding="utf-8").write(txt)
    print(txt)


if __name__ == "__main__":
    main()
