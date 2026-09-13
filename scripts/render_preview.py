# -*- coding: utf-8 -*-
"""Replay modFlow_Content*.bas (parses AddNode / AddPath), render a preview PNG
side by side with ANY source figure.

Usage: python render_preview.py <output_dir> <source_image>
"""
import math
import os
import re
import sys

from PIL import Image, ImageDraw, ImageFont

OUT = os.path.abspath(sys.argv[1])
SRC = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else None
VBA = sorted(f for f in os.listdir(OUT) if f.lower().endswith(".bas")
             and "content" in f.lower())


def _read_bas(path):
    """Read a .bas file trying UTF-8 then GBK (zh-CN VBE), then latin-1."""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "gbk"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("latin-1")


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
    if cur:
        out.append("".join(cur).strip())
    return out


def untext(t):
    t = t.strip()
    parts = re.split(r'"\s*&\s*vbLf\s*&\s*"', t, flags=re.I)
    return "\n".join(p.strip().strip('"') for p in parts).strip()


# ---------------------------------------------------------------- parse
pal = {}
nodes, paths = [], []
canvas = [1280, 1120]
FONT_NAME = "Arial"
PREFER_YAHEI = False
formulas = []
for fn in VBA:
    raw = _read_bas(os.path.join(OUT, fn))
    if "FONT_NAME" in raw and ("YaHei" in raw or "SimSun" in raw or "SimHei" in raw):
        PREFER_YAHEI = True
    for m in re.finditer(r'Public Const FONT_NAME As String\s*=\s*"([^"]*)"', raw):
        FONT_NAME = m.group(1)
    for m in re.finditer(r"Public Const (\w+)\s+As Long\s*=\s*RGB\((\d+),\s*(\d+),\s*(\d+)\)", raw):
        pal[m.group(1)] = (int(m.group(2)), int(m.group(3)), int(m.group(4)))
    for m in re.finditer(r"Public Const CANVAS_W_PX As Single = ([\d.]+)", raw):
        canvas[0] = float(m.group(1))
    for m in re.finditer(r"Public Const CANVAS_H_PX As Single = ([\d.]+)", raw):
        canvas[1] = float(m.group(1))
    for m in re.finditer(r"Public Const PX_TO_IN As Single = ([\d.]+)", raw):
        px_per_pt = 1.0 / (float(m.group(1)) * 72.0)
    logical = re.sub(r"_\s*\r?\n\s*", " ", raw)  # join VBA continuations
    for line in logical.splitlines():
        line = line.strip()
        if line.startswith("AddNode "):
            args = split_args(line[8:])[1:]  # drop sld
            ident = args[0].strip('"')
            kind = args[1].strip('"')
            x, y, w, h = (float(args[i]) for i in range(2, 6))
            fill, linec, lw, dash = args[6], args[7], float(args[8]), args[9]
            rest = args[10:]
            text = untext(rest[0]) if rest else ""
            pt = float(rest[1]) if len(rest) > 1 and rest[1] else 11
            bold = len(rest) > 2 and rest[2].lower() == "true"
            ital = len(rest) > 3 and rest[3].lower() == "true"
            fc = rest[4] if len(rest) > 4 and rest[4] else "-1"
            corner = float(rest[5]) if len(rest) > 5 and rest[5] else -1
            fname = rest[6].strip('"') if len(rest) > 6 and rest[6] else ""
            adj2 = float(rest[7]) if len(rest) > 7 and rest[7] else -1
            nodes.append(dict(id=ident, kind=kind.lower(), x=x, y=y, w=w, h=h,
                              fill=fill, line=linec, lw=lw, dash=dash, text=text,
                              pt=pt, bold=bold, ital=ital, fc=fc, corner=corner,
                              fname=fname, adj2=adj2, order=len(nodes)))
        elif line.startswith("AddPath "):
            args = split_args(line[8:])
            ident = args[1].strip('"')
            pts = [tuple(float(v) for v in p.split(",")) for p in args[2].strip('"').split(";")]
            color, lw, dash = args[3], float(args[4]), args[5]
            arrow = args[6].lower() == "true"
            paths.append(dict(id=ident, pts=pts, color=color, lw=lw, dash=dash,
                              arrow=arrow, order=len(paths)))
        elif line.startswith("AddFormula "):
            args = split_args(line[11:])
            # sld, id, x, y, w, h, "tex", "fallback", pt, fontC
            formulas.append(dict(id=args[1].strip('"'),
                                 x=float(args[2]), y=float(args[3]),
                                 w=float(args[4]), h=float(args[5]),
                                 tex=args[6].strip('"'),
                                 fb=args[7].strip('"'),
                                 pt=float(args[8]) if len(args) > 8 and args[8] else 16,
                                 fc=args[9] if len(args) > 9 else "-1",
                                 order=len(formulas)))
CW, CH = int(canvas[0]), int(canvas[1])
img = Image.new("RGB", (CW, CH), (255, 255, 255))
d = ImageDraw.Draw(img)


def col(c):
    c = c.strip()
    if c in pal:
        return pal[c]
    m = re.match(r"RGB\((\d+),\s*(\d+),\s*(\d+)\)", c)
    if m:
        return tuple(int(g) for g in m.groups())
    if c in ("-1", ""):
        return None
    return (0, 0, 0)


FONT_CANDIDATES = [
    r"C:\Windows\Fonts\times.ttf",
    r"C:\Windows\Fonts\timesbd.ttf",
    r"C:\Windows\Fonts\timesi.ttf",
    r"C:\Windows\Fonts\timesbi.ttf",
]
FONTS = {}


CJK_FONT_FILES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttf",
]


def has_cjk(s):
    """Return True if *s* contains CJK ideographs / full-width punctuation."""
    return any('\u3000' <= c <= '\u303f' or '\u3400' <= c <= '\u9fff'
               or '\uff00' <= c <= '\uffef' for c in s)


def cjk_font_for(family, bold):
    """Map the content module's FONT_NAME to a matching CJK font file.

    - sans-serif / YaHei / Arial        -> Microsoft YaHei (bold -> msyhbd.ttc)
    - serif / Times / SimSun / 宋体     -> SimSun
    - SimHei / 黑体                     -> SimHei
    This keeps the preview's Chinese typeface consistent with the source figure
    instead of always defaulting to YaHei.
    """
    f = (family or "").lower()
    if "simhei" in f or "黑" in f or "hei" in f:
        return r"C:\Windows\Fonts\simhei.ttf"
    if "simsun" in f or "宋" in f or "sun" in f or "serif" in f or "times" in f:
        return r"C:\Windows\Fonts\simsun.ttc"
    if bold:
        return r"C:\Windows\Fonts\msyhbd.ttc"
    return r"C:\Windows\Fonts\msyh.ttc"


def getfont(pt, bold, ital, cjk=False, family=""):
    key = (round(pt, 1), bold, ital, cjk, family.lower())
    if key not in FONTS:
        if cjk:
            cands = [cjk_font_for(family, bold)] + list(CJK_FONT_FILES)
        elif PREFER_YAHEI:
            cands = []
            if bold:
                cands.append(r"C:\Windows\Fonts\msyhbd.ttc")
            cands.append(r"C:\Windows\Fonts\msyh.ttc")
            cands.append(r"C:\Windows\Fonts\simsun.ttc")
        else:
            f = (r"C:\Windows\Fonts\timesbi.ttf" if bold and ital else
                 r"C:\Windows\Fonts\timesbd.ttf" if bold else
                 r"C:\Windows\Fonts\timesi.ttf" if ital else
                 r"C:\Windows\Fonts\times.ttf")
            cands = [f]
        got = None
        for f in cands:
            try:
                got = ImageFont.truetype(f, max(6, int(round(pt * px_per_pt))))
                break
            except OSError:
                continue
        FONTS[key] = got if got is not None else ImageFont.load_default()
    return FONTS[key]


def dashed(p0, p1, color, width, dash=10, gap=7):
    (x0, y0), (x1, y1) = p0, p1
    L = math.hypot(x1 - x0, y1 - y0)
    if L == 0:
        return
    ux, uy = (x1 - x0) / L, (y1 - y0) / L
    t = 0.0
    while t < L:
        e = min(t + dash, L)
        d.line([(x0 + ux * t, y0 + uy * t), (x0 + ux * e, y0 + uy * e)],
               fill=color, width=max(1, int(round(width))))
        t = e + gap


def head(p, ang, color, size):
    a1, a2 = ang + math.radians(155), ang - math.radians(155)
    p1 = (p[0] + size * math.cos(a1), p[1] + size * math.sin(a1))
    p2 = (p[0] + size * math.cos(a2), p[1] + size * math.sin(a2))
    d.polygon([p, p1, p2], fill=color)


def poly(pts, fill, linec, lw, closed=True):
    if fill:
        d.polygon(pts, fill=fill)
    if linec:
        w = max(1, int(round(lw)))
        seq = pts + [pts[0]] if closed else pts
        for i in range(len(seq) - 1):
            d.line([seq[i], seq[i + 1]], fill=linec, width=w)


def node_shape(n):
    """returns (kind_key, draw_fn) — polygon kinds get explicit outlines."""
    k, x, y, w, h = n["kind"], n["x"], n["y"], n["w"], n["h"]
    fill, linec = col(n["fill"]), col(n["line"])
    lw = n["lw"]
    dash = n["dash"] == "LINE_DASH"
    box = [x, y, x + w, y + h]

    if k in ("rect", "rectangle", "box"):
        if dash and linec:
            for a, b in (((x, y), (x + w, y)), ((x + w, y), (x + w, y + h)),
                         ((x + w, y + h), (x, y + h)), ((x, y + h), (x, y))):
                dashed(a, b, linec, lw)
            if fill:
                d.rectangle(box, fill=fill)
        else:
            if fill:
                d.rectangle(box, fill=fill)
            if linec:
                d.rectangle(box, outline=linec, width=max(1, int(round(lw))))
    elif k in ("round_rect", "rounded", "rounded_rect", "stadium", "pill", "terminator"):
        r = h / 2 if k in ("stadium", "pill", "terminator") else \
            max(2, min(w, h) * (n["corner"] if n["corner"] > 0 else 0.14))
        if fill:
            d.rounded_rectangle(box, radius=r, fill=fill)
        if linec:
            d.rounded_rectangle(box, radius=r, outline=linec, width=max(1, int(round(lw))))
    elif k in ("oval", "ellipse", "circle"):
        if fill:
            d.ellipse(box, fill=fill)
        if linec:
            d.ellipse(box, outline=linec, width=max(1, int(round(lw))))
    elif k == "diamond":
        poly([(x + w / 2, y), (x + w, y + h / 2), (x + w / 2, y + h), (x, y + h / 2)],
             fill, linec, lw)
    elif k == "parallelogram":
        off = w * 0.2
        poly([(x + off, y), (x + w, y), (x + w - off, y + h), (x, y + h)], fill, linec, lw)
    elif k == "trapezoid":
        off = w * 0.2
        poly([(x, y), (x + w, y), (x + w - off, y + h), (x + off, y + h)], fill, linec, lw)
    elif k == "pentagon":  # home plate: flat left, pointed right
        tip = h * 0.5
        poly([(x, y), (x + w - tip, y), (x + w, y + h / 2), (x + w - tip, y + h), (x, y + h)],
             fill, linec, lw)
    elif k == "hexagon":  # double-pointed banner
        tip = min(w, h) * (n["corner"] if n["corner"] > 0 else 0.25)
        poly([(x + tip, y), (x + w - tip, y), (x + w, y + h / 2),
              (x + w - tip, y + h), (x + tip, y + h), (x, y + h / 2)], fill, linec, lw)
    elif k == "chevron":
        tip = h * 0.5
        poly([(x, y), (x + w - tip, y), (x + w, y + h / 2), (x + w - tip, y + h),
              (x, y + h), (x + tip, y + h / 2)], fill, linec, lw)
    elif k == "right_arrow":
        hd = w * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bh = h * ((n["corner"] if 0 < n["corner"] < 1 else 0.5))
        cy = y + h / 2
        poly([(x, cy - bh / 2), (x + w - hd, cy - bh / 2), (x + w - hd, y),
              (x + w, cy), (x + w - hd, y + h), (x + w - hd, cy + bh / 2),
              (x, cy + bh / 2)], fill, linec, lw)
    elif k == "left_arrow":
        hd = w * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bh = h * ((n["corner"] if 0 < n["corner"] < 1 else 0.5))
        cy = y + h / 2
        poly([(x + w, cy - bh / 2), (x + hd, cy - bh / 2), (x + hd, y),
              (x, cy), (x + hd, y + h), (x + hd, cy + bh / 2),
              (x + w, cy + bh / 2)], fill, linec, lw)
    elif k == "up_arrow":
        hd = h * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bw = w * ((n["corner"] if 0 < n["corner"] < 1 else 0.5))
        cx = x + w / 2
        poly([(cx - bw / 2, y + h), (cx - bw / 2, y + hd), (x, y + hd),
              (cx, y), (x + w, y + hd), (cx + bw / 2, y + hd),
              (cx + bw / 2, y + h)], fill, linec, lw)
    elif k == "down_arrow":
        hd = h * (n["adj2"] if n["adj2"] > 0 else 0.5)
        bw = w * ((n["corner"] if 0 < n["corner"] < 1 else 0.5))
        cx = x + w / 2
        poly([(cx - bw / 2, y), (cx - bw / 2, y + h - hd), (x, y + h - hd),
              (cx, y + h), (x + w, y + h - hd), (cx + bw / 2, y + h - hd),
              (cx + bw / 2, y)], fill, linec, lw)
    else:  # cylinder / document / unknown -> rectangle fallback
        if fill:
            d.rectangle(box, fill=fill)
        if linec:
            d.rectangle(box, outline=linec, width=max(1, int(round(lw))))


# second pass: decoration calls (SetPara / SetPartColor) after ALL nodes exist
for fn in VBA:
    raw = _read_bas(os.path.join(OUT, fn))
    logical = re.sub(r"_\s*\r?\n\s*", " ", raw)
    for line in logical.splitlines():
        line = line.strip()
        if line.startswith("SetPara "):
            args = split_args(line[8:])
            nid, align = args[1].strip('"'), args[2].strip('"').lower()
            mg = float(args[3] or 0) if len(args) > 3 else 0.0
            for n in nodes:
                if n["id"] == nid:
                    n["align"], n["ml"] = align, mg
        elif line.startswith("SetPartColor "):
            args = split_args(line[8:])
            nid, nc = args[1].strip('"'), int(float(args[2]))
            token = args[3] if len(args) > 3 else "RGB(200, 30, 40)"
            for n in nodes:
                if n["id"] == nid:
                    n["prefix"] = (nc, token)

for n in nodes:
    node_shape(n)
    if n["text"]:
        cjk = has_cjk(n["text"])
        fam = n.get("fname") or FONT_NAME  # per-node fontName overrides FONT_NAME
        f = getfont(n["pt"], n["bold"], n["ital"], cjk, fam)
        tc = col(n["fc"]) or (31, 31, 31)
        left = n.get("align") == "left"
        ml = n.get("ml", 0) or 0
        # CJK 文本没有空格 -> 按字符折行；西文按单词折行。
        lines = []
        for para in n["text"].split("\n"):
            units = list(para) if cjk else para.split()
            sep = "" if cjk else " "
            lim = n["w"] - 6 - (ml if left else 0)
            cur = ""
            for wd in units:
                t = (cur + sep + wd).strip() if sep else (cur + wd)
                if d.textlength(t, font=f) <= lim or not cur:
                    cur = t
                else:
                    lines.append(cur)
                    cur = wd
            lines.append(cur)
        lh = n["pt"] * px_per_pt
        total = lh * len(lines)
        ty = n["y"] + n["h"] / 2 - total / 2
        prefix = n.get("prefix")
        px_chars = prefix[0] if prefix else 0
        pcol = col(prefix[1]) if prefix else tc
        fb = getfont(n["pt"], True, n["ital"], cjk, fam)
        for li, ln in enumerate(lines):
            if left:
                x0 = n["x"] + 3 + ml
                d.text((x0, ty), ln, font=f, fill=tc, anchor="la")
            elif px_chars > 0 and li == 0 and px_chars < len(ln):
                # red bold prefix on the FIRST line only, rest in base colour
                pre = ln[:px_chars]
                rest = ln[px_chars:]
                x0 = n["x"] + n["w"] / 2 - (d.textlength(ln, font=f)) / 2
                d.text((x0, ty), pre, font=fb, fill=pcol, anchor="la")
                d.text((x0 + d.textlength(pre, font=fb), ty), rest, font=f,
                       fill=tc, anchor="la")
            else:
                d.text((n["x"] + n["w"] / 2, ty), ln, font=f, fill=tc, anchor="ma")
            ty += lh

for p in paths:
    c = col(p["color"]) or (0, 0, 0)
    wdt = max(1, int(round(p["lw"])))
    is_dash = (p.get("dash") or "LINE_SOLID") != "LINE_SOLID"
    for i in range(len(p["pts"]) - 1):
        a, b = p["pts"][i], p["pts"][i + 1]
        if is_dash:
            if p["dash"] == "LINE_DOT":
                dashed(a, b, c, wdt, dash=2, gap=4)
            elif p["dash"] == "LINE_DASHDOT":
                dashed(a, b, c, wdt, dash=8, gap=5)
            else:
                dashed(a, b, c, wdt, dash=9, gap=6)
        else:
            d.line([a, b], fill=c, width=wdt)
    if p["arrow"]:
        a, b = p["pts"][-2], p["pts"][-1]
        ang = math.atan2(b[1] - a[1], b[0] - a[0])
        head(b, ang, c, 13)

# ---- LaTeX formulas: native OMML in the pptx; preview via mathtext traceback


def _render_tex_png(tex: str, size_pt: float, col_rgb):
    """Render a LaTeX string to a cropped RGBA PIL image sized for the canvas."""
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib import figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    fig = figure.Figure(figsize=(8, 2.5), dpi=72)
    FigureCanvasAgg(fig)
    fig.patch.set_alpha(0.0)
    safe = mathtext_safe(tex)
    col01 = tuple(c / 255.0 for c in (col_rgb or (31, 31, 31)))
    t = fig.text(0.5, 0.5, "$%s$" % safe, fontsize=size_pt, color=col01,
                 math_fontfamily="stix", ha="center", va="center")
    fig.canvas.draw()
    buf = bytearray(fig.canvas.buffer_rgba())
    W_, H_ = fig.canvas.get_width_height()
    arr = Image.frombytes("RGBA", (W_, H_), bytes(buf))
    alpha = arr.split()[3]
    bbox = alpha.getbbox()
    if not bbox:
        return None
    img2 = arr.crop(bbox)
    # rendered at dpi=72: 1pt == 1px; scale to canvas px-per-pt
    k = px_per_pt
    nw, nh = max(1, int(img2.width * k)), max(1, int(img2.height * k))
    return img2.resize((nw, nh), Image.LANCZOS)


def mathtext_safe(tex: str) -> str:
    """Trim to a subset matplotlib mathtext can parse."""
    s = tex.replace("\n", " ").strip()
    s = s.replace("\\,", " ").replace("\\;", " ").replace("\\!", "")
    s = re.sub(r"\s+", " ", s)
    return s.replace(" - ", " - ")


def _draw_formula_linear(f, d):
    cjk = has_cjk(f["fb"])
    fnt = getfont(max(6, f["pt"] * 0.6), False, True, cjk, FONT_NAME)
    tc = col(f["fc"]) or (31, 31, 31)
    d.text((f["x"] + f["w"] / 2, f["y"] + f["h"] / 2), f["fb"], font=fnt,
           fill=tc, anchor="mm")


def _render_tex_safe(f, tc):
    return _render_tex_png(f["tex"], f["pt"], tc)


for f in formulas:
    tc = col(f["fc"]) or (31, 31, 31)
    try:
        rendered = _render_tex_safe(f, tc)
        if rendered is None:
            raise RuntimeError("empty")
        px = int(f["x"] + f["w"] / 2 - rendered.width / 2)
        py = int(f["y"] + f["h"] / 2 - rendered.height / 2)
        img.paste(rendered, (px, py), rendered)
    except Exception as exc:
        import traceback
        traceback.print_exc()
        _draw_formula_linear(f, d)

img.save(os.path.join(OUT, "preview_render.png"))

if SRC and os.path.exists(SRC):
    src = Image.open(SRC).convert("RGB").resize((CW, CH))
    cmp_img = Image.new("RGB", (CW, CH * 2 + 16), (225, 225, 225))
    cmp_img.paste(src, (0, 0))
    cmp_img.paste(img, (0, CH + 16))
    scale = 1500.0 / CW
    cmp_img = cmp_img.resize((1500, int((CH * 2 + 16) * scale)), Image.LANCZOS)
    cmp_img.save(os.path.join(OUT, "preview_compare.png"))

print("nodes=%d paths=%d canvas=%dx%d" % (len(nodes), len(paths), CW, CH))
xs = [n["x"] for n in nodes] + [n["x"] + n["w"] for n in nodes]
ys = [n["y"] for n in nodes] + [n["y"] + n["h"] for n in nodes]
print("extent x %.0f..%.0f  y %.0f..%.0f" % (min(xs), max(xs), min(ys), max(ys)))
