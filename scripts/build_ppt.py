# -*- coding: utf-8 -*-
"""sci-flowchart-vba step 5 -- turn the delivered .bas modules into a real .pptx.

Usage:
    python build_ppt.py <output_dir> [--out name.pptx] [--no-run] [--engine PATH]

Two paths, tried in this order:

  A. COM path (preferred, Windows + PowerPoint installed)
     Import modFlow_Engine + modFlow_Content* into a blank presentation and
     actually invoke BuildFlowchart(), so PowerPoint itself draws the shapes.
     Result is pixel-identical to what the user would get by pressing F5.

  B. Replay path (fallback, no Office needed)
     Parse AddNode / AddPath calls out of the .bas with the same grammar the
     skill uses for previews, then emit equivalent native shapes via python-pptx.
     Coordinates, colours, corner radii and arrowheads match the VBA semantics.

Exits non-zero if neither path produced a slide with shapes.
"""
from __future__ import annotations

import argparse
import os
import re
import sys

# ---------------------------------------------------------------- .bas parsing

PALETTE_RE = re.compile(
    r"Public\s+Const\s+(\w+)\s+As\s+Long\s*=\s*RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)",
    re.I,
)
SCALAR_RE = re.compile(
    r"Public\s+Const\s+(\w+)\s+As\s+(?:Single|Double|Boolean|String|Long)\s*=\s*(.+)",
    re.I,
)

# name -> msoAutoShapeType, mirrors ShapeTypeFor() in modFlow_Engine.bas
KIND_TO_Mso = {
    "rect": 1, "rectangle": 1, "box": 1,
    "round_rect": 5, "rounded": 5, "rounded_rect": 5,
    "oval": 9, "ellipse": 9, "circle": 9,
    "diamond": 4, "decision": 4,
    "parallelogram": 2,
    "trapezoid": 3,
    "pentagon": 51,
    "hexagon": 10,
    "stadium": 5, "pill": 5, "terminator": 5,
    "cylinder": 21,
    "document": 61,
    "right_arrow": 33, "arrow_right": 33,
    "left_arrow": 34, "arrow_left": 34,
    "up_arrow": 35, "arrow_up": 35,
    "down_arrow": 36, "arrow_down": 36,
    "chevron": 52,
}


def split_args(s: str) -> list[str]:
    """Split on top-level commas, ignoring commas inside parens or quotes."""
    out, cur, depth, q = [], [], 0, False
    for ch in s:
        if ch == '"':
            q = not q
        elif not q and ch == "(":
            depth += 1
        elif not q and ch == ")":
            depth -= 1
        if ch == "," and depth == 0 and not q:
            out.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if cur:
        out.append("".join(cur).strip())
    return out


def untext(t: str) -> str:
    """Rebuild "A" & vbLf & "B" into real newlines."""
    t = t.strip()
    parts = re.split(r'"\s*&\s*vbLf\s*&\s*"', t, flags=re.I)
    return "\n".join(p.strip().strip('"') for p in parts).strip()


def load_bas(out_dir: str):
    """Read every .bas in out_dir; return (palette, scalars, nodes, paths)."""
    files = sorted(f for f in os.listdir(out_dir) if f.lower().endswith(".bas"))
    # engine first so its constants never shadow content ones incorrectly;
    # content files sorted by trailing digit (Content < Content2 < Content3).
    def key(fn: str):
        stem = os.path.splitext(fn)[0]
        m = re.search(r"(\d+)$", stem)
        return (0 if "engine" in stem.lower() else 1, int(m.group(1)) if m else 1, stem)
    files.sort(key=key)

    pal: dict[str, tuple[int, int, int]] = {}
    scal: dict[str, str] = {}
    nodes, paths = [], []
    const_names: set[str] = {
        "LINE_SOLID", "LINE_DASH", "LINE_DOT", "LINE_DASHDOT",
        "True", "False", "Nothing",
    }

    for fn in files:
        raw = open(os.path.join(out_dir, fn), "rb").read().decode("latin-1")
        for m in PALETTE_RE.finditer(raw):
            pal[m.group(1).lower()] = (int(m.group(2)), int(m.group(3)), int(m.group(4)))
        for m in SCALAR_RE.finditer(raw):
            scal[m.group(1).lower()] = m.group(2).strip()
        # join VBA line continuations ("_" at end of line) so multi-line
        # AddNode calls parse as a single logical line
        logical = re.sub(r"_\s*\r?\n\s*", " ", raw)
        for line in logical.splitlines():
            line = line.strip()
            if line.startswith("AddNode "):
                nodes.append(parse_node(split_args(line[8:])[1:]))
            elif line.startswith("AddPath "):
                paths.append(parse_path(split_args(line[8:])))
    return pal, scal, nodes, paths


def num(tok: str, default: float = 0.0) -> float:
    tok = tok.strip()
    m = re.match(r"^[+-]?[\d.]+", tok)
    return float(m.group(0)) if m else default


def parse_node(args: list[str]) -> dict:
    """args = [id, kind, x, y, w, h, fill, line, lw, dash, text?, pt?, bold?, ...]"""
    n = dict(
        id=args[0].strip('"'), kind=args[1].strip('"').lower(),
        x=num(args[2]), y=num(args[3]), w=num(args[4]), h=num(args[5]),
        fill=args[6] if len(args) > 6 else "-1",
        line=args[7] if len(args) > 7 else "-1",
        lw=num(args[8], 1.0) if len(args) > 8 else 1.0,
        dash=args[9] if len(args) > 9 else "LINE_SOLID",
    )
    rest = args[10:]
    n["text"] = untext(rest[0]) if rest and rest[0] else ""
    n["pt"] = num(rest[1], 11.0) if len(rest) > 1 and rest[1] else 11.0
    n["bold"] = len(rest) > 2 and rest[2].strip().lower() == "true"
    n["ital"] = len(rest) > 3 and rest[3].strip().lower() == "true"
    n["fc"] = rest[4] if len(rest) > 4 and rest[4] else "-1"
    n["corner"] = num(rest[5], -1.0) if len(rest) > 5 and rest[5] else -1.0
    n["adj2"] = num(rest[7], -1.0) if len(rest) > 7 and rest[7] else -1.0
    return n


def parse_path(args: list[str]) -> dict:
    """args = [sld, id, "x1,y1;x2,y2", color, lw, dash, arrowEnd]"""
    pts = []
    for pair in args[2].strip('"').split(";"):
        xy = split_args(pair)
        if len(xy) >= 2:
            pts.append((num(xy[0]), num(xy[1])))
    return dict(
        id=args[1].strip('"'), pts=pts,
        color=args[3] if len(args) > 3 else "-1",
        lw=num(args[4], 1.0) if len(args) > 4 else 1.0,
        arrow=len(args) > 6 and args[6].strip().lower() == "true",
    )


# ---------------------------------------------------------------- geometry

def resolve_color(tok: str, pal: dict):
    tok = (tok or "").strip()
    low = tok.lower()
    if low in pal:
        return pal[low]
    m = re.match(r"RGB\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)", tok, re.I)
    if m:
        return tuple(int(g) for g in m.groups())
    if tok in ("-1", ""):
        return None
    return (0, 0, 0)


def resolve_scalar(tok: str, scal: dict, default: float) -> float:
    """Resolve a numeric literal that may reference another Public Const."""
    tok = (tok or "").strip()
    if not tok:
        return default
    if tok.lower() in scal:
        return resolve_scalar(scal[tok.lower()], scal, default)
    return num(tok, default)


def canvas_geometry(scal: dict):
    """Reproduce the 9 geometry constants the engine relies on."""
    cw = resolve_scalar(scal.get("canvas_w_px", "1200"), scal, 1200.0)
    ch = resolve_scalar(scal.get("canvas_h_px", "675"), scal, 675.0)
    slide_w_in = resolve_scalar(scal.get("slide_w_in", "13.333"), scal, 13.333)
    slide_h_in = resolve_scalar(scal.get("slide_h_in", "7.5"), scal, 7.5)
    custom = scal.get("custom_size", "False").strip().lower() == "true"
    margin = 0.35
    avail_w = slide_w_in - 2 * margin
    avail_h = slide_h_in - 2 * margin
    px_to_in = resolve_scalar(scal.get("px_to_in", ""), scal, 0.0)
    if px_to_in <= 0:
        px_to_in = min(avail_w / cw, avail_h / ch)
    off_x = resolve_scalar(scal.get("offset_x_in", ""), scal, -1.0)
    off_y = resolve_scalar(scal.get("offset_y_in", ""), scal, -1.0)
    if off_x < 0:
        off_x = margin + (avail_w - cw * px_to_in) / 2
    if off_y < 0:
        off_y = margin + (avail_h - ch * px_to_in) / 2
    font_name = scal.get("font_name", '"Arial"').strip().strip('"') or "Arial"
    return dict(cw=cw, ch=ch, sw=slide_w_in, sh=slide_h_in, custom=custom,
                s=px_to_in, ox=off_x, oy=off_y, font=font_name)


# ---------------------------------------------------------------- path A: COM

ENGINE_NAME = "modFlow_Engine"
CONTENT_PREFIX = "modFlow_Content"


def is_powerpoint_available() -> bool:
    if os.name != "nt":
        return False
    # do not let gen_py cache shadow the real typelib
    import importlib
    try:
        mod = importlib.import_module("win32com.client")
    except ImportError:
        return False
    try:
        mod.gencache.EnsureDispatch("PowerPoint.Application")
        return True
    except Exception:
        return False


def bas_text(path: str) -> str:
    return open(path, "rb").read().decode("latin-1")


def module_name(text: str, fallback: str) -> str:
    m = re.search(r"Attribute\s+VB_Name\s*=\s*\"([^\"]+)\"", text)
    return m.group(1) if m else fallback


def build_with_com(out_dir: str, out_path: str, bas_files: list[str], run: bool) -> int:
    import win32com.client as w32

    app = w32.gencache.EnsureDispatch("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Add()
    # start from a blank 16:9 board; BuildFlowchart resizes when CUSTOM_SIZE
    pres.PageSetup.SlideWidth = 13.333 * 72
    pres.PageSetup.SlideHeight = 7.5 * 72

    # order matters: engine first, then content modules by trailing index
    def key(p: str):
        stem = os.path.splitext(os.path.basename(p))[0]
        m = re.search(r"(\d+)$", stem)
        return (0 if "engine" in stem.lower() else 1, int(m.group(1)) if m else 1, stem)
    bas_files = sorted(bas_files, key=key)

    vbproj = pres.VBProject
    for p in bas_files:
        text = bas_text(p)
        name = module_name(text, os.path.splitext(os.path.basename(p))[0])
        comp = vbproj.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = name
        # strip the Attribute line: a component we just created already carries it
        body = re.sub(r"^Attribute\s+VB_Name\s*=.*\r?\n", "", text, count=1)
        comp.CodeModule.AddFromString(body)
        print("  COM: imported %s" % name)

    if not run:
        pres.SaveAs(out_path)
        print("  COM: saved template (BuildFlowchart NOT executed)")
        return 0

    try:
        app.Run("BuildFlowchart")
    except Exception as exc:  # surface the VBA error instead of hiding it
        print("  COM: BuildFlowchart raised -> %s" % exc, file=sys.stderr)
        pres.SaveAs(out_path)
        return 2

    pptm = os.path.splitext(out_path)[0] + ".pptm"
    try:
        pres.SaveAs(pptm)
    except Exception as exc:
        print("  COM: SaveAs .pptm failed -> %s" % exc, file=sys.stderr)
    try:
        if os.path.exists(pptm):
            tmp = app.Presentations.Open(pptm, WithWindow=False)
            tmp.SaveAs(out_path, 24)  # ppSaveAsOpenXMLPresentation
            tmp.Close()
    except Exception as exc:
        print("  COM: .pptx conversion failed -> %s" % exc, file=sys.stderr)
    return 0


# ---------------------------------------------------------------- path B: replay

def build_with_replay(out_dir: str, out_path: str) -> int:
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Inches, Pt

    pal, scal, nodes, paths = load_bas(out_dir)
    g = canvas_geometry(scal)
    if nodes or paths:
        print("  replay: %d nodes, %d paths" % (len(nodes), len(paths)))

    prs = Presentation()
    prs.slide_width = Inches(g["sw"])
    prs.slide_height = Inches(g["sh"])
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    def X(px):
        return Inches(g["ox"] + px * g["s"])

    def Y(py):
        return Inches(g["oy"] + py * g["s"])

    def L(v):
        return Inches(v * g["s"])

    for n in nodes:
        mso_id = KIND_TO_Mso.get(n["kind"], 1)
        try:
            shape_type = MSO_SHAPE(mso_id)
        except ValueError:
            shape_type = MSO_SHAPE.RECTANGLE
        try:
            shp = slide.shapes.add_shape(
                shape_type, X(n["x"]), Y(n["y"]), L(n["w"]), L(n["h"]))
        except Exception:
            # stadium/pill on some python-pptx versions: fall back to round rect
            shp = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, X(n["x"]), Y(n["y"]), L(n["w"]), L(n["h"]))

        fill = resolve_color(n["fill"], pal)
        if fill is None:
            shp.fill.background()
        else:
            shp.fill.solid()
            shp.fill.fore_color.rgb = RGBColor(*fill)

        linec = resolve_color(n["line"], pal)
        if linec is None:
            shp.line.fill.background()
        else:
            shp.line.color.rgb = RGBColor(*linec)
            shp.line.width = Pt(max(0.5, n["lw"]))
            if n["dash"].upper() in ("LINE_DASH", "LINE_DASHDOT", "LINE_DOT"):
                from pptx.enum.dml import MSO_LINE_DASH_STYLE
                shp.line.dash_style = (
                    MSO_LINE_DASH_STYLE.DASH if n["dash"].upper() == "LINE_DASH"
                    else MSO_LINE_DASH_STYLE.ROUND_DOT if n["dash"].upper() == "LINE_DOT"
                    else MSO_LINE_DASH_STYLE.DASH_DOT)

        corner = 0.5 if n["kind"] in ("stadium", "pill", "terminator") else n["corner"]
        if corner and corner > 0:
            try:
                shp.adjustments[0] = max(0.0, min(0.5, corner))
            except Exception:
                pass
        if n["adj2"] and n["adj2"] > 0:
            try:
                shp.adjustments[1] = max(0.0, min(1.0, n["adj2"]))
            except Exception:
                pass

        tf = shp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Pt(1)
        tf.margin_top = tf.margin_bottom = Pt(0)
        if n["text"]:
            tf.text = n["text"]
            tc = resolve_color(n["fc"], pal) or (31, 31, 31)
            for i, para in enumerate(tf.paragraphs):
                para.alignment = PP_ALIGN.CENTER
                for r in para.runs:
                    r.font.size = Pt(n["pt"])
                    r.font.bold = n["bold"]
                    r.font.italic = n["ital"]
                    r.font.name = g["font"]
                    r.font.color.rgb = RGBColor(*tc)

    for p in paths:
        pts = p["pts"]
        c = resolve_color(p["color"], pal) or (0, 0, 0)
        for i in range(len(pts) - 1):
            (x0, y0), (x1, y1) = pts[i], pts[i + 1]
            conn = slide.shapes.add_connector(
                MSO_CONNECTOR.STRAIGHT, X(x0), Y(y0), X(x1), Y(y1))
            conn.line.color.rgb = RGBColor(*c)
            conn.line.width = Pt(max(0.5, p["lw"]))
            if p["arrow"] and i == len(pts) - 2:
                # python-pptx exposes no arrowhead API: set it via the XML
                ln = conn.line._get_or_add_ln()
                from pptx.oxml.ns import qn
                tail = ln.makeelement(qn("a:tailEnd"),
                                      {"type": "triangle", "w": "med", "len": "med"})
                ln.append(tail)

    prs.save(out_path)
    return 0


# ---------------------------------------------------------------- driver

def slide_shape_count(pptx_path: str) -> int:
    try:
        from pptx import Presentation
    except ImportError:
        return -1
    prs = Presentation(pptx_path)
    return sum(len(s.shapes) for s in prs.slides)


def main() -> int:
    ap = argparse.ArgumentParser(description="sci-flowchart-vba: .bas -> .pptx")
    ap.add_argument("out_dir", help="directory holding modFlow_Engine.bas + modFlow_Content*.bas")
    ap.add_argument("--out", default="flowchart.pptx", help="output pptx file name")
    ap.add_argument("--no-run", action="store_true",
                    help="COM path only: build the file without executing BuildFlowchart")
    ap.add_argument("--replay", action="store_true",
                    help="force the no-Office replay path even if PowerPoint is installed")
    args = ap.parse_args()

    out_dir = os.path.abspath(args.out_dir)
    if not os.path.isdir(out_dir):
        print("output dir not found: %s" % out_dir, file=sys.stderr)
        return 2
    bas_files = [os.path.join(out_dir, f) for f in sorted(os.listdir(out_dir))
                 if f.lower().endswith(".bas")]
    if not bas_files:
        print("no .bas files in %s" % out_dir, file=sys.stderr)
        return 2
    engine = [p for p in bas_files if "engine" in os.path.basename(p).lower()]
    content = [p for p in bas_files if "content" in os.path.basename(p).lower()]
    if not engine:
        print("warning: modFlow_Engine.bas missing from %s" % out_dir, file=sys.stderr)
    if not content:
        print("warning: no modFlow_Content*.bas found in %s" % out_dir, file=sys.stderr)

    out_path = args.out if os.path.isabs(args.out) else os.path.join(out_dir, args.out)
    print("sci-flowchart-vba :: %s -> %s" % (out_dir, out_path))

    used = None
    if not args.replay and is_powerpoint_available():
        print("[path A] PowerPoint COM automation")
        try:
            rc = build_with_com(out_dir, out_path, bas_files, run=not args.no_run)
            used = "COM"
            if rc != 0 and not os.path.exists(out_path):
                used = None
        except Exception as exc:
            print("  COM path failed -> %s" % exc, file=sys.stderr)
            used = None

    if used is None:
        print("[path B] .bas replay via python-pptx (no PowerPoint needed)")
        try:
            build_with_replay(out_dir, out_path)
            used = "replay"
        except ImportError as exc:
            print("python-pptx is required for the replay path: %s" % exc, file=sys.stderr)
            return 2

    if not os.path.exists(out_path):
        print("FAILED: %s was not written" % out_path, file=sys.stderr)
        return 1
    count = slide_shape_count(out_path)
    size_kb = os.path.getsize(out_path) / 1024.0
    print("OK  [%s]  %s  (%.1f KB, %d shapes)" % (used, out_path, size_kb, count))
    if count == 0:
        print("FAILED: the slide has no shapes -- the flowchart was not drawn", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
