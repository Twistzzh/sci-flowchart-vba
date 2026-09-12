# sci-flowchart-vba

A WorkBuddy skill that turns SCI-style flowchart figures (and paper text) into
**editable PowerPoint files** — a real `.pptx` plus the VBA source that can
rebuild it at any time.

## What it does

Given a flowchart image, the skill produces VBA code that draws native,
editable PowerPoint shapes (not a picture). Paste the `.bas` modules into
PowerPoint's VBA editor and run `BuildFlowchart`, or just open the generated
`.pptx` — the shapes are already drawn and fully editable.

Two modes:

- **Mode 1 — plan + generate**: supply a style template figure + a paper
  paragraph. The skill borrows the template's visual style (palette, rounded
  corners, stroke, font) and the paragraph's content to author a new flowchart.
- **Mode 2 — replicate**: supply an existing flowchart image; the skill rebuilds
  its logical structure in SCI style.

## Why vision-direct

The old pipeline rasterized the figure to SVG first, then parsed it through
three scripts. That intermediate step lost shape semantics, exact colors, and
text. This skill drops SVG entirely: the image goes straight to a multimodal
model, which emits the *content module* (`modFlow_Content.bas`). A fixed engine
(`assets/modFlow_Engine.bas`) owns all the hard parts — coordinate mapping,
shape creation, connectors, text, idempotent cleanup, and the entry point.

## Layout

```
sci-flowchart-vba/
  SKILL.md              # skill manifest + full contract for the code model
  assets/
    modFlow_Engine.bas  # fixed engine (do not rewrite)
  references/
    example_content.bas # runnable content-module example
    vba-module-map.md   # engine API table + content-module contract
  scripts/
    build_ppt.py        # final step: .bas -> real .pptx (COM, replay fallback)
    lint_vba.py         # pre-delivery structural self-check
```

## Quick start (as a user)

1. Get the `.pptx` (the skill already generates it via `build_ppt.py`).
2. Or rebuild yourself: in PowerPoint press `Alt+F11`, right-click the project,
   **Import File** each `.bas` (engine + content), then `F5` → `BuildFlowchart`.
3. All shapes are editable. Run `RemoveFlowchart` to clear and redraw.

## Contract highlights (for the code model)

- Content module must declare 9 geometry constants and a single
  `Public Sub DrawAll(sld As Slide)`.
- Colors are `RGB(r,g,b)` literals; no `vbRed` etc.
- `.bas` files are **pure ASCII + CRLF** — Chinese in strings or comments
  corrupts the file when imported into the VBA editor.
- Split into `modFlow_Content2.bas` when nodes > 18 or a module > 300 lines.

See `SKILL.md` for the full code-generation contract.

## License

MIT — do whatever you like; attribution appreciated.
