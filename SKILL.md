---
name: bruce-svg-to-pptx
description: Use when the user wants to convert SVG files into native editable PowerPoint shapes on Windows, or wants to insert SVG-based content into an existing branded PPTX template (preserving cover/TOC/layout slides while adding or replacing content pages). Triggers include "SVG → 形状", "SVG 转 PPT 可编辑形状", "插入到现有 PPT", "基于模板生成", batch-converting SVG icons, expanding a deck with more content pages, or any request to make SVG content editable inside PowerPoint. Windows only.
---

# SVG to editable PowerPoint shapes (Windows)

## What this skill does

Drives Microsoft PowerPoint on Windows via COM to perform the same operation a
user would do manually:

1. Insert each SVG onto a slide as a `msoGraphic` shape.
2. Select it and call `CommandBars.ExecuteMso("SVGEdit")` — PowerPoint's
   internal "Convert to Shape" Ribbon command. There is no first-class VBA
   method; this Mso command is the only programmatic hook Microsoft exposes.
3. Save the deck. The result is identical to right-click → Convert to Shape:
   each SVG becomes a group of native DrawingML shapes that can be ungrouped,
   recoloured, animated, and edited individually.

The skill ships **two scripts**:

| Script | Purpose |
|---|---|
| `scripts/Convert-SvgToShapes.ps1` | Create a **new** .pptx from one or more SVGs |
| `scripts/Edit-ExistingPptx.ps1` | Insert SVGs into an **existing** branded template |

## When to use which script

```
User wants a brand-new deck from SVGs?
  → Convert-SvgToShapes.ps1

User has a branded template (cover / TOC / content slides) and wants to
fill it with SVG content, or expand it with more content pages?
  → Edit-ExistingPptx.ps1
```

## Prerequisites

- Windows 10 / 11
- Microsoft PowerPoint 2016 build 1712 or later (any current Microsoft 365 /
  Office 2019 / 2021 / 2024 install qualifies)
- PowerShell 5.1+ (ships with Windows) or PowerShell 7

If PowerPoint is not installed the script exits with a clear error message.

---

## Script 1 — `Convert-SvgToShapes.ps1`

Creates a new .pptx with one slide per SVG, or appends slides to an existing deck.

### Parameters

| Parameter | Required | Description |
|---|---|---|
| `-SvgPath` | ✓ | `.svg` file(s) or a directory. Array accepted. |
| `-OutputPath` | ✓ | Destination `.pptx`. Created if absent. |
| `-Append` | | Append to existing `-OutputPath` instead of overwriting. |
| `-Force` | | Skip overwrite confirmation. |

### Usage

```powershell
# Single SVG → new deck
pwsh -File scripts\Convert-SvgToShapes.ps1 `
    -SvgPath .\icon.svg `
    -OutputPath .\out.pptx

# Batch: folder of SVGs, one slide each
pwsh -File scripts\Convert-SvgToShapes.ps1 `
    -SvgPath .\icons\ `
    -OutputPath .\icons.pptx

# Append to an existing deck
pwsh -File scripts\Convert-SvgToShapes.ps1 `
    -SvgPath .\new-icons\ `
    -OutputPath .\existing.pptx `
    -Append
```

Use `powershell.exe` instead of `pwsh` when PowerShell 7 is not installed.

---

## Script 2 — `Edit-ExistingPptx.ps1`

Inserts SVG-converted shapes into an existing branded PPTX, preserving all
template formatting (backgrounds, master layouts, cover slide, TOC slide,
footers, logos, etc.).

### Two operation modes

**INSERT mode** — target specific existing slides (one SVG per slide).
Triggered when `-TargetSlide` is provided.

**EXPAND mode** (default) — duplicate a "content template" slide once per SVG,
place the copies at the desired position, then insert one SVG into each copy.
Triggered automatically when `-TargetSlide` is omitted.

### Parameters

| Parameter | Mode | Default | Description |
|---|---|---|---|
| `-TemplatePath` | both | — | Existing `.pptx` to edit (required). |
| `-SvgPath` | both | — | `.svg` file(s) or directory (required). |
| `-OutputPath` | both | TemplatePath | Destination `.pptx`. Omit to edit in-place. |
| `-TargetSlide` | INSERT | — | First slide index (1-based) to insert into. Additional SVGs → consecutive slides. |
| `-ContentSlide` | EXPAND | last slide | Slide index to duplicate as the content page template. |
| `-InsertAfterSlide` | EXPAND | last slide | New slides are inserted after this index. |
| `-ContentZone` | both | auto-detect | SVG placement area: `"Left,Top,Width,Height"` in points. |
| `-SlideTitles` | both | — | Array of title strings, one per SVG. Sets each slide's title placeholder. Blank/null entries keep the template title. |
| `-ClearContent` | both | off | Delete all non-structural shapes before inserting SVG. Keeps title, footer, date, slide-number placeholders. |
| `-NoBackup` | both | off | Skip creating `.bak.pptx` backup of the original. |
| `-Force` | both | off | Skip overwrite confirmation for `-OutputPath`. |

### Content zone auto-detection

When `-ContentZone` is not provided, the script searches the reference slide for
a content placeholder (body, object, chart, picture, media). If none is found it
falls back to:

```
Left  = 36 pt  (0.5")
Top   = 108 pt (1.5" — clears a standard title bar)
Width  = SlideWidth  − 72 pt
Height = SlideHeight − 126 pt
```

Override with explicit points whenever the template uses a non-standard layout.
To find point values: in PowerPoint, right-click a shape → Size and Position;
multiply inches × 72 to get points.

### Backup behaviour

A `.bak.pptx` copy of the original template is created automatically before any
edit. Suppress with `-NoBackup`. The backup lives next to the original:

```
branded-template.pptx      ← original (will be edited / overwritten if OutputPath matches)
branded-template.bak.pptx  ← backup created by the script
```

### Usage examples

```powershell
# ── EXPAND mode (most common) ────────────────────────────────
# Template has: slide 1 = cover, slide 2 = TOC, slide 3 = content template.
# Insert 5 SVGs as new content pages, duplicating slide 3, after slide 3.

pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -SvgPath      .\content-slides\ `
    -ContentSlide 3 `
    -InsertAfterSlide 3 `
    -ClearContent

# ── EXPAND mode with separate output ────────────────────────
# Keep the original template untouched; write result to a new file.

pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -SvgPath      .\slides\ `
    -ContentSlide 3 `
    -OutputPath   .\presentation-final.pptx `
    -ClearContent

# ── EXPAND mode — append after last slide (defaults) ────────
# Simplest invocation: duplicate the last slide for each SVG.

pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -SvgPath      .\slides\

# ── INSERT mode — put one SVG into an existing slide ─────────
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\deck.pptx `
    -SvgPath      .\chart.svg `
    -TargetSlide  5

# ── INSERT mode — replace content in slides 3, 4, 5 ─────────
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\deck.pptx `
    -SvgPath      .\slide1.svg, .\slide2.svg, .\slide3.svg `
    -TargetSlide  3 `
    -ClearContent

# ── Custom content zone (explicit points) ───────────────────
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -SvgPath      .\slides\ `
    -ContentZone  "54,126,852,378" `
    -ClearContent

# ── Set slide titles for each duplicated page ────────────────
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -SvgPath      .\slides\ `
    -ContentSlide    3 `
    -InsertAfterSlide 2 `
    -ClearContent `
    -SlideTitles  "市场概述", "竞争格局", "战略规划", "执行路径"
```

### Typical workflow for branded templates

A common PPTX structure and recommended invocation:

```
Slide 1  — Cover (封面)
Slide 2  — Table of contents (目录)
Slide 3  — Content template (正文内容，含标题占位符和内容区占位符)
```

```powershell
# Step 1: Generate SVG files for each content page (Claude or any tool).
# Step 2: Run the script to expand the deck.

pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\company-deck.pptx `
    -SvgPath         .\generated-slides\ `
    -ContentSlide    3 `
    -InsertAfterSlide 2 `
    -ClearContent `
    -SlideTitles "第一章：市场概述", "第二章：竞争格局", "第三章：战略规划"
# Result:
# Slide 1 — Cover   (unchanged)
# Slide 2 — TOC     (unchanged)
# Slide 3 — SVG 1   (duplicated from original slide 3, content replaced)
# Slide 4 — SVG 2   (duplicated from original slide 3, content replaced)
# ...
# Slide N — original slide 3 template (now at end, can be deleted manually)
```

> **Tip:** To discard the original content template after expansion, insert
> after the last real content slide (e.g. `-InsertAfterSlide 2`). The original
> template slide migrates to the end; delete it manually in PowerPoint if
> unneeded.

---

## How Claude should invoke these scripts

1. Identify which script applies (new deck vs. existing template).
2. Resolve all file paths to **absolute paths** before passing to the script;
   COM is unreliable with relative paths.
3. Run via the Bash tool using PowerShell:
   ```powershell
   pwsh -File "C:\path\scripts\Edit-ExistingPptx.ps1" -TemplatePath "C:\..." ...
   # or for PowerShell 5.1:
   powershell.exe -File "C:\path\scripts\Edit-ExistingPptx.ps1" ...
   ```
4. Surface `stderr` verbatim if the script fails. The error messages distinguish
   between "PowerPoint not installed", "file not found", "slide out of range",
   and "ExecuteMso failed".
5. After a successful run, confirm the output path and mention the fidelity
   caveat (see Known limitations).

---

## Known limitations (both scripts)

- **PowerPoint window must be visible.** `ExecuteMso` is a Ribbon command;
  it is unreliable when the window is hidden. Both scripts set `Visible = True`
  and close cleanly at the end.
- **Close other presentations first.** COM attaches to the running instance.
  Open presentations can interfere with selection state. The scripts warn but
  do not abort.
- **No headless / CI support.** The visible-window requirement prevents
  server/headless use. For headless conversion use the `svg2pptx` Python library
  or a pure-DrawingML approach instead.
- **SVG fidelity matches PowerPoint's built-in converter.** Paths, basic shapes,
  and solid/simple-gradient fills convert well. Complex filters, `clipPath`
  chains, embedded raster images, and advanced gradient stops may be flattened
  or dropped. This is a PowerPoint limitation; simplify the SVG (re-export
  without effects from Figma/Illustrator) if quality is poor.
- **ExecuteMso timing.** A 300 ms pause is inserted before `ExecuteMso("SVGEdit")`
  to let the selection settle. On slower machines a shape may silently remain
  unconverted. If this happens, increase `Start-Sleep -Milliseconds 300` to
  `500` or more in the script.
- **One SVG per slide.** Each SVG occupies its own slide. Multi-SVG layouts on
  a single slide are out of scope.
- **Windows only.** Both scripts require Windows + COM automation. There is no
  macOS/Linux equivalent; suggest `svg2pptx` or DrawingML libraries for
  cross-platform needs.

## Verifying the output

Open the `.pptx` and right-click any inserted shape. If the context menu shows
**Edit Points** or the shape can be ungrouped into individual sub-shapes, the
conversion succeeded. If the menu still shows **Convert to Shape**, that shape
did not convert (usually a malformed SVG); the script's stderr will have noted it.
