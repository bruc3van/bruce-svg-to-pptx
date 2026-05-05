---
name: bruce-svg-to-pptx
description: Use when the user wants to convert SVG files into native editable PowerPoint shapes on Windows, or wants to insert SVG-based content into an existing branded PPTX template (preserving cover/TOC/layout slides while editing text or adding content pages). Triggers include "SVG → 形状", "SVG 转 PPT 可编辑形状", "插入到现有 PPT", "基于模板生成", batch-converting SVG icons, expanding a deck with more content pages, editing only the text on cover/TOC pages, or any request to make SVG content editable inside PowerPoint. Windows only.
---

# SVG to editable PowerPoint shapes (Windows)

## Mental model

The skill's core idea: the LLM lays out a full content page **as an SVG**
(graphics + text), PowerPoint imports that SVG and converts it into native
editable shapes via `CommandBars.ExecuteMso("SVGEdit")`. The result is a real
PowerPoint group of DrawingML shapes — ungroup-able, recolour-able,
animate-able, fully editable.

When the user has a **branded template** (corporate cover, TOC, content
layouts), the goal shifts: the existing branded slides should be preserved.
Different page types take different paths:

```
封面 / 标题页 / 目录页    →  TEXT mode
                          仅替换占位符里的文字（标题、副标题、目录条目…）
                          不插入 SVG。整页布局来自模板。

正文 / 内容页             →  EXPAND mode
                          复制模板的"内容页"做底稿，保留它的标题栏、背景、
                          页脚等品牌元素；在中间的安全区里嵌入一张较小的
                          LLM 生成 SVG，再把它转换成可编辑形状。

完整一份演示稿            →  MANIFEST mode
                          一份 JSON 描述每一页的编辑（封面文字 + 目录文字
                          + N 张正文页插图），单次调用一次性产出。
```

## What this skill does

Drives Microsoft PowerPoint on Windows via COM to do the same operation a
user would do manually:

1. Insert each SVG onto a slide as a `msoGraphic` shape.
2. Select it and call `CommandBars.ExecuteMso("SVGEdit")` — PowerPoint's
   internal "Convert to Shape" Ribbon command. There is no first-class VBA
   method; this Mso command is the only programmatic hook Microsoft exposes.
3. Save the deck.

The skill ships **two scripts**:

| Script | Purpose |
|---|---|
| `scripts/Convert-SvgToShapes.ps1` | Create a **new** .pptx from one or more SVGs (one full-slide SVG per slide). |
| `scripts/Edit-ExistingPptx.ps1`   | Edit text and/or insert SVG content pages into an **existing** branded template. Four modes: TEXT, INSERT, EXPAND, MANIFEST. |

## Decision tree — picking the right invocation

```
User request                                                  → Use
────────────────────────────────────────────────────────────────────────
"Generate a brand-new deck from these SVGs"                   → Convert-SvgToShapes.ps1
"Batch-convert these SVG icons into one deck"                 → Convert-SvgToShapes.ps1

"I have a branded template; I only need to change the text
 on the cover / TOC / title page"                             → Edit-ExistingPptx.ps1   (TEXT mode)

"I have a branded template; replace the content on slide N
 with this SVG"                                               → Edit-ExistingPptx.ps1   (INSERT mode)

"I have a branded template; duplicate the content layout
 N times, one SVG per duplicated page"                        → Edit-ExistingPptx.ps1   (EXPAND mode)

"Generate a full deck from a branded template — cover text +
 TOC text + N content pages with SVGs — in one shot"          → Edit-ExistingPptx.ps1   (MANIFEST mode)
```

## SVG generation guidelines (for the LLM)

PowerPoint's built-in SVG → shape converter is the bottleneck. To get clean,
fully-editable output, the SVG you (or the LLM) produce **must** stay within
its supported subset:

**Safe to use**

- `<path>`, `<rect>`, `<circle>`, `<ellipse>`, `<line>`, `<polygon>`, `<polyline>`
- `<text>` with `font-family`, `font-size`, `font-weight`, `fill`
- Solid `fill` and `stroke` (hex, rgb)
- Simple `linearGradient` (2–3 stops)
- `<g>` grouping with basic transforms (`translate`, `scale`, `rotate`)

**Avoid (will be flattened, dropped, or fail conversion)**

- `<filter>` (drop-shadow, blur, etc.) — flattens to a raster
- `<clipPath>`, `<mask>` — often dropped silently
- `<foreignObject>` — never supported
- `<image href="…">` (embedded raster) — defeats the purpose; the result becomes a picture, not editable shapes
- `radialGradient`, mesh gradients, gradient with >3 stops
- CSS in `<style>` blocks — inline `style=` or attributes only
- Web fonts / `@font-face` — PowerPoint won't substitute; stick to system fonts (Arial, Calibri, 思源黑体, 微软雅黑) or convert text to paths
- `pattern` fills — frequently lost

**Layout rule for content-page SVGs (EXPAND mode)**

The SVG is inserted into a **safe content zone**, not the full slide. By
default that zone is the body placeholder of the content-template slide, or a
margined area below the title bar if there is no body placeholder. **Design
the SVG to fill its own viewBox without assuming a specific slide size** — the
script preserves aspect ratio (letterbox) and centres the SVG inside the zone.

## Prerequisites

- Windows 10 / 11
- Microsoft PowerPoint 2016 build 1712+ (any current Microsoft 365 / Office
  2019 / 2021 / 2024 install qualifies)
- PowerShell 5.1 (ships with Windows) or PowerShell 7 (`pwsh`, **recommended**
  for Chinese strings — UTF-8 by default)

If PowerPoint is not installed the script exits with a clear error.

---

## Script 1 — `Convert-SvgToShapes.ps1`

Creates a new .pptx with one slide per SVG, or appends slides to an existing
deck. Each SVG fills its slide (letterboxed).

| Parameter | Required | Description |
|---|---|---|
| `-SvgPath`    | ✓ | `.svg` file(s) or a directory. Array accepted. |
| `-OutputPath` | ✓ | Destination `.pptx`. Created if absent. |
| `-Append`     |   | Append to existing `-OutputPath` instead of overwriting. |
| `-Force`      |   | Skip overwrite confirmation. |

```powershell
# Single SVG → new deck
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\icon.svg -OutputPath .\out.pptx

# Folder of SVGs, one slide each
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\icons\ -OutputPath .\icons.pptx

# Append to an existing deck
pwsh -File scripts\Convert-SvgToShapes.ps1 -SvgPath .\new-icons\ -OutputPath .\existing.pptx -Append
```

---

## Script 2 — `Edit-ExistingPptx.ps1`

Four operation modes, picked automatically from the parameters supplied.

### Common parameters (apply to all modes)

| Parameter        | Description |
|---|---|
| `-TemplatePath`  | Existing `.pptx` to edit (required). |
| `-OutputPath`    | Destination. Defaults to TemplatePath (in-place edit, after backup). |
| `-NoBackup`      | Skip creating `<TemplatePath>.bak.pptx`. |
| `-Force`         | Skip overwrite confirmation when OutputPath differs from TemplatePath. |

### TEXT mode — edit text on existing slides only (no SVG)

Triggered when `-SlideTexts` is supplied **without** `-SvgPath`. Use this for
covers, title pages, and TOCs where you only need to change wording.

`-SlideTexts` is a hashtable keyed by 1-based slide index. Each value is a
hashtable describing what to write into which placeholder type:

| Key        | Type       | Target placeholder |
|------------|------------|--------------------|
| `Title`    | `string`   | first PP_TITLE / PP_CENTER_TITLE |
| `Subtitle` | `string`   | first PP_SUBTITLE |
| `Body`     | `string[]` | Nth PP_BODY / PP_OBJECT (in slide order) |
| `Date`     | `string`   | first PP_DATE |
| `Footer`   | `string`   | first PP_FOOTER |

```powershell
# Edit cover (slide 1) and TOC (slide 2) only — no SVG
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -OutputPath   .\out.pptx `
    -SlideTexts @{
        1 = @{ Title = '2026 战略报告'; Subtitle = '董事会汇报'; Date = '2026-05' }
        2 = @{ Title = '目录'; Body = @('市场概述','竞争格局','战略规划','执行路径') }
    }
```

### INSERT mode — replace one slide's content with one SVG

Triggered by `-TargetSlide`. Multiple SVGs map to consecutive slides starting
at TargetSlide.

| Parameter        | Description |
|---|---|
| `-SvgPath`       | `.svg` file(s) or directory. |
| `-TargetSlide`   | 1-based first-target slide. |
| `-ContentZone`   | Override SVG zone: `"Left,Top,Width,Height"` in points. |
| `-ClearContent`  | Strip non-structural shapes before inserting. |
| `-SlideTitles`   | One title per SVG (legacy shortcut). |
| `-SlideTexts`    | Per-SVG hashtable (key = 1-based SVG index). |

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\deck.pptx `
    -SvgPath      .\chart.svg `
    -TargetSlide  5 `
    -ClearContent
```

### EXPAND mode — duplicate a content-template slide N times

Triggered when `-SvgPath` is supplied **without** `-TargetSlide`. Each SVG
becomes a new slide that copies all the branded furniture (title, background,
footer) from the chosen content-template slide; the SVG is centred inside the
content zone (placeholder-detected or default margin below the title).

| Parameter           | Default     | Description |
|---|---|---|
| `-SvgPath`          | —           | `.svg` file(s) or directory. |
| `-ContentSlide`     | last slide  | Index to duplicate as the content template. |
| `-InsertAfterSlide` | last slide  | New slides go after this index. |
| `-ContentZone`      | auto-detect | Override placement: `"Left,Top,Width,Height"` in points. |
| `-ClearContent`     | off         | Strip non-structural shapes from each duplicate before inserting. |
| `-SlideTitles`      | —           | One title per SVG (legacy shortcut). |
| `-SlideTexts`       | —           | Per-SVG hashtable (key = 1-based SVG index). |

```powershell
# Duplicate slide 3 once per SVG, insert after slide 2, set rich text on each
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath     .\branded.pptx `
    -SvgPath          .\generated\ `
    -ContentSlide     3 `
    -InsertAfterSlide 2 `
    -ClearContent `
    -SlideTexts @{
        1 = @{ Title = '第一章：市场概述' }
        2 = @{ Title = '第二章：竞争格局' }
        3 = @{ Title = '第三章：战略规划' }
    }
```

### MANIFEST mode — whole-deck workflow from a JSON file

The recommended path for "edit cover + edit TOC + add N content pages" in a
single invocation. Triggered by `-Manifest`. SVG paths inside the manifest are
resolved relative to the manifest file's directory.

```powershell
pwsh -File scripts\Edit-ExistingPptx.ps1 `
    -TemplatePath .\branded.pptx `
    -Manifest     .\deck.json `
    -OutputPath   .\out.pptx
```

**Manifest schema** (see `examples/deck.json` for a full example):

```json
{
  "edits": [
    {
      "type":     "text",
      "slide":    1,
      "title":    "2026 年度战略报告",
      "subtitle": "董事会汇报",
      "date":     "2026-05"
    },
    {
      "type":  "text",
      "slide": 2,
      "title": "目录",
      "body":  ["市场概述", "竞争格局", "战略规划", "执行路径"]
    },
    {
      "type":          "expand",
      "templateSlide": 3,
      "insertAfter":   2,
      "clearContent":  true,
      "items": [
        { "svg": "slides/s1.svg", "title": "第一章：市场概述" },
        { "svg": "slides/s2.svg", "title": "第二章：竞争格局" },
        { "svg": "slides/s3.svg", "title": "第三章：战略规划" }
      ]
    },
    {
      "type":         "insert",
      "slide":        9,
      "svg":          "slides/closing-chart.svg",
      "title":        "结语",
      "clearContent": true
    }
  ]
}
```

Edit types and their fields:

| `type`     | Required fields              | Optional fields                                                   |
|------------|------------------------------|-------------------------------------------------------------------|
| `"text"`   | `slide`                      | `title`, `subtitle`, `body[]`, `date`, `footer`                   |
| `"insert"` | `slide`, `svg`               | `title`, `subtitle`, `body[]`, `date`, `footer`, `clearContent`, `contentZone` |
| `"expand"` | `items[]` (each needs `svg`) | `templateSlide`, `insertAfter`, `clearContent`, `contentZone`; per-item: `title`, `subtitle`, `body[]`, `date`, `footer` |

### Content-zone auto-detection

When `-ContentZone` (or manifest `contentZone`) is not provided:

1. The script searches the reference slide for a content placeholder (body,
   object, chart, picture, media). If found, uses its bounds.
2. Otherwise falls back to a default below the title bar (~36 pt side margin,
   ~108 pt top, ~18 pt bottom).

To find points manually: in PowerPoint, right-click a shape → Size and
Position; multiply inches × 72.

### Backup behaviour

A `.bak.pptx` of the original template is created automatically on every run
(unless `-NoBackup`). It lives next to the original.

---

## How Claude should invoke these scripts

1. **Pick the right mode using the decision tree above.** Cover/TOC = TEXT;
   one slide replacement = INSERT; many duplicated content pages = EXPAND;
   whole deck = MANIFEST.
2. Resolve all file paths to **absolute paths** before passing in. COM is
   unreliable with relative paths.
3. **Prefer `pwsh` over `powershell.exe`** when the user has PowerShell 7.
   It defaults to UTF-8, which avoids mangling Chinese strings on the command
   line.
4. For complex jobs (Chinese text + many slides), write a manifest JSON file
   to disk and call `-Manifest`. This sidesteps shell-quoting issues entirely.
5. Run via Bash:
   ```powershell
   pwsh -File "C:\path\scripts\Edit-ExistingPptx.ps1" -TemplatePath "C:\..." ...
   ```
6. Surface stderr verbatim if the script fails. Errors distinguish "PowerPoint
   not installed", "file not found", "slide out of range", "ExecuteMso failed".

---

## Known limitations

- **PowerPoint window must be visible.** `ExecuteMso` is a Ribbon command;
  unreliable when hidden. Both scripts set `Visible = True`.
- **Close other presentations first.** COM attaches to the running instance;
  open presentations can interfere with selection state. Scripts warn but
  don't abort.
- **No headless / CI support.** For headless conversion use `svg2pptx` or a
  pure-DrawingML library instead.
- **SVG fidelity matches PowerPoint's built-in converter.** See "SVG
  generation guidelines" above.
- **ExecuteMso timing.** A 300 ms pause is inserted before `ExecuteMso`. On
  slow machines, increase to 500 ms in the script.
- **One SVG per slide.** Multi-SVG layouts on a single slide are out of
  scope.
- **Replacing placeholder text** uses `TextRange.Text = …`, which preserves
  layout-inherited formatting but may flatten run-level overrides (e.g. a
  word that was specifically coloured red). For pristine fidelity, edit such
  slides manually.
- **Windows only.**

## Verifying the output

Open the `.pptx` and right-click any inserted shape. If the context menu
shows **Edit Points** or the shape can be ungrouped into individual
sub-shapes, the conversion succeeded. If it still shows **Convert to Shape**,
that shape did not convert (usually a malformed or out-of-subset SVG); the
script's stderr will have noted it.
