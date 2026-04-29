---
name: bruce-svg-to-pptx
description: Convert SVG files into native, editable PowerPoint shapes (DrawingML) on Windows by automating PowerPoint via COM. Use this skill whenever the user wants to turn one or more .svg files into a .pptx where each SVG is broken into editable shapes (the same result as right-clicking an inserted SVG and choosing "Convert to Shape"), or asks to batch-convert SVGs into editable PowerPoint graphics, or mentions "SVG → 形状", "SVG 转 PPT 可编辑形状", or similar. Trigger this even when the user only describes the goal ("I want my SVG icons to be editable inside the deck") without naming the conversion explicitly. Requires Windows with Microsoft PowerPoint 2016 build 1712 or later.
---

# SVG to editable PowerPoint shapes (Windows)

## What this skill does

Drives Microsoft PowerPoint on Windows via COM to perform the same operation a user would do manually:

1. Insert each SVG onto a blank slide as a `msoGraphic` shape.
2. Select that graphic and invoke `CommandBars.ExecuteMso("SVGEdit")` — PowerPoint's internal "Convert to Shape" command. There is no first-class VBA method for this conversion; the Mso command is the only programmatic hook Microsoft exposes.
3. Save the deck. The result is identical to the manual right-click flow: each SVG becomes a group of native DrawingML shapes that can be ungrouped, recolored, animated, and edited individually.

The conversion runs inside PowerPoint's own engine, so fidelity matches what the user would see manually — including PowerPoint's known limitations with complex filters, masks, and certain gradient types.

## When to use this skill

Use this skill when the user asks to convert SVG(s) into editable PowerPoint shapes, batch-convert a folder of icons, or build a deck where SVG content needs to be edited inside PowerPoint after the fact. **Do not** use this skill for:

- Embedding SVG as a static picture (just use `python-pptx` or insert manually).
- Cross-platform (macOS/Linux) workflows — this skill requires Windows COM. Suggest the pure-DrawingML route (`svg2pptx` library or similar) instead.

## Prerequisites

- Windows 10/11.
- Microsoft PowerPoint, version 2016 build 1712 or later (any current Microsoft 365 / Office 2019 / Office 2021 / Office 2024 install qualifies).
- PowerShell 5.1+ (ships with Windows) or PowerShell 7.

If PowerPoint is not installed, the script will exit with a clear error. Tell the user the skill cannot run in that environment.

## Usage

The skill ships with one script: `scripts/Convert-SvgToShapes.ps1`.

Invoke it from PowerShell:

```powershell
# Single SVG into a new deck
pwsh -File scripts\Convert-SvgToShapes.ps1 `
    -SvgPath .\icon.svg `
    -OutputPath .\out.pptx

# Batch-convert a folder of SVGs (one per slide)
pwsh -File scripts\Convert-SvgToShapes.ps1 `
    -SvgPath .\icons\ `
    -OutputPath .\icons.pptx

# Append to an existing deck instead of creating a new one
pwsh -File scripts\Convert-SvgToShapes.ps1 `
    -SvgPath .\new-icons\ `
    -OutputPath .\existing-deck.pptx `
    -Append
```

Use `powershell.exe` instead of `pwsh` if PowerShell 7 is not installed.

### Parameters

- `-SvgPath` (required): one or more `.svg` file paths, or a directory. Accepts an array. If a directory is given, all `*.svg` files inside (non-recursive) are processed.
- `-OutputPath` (required): destination `.pptx`. If it does not exist, a new deck is created. If it exists and `-Append` is set, slides are appended; otherwise the existing file is overwritten after a confirmation prompt (skipped with `-Force`).
- `-Append` (switch): append to the existing `OutputPath` deck rather than overwriting.
- `-Force` (switch): skip the overwrite confirmation.

## How to invoke from Claude Code

When the user is on Windows and asks to convert SVGs, run the script with bash/PowerShell tool. The script resolves relative paths to absolute internally, but passing absolute paths is safer and avoids ambiguity. Surface the script's stderr to the user verbatim if it fails; the error messages distinguish between "PowerPoint not installed", "file not found", and "ExecuteMso failed for this SVG".

Even if the script fails mid-run, the `finally` block guarantees PowerPoint is always closed and the COM object released — the user does not need to manually kill any PowerPoint process.

After a successful run, confirm the output path and mention the fidelity caveat: PowerPoint's built-in converter handles paths, basic shapes, and solid/simple-gradient fills well, but may flatten or drop complex SVG features (filters, clipPath chains, embedded raster images, certain advanced gradient stops). If the user reports a bad conversion on a specific SVG, that's a PowerPoint limitation, not a script bug — suggest simplifying the SVG (e.g., re-exporting from Figma/Illustrator without effects) or falling back to a pure-DrawingML library.

## Known limitations

- **PowerPoint window must be visible during conversion.** `CommandBars.ExecuteMso` is a Ribbon command and unreliable when the PowerPoint window is hidden. The script sets `Visible = True` accordingly. The window is closed cleanly at the end.
- **Cannot run alongside another PowerPoint session reliably.** If the user has PowerPoint open, COM may attach to that instance and produce surprising results. The script detects this and warns; recommend closing other PowerPoint windows first.
- **One SVG per slide.** v1 places each SVG centered on its own blank slide. Multi-SVG layouts are out of scope.
- **No headless mode.** Server/CI use is not supported. For headless, recommend the pure-Python `svg2pptx` library route.
- **ExecuteMso timing sensitivity.** The script inserts a 250 ms pause before calling `ExecuteMso("SVGEdit")` to let PowerPoint's selection state settle. On slower machines a shape may silently remain unconverted (the script output will still show it as "failed"). If this happens, increase the `Start-Sleep -Milliseconds 250` line in the script to 500 ms or more.
- **Slide dimensions follow the presentation's page setup.** When creating a new deck, PowerPoint uses its default slide size (typically 10"×7.5" for standard or 13.33"×7.5" for widescreen, depending on the Office version). The SVG is scaled to fill that canvas while preserving its aspect ratio (letterboxed if needed). There is currently no parameter to specify a custom slide size.

## Verifying the output

After running, the user can confirm the conversion worked by opening the `.pptx` and right-clicking any inserted graphic — if the context menu shows "Edit Points" or the shape can be ungrouped into individual sub-shapes, the conversion succeeded. If the menu still shows "Convert to Shape" for a graphic, that one did not convert (rare — usually a malformed SVG); the script's stderr will have noted it.
