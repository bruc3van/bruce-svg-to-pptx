#Requires -Version 5.1
<#
.SYNOPSIS
    Insert SVG-converted shapes into an existing PowerPoint presentation, with optional
    slide duplication to expand the deck while preserving its branded template.

.DESCRIPTION
    Opens an existing .pptx (e.g., a corporate template with cover / TOC / content slides)
    and operates in one of two modes:

    INSERT mode  : Inserts each SVG into one or more already-existing slides.
                   Triggered by -TargetSlide. Multiple SVGs map to consecutive slides
                   starting at TargetSlide.

    EXPAND mode  : Duplicates a "content template" slide once per SVG, appends the
                   copies after a specified position, and inserts one SVG into each copy.
                   Triggered automatically when -TargetSlide is omitted.
                   The template slide keeps all branded furniture (title bar, footer,
                   background, logo, etc.) intact.

    In both modes the SVG is converted to native editable shapes via PowerPoint's
    "Convert to Shape" command (CommandBars.ExecuteMso("SVGEdit")), producing the
    same result as right-clicking an SVG and choosing Convert to Shape.

.PARAMETER TemplatePath
    Path to the existing .pptx to edit. Must already exist.

.PARAMETER SvgPath
    One or more .svg file paths, or a directory of .svg files (non-recursive).

.PARAMETER OutputPath
    Destination .pptx. Defaults to TemplatePath (in-place edit after backup).
    If different from TemplatePath, the original is not modified.

.PARAMETER TargetSlide
    INSERT mode. Slide index (1-based) for the first SVG. Additional SVGs go into
    TargetSlide+1, TargetSlide+2, etc. All target slides must already exist.
    Omit to enter EXPAND mode.

.PARAMETER ContentSlide
    EXPAND mode. Slide index (1-based) to duplicate as the content page template.
    Defaults to the last slide in the presentation.

.PARAMETER InsertAfterSlide
    EXPAND mode. New slides are inserted after this index.
    Defaults to the last slide (append to end).

.PARAMETER ContentZone
    Override the area used for SVG placement: "Left,Top,Width,Height" in PowerPoint
    points (1 pt = 1/72 inch).
    Example: "36,108,888,396" for a standard widescreen content area.
    If omitted, the script auto-detects a content placeholder on the template/target
    slide, falling back to a sensible default below the title.

.PARAMETER SlideTitles
    Optional array of title strings, one per SVG (in the same order).
    When provided, each duplicated or target slide's title placeholder text is
    set to the corresponding string. Titles that are empty or $null are skipped
    (the template title is preserved).
    Example: -SlideTitles "市场分析","竞争格局","战略规划"

.PARAMETER ClearContent
    Before inserting the SVG, delete all shapes from the slide that are NOT title,
    subtitle, footer, slide-number, or date placeholders. Use this when duplicating
    a template slide that contains sample content you want replaced.

.PARAMETER NoBackup
    Skip creating a .bak.pptx backup copy of the original TemplatePath.

.PARAMETER Force
    Skip the overwrite confirmation when OutputPath already exists and differs from
    TemplatePath.

.EXAMPLE
    # EXPAND mode: duplicate the last slide for each SVG, append after the last slide
    .\Edit-ExistingPptx.ps1 `
        -TemplatePath .\branded-template.pptx `
        -SvgPath .\content-slides\ `
        -ClearContent

.EXAMPLE
    # EXPAND mode: use slide 3 as the content template, insert after slide 4
    .\Edit-ExistingPptx.ps1 `
        -TemplatePath .\deck.pptx `
        -SvgPath .\slides\ `
        -ContentSlide 3 `
        -InsertAfterSlide 4 `
        -ClearContent `
        -OutputPath .\deck-filled.pptx

.EXAMPLE
    # INSERT mode: put icon.svg into slide 5
    .\Edit-ExistingPptx.ps1 `
        -TemplatePath .\deck.pptx `
        -SvgPath .\icon.svg `
        -TargetSlide 5

.EXAMPLE
    # INSERT mode: replace content in slides 3, 4, 5 with three SVGs
    .\Edit-ExistingPptx.ps1 `
        -TemplatePath .\deck.pptx `
        -SvgPath .\slide1.svg, .\slide2.svg, .\slide3.svg `
        -TargetSlide 3 `
        -ClearContent

.EXAMPLE
    # EXPAND mode with custom titles for each duplicated slide
    .\Edit-ExistingPptx.ps1 `
        -TemplatePath .\branded.pptx `
        -SvgPath      .\slides\ `
        -ContentSlide    3 `
        -InsertAfterSlide 2 `
        -ClearContent `
        -SlideTitles "市场概述", "竞争格局", "战略规划", "执行路径"

.NOTES
    Requires Windows + Microsoft PowerPoint 2016 build 1712 or later.
    PowerPoint window is shown during conversion (required by ExecuteMso) and
    closed automatically on completion or error.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TemplatePath,

    [Parameter(Mandatory)]
    [string[]]$SvgPath,

    [string]$OutputPath,

    # INSERT mode
    [int]$TargetSlide      = 0,

    # EXPAND mode
    [int]$ContentSlide     = 0,
    [int]$InsertAfterSlide = 0,

    [string]$ContentZone,

    [string[]]$SlideTitles,

    [switch]$ClearContent,
    [switch]$NoBackup,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ──────────────────────────────────────────────────────────────
# PowerPoint / Office constants
# ──────────────────────────────────────────────────────────────
$msoFalse                    = 0
$msoTrue                     = -1
$msoGraphic                  = 28
$ppSaveAsOpenXMLPresentation = 24

# Placeholder type IDs
$PP_TITLE          = 1
$PP_BODY           = 2
$PP_CENTER_TITLE   = 3
$PP_SUBTITLE       = 4
$PP_VERT_BODY      = 5
$PP_VERT_TITLE     = 6
$PP_DATE           = 10
$PP_FOOTER         = 11
$PP_SLIDE_NUMBER   = 12
$PP_OBJECT         = 14

# Types that stay when -ClearContent is used (structural chrome)
$KEEPER_TYPES = @($PP_TITLE, $PP_CENTER_TITLE, $PP_SUBTITLE,
                  $PP_VERT_BODY, $PP_VERT_TITLE,
                  $PP_DATE, $PP_FOOTER, $PP_SLIDE_NUMBER)

# Types considered "content zone" for auto-detection
$CONTENT_PH_TYPES = @($PP_BODY, $PP_OBJECT, 15, 16, 17, 18, 19)

# ──────────────────────────────────────────────────────────────
# Helper: enumerate .svg files from path list
# ──────────────────────────────────────────────────────────────
function Resolve-SvgFiles {
    param([string[]]$Inputs)
    $files = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $Inputs) {
        if (-not (Test-Path -LiteralPath $p)) { throw "Path not found: $p" }
        $item = Get-Item -LiteralPath $p
        if ($item.PSIsContainer) {
            $found = Get-ChildItem -LiteralPath $item.FullName -Filter *.svg -File |
                     Sort-Object Name | Select-Object -ExpandProperty FullName
            if ($found.Count -eq 0) { Write-Warning "No .svg files in: $($item.FullName)" }
            $found | ForEach-Object { $files.Add($_) }
        } else {
            if ($item.Extension -ine '.svg') { throw "Not an .svg file: $($item.FullName)" }
            $files.Add($item.FullName)
        }
    }
    if ($files.Count -eq 0) { throw "No .svg files to process." }
    return $files.ToArray()
}

# ──────────────────────────────────────────────────────────────
# Helper: detect content zone on a slide (placeholder or default)
# ──────────────────────────────────────────────────────────────
function Get-ContentZone {
    param(
        $Slide,
        [string]$Override,
        [double]$SlideWidth,
        [double]$SlideHeight
    )

    # 1. Explicit override
    if (-not [string]::IsNullOrWhiteSpace($Override)) {
        $parts = $Override -split ',' | ForEach-Object { $_.Trim() }
        if ($parts.Count -ne 4) { throw "-ContentZone must be 'Left,Top,Width,Height' in points." }
        return @{
            Left   = [double]$parts[0]
            Top    = [double]$parts[1]
            Width  = [double]$parts[2]
            Height = [double]$parts[3]
        }
    }

    # 2. Auto-detect from content placeholder (body, object, chart, picture, etc.)
    foreach ($shape in $Slide.Shapes) {
        try {
            $phType = $shape.PlaceholderFormat.Type
            if ($CONTENT_PH_TYPES -contains $phType) {
                return @{
                    Left   = [double]$shape.Left
                    Top    = [double]$shape.Top
                    Width  = [double]$shape.Width
                    Height = [double]$shape.Height
                }
            }
        } catch { }
    }

    # 3. Default: full-width content area below a typical title bar
    #    0.5" (36pt) side margins, 1.2" (86.4pt) title height, 0.3" (21.6pt) gap
    $hMargin = 36.0
    $titleH  = 86.4
    $gap     = 21.6
    $vMargin = 18.0
    return @{
        Left   = $hMargin
        Top    = $titleH + $gap
        Width  = $SlideWidth  - 2 * $hMargin
        Height = $SlideHeight - $titleH - $gap - $vMargin
    }
}

# ──────────────────────────────────────────────────────────────
# Helper: scale SVG to fit inside a zone (letterbox)
# ──────────────────────────────────────────────────────────────
function Get-SvgFit {
    param(
        [string]$Path,
        [double]$ZoneLeft,
        [double]$ZoneTop,
        [double]$ZoneWidth,
        [double]$ZoneHeight
    )

    $fallback = @{
        Left   = $ZoneLeft
        Top    = $ZoneTop
        Width  = $ZoneWidth
        Height = $ZoneHeight
    }

    try {
        [xml]$xml = Get-Content -LiteralPath $Path -Raw
        $svg = $xml.svg
        if (-not $svg) { return $fallback }

        $w = $null; $h = $null

        if ($svg.viewBox) {
            $parts = $svg.viewBox.Trim() -split '[\s,]+' | Where-Object { $_ -ne '' }
            if ($parts.Count -eq 4) { $w = [double]$parts[2]; $h = [double]$parts[3] }
        }
        if ((-not $w) -or (-not $h)) {
            if ($svg.width -and $svg.height) {
                $w = [double]([regex]::Match($svg.width,  '[\d.]+').Value)
                $h = [double]([regex]::Match($svg.height, '[\d.]+').Value)
            }
        }
        if ((-not $w) -or (-not $h) -or $w -le 0 -or $h -le 0) { return $fallback }

        $scale = [Math]::Min($ZoneWidth / $w, $ZoneHeight / $h)
        $fitW  = $w * $scale
        $fitH  = $h * $scale
        return @{
            Left   = $ZoneLeft + ($ZoneWidth  - $fitW) / 2.0
            Top    = $ZoneTop  + ($ZoneHeight - $fitH) / 2.0
            Width  = $fitW
            Height = $fitH
        }
    } catch {
        return $fallback
    }
}

# ──────────────────────────────────────────────────────────────
# Helper: remove non-structural shapes from a slide
# ──────────────────────────────────────────────────────────────
function Clear-SlideContent {
    param($Slide)
    # Iterate backwards so deletion doesn't shift indices
    for ($i = $Slide.Shapes.Count; $i -ge 1; $i--) {
        $shape = $Slide.Shapes.Item($i)
        $keep  = $false
        try {
            $phType = $shape.PlaceholderFormat.Type
            $keep   = ($KEEPER_TYPES -contains $phType)
        } catch { }
        if (-not $keep) {
            try { $shape.Delete() } catch { }
        }
    }
}

# ──────────────────────────────────────────────────────────────
# Helper: set the title placeholder text on a slide
# ──────────────────────────────────────────────────────────────
function Set-SlideTitle {
    param($Slide, [string]$Title)
    $titleTypes = @($PP_TITLE, $PP_CENTER_TITLE)
    foreach ($shape in $Slide.Shapes) {
        try {
            $phType = $shape.PlaceholderFormat.Type
            if ($titleTypes -contains $phType) {
                $shape.TextFrame.TextRange.Text = $Title
                return
            }
        } catch { }
    }
    Write-Warning "    No title placeholder found on this slide; title '$Title' was not applied."
}

# ──────────────────────────────────────────────────────────────
# Helper: release COM object
# ──────────────────────────────────────────────────────────────
function Release-Com {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch { }
    }
}

# ══════════════════════════════════════════════════════════════
# Pre-flight
# ══════════════════════════════════════════════════════════════

$svgFiles = Resolve-SvgFiles -Inputs $SvgPath

# Resolve TemplatePath to absolute
if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "TemplatePath not found: $TemplatePath"
}
$templateAbs = (Resolve-Path -LiteralPath $TemplatePath).Path
if ([IO.Path]::GetExtension($templateAbs) -ine '.pptx') {
    throw "TemplatePath must be a .pptx file (got: $templateAbs)"
}

# Resolve OutputPath to absolute (default = TemplatePath = in-place edit)
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $templateAbs
}
$outDir = Split-Path -Path $OutputPath -Parent
if ([string]::IsNullOrEmpty($outDir)) { $outDir = (Get-Location).Path }
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
$OutputPath = Join-Path (Resolve-Path -LiteralPath $outDir).Path (Split-Path -Path $OutputPath -Leaf)
if ([IO.Path]::GetExtension($OutputPath) -ine '.pptx') {
    throw "OutputPath must end in .pptx (got: $OutputPath)"
}

# Overwrite confirmation (only when output differs from template and already exists)
$outputIsSameAsTemplate = ($OutputPath -ieq $templateAbs)
if (-not $outputIsSameAsTemplate -and (Test-Path -LiteralPath $OutputPath) -and -not $Force) {
    $resp = Read-Host "Overwrite existing $OutputPath? [y/N]"
    if ($resp -notmatch '^[Yy]') { Write-Host "Aborted."; return }
}

# Backup the original template (always, unless -NoBackup)
if (-not $NoBackup) {
    $bakPath = [IO.Path]::ChangeExtension($templateAbs, '.bak.pptx')
    Copy-Item -LiteralPath $templateAbs -Destination $bakPath -Force
    Write-Host "Backup: $bakPath"
}

# ══════════════════════════════════════════════════════════════
# Launch PowerPoint
# ══════════════════════════════════════════════════════════════
Write-Host "Starting PowerPoint..."
$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
} catch {
    throw "Cannot start PowerPoint via COM. Is Microsoft PowerPoint installed? Error: $($_.Exception.Message)"
}
$ppt.Visible = $msoTrue

if ($ppt.Presentations.Count -gt 0) {
    Write-Warning "PowerPoint has $($ppt.Presentations.Count) open presentation(s). Close them for reliable results."
}

$pres           = $null
$convertedCount = 0
$failedItems    = @()

try {
    # Open template (ReadOnly=$false so we can save back to it; AddToMRU=$false)
    $pres = $ppt.Presentations.Open($templateAbs, $msoFalse, $msoFalse, $msoTrue)

    $slideWidth  = $pres.PageSetup.SlideWidth
    $slideHeight = $pres.PageSetup.SlideHeight
    $totalSlides = $pres.Slides.Count

    Write-Host "Opened: $templateAbs  ($totalSlides slides, ${slideWidth}x${slideHeight} pts)"

    # ──────────────────────────────────────────────────────────
    # Determine mode and validate parameters
    # ──────────────────────────────────────────────────────────
    $expandMode = ($TargetSlide -eq 0)

    if ($expandMode) {
        # EXPAND mode: resolve content template index
        $templateIdx = if ($ContentSlide -gt 0) { $ContentSlide } else { $totalSlides }
        if ($templateIdx -lt 1 -or $templateIdx -gt $totalSlides) {
            throw "ContentSlide $templateIdx is out of range (presentation has $totalSlides slides)."
        }

        # Resolve base insert-after index
        $baseInsertAfter = if ($InsertAfterSlide -gt 0) { $InsertAfterSlide } else { $totalSlides }
        if ($baseInsertAfter -lt 0 -or $baseInsertAfter -gt $totalSlides) {
            throw "InsertAfterSlide $baseInsertAfter is out of range."
        }

        Write-Host "Mode: EXPAND  ContentSlide=$templateIdx  InsertAfter=$baseInsertAfter"
        Write-Host "Note: title placeholders on duplicated slides keep the template's title text." `
                   "Update slide titles manually in PowerPoint after generation."
    } else {
        # INSERT mode: check that all target slides exist
        $lastTargetSlide = $TargetSlide + $svgFiles.Count - 1
        if ($TargetSlide -lt 1 -or $lastTargetSlide -gt $totalSlides) {
            throw "INSERT mode requires slides $TargetSlide–$lastTargetSlide to exist (presentation has $totalSlides slides)."
        }
        Write-Host "Mode: INSERT  Slides $TargetSlide–$lastTargetSlide"
    }

    # ──────────────────────────────────────────────────────────
    # Auto-detect (or parse) the content zone from the reference slide
    # ──────────────────────────────────────────────────────────
    $refSlideIdx = if ($expandMode) { $templateIdx } else { $TargetSlide }
    $zone = Get-ContentZone `
        -Slide    $pres.Slides.Item($refSlideIdx) `
        -Override $ContentZone `
        -SlideWidth  $slideWidth `
        -SlideHeight $slideHeight

    Write-Host ("Content zone: L={0:F1}  T={1:F1}  W={2:F1}  H={3:F1} pts" -f `
        $zone.Left, $zone.Top, $zone.Width, $zone.Height)

    # ──────────────────────────────────────────────────────────
    # Process each SVG
    # ──────────────────────────────────────────────────────────
    for ($i = 0; $i -lt $svgFiles.Count; $i++) {
        $svg = $svgFiles[$i]
        Write-Host ("  [{0}/{1}] {2}" -f ($i + 1), $svgFiles.Count, $svg)

        $workSlide = $null

        if ($expandMode) {
            # ── Duplicate the content template slide ──────────
            # Duplicate() officially returns SlideRange; .Item(1) retrieves the Slide.
            # A small number of builds return the Slide object directly — handle both.
            $duped    = $pres.Slides.Item($templateIdx).Duplicate()
            $newSlide = $null
            try   { $newSlide = $duped.Item(1) }
            catch { $newSlide = $duped }

            # Target position in the final deck (0-based offset from baseInsertAfter)
            $targetPos = $baseInsertAfter + $i + 1

            # Move the duplicate from its temporary position to targetPos.
            # If targetPos is before or at templateIdx the template will shift right.
            $newSlide.MoveTo($targetPos)
            if ($targetPos -le $templateIdx) { $templateIdx++ }

            # Re-fetch by index (safer than holding stale reference after MoveTo)
            $workSlide = $pres.Slides.Item($targetPos)

        } else {
            # ── Use the existing slide ────────────────────────
            $workSlide = $pres.Slides.Item($TargetSlide + $i)
        }

        # ── Optionally strip content before inserting ─────────
        if ($ClearContent) {
            Clear-SlideContent -Slide $workSlide
        }

        # ── Optionally set slide title ─────────────────────────
        if ($SlideTitles -and $i -lt $SlideTitles.Count) {
            $titleText = $SlideTitles[$i]
            if (-not [string]::IsNullOrWhiteSpace($titleText)) {
                Set-SlideTitle -Slide $workSlide -Title $titleText
            }
        }

        # ── Scale SVG to fit inside the content zone ─────────
        $fit = Get-SvgFit `
            -Path      $svg `
            -ZoneLeft  $zone.Left `
            -ZoneTop   $zone.Top `
            -ZoneWidth $zone.Width `
            -ZoneHeight $zone.Height

        # ── Insert SVG as an embedded picture ─────────────────
        $shape = $null
        try {
            $shape = $workSlide.Shapes.AddPicture(
                $svg,
                $msoFalse,   # LinkToFile  = false
                $msoTrue,    # SaveWithDocument = true
                $fit.Left,
                $fit.Top,
                $fit.Width,
                $fit.Height
            )
        } catch {
            Write-Warning "    AddPicture failed: $($_.Exception.Message)"
            $failedItems += $svg
            continue
        }

        # Verify PowerPoint recognised it as SVG (msoGraphic)
        if ($shape.Type -ne $msoGraphic) {
            Write-Warning ("    Inserted as Type={0}, expected msoGraphic ({1}). " +
                           "Conversion skipped (possibly not a valid SVG)." -f $shape.Type, $msoGraphic)
            $failedItems += $svg
            continue
        }

        # ── Execute "Convert to Shape" ─────────────────────────
        try {
            $workSlide.Select()
            $shape.Select()
            # Brief pause for PowerPoint's selection state to settle.
            # On slow machines increase this to 500 ms if conversions silently fail.
            Start-Sleep -Milliseconds 300
            $ppt.CommandBars.ExecuteMso('SVGEdit')
            $convertedCount++
        } catch {
            Write-Warning "    ExecuteMso('SVGEdit') failed: $($_.Exception.Message)"
            $failedItems += $svg
        }
    }

    # ──────────────────────────────────────────────────────────
    # Save
    # ──────────────────────────────────────────────────────────
    Write-Host "Saving to $OutputPath..."
    if ($outputIsSameAsTemplate) {
        # In-place edit: use Save() to overwrite the already-open file path.
        $pres.Save()
    } else {
        $pres.SaveAs($OutputPath, $ppSaveAsOpenXMLPresentation)
    }
    $pres.Close()

} finally {
    if ($null -ne $pres) { Release-Com $pres }
    try { $ppt.Quit() } catch { }
    Release-Com $ppt
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# ══════════════════════════════════════════════════════════════
# Summary
# ══════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "Done. Converted $convertedCount of $($svgFiles.Count) SVG(s) → $OutputPath"

if ($failedItems.Count -gt 0) {
    Write-Host "Failed ($($failedItems.Count)):"
    $failedItems | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
