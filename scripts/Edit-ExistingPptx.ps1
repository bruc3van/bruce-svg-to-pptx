#Requires -Version 5.1
<#
.SYNOPSIS
    Edit an existing PPTX template: replace text on cover/TOC pages, and/or
    insert SVG-derived editable shapes into content pages.

.DESCRIPTION
    Four modes of operation, picked automatically from the parameters supplied:

    TEXT mode     Edit text on existing slides only — no SVG. Use this for
                  cover, title, and TOC slides where you only need to change
                  wording. Triggered when -SlideTexts is provided without
                  -SvgPath.

    INSERT mode   Insert each SVG into one or more already-existing slides
                  (one SVG per slide). Triggered by -TargetSlide.

    EXPAND mode   Duplicate a "content template" slide once per SVG, append the
                  copies after a chosen position, and insert one SVG into each.
                  The duplicates keep the branded title bar / background; the
                  SVG goes inside a safe content zone. Triggered when -SvgPath
                  is given without -TargetSlide.

    MANIFEST mode Drive a whole-deck workflow from a JSON file. Each entry in
                  the manifest is a TEXT, INSERT, or EXPAND edit. Triggered by
                  -Manifest. This is the recommended path for the full
                  "edit cover + edit TOC + add N content pages" flow in a
                  single invocation.

    SVG-to-shape conversion uses PowerPoint's CommandBars.ExecuteMso("SVGEdit")
    — the Ribbon "Convert to Shape" command — producing native editable shapes
    identical to right-clicking an SVG and choosing Convert to Shape.

.PARAMETER TemplatePath
    Path to the existing .pptx to edit. Must already exist.

.PARAMETER SvgPath
    Optional. One or more .svg file paths or a directory of .svg files
    (non-recursive). Omit to enter TEXT mode.

.PARAMETER OutputPath
    Destination .pptx. Defaults to TemplatePath (in-place edit after backup).

.PARAMETER Manifest
    Path to a JSON file describing per-slide edits. Mutually exclusive with
    INSERT/EXPAND/TEXT command-line modes. See "Manifest schema" below.

.PARAMETER TargetSlide
    INSERT mode. 1-based slide index for the first SVG.

.PARAMETER ContentSlide
    EXPAND mode. 1-based slide index to duplicate as the content template.
    Defaults to the last slide.

.PARAMETER InsertAfterSlide
    EXPAND mode. New slides are inserted after this index. Defaults to last.

.PARAMETER ContentZone
    Override the SVG placement area: "Left,Top,Width,Height" in PowerPoint
    points. If omitted, the script auto-detects a content placeholder, falling
    back to a sensible default below the title.

.PARAMETER SlideTitles
    Backwards-compatibility shortcut. Array of strings, one per SVG, that sets
    each new slide's title placeholder. For richer text edits (subtitle, body,
    date, footer) use -SlideTexts or -Manifest.

.PARAMETER SlideTexts
    Hashtable for richer per-slide text edits.

    Top-level keys are 1-based slide indices (or, in EXPAND/INSERT mode, 1-based
    SVG indices — see notes). Each value is a hashtable describing what text to
    replace, with the following recognised keys:

        Title     string         first PP_TITLE / PP_CENTER_TITLE placeholder
        Subtitle  string         first PP_SUBTITLE placeholder
        Body      string[]       Nth PP_BODY / PP_OBJECT placeholder (in order)
        Date      string         first PP_DATE placeholder
        Footer    string         first PP_FOOTER placeholder

    Example (TEXT mode — fixed slide indices):

        -SlideTexts @{
            1 = @{ Title = '2026 战略报告'; Subtitle = '董事会汇报' }
            2 = @{ Title = '目录'; Body = @('市场概述','竞争格局','战略规划') }
        }

    In EXPAND or INSERT mode the indices refer to SVG order (1-based), since
    the destination slide indices are computed by the script.

.PARAMETER ClearContent
    Before inserting SVG, delete non-structural shapes (anything that isn't a
    title / subtitle / footer / slide-number / date placeholder).

.PARAMETER NoBackup
    Skip creating .bak.pptx of the original.

.PARAMETER Force
    Skip overwrite confirmation when OutputPath differs from TemplatePath.

.NOTES
    Manifest schema (JSON):

        {
          "edits": [
            { "type": "text",   "slide": 1,
              "title": "...", "subtitle": "...", "body": ["..."],
              "date": "...", "footer": "..." },

            { "type": "expand", "templateSlide": 3, "insertAfter": 2,
              "clearContent": true,
              "items": [
                { "svg": "s1.svg",
                  "title": "...", "subtitle": "...", "body": ["..."] }
              ] },

            { "type": "insert", "slide": 5, "svg": "chart.svg",
              "clearContent": true,
              "title": "..." }
          ]
        }

    SVG paths inside the manifest are resolved relative to the manifest file's
    directory.

    Requires Windows + Microsoft PowerPoint 2016 build 1712 or later.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TemplatePath,

    [string[]]$SvgPath,

    [string]$OutputPath,

    [string]$Manifest,

    # INSERT mode
    [int]$TargetSlide      = 0,

    # EXPAND mode
    [int]$ContentSlide     = 0,
    [int]$InsertAfterSlide = 0,

    [string]$ContentZone,

    [string[]]$SlideTitles,

    [hashtable]$SlideTexts,

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

# Placeholder type IDs (PpPlaceholderType)
$PP_TITLE          = 1
$PP_BODY           = 2
$PP_CENTER_TITLE   = 3
$PP_SUBTITLE       = 4
$PP_VERT_TITLE     = 5
$PP_VERT_BODY      = 6
$PP_DATE           = 16
$PP_FOOTER         = 15
$PP_SLIDE_NUMBER   = 13
$PP_OBJECT         = 7

# Types that stay when -ClearContent is used (structural chrome)
$KEEPER_TYPES = @($PP_TITLE, $PP_CENTER_TITLE, $PP_SUBTITLE,
                  $PP_VERT_BODY, $PP_VERT_TITLE,
                  $PP_DATE, $PP_FOOTER, $PP_SLIDE_NUMBER)

# Types considered "content zone" for auto-detection
$CONTENT_PH_TYPES = @($PP_BODY, $PP_OBJECT, 8, 9, 10, 11, 12)

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

    # Default: full-width content area below a typical title bar
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
        $raw = Get-Content -LiteralPath $Path -Raw
        $w = $null; $h = $null

        # Prefer viewBox via regex (handles namespaced <svg:svg> and odd whitespace)
        $m = [regex]::Match($raw, 'viewBox\s*=\s*"([^"]+)"')
        if ($m.Success) {
            $parts = $m.Groups[1].Value.Trim() -split '[\s,]+' | Where-Object { $_ -ne '' }
            if ($parts.Count -eq 4) {
                $w = [double]$parts[2]
                $h = [double]$parts[3]
            }
        }
        if ((-not $w) -or (-not $h)) {
            $mw = [regex]::Match($raw, '<svg[^>]*\swidth\s*=\s*"([^"]+)"')
            $mh = [regex]::Match($raw, '<svg[^>]*\sheight\s*=\s*"([^"]+)"')
            if ($mw.Success -and $mh.Success) {
                $w = [double]([regex]::Match($mw.Groups[1].Value, '[\d.]+').Value)
                $h = [double]([regex]::Match($mh.Groups[1].Value, '[\d.]+').Value)
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
# Helper: write text into a placeholder, preserving format inheritance
# ──────────────────────────────────────────────────────────────
function Set-PhText {
    param($Shape, [string]$Text, [string]$Label)
    if ($null -eq $Shape) {
        Write-Warning "    No $Label placeholder; '$Text' skipped."
        return
    }
    try {
        if ($Shape.HasTextFrame -ne $msoTrue) {
            Write-Warning "    $Label placeholder has no text frame; skipped."
            return
        }
        $Shape.TextFrame.TextRange.Text = $Text
    } catch {
        Write-Warning "    Failed to set $Label text: $($_.Exception.Message)"
    }
}

# ──────────────────────────────────────────────────────────────
# Helper: apply a -SlideTexts hashtable entry to a slide
#   Entry shape: @{ Title=..; Subtitle=..; Body=@(..); Date=..; Footer=.. }
# ──────────────────────────────────────────────────────────────
function Set-SlideTexts {
    param($Slide, $Texts)

    if (-not $Texts) { return }

    $titleTypes = @($PP_TITLE, $PP_CENTER_TITLE, $PP_VERT_TITLE)
    $bodyTypes  = @($PP_BODY,  $PP_OBJECT,      $PP_VERT_BODY)

    $titlePh    = $null
    $subtitlePh = $null
    $bodyPhs    = New-Object System.Collections.Generic.List[object]
    $datePh     = $null
    $footerPh   = $null

    foreach ($shape in $Slide.Shapes) {
        try {
            $phType = $shape.PlaceholderFormat.Type
        } catch { continue }
        if ($titleTypes -contains $phType) {
            if (-not $titlePh) { $titlePh = $shape }
            else { $bodyPhs.Add($shape) }   # extra title-like shapes treated as body fallbacks
            continue
        }
        if ($phType -eq $PP_SUBTITLE -and -not $subtitlePh) { $subtitlePh = $shape; continue }
        if ($bodyTypes -contains $phType)                    { $bodyPhs.Add($shape); continue }
        if ($phType -eq $PP_DATE   -and -not $datePh)        { $datePh   = $shape; continue }
        if ($phType -eq $PP_FOOTER -and -not $footerPh)      { $footerPh = $shape; continue }
    }

    # Hashtable accessor that works for both PS hashtable and PSCustomObject (from JSON)
    $get = {
        param($obj, [string]$key)
        if ($null -eq $obj) { return $null }
        if ($obj -is [hashtable]) {
            foreach ($k in $obj.Keys) { if ([string]$k -ieq $key) { return $obj[$k] } }
            return $null
        }
        $prop = $obj.PSObject.Properties | Where-Object { $_.Name -ieq $key } | Select-Object -First 1
        if ($prop) { return $prop.Value }
        return $null
    }

    $title    = & $get $Texts 'Title'
    $subtitle = & $get $Texts 'Subtitle'
    $body     = & $get $Texts 'Body'
    $date     = & $get $Texts 'Date'
    $footer   = & $get $Texts 'Footer'

    if ($null -ne $title    -and "$title"    -ne '') { Set-PhText -Shape $titlePh    -Text "$title"    -Label 'Title' }
    if ($null -ne $subtitle -and "$subtitle" -ne '') { Set-PhText -Shape $subtitlePh -Text "$subtitle" -Label 'Subtitle' }
    if ($null -ne $date     -and "$date"     -ne '') { Set-PhText -Shape $datePh     -Text "$date"     -Label 'Date' }
    if ($null -ne $footer   -and "$footer"   -ne '') { Set-PhText -Shape $footerPh   -Text "$footer"   -Label 'Footer' }

    if ($null -ne $body) {
        $arr = @($body)
        for ($j = 0; $j -lt $arr.Count; $j++) {
            $val = "$($arr[$j])"
            if ($val -eq '') { continue }
            if ($j -lt $bodyPhs.Count) {
                Set-PhText -Shape $bodyPhs[$j] -Text $val -Label "Body[$j]"
            } else {
                Write-Warning "    No body placeholder #$j on this slide; '$val' skipped."
            }
        }
    }
}

# ──────────────────────────────────────────────────────────────
# Helper: insert one SVG into a slide and convert to shapes
# ──────────────────────────────────────────────────────────────
function Insert-SvgIntoSlide {
    param(
        $PptApp,
        $Slide,
        [string]$SvgFile,
        $Zone
    )

    $fit = Get-SvgFit `
        -Path       $SvgFile `
        -ZoneLeft   $Zone.Left `
        -ZoneTop    $Zone.Top `
        -ZoneWidth  $Zone.Width `
        -ZoneHeight $Zone.Height

    $shape = $null
    try {
        $shape = $Slide.Shapes.AddPicture(
            $SvgFile, $msoFalse, $msoTrue,
            $fit.Left, $fit.Top, $fit.Width, $fit.Height
        )
    } catch {
        Write-Warning "    AddPicture failed: $($_.Exception.Message)"
        return $false
    }

    if ($shape.Type -ne $msoGraphic) {
        Write-Warning ("    Inserted as Type={0}, expected msoGraphic ({1}). Conversion skipped." `
                       -f $shape.Type, $msoGraphic)
        return $false
    }

    try {
        $Slide.Select()
        $shape.Select()
        Start-Sleep -Milliseconds 300
        $PptApp.CommandBars.ExecuteMso('SVGEdit')
        return $true
    } catch {
        Write-Warning "    ExecuteMso('SVGEdit') failed: $($_.Exception.Message)"
        return $false
    }
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
# Pre-flight: validate paths, decide mode
# ══════════════════════════════════════════════════════════════

if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "TemplatePath not found: $TemplatePath"
}
$templateAbs = (Resolve-Path -LiteralPath $TemplatePath).Path
if ([IO.Path]::GetExtension($templateAbs) -ine '.pptx') {
    throw "TemplatePath must be a .pptx file (got: $templateAbs)"
}

# Decide top-level mode
$manifestData = $null
$manifestDir  = $null
if ($Manifest) {
    if (-not (Test-Path -LiteralPath $Manifest)) { throw "Manifest not found: $Manifest" }
    $manifestAbs  = (Resolve-Path -LiteralPath $Manifest).Path
    $manifestDir  = Split-Path -Parent $manifestAbs
    try {
        $manifestData = Get-Content -LiteralPath $manifestAbs -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        throw "Failed to parse manifest JSON: $($_.Exception.Message)"
    }
    if (-not $manifestData.edits) { throw "Manifest has no 'edits' array." }
}

$svgFiles = @()
if (-not $Manifest -and $SvgPath) {
    $svgFiles = Resolve-SvgFiles -Inputs $SvgPath
}

# Mode resolution (only one of these is true)
$mode = $null
if ($Manifest) {
    $mode = 'MANIFEST'
} elseif ($svgFiles.Count -gt 0) {
    $mode = if ($TargetSlide -gt 0) { 'INSERT' } else { 'EXPAND' }
} elseif ($SlideTexts) {
    $mode = 'TEXT'
} else {
    throw "Nothing to do. Provide -Manifest, -SvgPath, or -SlideTexts."
}

# Resolve OutputPath (default = TemplatePath = in-place edit)
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

$outputIsSameAsTemplate = ($OutputPath -ieq $templateAbs)
if (-not $outputIsSameAsTemplate -and (Test-Path -LiteralPath $OutputPath) -and -not $Force) {
    $resp = Read-Host "Overwrite existing $OutputPath? [y/N]"
    if ($resp -notmatch '^[Yy]') { Write-Host "Aborted."; return }
}

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

$pres            = $null
$convertedCount  = 0
$svgAttempted    = 0
$failedItems     = @()
$textEditCount   = 0

try {
    $pres = $ppt.Presentations.Open($templateAbs, $msoFalse, $msoFalse, $msoTrue)
    $slideWidth  = $pres.PageSetup.SlideWidth
    $slideHeight = $pres.PageSetup.SlideHeight
    Write-Host ("Opened: {0}  ({1} slides, {2}x{3} pts)" -f $templateAbs, $pres.Slides.Count, $slideWidth, $slideHeight)
    Write-Host "Mode: $mode"

    switch ($mode) {

        # ──────────────────────────────────────────────────────
        'TEXT' {
            foreach ($key in $SlideTexts.Keys) {
                $idx = [int]$key
                if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) {
                    Write-Warning "  Slide $idx out of range (1..$($pres.Slides.Count)); skipped."
                    continue
                }
                Write-Host "  Editing slide $idx text…"
                Set-SlideTexts -Slide $pres.Slides.Item($idx) -Texts $SlideTexts[$key]
                $textEditCount++
            }
        }

        # ──────────────────────────────────────────────────────
        'INSERT' {
            $svgAttempted = $svgFiles.Count
            $lastTargetSlide = $TargetSlide + $svgFiles.Count - 1
            if ($TargetSlide -lt 1 -or $lastTargetSlide -gt $pres.Slides.Count) {
                throw "INSERT mode requires slides $TargetSlide–$lastTargetSlide to exist."
            }

            $zone = Get-ContentZone -Slide $pres.Slides.Item($TargetSlide) -Override $ContentZone `
                                    -SlideWidth $slideWidth -SlideHeight $slideHeight
            Write-Host ("Content zone: L={0:F1} T={1:F1} W={2:F1} H={3:F1}" -f $zone.Left,$zone.Top,$zone.Width,$zone.Height)

            for ($i = 0; $i -lt $svgFiles.Count; $i++) {
                $svg = $svgFiles[$i]
                Write-Host ("  [{0}/{1}] {2}" -f ($i+1), $svgFiles.Count, $svg)
                $workSlide = $pres.Slides.Item($TargetSlide + $i)

                if ($ClearContent) { Clear-SlideContent -Slide $workSlide }

                # Apply per-SVG texts (1-based key by SVG order) or legacy -SlideTitles
                $perSlide = $null
                if ($SlideTexts -and $SlideTexts.ContainsKey($i + 1)) {
                    $perSlide = $SlideTexts[$i + 1]
                } elseif ($SlideTitles -and $i -lt $SlideTitles.Count -and $SlideTitles[$i]) {
                    $perSlide = @{ Title = $SlideTitles[$i] }
                }
                if ($perSlide) { Set-SlideTexts -Slide $workSlide -Texts $perSlide }

                if (Insert-SvgIntoSlide -PptApp $ppt -Slide $workSlide -SvgFile $svg -Zone $zone) {
                    $convertedCount++
                } else {
                    $failedItems += $svg
                }
            }
        }

        # ──────────────────────────────────────────────────────
        'EXPAND' {
            $svgAttempted = $svgFiles.Count
            $totalSlides  = $pres.Slides.Count
            $templateIdx  = if ($ContentSlide     -gt 0) { $ContentSlide     } else { $totalSlides }
            $baseInsAfter = if ($InsertAfterSlide -gt 0) { $InsertAfterSlide } else { $totalSlides }
            if ($templateIdx -lt 1 -or $templateIdx -gt $totalSlides) {
                throw "ContentSlide $templateIdx out of range (1..$totalSlides)."
            }
            if ($baseInsAfter -lt 0 -or $baseInsAfter -gt $totalSlides) {
                throw "InsertAfterSlide $baseInsAfter out of range."
            }

            $zone = Get-ContentZone -Slide $pres.Slides.Item($templateIdx) -Override $ContentZone `
                                    -SlideWidth $slideWidth -SlideHeight $slideHeight
            Write-Host ("Content zone: L={0:F1} T={1:F1} W={2:F1} H={3:F1}" -f $zone.Left,$zone.Top,$zone.Width,$zone.Height)
            Write-Host "  ContentSlide=$templateIdx  InsertAfter=$baseInsAfter"

            for ($i = 0; $i -lt $svgFiles.Count; $i++) {
                $svg = $svgFiles[$i]
                Write-Host ("  [{0}/{1}] {2}" -f ($i+1), $svgFiles.Count, $svg)

                $duped = $pres.Slides.Item($templateIdx).Duplicate()
                $newSlide = $null
                try   { $newSlide = $duped.Item(1) } catch { $newSlide = $duped }

                $targetPos = $baseInsAfter + $i + 1
                $newSlide.MoveTo($targetPos)
                if ($targetPos -le $templateIdx) { $templateIdx++ }

                $workSlide = $pres.Slides.Item($targetPos)
                if ($ClearContent) { Clear-SlideContent -Slide $workSlide }

                $perSlide = $null
                if ($SlideTexts -and $SlideTexts.ContainsKey($i + 1)) {
                    $perSlide = $SlideTexts[$i + 1]
                } elseif ($SlideTitles -and $i -lt $SlideTitles.Count -and $SlideTitles[$i]) {
                    $perSlide = @{ Title = $SlideTitles[$i] }
                }
                if ($perSlide) { Set-SlideTexts -Slide $workSlide -Texts $perSlide }

                if (Insert-SvgIntoSlide -PptApp $ppt -Slide $workSlide -SvgFile $svg -Zone $zone) {
                    $convertedCount++
                } else {
                    $failedItems += $svg
                }
            }
        }

        # ──────────────────────────────────────────────────────
        'MANIFEST' {
            foreach ($edit in $manifestData.edits) {
                $type = "$($edit.type)".ToLowerInvariant()
                switch ($type) {

                    'text' {
                        $idx = [int]$edit.slide
                        if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) {
                            Write-Warning "  [manifest:text] slide $idx out of range; skipped."
                            continue
                        }
                        Write-Host "  [text] slide $idx"
                        Set-SlideTexts -Slide $pres.Slides.Item($idx) -Texts $edit
                        $textEditCount++
                    }

                    'insert' {
                        $idx = [int]$edit.slide
                        if ($idx -lt 1 -or $idx -gt $pres.Slides.Count) {
                            Write-Warning "  [manifest:insert] slide $idx out of range; skipped."
                            continue
                        }
                        if (-not $edit.svg) {
                            Write-Warning "  [manifest:insert] missing 'svg'; skipped."
                            continue
                        }
                        $svgFile = $edit.svg
                        if (-not [IO.Path]::IsPathRooted($svgFile)) {
                            $svgFile = Join-Path $manifestDir $svgFile
                        }
                        if (-not (Test-Path -LiteralPath $svgFile)) {
                            Write-Warning "  [manifest:insert] svg not found: $svgFile; skipped."
                            $failedItems += $svgFile
                            continue
                        }
                        Write-Host "  [insert] slide $idx  $svgFile"
                        $svgAttempted++
                        $workSlide = $pres.Slides.Item($idx)

                        $zoneOverride = $null
                        if ($edit.contentZone) { $zoneOverride = "$($edit.contentZone)" }
                        $zone = Get-ContentZone -Slide $workSlide -Override $zoneOverride `
                                                -SlideWidth $slideWidth -SlideHeight $slideHeight

                        if ($edit.clearContent) { Clear-SlideContent -Slide $workSlide }
                        Set-SlideTexts -Slide $workSlide -Texts $edit

                        if (Insert-SvgIntoSlide -PptApp $ppt -Slide $workSlide -SvgFile $svgFile -Zone $zone) {
                            $convertedCount++
                        } else {
                            $failedItems += $svgFile
                        }
                    }

                    'expand' {
                        if (-not $edit.items -or $edit.items.Count -eq 0) {
                            Write-Warning "  [manifest:expand] empty 'items'; skipped."
                            continue
                        }
                        $totalSlides  = $pres.Slides.Count
                        $tplIdx       = if ($edit.templateSlide) { [int]$edit.templateSlide } else { $totalSlides }
                        $baseInsAfter = if ($edit.insertAfter)   { [int]$edit.insertAfter   } else { $totalSlides }
                        if ($tplIdx -lt 1 -or $tplIdx -gt $totalSlides) {
                            Write-Warning "  [manifest:expand] templateSlide $tplIdx out of range; skipped."
                            continue
                        }

                        $zoneOverride = $null
                        if ($edit.contentZone) { $zoneOverride = "$($edit.contentZone)" }
                        $zone = Get-ContentZone -Slide $pres.Slides.Item($tplIdx) -Override $zoneOverride `
                                                -SlideWidth $slideWidth -SlideHeight $slideHeight
                        Write-Host ("  [expand] templateSlide={0} insertAfter={1} count={2}" -f $tplIdx,$baseInsAfter,$edit.items.Count)

                        for ($i = 0; $i -lt $edit.items.Count; $i++) {
                            $item   = $edit.items[$i]
                            $svgRel = $item.svg
                            if (-not $svgRel) {
                                Write-Warning "    item $i missing 'svg'; skipped."
                                continue
                            }
                            $svgFile = $svgRel
                            if (-not [IO.Path]::IsPathRooted($svgFile)) {
                                $svgFile = Join-Path $manifestDir $svgFile
                            }
                            if (-not (Test-Path -LiteralPath $svgFile)) {
                                Write-Warning "    svg not found: $svgFile; skipped."
                                $failedItems += $svgFile
                                continue
                            }
                            $svgAttempted++
                            Write-Host ("    [{0}/{1}] {2}" -f ($i+1), $edit.items.Count, $svgFile)

                            $duped    = $pres.Slides.Item($tplIdx).Duplicate()
                            $newSlide = $null
                            try   { $newSlide = $duped.Item(1) } catch { $newSlide = $duped }

                            $targetPos = $baseInsAfter + $i + 1
                            $newSlide.MoveTo($targetPos)
                            if ($targetPos -le $tplIdx) { $tplIdx++ }

                            $workSlide = $pres.Slides.Item($targetPos)
                            if ($edit.clearContent) { Clear-SlideContent -Slide $workSlide }
                            Set-SlideTexts -Slide $workSlide -Texts $item

                            if (Insert-SvgIntoSlide -PptApp $ppt -Slide $workSlide -SvgFile $svgFile -Zone $zone) {
                                $convertedCount++
                            } else {
                                $failedItems += $svgFile
                            }
                        }
                    }

                    default {
                        Write-Warning "  [manifest] unknown edit type '$type'; skipped."
                    }
                }
            }
        }
    }

    Write-Host "Saving to $OutputPath..."
    if ($outputIsSameAsTemplate) { $pres.Save() }
    else                          { $pres.SaveAs($OutputPath, $ppSaveAsOpenXMLPresentation) }
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
if ($mode -eq 'TEXT') {
    Write-Host "Done. Edited text on $textEditCount slide(s) → $OutputPath"
} else {
    Write-Host ("Done. Converted {0} of {1} SVG(s); text edits on {2} slide(s) → {3}" `
                -f $convertedCount, $svgAttempted, $textEditCount, $OutputPath)
}

if ($failedItems.Count -gt 0) {
    Write-Host "Failed ($($failedItems.Count)):"
    $failedItems | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
