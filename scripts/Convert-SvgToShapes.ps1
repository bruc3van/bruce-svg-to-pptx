#Requires -Version 5.1
<#
.SYNOPSIS
    Convert SVG files into native, editable PowerPoint shapes via COM automation.

.DESCRIPTION
    Drives Microsoft PowerPoint on Windows to:
      1. Insert each SVG onto a blank slide as a msoGraphic shape.
      2. Select that graphic and call CommandBars.ExecuteMso("SVGEdit"),
         which is the Ribbon command for "Convert to Shape".
      3. Save the deck as .pptx.

    This is the same operation as right-clicking an inserted SVG in PowerPoint
    and choosing "Convert to Shape" — Microsoft does not expose a first-class
    VBA method for the conversion, so the Mso command is the only hook.

.PARAMETER SvgPath
    One or more .svg file paths, or a directory containing .svg files.
    Directory mode is non-recursive.

.PARAMETER OutputPath
    Destination .pptx path. Created if it does not exist.

.PARAMETER Append
    Append to an existing OutputPath deck instead of overwriting.

.PARAMETER Force
    Skip the overwrite confirmation when OutputPath already exists.

.EXAMPLE
    .\Convert-SvgToShapes.ps1 -SvgPath .\icon.svg -OutputPath .\out.pptx

.EXAMPLE
    .\Convert-SvgToShapes.ps1 -SvgPath .\icons\ -OutputPath .\icons.pptx

.EXAMPLE
    .\Convert-SvgToShapes.ps1 -SvgPath .\new\ -OutputPath .\deck.pptx -Append

.NOTES
    Requires Windows + Microsoft PowerPoint 2016 build 1712 or later.
    The PowerPoint window is visible during conversion (required by
    CommandBars.ExecuteMso) and closes automatically when finished.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string[]]$SvgPath,

    [Parameter(Mandatory)]
    [string]$OutputPath,

    [switch]$Append,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ---------- PowerPoint / Office constants ----------
$msoFalse                       = 0
$msoTrue                        = -1
$msoGraphic                     = 28
$ppLayoutBlank                  = 12
$ppSaveAsOpenXMLPresentation    = 24

# ---------- Helpers ----------

function Resolve-SvgFiles {
    param([string[]]$Inputs)

    $files = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Inputs) {
        if (-not (Test-Path -LiteralPath $p)) {
            throw "Path not found: $p"
        }
        $item = Get-Item -LiteralPath $p
        if ($item.PSIsContainer) {
            $found = Get-ChildItem -LiteralPath $item.FullName -Filter *.svg -File |
                Sort-Object Name |
                Select-Object -ExpandProperty FullName
            if ($found.Count -eq 0) {
                Write-Warning "No .svg files found in directory: $($item.FullName)"
            }
            $found | ForEach-Object { $files.Add($_) }
        } else {
            if ($item.Extension -ine '.svg') {
                throw "Not an .svg file: $($item.FullName)"
            }
            $files.Add($item.FullName)
        }
    }
    if ($files.Count -eq 0) { throw "No .svg files to process." }
    return $files.ToArray()
}

function Get-SvgIntrinsicSize {
    <#
        Parse the SVG to recover its intrinsic aspect ratio, then scale to
        fill the slide while preserving aspect ratio (letterbox if needed).
        Returns @{ Width; Height; Left; Top } in PowerPoint points.
        Falls back to filling the slide if parsing fails.
    #>
    param(
        [string]$Path,
        [double]$SlideWidth  = 720.0,
        [double]$SlideHeight = 540.0
    )

    $fallback = @{
        Width  = $SlideWidth
        Height = $SlideHeight
        Left   = 0.0
        Top    = 0.0
    }

    try {
        [xml]$xml = Get-Content -LiteralPath $Path -Raw
        $svg = $xml.svg
        if (-not $svg) { return $fallback }

        $w = $null; $h = $null

        if ($svg.viewBox) {
            $parts = $svg.viewBox.Trim() -split '[\s,]+' | Where-Object { $_ -ne '' }
            if ($parts.Count -eq 4) {
                $w = [double]$parts[2]
                $h = [double]$parts[3]
            }
        }

        if (-not $w -or -not $h) {
            if ($svg.width -and $svg.height) {
                $w = [double]([regex]::Match($svg.width,  '[\d.]+').Value)
                $h = [double]([regex]::Match($svg.height, '[\d.]+').Value)
            }
        }

        if (-not $w -or -not $h -or $w -le 0 -or $h -le 0) { return $fallback }

        # Scale to fill slide, maintaining aspect ratio (letterbox if needed)
        $scale  = [Math]::Min($SlideWidth / $w, $SlideHeight / $h)
        $fitW   = [double]($w * $scale)
        $fitH   = [double]($h * $scale)
        return @{
            Width  = $fitW
            Height = $fitH
            Left   = ($SlideWidth  - $fitW) / 2.0
            Top    = ($SlideHeight - $fitH) / 2.0
        }
    }
    catch {
        return $fallback
    }
}

function Release-Com {
    param([object]$ComObject)
    if ($null -ne $ComObject) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) } catch {}
    }
}

# ---------- Pre-flight ----------

$svgFiles = Resolve-SvgFiles -Inputs $SvgPath

# Resolve OutputPath to absolute. PowerPoint COM is unreliable with relative paths.
$outputDir = Split-Path -Path $OutputPath -Parent
if ([string]::IsNullOrEmpty($outputDir)) { $outputDir = (Get-Location).Path }
if (-not (Test-Path -LiteralPath $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
$outputDirAbs = (Resolve-Path -LiteralPath $outputDir).Path
$OutputPath   = Join-Path $outputDirAbs (Split-Path -Path $OutputPath -Leaf)

# Reject .ppt — we only output .pptx
if ([System.IO.Path]::GetExtension($OutputPath) -ine '.pptx') {
    throw "OutputPath must end in .pptx (got: $OutputPath)"
}

$outputExists = Test-Path -LiteralPath $OutputPath

if ($outputExists -and -not $Append -and -not $Force) {
    $resp = Read-Host "Overwrite existing $OutputPath? [y/N]"
    if ($resp -notmatch '^[Yy]') {
        Write-Host "Aborted."
        return
    }
}

# Try to instantiate PowerPoint. Fail fast with a clear message if it isn't installed.
Write-Host "Starting PowerPoint..."
try {
    $ppt = New-Object -ComObject PowerPoint.Application
} catch {
    throw "Could not start PowerPoint via COM. Is Microsoft PowerPoint installed on this machine? Underlying error: $($_.Exception.Message)"
}

# ExecuteMso requires Visible = True. There is no reliable hidden mode for this workflow.
$ppt.Visible = $msoTrue

# Warn if other presentations are already open — they can interfere with selection state.
if ($ppt.Presentations.Count -gt 0) {
    Write-Warning "PowerPoint already has $($ppt.Presentations.Count) presentation(s) open. Close them for best results."
}

$pres = $null
$convertedCount = 0
$failedItems    = @()

try {
    if ($outputExists -and $Append) {
        Write-Host "Opening existing $OutputPath for append..."
        $pres = $ppt.Presentations.Open($OutputPath, $msoFalse, $msoFalse, $msoTrue)
    } else {
        Write-Host "Creating new presentation..."
        $pres = $ppt.Presentations.Add($msoTrue)
    }

    $slideWidth  = $pres.PageSetup.SlideWidth
    $slideHeight = $pres.PageSetup.SlideHeight

    foreach ($svg in $svgFiles) {
        Write-Host "  [$($convertedCount + 1)/$($svgFiles.Count)] $svg"

        $size = Get-SvgIntrinsicSize -Path $svg -SlideWidth $slideWidth -SlideHeight $slideHeight
        $left = $size.Left
        $top  = $size.Top

        $slideIndex = $pres.Slides.Count + 1
        $slide      = $pres.Slides.Add($slideIndex, $ppLayoutBlank)

        $shape = $null
        try {
            $shape = $slide.Shapes.AddPicture(
                $svg,
                $msoFalse,    # LinkToFile
                $msoTrue,     # SaveWithDocument
                $left,
                $top,
                $size.Width,
                $size.Height
            )
        } catch {
            Write-Warning "    AddPicture failed: $($_.Exception.Message)"
            $failedItems += $svg
            continue
        }

        if ($shape.Type -ne $msoGraphic) {
            Write-Warning "    Inserted as Type=$($shape.Type), expected msoGraphic ($msoGraphic). PowerPoint may not have recognized this as SVG; conversion will be skipped."
            $failedItems += $svg
            continue
        }

        # Activate the slide and select the shape so ExecuteMso has a target.
        try {
            $slide.Select()
            $shape.Select()
            # Tiny pause: ExecuteMso occasionally races the selection update.
            Start-Sleep -Milliseconds 250

            $ppt.CommandBars.ExecuteMso("SVGEdit")
            $convertedCount++
        } catch {
            Write-Warning "    ExecuteMso('SVGEdit') failed: $($_.Exception.Message)"
            $failedItems += $svg
        }
    }

    Write-Host "Saving to $OutputPath..."
    if ($outputExists -and -not $Append) {
        # Overwriting an existing file: SaveAs handles this with the format arg.
        $pres.SaveAs($OutputPath, $ppSaveAsOpenXMLPresentation)
    } elseif ($outputExists -and $Append) {
        $pres.Save()
    } else {
        $pres.SaveAs($OutputPath, $ppSaveAsOpenXMLPresentation)
    }
    $pres.Close()
}
finally {
    if ($null -ne $pres) { Release-Com $pres }
    try { $ppt.Quit() } catch {}
    Release-Com $ppt
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Done. Converted $convertedCount of $($svgFiles.Count) SVG(s)."
if ($failedItems.Count -gt 0) {
    Write-Host "Failed:"
    $failedItems | ForEach-Object { Write-Host "  - $_" }
    exit 1
}
