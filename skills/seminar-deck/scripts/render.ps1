# Render every slide of a .pptx to PNG through PowerPoint (COM). Never quits PowerPoint:
# the user's own PowerPoint window may be open and must survive.
#   powershell -NoProfile -File render.ps1 -src deck.pptx -outdir png -width 1280
param([string]$src, [string]$outdir, [int]$width = 1280)
$ErrorActionPreference = "Stop"
$src = (Resolve-Path $src).Path
New-Item -ItemType Directory -Force $outdir | Out-Null
$outdir = (Resolve-Path $outdir).Path
$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open($src, $true, $false, $false)   # read-only, untitled, no window
$h = [int]($width * $p.PageSetup.SlideHeight / $p.PageSetup.SlideWidth)
$n = $p.Slides.Count
$pad = if ($n -ge 100) { "D3" } else { "D2" }
foreach ($s in $p.Slides) { $s.Export((Join-Path $outdir ("s{0:$pad}.png" -f $s.SlideIndex)), "PNG", $width, $h) }
"slides: $n  size: " + $p.PageSetup.SlideWidth + "x" + $p.PageSetup.SlideHeight
$p.Close()
