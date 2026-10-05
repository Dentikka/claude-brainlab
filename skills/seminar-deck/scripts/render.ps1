# Render every slide of a .pptx to PNG through PowerPoint (COM) and report layout problems.
# Never quits PowerPoint: the user's own PowerPoint window may be open and must survive.
#   powershell -NoProfile -File render.ps1 -src deck.pptx -outdir png -width 1280
#   powershell -NoProfile -File render.ps1 -src deck.pptx -outdir png -pdf deck.pdf
#   powershell -NoProfile -File render.ps1 -src deck.pptx -pdf deck.pdf -nocheck   # PDF only
# Layout report (lines starting with CHECK), measured by PowerPoint itself:
#   - text of several lines that does not fit its box (spills onto what lies below);
#   - a title (20 pt and larger) that wraps, except on the first slide;
#   - a short label (up to three words) that wraps;
#   - a shape that sticks out of the slide.
# A line broken by hand is not a wrap: meme captions and two-line labels pass.
param([string]$src, [string]$outdir = "", [int]$width = 1280, [string]$pdf = "", [switch]$nocheck)
$ErrorActionPreference = "Stop"
$src = (Resolve-Path $src).Path
$pp = New-Object -ComObject PowerPoint.Application
$p = $pp.Presentations.Open($src, $true, $false, $false)   # read-only, untitled, no window
try {
    $sw, $sh_ = $p.PageSetup.SlideWidth, $p.PageSetup.SlideHeight
    $n = $p.Slides.Count
    if ($outdir) {
        New-Item -ItemType Directory -Force $outdir | Out-Null
        $outdir = (Resolve-Path $outdir).Path
        $h = [int]($width * $sh_ / $sw)
        $pad = if ($n -ge 100) { "D3" } else { "D2" }
        foreach ($s in $p.Slides) { $s.Export((Join-Path $outdir ("s{0:$pad}.png" -f $s.SlideIndex)), "PNG", $width, $h) }
    }
    "slides: $n  size: $sw x $sh_"
    if (-not $nocheck) {
        $issues = 0
        foreach ($s in $p.Slides) {
            $i = $s.SlideIndex
            foreach ($shape in $s.Shapes) {
                $w, $h = $shape.Width, $shape.Height
                $r = [math]::Abs($shape.Rotation % 180)
                if ($r -ge 45 -and $r -le 135) { $w, $h = $h, $w }      # turned on its side
                $cx, $cy = ($shape.Left + $shape.Width / 2), ($shape.Top + $shape.Height / 2)
                if ($cx - $w / 2 -lt -1 -or $cy - $h / 2 -lt -1 -or $cx + $w / 2 -gt $sw + 1 -or $cy + $h / 2 -gt $sh_ + 1) {
                    "CHECK s{0:D2} off the slide: {1}" -f $i, $shape.Name; $issues++
                }
                if (-not $shape.HasTextFrame) { continue }
                if (-not $shape.TextFrame.HasText) { continue }
                $tf2 = $shape.TextFrame2
                $tr = $shape.TextFrame.TextRange
                $lines = $tr.Lines().Count
                $manual = ($tr.Text -split "`r`n|[`r`n`v]").Count         # lines broken by hand
                $wraps = $lines -gt $manual                                # PowerPoint wrapped a line
                $size = $tr.Font.Size
                $text = ($tr.Text -replace "[`r`n`v]+", " / ").Trim()
                $short = if ($text.Length -gt 50) { $text.Substring(0, 50) + "..." } else { $text }
                $words = ($text -split "\s+" | Where-Object { $_ }).Count
                $room = $shape.Height - $tf2.MarginTop - $tf2.MarginBottom
                # AutoSize 1 (shape to fit text) counts too: PowerPoint resizes the shape only when
                # the text is edited, so a generated box keeps its size and the text spills out
                if ($tf2.AutoSize -ne 2 -and $lines -gt 1 -and $tf2.TextRange.BoundHeight -gt $room + 2) {
                    "CHECK s{0:D2} overflow: {1} lines need {2:N0} pt, box {3:N0} pt: `"{4}`"" -f $i, $lines,
                        $tf2.TextRange.BoundHeight, $room, $short; $issues++
                }
                if ($i -gt 1 -and $size -ge 20 -and $wraps) {
                    "CHECK s{0:D2} title wraps ({1} lines, {2} pt): `"{3}`"" -f $i, $lines, $size, $short; $issues++
                }
                elseif ($wraps -and $words -le 3) {
                    "CHECK s{0:D2} short label wraps ({1} lines): `"{2}`"" -f $i, $lines, $short; $issues++
                }
            }
        }
        if ($issues) { "layout: $issues issue(s) - look at these slides" } else { "layout: clean" }
    }
    if ($pdf) {
        $pdfPath = if ([System.IO.Path]::IsPathRooted($pdf)) { $pdf } else { Join-Path (Get-Location).Path $pdf }
        $p.SaveAs($pdfPath, 32)                                  # 32 = ppSaveAsPDF
        "pdf: $pdfPath"
    }
}
finally { $p.Close() }
