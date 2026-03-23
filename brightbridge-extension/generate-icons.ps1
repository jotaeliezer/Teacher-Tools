Add-Type -AssemblyName System.Drawing
$iconDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$iconDir = Join-Path $iconDir "icons"

$sizes = @(16, 48, 128)
foreach ($size in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)

    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(108, 63, 162))
    $g.FillEllipse($brush, 0, 0, ($size - 1), ($size - 1))

    $fontSize = [Math]::Max(6, [int]($size * 0.50))
    $font = New-Object System.Drawing.Font("Arial", $fontSize, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $rect = New-Object System.Drawing.RectangleF(0, 0, $size, $size)
    $g.DrawString("B", $font, [System.Drawing.Brushes]::White, $rect, $sf)

    $outPath = Join-Path $iconDir ("icon" + $size + ".png")
    $bmp.Save($outPath)
    $g.Dispose()
    $bmp.Dispose()
    Write-Host "Saved $outPath"
}
Write-Host "Done."
