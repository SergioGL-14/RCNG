# Genera un icono 16x16 para el elemento de menú DHCP (azul con texto "IP")
try {
    $bmp = New-Object System.Drawing.Bitmap(16, 16, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit

    # Fondo degradado azul
    $gradRect = New-Object System.Drawing.Rectangle(0, 0, 16, 16)
    $grad = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $gradRect,
        [System.Drawing.Color]::FromArgb(30, 130, 220),
        [System.Drawing.Color]::FromArgb(0, 90, 170),
        [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
    $g.FillRectangle($grad, $gradRect)
    $grad.Dispose()

    # Línea horizontal blanca (símbolo de red)
    $penW = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(200, 255, 255, 255), 1)
    $g.DrawLine($penW, 3, 8, 13, 8)
    $g.DrawLine($penW, 3, 6, 3, 10)
    $g.DrawLine($penW, 13, 6, 13, 10)
    $g.DrawLine($penW, 8, 5, 8, 11)
    $penW.Dispose()

    $g.Dispose()
    $ToolStripMenuItem_DHCP.Image = $bmp
} catch {}
