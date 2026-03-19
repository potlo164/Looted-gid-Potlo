Add-Type -AssemblyName System.Drawing

$waifu2x  = "C:\Users\Loipuyt\Desktop\looted Gid\8-- Looted - Copy\tools\waifu2x\waifu2x-ncnn-vulkan-20250915-windows\waifu2x-ncnn-vulkan.exe"
$base     = "C:\Users\Loipuyt\Desktop\looted Gid\8-- Looted - Copy\assets"
$iconBase = "https://diablo2.io/styles/zulu/theme/images/items/"
$cellSize = 56

# name, invfile, image URL fragment
$items = @(
    @{name="Defender's Bile";              inv="invdfb";  img="colossal_jewel1_graphic"},
    @{name="Defender's Fire";              inv="invdff";  img="colossal_jewel1_graphic"},
    @{name="Entropy Locket";               inv="inventl"; img="amu1_graphic"},
    @{name="Guardian's Light";             inv="invgdl";  img="colossal_jewel3_graphic"},
    @{name="Guardian's Thunder";           inv="invgdt";  img="colossal_jewel3_graphic"},
    @{name="Latent Black Cleft";           inv="invlbc";  img="blackcleft_graphic"},
    @{name="Latent Bone Break";            inv="invlbb2"; img="bonebreakcharm_graphic"},
    @{name="Latent Cold Rupture";          inv="invlcr";  img="coldrupture_graphic"},
    @{name="Latent Crack of the Heavens";  inv="invlch";  img="crackofheavens_graphic"},
    @{name="Latent Flame Rift";            inv="invlfr";  img="flamerift_graphic"},
    @{name="Latent Rotting Fissure";       inv="invlrf";  img="rottingfissure_graphic"},
    @{name="Opalvein";                     inv="invopal"; img="ring3_graphic"},
    @{name="Protector's Frost";            inv="invptf";  img="colossal_jewel2_graphic"},
    @{name="Protector's Stone";            inv="invpts";  img="colossal_jewel2_graphic"},
    @{name="Renewed Black Cleft";          inv="invrbc";  img="blackcleft_graphic"},
    @{name="Renewed Bone Break";           inv="invrbb";  img="bonebreakcharm_graphic"},
    @{name="Renewed Cold Rupture";         inv="invrcr";  img="coldrupture_graphic"},
    @{name="Renewed Crack of the Heavens"; inv="invrch";  img="crackofheavens_graphic"},
    @{name="Renewed Flame Rift";           inv="invrfr";  img="flamerift_graphic"},
    @{name="Renewed Rotting Fissure";      inv="invrrf";  img="rottingfissure_graphic"},
    @{name="Skull Collector";              inv="invskco"; img="skullcollector_graphic"},
    @{name="Sling";                        inv="invslng"; img="ring2_graphic"},
    @{name="Wraithstep";                   inv="invwrst"; img="lightplateboots_graphic"}
)

function Process-Icon {
    param([string]$Url, [string]$InvFile, [string]$ItemName)

    $tmpIn   = "$env:TEMP\${InvFile}_in.png"
    $tmpMid  = "$env:TEMP\${InvFile}_mid.png"
    $tmpOut  = "$env:TEMP\${InvFile}_out.png"

    Write-Host "`n[$InvFile] $ItemName" -ForegroundColor Cyan
    try {
        $bytes = (Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop).Content
        [IO.File]::WriteAllBytes($tmpIn, $bytes)
    } catch {
        Write-Host "  DOWNLOAD FAIL: $Url" -ForegroundColor Red
        return $false
    }

    $srcImg  = [System.Drawing.Image]::FromFile($tmpIn)
    $srcW = $srcImg.Width; $srcH = $srcImg.Height
    $srcImg.Dispose()
    $targetW = [Math]::Max(49, [Math]::Round($srcW / $cellSize) * 49)
    $targetH = [Math]::Max(49, [Math]::Round($srcH / $cellSize) * 49)
    Write-Host "  Source: ${srcW}x${srcH} -> Target: ${targetW}x${targetH}"

    # Passe 1 : cunet n=3 s=2 (2x)
    $p1 = Start-Process -FilePath $waifu2x -ArgumentList "-i `"$tmpIn`" -o `"$tmpMid`" -n 3 -s 2 -m models-cunet" -Wait -PassThru -NoNewWindow
    if ($p1.ExitCode -ne 0 -or -not (Test-Path $tmpMid)) {
        Write-Host "  WAIFU2X PASS1 FAIL - fallback upconv" -ForegroundColor Yellow
        $p1 = Start-Process -FilePath $waifu2x -ArgumentList "-i `"$tmpIn`" -o `"$tmpMid`" -n 1 -s 2 -m models-upconv_7_anime_style_art_rgb" -Wait -PassThru -NoNewWindow
        if ($p1.ExitCode -ne 0 -or -not (Test-Path $tmpMid)) {
            Write-Host "  WAIFU2X FAIL" -ForegroundColor Red
            Remove-Item $tmpIn -Force -ErrorAction SilentlyContinue
            return $false
        }
    }

    # Passe 2 : cunet n=3 s=2 (4x total)
    $p2 = Start-Process -FilePath $waifu2x -ArgumentList "-i `"$tmpMid`" -o `"$tmpOut`" -n 3 -s 2 -m models-cunet" -Wait -PassThru -NoNewWindow
    if ($p2.ExitCode -ne 0 -or -not (Test-Path $tmpOut)) {
        Write-Host "  WAIFU2X PASS2 FAIL - using pass1 result" -ForegroundColor Yellow
        Copy-Item $tmpMid $tmpOut -Force
    }
    Remove-Item $tmpMid -Force -ErrorAction SilentlyContinue

    Write-Host "  Upscaled 4x with cunet n=3" -ForegroundColor DarkGreen

    $img = [System.Drawing.Image]::FromFile($tmpOut)
    foreach ($folder in @("gfx_hd_low", "gfx_hd_hd")) {
        $dir = "$base\$folder\$InvFile"
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
        $bmp = New-Object System.Drawing.Bitmap($targetW, $targetH, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::Transparent)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.DrawImage($img, 0, 0, $targetW, $targetH)
        $g.Dispose()
        $bmp.Save("$dir\21.png", [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        Write-Host "  -> $folder\$InvFile\21.png (${targetW}x${targetH})" -ForegroundColor Green
    }
    $img.Dispose()
    Remove-Item $tmpIn, $tmpOut -Force -ErrorAction SilentlyContinue
    Remove-Item $tmpMid -Force -ErrorAction SilentlyContinue
    return $true
}

$ok   = 0
$fail = 0

foreach ($item in $items) {
    $url = $iconBase + $item.img + ".png"
    $res = Process-Icon -Url $url -InvFile $item.inv -ItemName $item.name
    if ($res) { $ok++ } else { $fail++ }
}

Write-Host "`n=== Done: $ok OK, $fail failed ===" -ForegroundColor White
