# Local OCR for the receipt pipeline (Windows built-in OCR, zh-Hans-CN).
# Usage: powershell -ExecutionPolicy Bypass -File ocr_local.ps1 <image> [-Device cpu]
# Output: recognized text lines wrapped between two rows of 60 '-' chars,
#         the format the main pipeline parser expects.
# Cache : if <BASE>\90_linshi\ocr_cache\<MD5 of image>.txt exists, its content is
#         emitted directly and the OCR engine is not called. The cache is keyed by
#         file content hash, so it stays valid across renames and reordering.
# NOTE  : this file is intentionally ASCII-only. Windows PowerShell 5.1 decodes a
#         BOM-less .ps1 as ANSI, which corrupts non-ASCII literals in regexes and
#         path fragments. All CJK characters are built from code points instead.
param(
    [Parameter(Position = 0)][string]$ImagePath,
    [string]$Device = "cpu"
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

$abs = (Resolve-Path -LiteralPath $ImagePath).Path

# Temp/cache directory: "90_" + U+4E34 U+65F6, built from code points.
$tempName = "90_" + [char]0x4E34 + [char]0x65F6
$baseDir = Split-Path -Parent $PSScriptRoot
$cacheDir = Join-Path $baseDir (Join-Path $tempName "ocr_cache")
$tmpDir = Join-Path $baseDir (Join-Path $tempName "ocr_tmp")
$hash = (Get-FileHash -LiteralPath $abs -Algorithm MD5).Hash
$cacheFile = Join-Path $cacheDir ($hash + ".txt")

if (Test-Path -LiteralPath $cacheFile) {
    Write-Output ("-" * 60)
    Get-Content -LiteralPath $cacheFile -Encoding UTF8 | ForEach-Object { Write-Output $_ }
    Write-Output ("-" * 60)
    exit 0
}

# Sidecar cache next to the original image: the pipeline copies each scan into
# ocr_tmp as img_NNNN.jpg before invoking us, which loses the original name. When
# that happens and no content-hash cache exists, fall back to "<original>.ocr.txt"
# recorded by the wrapper in paths.py (see OCR_SIDECAR_MAP env var).
$sidecarMap = if ($env:OCR_SIDECAR_MAP) { $env:OCR_SIDECAR_MAP } else { "" }
if ($sidecarMap -ne "" -and (Test-Path -LiteralPath $sidecarMap)) {
    $mapLine = Get-Content -LiteralPath $sidecarMap -Encoding UTF8 |
        Where-Object { $_ -like "$((Split-Path -Leaf $abs))=`t*" } | Select-Object -First 1
    if ($mapLine) {
        $orig = (Split-Path -Leaf $abs)
        $origPath = $mapLine.Substring($orig.Length + 1)
        $sidecar = $origPath + ".ocr.txt"
        if (Test-Path -LiteralPath $sidecar) {
            Write-Output ("-" * 60)
            Get-Content -LiteralPath $sidecar -Encoding UTF8 | ForEach-Object { Write-Output $_ }
            Write-Output ("-" * 60)
            exit 0
        }
    }
}

Add-Type -AssemblyName System.Runtime.WindowsRuntime

function Await {
    param($AsyncOp, [type]$ResultType)
    $m = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 1
    } | Select-Object -First 1
    $t = $m.MakeGenericMethod($ResultType).Invoke($null, @($AsyncOp))
    $t.Wait(-1) | Out-Null
    return $t.Result
}

[void][Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
[void][Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime]
[void][Windows.Media.Ocr.OcrEngine,Windows.Media,ContentType=WindowsRuntime]

# Characters the engine tends to emit in place of a decimal point.
$ideographicComma = [string][char]0x3001   # U+3001
$fullwidthDot = [string][char]0xFF0E       # U+FF0E
$digitLike = '\d' + [regex]::Escape($fullwidthDot) + [regex]::Escape($ideographicComma) + '\.'

try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($abs)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])

    $engine = $null
    try {
        $lang = [Windows.Globalization.Language]::new("zh-Hans-CN")
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
    } catch { }
    if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
    if (-not $engine) {
        Write-Output ("-" * 60)
        Write-Output "[ERROR] no OCR engine available"
        Write-Output ("-" * 60)
        exit 1
    }

    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    Write-Output ("-" * 60)
    foreach ($line in $result.Lines) {
        # Aggregate the words of one visual line into table cells: split when the
        # horizontal gap between consecutive words exceeds the threshold.
        $cells = @()
        $cur = ""
        $prevEnd = -1.0
        foreach ($word in $line.Words) {
            $x = $word.Rect.X
            if ($prevEnd -ge 0 -and ($x - $prevEnd) -gt 12 -and $cur -ne "") {
                $cells += $cur
                $cur = ""
            }
            $t = $word.Text
            if ($cur -ne "" -and $cur -notmatch '[\u4e00-\u9fa5]$' -and $t -notmatch '^[\u4e00-\u9fa5]' -and $t -notmatch ('^[' + $digitLike + ']') -and $cur -notmatch ('[' + $digitLike + ']$')) {
                $cur += " "
            }
            $cur += $t
            $prevEnd = $x + $word.Rect.Width
        }
        if ($cur -ne "") { $cells += $cur }
        foreach ($c in $cells) {
            Write-Output ($c.Replace($ideographicComma, ".").Replace($fullwidthDot, "."))
        }
    }
    Write-Output ("-" * 60)
} catch {
    Write-Output ("-" * 60)
    Write-Output ("[ERROR] " + $_.Exception.Message)
    Write-Output ("-" * 60)
    exit 1
}
