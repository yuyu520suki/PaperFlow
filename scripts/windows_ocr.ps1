param(
    [string]$ImagePath,
    [string]$Language,
    [switch]$Status,
    [ValidateRange(1, 120)][int]$TimeoutSeconds = 30
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
    [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
    [Windows.Globalization.Language, Windows.Globalization, ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.FileAccessMode, Windows.Storage, ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.BitmapPixelFormat, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.BitmapAlphaMode, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null

    $available = @([Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages | ForEach-Object {
        @{ tag = $_.LanguageTag; display_name = $_.DisplayName }
    })
    if ($Status) {
        @{
            engine = 'Windows.Media.Ocr'
            available = ($available.Count -gt 0)
            languages = $available
            max_image_dimension = [Windows.Media.Ocr.OcrEngine]::MaxImageDimension
            local_only = $true
            powershell_version = $PSVersionTable.PSVersion.ToString()
        } | ConvertTo-Json -Depth 8 -Compress
        exit 0
    }

    if ([string]::IsNullOrWhiteSpace($ImagePath)) { throw 'ImagePath is required' }
    $resolved = (Resolve-Path -LiteralPath $ImagePath).ProviderPath
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) { throw 'ImagePath is not a file' }

    $asTaskMethod = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
        $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and
        $_.GetParameters().Count -eq 1 -and
        $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
    } | Select-Object -First 1
    if ($null -eq $asTaskMethod) { throw 'WinRT task adapter is unavailable' }

    function Await-Operation {
        param($Operation, [Type]$ResultType)
        $task = $asTaskMethod.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
        if (-not $task.Wait($TimeoutSeconds * 1000)) { throw 'Windows OCR operation timed out' }
        return $task.GetAwaiter().GetResult()
    }

    if ([string]::IsNullOrWhiteSpace($Language)) {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    } else {
        $requested = New-Object Windows.Globalization.Language($Language)
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($requested)
    }
    if ($null -eq $engine) { throw 'Requested OCR language is not installed; no language pack was installed' }

    $storage = Await-Operation ([Windows.Storage.StorageFile]::GetFileFromPathAsync($resolved)) ([Windows.Storage.StorageFile])
    $stream = $null
    $bitmap = $null
    try {
        $stream = Await-Operation ($storage.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        $decoder = Await-Operation ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        if ($decoder.FrameCount -ne 1) { throw 'Multi-frame image requires per-frame extraction; only single-page images are accepted' }
        $bitmap = Await-Operation ($decoder.GetSoftwareBitmapAsync([Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8, [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied)) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $maxDimension = [Windows.Media.Ocr.OcrEngine]::MaxImageDimension
        if ($bitmap.PixelWidth -gt $maxDimension -or $bitmap.PixelHeight -gt $maxDimension) {
            throw 'Image exceeds Windows OCR MaxImageDimension; explicitly rescale a working copy first'
        }
        $result = Await-Operation ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $lines = @()
        foreach ($line in $result.Lines) {
            $words = @()
            $left = [double]::PositiveInfinity
            $top = [double]::PositiveInfinity
            $right = [double]::NegativeInfinity
            $bottom = [double]::NegativeInfinity
            foreach ($word in $line.Words) {
                $box = $word.BoundingRect
                $words += @{
                    text = $word.Text
                    bbox = @{ x = $box.X; y = $box.Y; width = $box.Width; height = $box.Height }
                }
                $left = [math]::Min($left, $box.X)
                $top = [math]::Min($top, $box.Y)
                $right = [math]::Max($right, $box.X + $box.Width)
                $bottom = [math]::Max($bottom, $box.Y + $box.Height)
            }
            $lineBox = $null
            if ($words.Count -gt 0) {
                $lineBox = @{ x = $left; y = $top; width = $right - $left; height = $bottom - $top }
            }
            $lines += @{ text = $line.Text; bbox = $lineBox; words = $words }
        }
        @{
            engine = 'Windows.Media.Ocr'
            language = $engine.RecognizerLanguage.LanguageTag
            source = $resolved
            text = $result.Text
            image = @{ width = $bitmap.PixelWidth; height = $bitmap.PixelHeight; frames = $decoder.FrameCount }
            text_angle_degrees = $result.TextAngle
            lines = $lines
            local_only = $true
            requires_review = $true
            review_reasons = @('OCR may misread scientific symbols', 'Equations are not reconstructed as editable mathematics', 'Table structure and reading order require review')
        } | ConvertTo-Json -Depth 12 -Compress
    } finally {
        if ($null -ne $bitmap) { $bitmap.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
} catch {
    @{ error = 'WindowsOcrError'; message = $_.Exception.Message } | ConvertTo-Json -Compress
    exit 1
}
