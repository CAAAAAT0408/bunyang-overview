# Windows 내장 OCR(한국어)로 그림 속 글자 읽기 — 평면도의 '기본형'·'확장형' 표기 찾기용
# powershell -File ocr.ps1 -List <목록.txt (UTF-8, 한 줄에 그림 경로 하나)> -Out <결과.jsonl>
param([string]$List, [string]$Out)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait() | Out-Null; $task.Result }
$lang = New-Object Windows.Globalization.Language 'ko'
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
$sw = New-Object System.IO.StreamWriter($Out, $false, (New-Object System.Text.UTF8Encoding($false)))
if (-not $engine) { $sw.WriteLine('{"error":"no-korean-ocr"}'); $sw.Close(); exit }
foreach ($path in [System.IO.File]::ReadAllLines($List, [System.Text.Encoding]::UTF8)) {
  if (-not $path.Trim()) { continue }
  $lines = @()
  try {
    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($path)) ([Windows.Storage.StorageFile])
    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    $result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
    foreach ($line in $result.Lines) {
      $r = $line.Words | ForEach-Object { $_.BoundingRect }
      $x = [int](($r | Measure-Object -Property X -Minimum).Minimum)
      $y = [int](($r | Measure-Object -Property Y -Minimum).Minimum)
      $lines += ,@($x, $y, $line.Text)
    }
    $stream.Dispose()
    $obj = @{ file = $path; w = $bitmap.PixelWidth; h = $bitmap.PixelHeight; lines = $lines }
  } catch {
    $obj = @{ file = $path; error = $_.Exception.Message; lines = @() }
  }
  $sw.WriteLine((ConvertTo-Json $obj -Compress -Depth 5))
  $sw.Flush()
}
$sw.Close()
