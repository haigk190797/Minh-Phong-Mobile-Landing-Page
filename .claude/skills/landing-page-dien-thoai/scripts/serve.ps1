param(
  [Parameter(Mandatory=$true)]
  [string]$Root,
  [int]$Port = 8765
)

# Static file server dùng để preview landing page cục bộ khi máy không có Node/Python.
# Luôn truyền -Root qua runtimeArgs trong .claude/launch.json (argv), KHÔNG hard-code
# đường dẫn có dấu tiếng Việt trong nội dung script này — powershell.exe có thể đọc sai
# encoding của file .ps1 và biến chuỗi tiếng Việt thành ký tự rác.

Add-Type -AssemblyName System.Net.HttpListener -ErrorAction SilentlyContinue

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()
Write-Host "Serving $Root on http://localhost:$Port/"

$mime = @{
  ".html"="text/html; charset=utf-8"
  ".css"="text/css"
  ".js"="application/javascript"
  ".png"="image/png"
  ".jpg"="image/jpeg"
  ".jpeg"="image/jpeg"
  ".svg"="image/svg+xml"
  ".ico"="image/x-icon"
}

while ($listener.IsListening) {
  $context = $listener.GetContext()
  $req = $context.Request
  $res = $context.Response
  try {
    $path = [System.Uri]::UnescapeDataString($req.Url.AbsolutePath)
    if ($path -eq "/") { $path = "/index.html" }
    $filePath = Join-Path $Root ($path.TrimStart('/'))
    if (Test-Path $filePath -PathType Leaf) {
      $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
      $ct = $mime[$ext]
      if (-not $ct) { $ct = "application/octet-stream" }
      $bytes = [System.IO.File]::ReadAllBytes($filePath)
      $res.ContentType = $ct
      $res.ContentLength64 = $bytes.Length
      $res.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
      $res.StatusCode = 404
      $msg = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found: $path")
      $res.OutputStream.Write($msg, 0, $msg.Length)
    }
  } catch {
    $res.StatusCode = 500
  } finally {
    $res.OutputStream.Close()
  }
}
