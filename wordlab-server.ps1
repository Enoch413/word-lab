param(
  [int]$Port = 8763,
  [string]$Root = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

function Get-ContentType {
  param([string]$Path)

  $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
  switch ($extension) {
    ".html" { "text/html; charset=utf-8"; break }
    ".css" { "text/css; charset=utf-8"; break }
    ".js" { "application/javascript; charset=utf-8"; break }
    ".json" { "application/json; charset=utf-8"; break }
    ".txt" { "text/plain; charset=utf-8"; break }
    ".svg" { "image/svg+xml"; break }
    ".png" { "image/png"; break }
    ".jpg" { "image/jpeg"; break }
    ".jpeg" { "image/jpeg"; break }
    ".gif" { "image/gif"; break }
    ".ico" { "image/x-icon"; break }
    ".woff" { "font/woff"; break }
    ".woff2" { "font/woff2"; break }
    ".ttf" { "font/ttf"; break }
    default { "application/octet-stream" }
  }
}

function Write-HttpResponse {
  param(
    [System.Net.Sockets.NetworkStream]$Stream,
    [int]$StatusCode,
    [string]$StatusText,
    [byte[]]$Body,
    [string]$ContentType = "text/plain; charset=utf-8",
    [bool]$IncludeBody = $true
  )

  $headers = @(
    "HTTP/1.1 $StatusCode $StatusText",
    "Content-Type: $ContentType",
    "Content-Length: $($Body.Length)",
    "Cache-Control: no-store",
    "X-Word-Lab-Root: $rootIdentity",
    "Connection: close",
    "",
    ""
  ) -join "`r`n"

  $headerBytes = [System.Text.Encoding]::ASCII.GetBytes($headers)
  $Stream.Write($headerBytes, 0, $headerBytes.Length)

  if ($IncludeBody -and $Body.Length -gt 0) {
    $Stream.Write($Body, 0, $Body.Length)
  }
}

function Resolve-RequestPath {
  param(
    [string]$RequestPath,
    [string]$RootPath
  )

  $path = [System.Uri]::UnescapeDataString($RequestPath.Split("?")[0])
  if ([string]::IsNullOrWhiteSpace($path) -or $path -eq "/") {
    $path = "/index.html"
  }

  $publicPaths = @(
    "/index.html", "/app/index.html", "/app/app.js", "/app/styles.css",
    "/app/firebase-config.js", "/app/vendor/html2canvas.min.js", "/app/vendor/jspdf.umd.min.js"
  )
  if ($publicPaths -notcontains $path) {
    return $null
  }

  $relativePath = $path.TrimStart("/") -replace "/", [System.IO.Path]::DirectorySeparatorChar
  $rootFullPath = [System.IO.Path]::GetFullPath($RootPath).TrimEnd("\", "/")
  $fullPath = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($rootFullPath, $relativePath))
  $rootPrefix = $rootFullPath + [System.IO.Path]::DirectorySeparatorChar

  if ($fullPath -ne $rootFullPath -and -not $fullPath.StartsWith($rootPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
    return $null
  }

  if ([System.IO.Directory]::Exists($fullPath)) {
    $fullPath = [System.IO.Path]::Combine($fullPath, "index.html")
  }

  return $fullPath
}

function Read-RequestLine {
  param([System.Net.Sockets.NetworkStream]$Stream)

  $buffer = New-Object 'System.Collections.Generic.List[byte]'
  $timer = [System.Diagnostics.Stopwatch]::StartNew()
  while ($buffer.Count -lt 32768) {
    if ($timer.ElapsedMilliseconds -ge 1000) {
      throw [System.TimeoutException]::new("Request headers timed out")
    }
    $Stream.ReadTimeout = [int][Math]::Max(1, [Math]::Min(500, 1000 - $timer.ElapsedMilliseconds))
    $next = $Stream.ReadByte()
    if ($next -lt 0) { return "" }
    $buffer.Add([byte]$next)
    $length = $buffer.Count
    if ($length -ge 4 -and $buffer[$length - 4] -eq 13 -and $buffer[$length - 3] -eq 10 -and $buffer[$length - 2] -eq 13 -and $buffer[$length - 1] -eq 10) {
      $headers = [System.Text.Encoding]::ASCII.GetString($buffer.ToArray())
      return $headers.Split("`r`n")[0]
    }
  }
  throw [System.IO.InvalidDataException]::new("Request headers too large")
}

$rootFullPath = [System.IO.Path]::GetFullPath($Root).TrimEnd("\", "/")
$rootIdentity = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($rootFullPath))
$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()

Write-Host "WORD LAB local server"
Write-Host "Root: $rootFullPath"
Write-Host "URL : http://127.0.0.1:$Port/app/index.html"
Write-Host "Close this window to stop the server."

while ($true) {
  $client = $listener.AcceptTcpClient()
  $stream = $null

  try {
    $stream = $client.GetStream()
    $stream.WriteTimeout = 2000
    $requestLine = Read-RequestLine $stream

    if ([string]::IsNullOrWhiteSpace($requestLine)) {
      continue
    }

    $parts = $requestLine.Split(" ")
    $method = $parts[0]
    $requestPath = if ($parts.Length -gt 1) { $parts[1] } else { "/" }
    $includeBody = $method -ne "HEAD"

    if ($method -ne "GET" -and $method -ne "HEAD") {
      $body = [System.Text.Encoding]::UTF8.GetBytes("Method Not Allowed")
      Write-HttpResponse $stream 405 "Method Not Allowed" $body "text/plain; charset=utf-8" $includeBody
      continue
    }

    $filePath = Resolve-RequestPath $requestPath $rootFullPath
    if ($null -eq $filePath) {
      $body = [System.Text.Encoding]::UTF8.GetBytes("Forbidden")
      Write-HttpResponse $stream 403 "Forbidden" $body "text/plain; charset=utf-8" $includeBody
      continue
    }

    if (-not [System.IO.File]::Exists($filePath)) {
      $body = [System.Text.Encoding]::UTF8.GetBytes("Not Found")
      Write-HttpResponse $stream 404 "Not Found" $body "text/plain; charset=utf-8" $includeBody
      continue
    }

    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    Write-HttpResponse $stream 200 "OK" $bytes (Get-ContentType $filePath) $includeBody
  } catch {
    try {
      $statusCode = 408
      $statusText = "Request Timeout"
      if ($_.Exception -is [System.IO.InvalidDataException]) {
        $statusCode = 431
        $statusText = "Request Header Fields Too Large"
      } elseif (-not ($_.Exception -is [System.IO.IOException]) -and -not ($_.Exception -is [System.TimeoutException])) {
        $statusCode = 500
        $statusText = "Internal Server Error"
      }
      $body = [System.Text.Encoding]::UTF8.GetBytes($statusText)
      Write-HttpResponse $stream $statusCode $statusText $body "text/plain; charset=utf-8" $true
      if ($statusCode -eq 431) {
        # Consume bounded buffered excess so Windows delivers the response before closing.
        $discard = New-Object byte[] 4096
        $drained = 0
        while ($stream.DataAvailable -and $drained -lt 32768) {
          $drained += $stream.Read($discard, 0, $discard.Length)
        }
      }
    } catch {
      # The client may have disconnected before the response could be written.
    }
  } finally {
    if ($null -ne $stream) {
      $stream.Close()
    }
    $client.Close()
  }
}
