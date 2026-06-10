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

$rootFullPath = [System.IO.Path]::GetFullPath($Root)
$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
$listener.Start()

Write-Host "WORD LAB local server"
Write-Host "Root: $rootFullPath"
Write-Host "URL : http://127.0.0.1:$Port/app/index.html"
Write-Host "Close this window to stop the server."

while ($true) {
  $client = $listener.AcceptTcpClient()

  try {
    $stream = $client.GetStream()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::ASCII, $false, 4096, $true)
    $requestLine = $reader.ReadLine()

    while ($true) {
      $line = $reader.ReadLine()
      if ($null -eq $line -or $line -eq "") {
        break
      }
    }

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
      $body = [System.Text.Encoding]::UTF8.GetBytes("Server Error")
      Write-HttpResponse $stream 500 "Internal Server Error" $body "text/plain; charset=utf-8" $true
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
