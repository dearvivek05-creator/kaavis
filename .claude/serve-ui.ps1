$root = Join-Path $PSScriptRoot "..\ui"
$port = if ($env:KAAVIS_PORT) { $env:KAAVIS_PORT } else { 5500 }
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Serving $root on http://localhost:$port/"
while ($listener.IsListening) {
    $context = $listener.GetContext()
    $req = $context.Request
    $res = $context.Response
    $path = $req.Url.LocalPath.TrimStart('/')
    if ([string]::IsNullOrEmpty($path)) { $path = "index.html" }
    $filePath = Join-Path $root $path
    if (Test-Path $filePath -PathType Leaf) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        if ($filePath -like "*.html") { $res.ContentType = "text/html; charset=utf-8" }
        elseif ($filePath -like "*.md") { $res.ContentType = "text/plain; charset=utf-8" }
        $res.ContentLength64 = $bytes.Length
        $res.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $res.StatusCode = 404
    }
    $res.OutputStream.Close()
}
