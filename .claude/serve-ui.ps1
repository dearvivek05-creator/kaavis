$root = Join-Path $PSScriptRoot "..\ui"
$port = if ($env:KAAVIS_PORT) { $env:KAAVIS_PORT } else { 5550 }

# Friendly names -> the executable/URI Start-Process should launch.
# Anything not listed here is passed through as-is (works for anything already
# resolvable via PATH or the Windows "App Paths" registry, e.g. "spotify").
$appAliases = @{
    "notepad" = "notepad"
    "calculator" = "calc"; "calc" = "calc"
    "chrome" = "chrome"; "google chrome" = "chrome"
    "edge" = "msedge"; "microsoft edge" = "msedge"
    "firefox" = "firefox"
    "explorer" = "explorer"; "file explorer" = "explorer"; "files" = "explorer"
    "word" = "winword"; "microsoft word" = "winword"
    "excel" = "excel"; "microsoft excel" = "excel"
    "powerpoint" = "powerpnt"
    "vscode" = "code"; "vs code" = "code"; "visual studio code" = "code"
    "paint" = "mspaint"
    "terminal" = "wt"; "command prompt" = "cmd"; "powershell" = "powershell"
    "settings" = "ms-settings:"
    "spotify" = "spotify"
}

function Write-JsonResponse($res, $obj, $status = 200) {
    $res.StatusCode = $status
    $res.ContentType = "application/json; charset=utf-8"
    $res.Headers.Add("Access-Control-Allow-Origin", "*")
    $json = $obj | ConvertTo-Json -Compress
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $res.ContentLength64 = $bytes.Length
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.OutputStream.Close()
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Serving $root on http://localhost:$port/ (with /api/open and /api/ping)"

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $req = $context.Request
    $res = $context.Response
    $path = $req.Url.LocalPath

    if ($path -eq "/api/ping") {
        Write-JsonResponse $res @{ ok = $true }
        continue
    }

    if ($path -eq "/api/open") {
        $appName = $req.QueryString["app"]
        if ([string]::IsNullOrWhiteSpace($appName)) {
            Write-JsonResponse $res @{ ok = $false; error = "No app specified" } 400
            continue
        }
        $key = $appName.Trim().ToLower()
        $target = if ($appAliases.ContainsKey($key)) { $appAliases[$key] } else { $appName.Trim() }
        try {
            $proc = Start-Process $target -PassThru -ErrorAction Stop
            Write-JsonResponse $res @{ ok = $true; app = $target; pid = $proc.Id }
        } catch {
            Write-JsonResponse $res @{ ok = $false; error = $_.Exception.Message }
        }
        continue
    }

    $relPath = $path.TrimStart('/')
    if ([string]::IsNullOrEmpty($relPath)) { $relPath = "index.html" }
    $filePath = Join-Path $root $relPath
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
