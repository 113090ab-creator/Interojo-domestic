$ErrorActionPreference = "SilentlyContinue"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectDir

$port = 8514
$pidFile = Join-Path $projectDir ".dashboard_pid"

function Stop-Pid([int]$pidNum) {
    cmd /c "taskkill /PID $pidNum /F >nul 2>&1" | Out-Null
}

if (Test-Path $pidFile) {
    $savedPid = (Get-Content -LiteralPath $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
    $pidNum = 0
    if ([int]::TryParse($savedPid, [ref]$pidNum)) {
        Stop-Pid -pidNum $pidNum
        Write-Host "[INFO] Stopped dashboard PID $pidNum."
    }
    Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue
}

$lines = cmd /c "netstat -ano -p tcp | findstr /r /c:"":$port .*LISTENING""" 2>$null
$killed = @{}
foreach ($line in $lines) {
    $parts = ($line -split "\s+") | Where-Object { $_ -ne "" }
    if ($parts.Count -ge 5) {
        $pidText = $parts[$parts.Count - 1]
        $pidNum = 0
        if ([int]::TryParse($pidText, [ref]$pidNum)) {
            if (-not $killed.ContainsKey($pidNum)) {
                Stop-Pid -pidNum $pidNum
                $killed[$pidNum] = $true
                Write-Host "[INFO] Stopped process on port $port (PID $pidNum)."
            }
        }
    }
}

Write-Host "[DONE] Dashboard stop routine finished."
