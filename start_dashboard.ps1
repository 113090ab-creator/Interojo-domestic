$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectDir

$preferredPort = 8514
$fallbackPorts = @()
$pidFile = Join-Path $projectDir ".dashboard_pid"
$portFile = Join-Path $projectDir ".dashboard_port"
$outLog = Join-Path $projectDir "streamlit_8514_out.log"
$errLog = Join-Path $projectDir "streamlit_8514_err.log"

function Get-PythonExe {
    $venvPy = Join-Path $projectDir ".venv\Scripts\python.exe"
    if (Test-Path $venvPy) { return $venvPy }
    $py = Get-Command python -ErrorAction SilentlyContinue
    if ($py) { return $py.Source }
    throw "Python executable not found. Install Python or create .venv."
}

function Get-ListeningPid([int]$port) {
    $lines = cmd /c "netstat -ano -p tcp | findstr /r /c:"":$port .*LISTENING""" 2>$null
    foreach ($line in $lines) {
        $parts = ($line -split "\s+") | Where-Object { $_ -ne "" }
        if ($parts.Count -ge 5) {
            $parsedPid = 0
            if ([int]::TryParse($parts[$parts.Count - 1], [ref]$parsedPid)) {
                return $parsedPid
            }
        }
    }
    return $null
}

function Start-DashboardProcess([string]$pythonExe, [int]$port) {
    if (Test-Path $outLog) { Remove-Item -LiteralPath $outLog -Force }
    if (Test-Path $errLog) { Remove-Item -LiteralPath $errLog -Force }

    $args = @(
        "-m", "streamlit", "run", "dashboard.py",
        "--server.port", "$port",
        "--server.address", "localhost",
        "--server.headless", "true"
    )

    return Start-Process -FilePath $pythonExe -ArgumentList $args -WorkingDirectory $projectDir -WindowStyle Hidden -RedirectStandardOutput $outLog -RedirectStandardError $errLog -PassThru
}

function Wait-Ready([int]$port, [int]$timeoutSec) {
    $deadline = (Get-Date).AddSeconds($timeoutSec)
    while ((Get-Date) -lt $deadline) {
        if (Get-ListeningPid -port $port) { return $true }
        Start-Sleep -Milliseconds 700
    }
    return $false
}

function Resolve-TargetPort {
    $savedPid = 0
    $savedPort = 0

    if ((Test-Path $pidFile) -and (Test-Path $portFile)) {
        $pidRaw = (Get-Content -LiteralPath $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
        $portRaw = (Get-Content -LiteralPath $portFile -ErrorAction SilentlyContinue | Select-Object -First 1).Trim()
        if ([int]::TryParse($pidRaw, [ref]$savedPid) -and [int]::TryParse($portRaw, [ref]$savedPort)) {
            $listeningPid = Get-ListeningPid -port $savedPort
            if ($listeningPid -and $listeningPid -eq $savedPid) {
                return @{ mode = "reuse"; port = $savedPort; pid = $savedPid }
            }
        }
    }

    if (Test-Path $pidFile) { Remove-Item -LiteralPath $pidFile -Force -ErrorAction SilentlyContinue }
    if (Test-Path $portFile) { Remove-Item -LiteralPath $portFile -Force -ErrorAction SilentlyContinue }

    $candidates = @($preferredPort) + $fallbackPorts
    foreach ($port in $candidates) {
        if (-not (Get-ListeningPid -port $port)) {
            return @{ mode = "new"; port = $port }
        }
    }

    throw "No free dashboard port found. Checked: $($candidates -join ', ')"
}

try {
    $target = Resolve-TargetPort

    if ($target.mode -eq "reuse") {
        Start-Process "http://localhost:$($target.port)/?v=$([int](Get-Random))"
        Write-Host "[INFO] Dashboard already running on PID $($target.pid) (port $($target.port))."
        exit 0
    }

    $pythonExe = Get-PythonExe
    Write-Host "[INFO] Starting dashboard with: $pythonExe"
    $proc = Start-DashboardProcess -pythonExe $pythonExe -port $target.port

    if (Wait-Ready -port $target.port -timeoutSec 30) {
        Set-Content -LiteralPath $pidFile -Value "$($proc.Id)" -Encoding ascii
        Set-Content -LiteralPath $portFile -Value "$($target.port)" -Encoding ascii
        Start-Process "http://localhost:$($target.port)/?v=$([int](Get-Random))"
        Write-Host "[DONE] Dashboard is ready at http://localhost:$($target.port) (PID $($proc.Id))."
        exit 0
    }

    $alive = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue
    if (-not $alive) {
        Write-Host "[ERROR] Dashboard process exited before startup."
    } else {
        Write-Host "[ERROR] Dashboard did not become ready in time."
    }
    Write-Host "[HINT] Check logs:"
    Write-Host "  $outLog"
    Write-Host "  $errLog"
    exit 1
} catch {
    Write-Host "[ERROR] $($_.Exception.Message)"
    exit 1
}
