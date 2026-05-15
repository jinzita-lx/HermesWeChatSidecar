[CmdletBinding()]
param(
    [Parameter(Position=0)]
    [ValidateSet('start','stop','restart','status','logs','help')]
    [string]$Command = 'status',

    [int]$Tail = 30,
    [switch]$Follow
)

$ErrorActionPreference = 'Stop'
$Root   = 'C:\HermesWeChatSidecar'
$Py     = Join-Path $Root '.venv\Scripts\pythonw.exe'
$LogDir = Join-Path $Root 'logs'

function Get-SidecarProcs {
    Get-CimInstance Win32_Process -Filter "Name='pythonw.exe' OR Name='python.exe'" |
        Where-Object { $_.CommandLine -match 'src\.main' }
}

function Get-LatestLogFile {
    if (-not (Test-Path $LogDir)) { return $null }
    Get-ChildItem -Path $LogDir -Filter '*.log' -Recurse -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

function Show-Help {
    @'
Usage: hs <command> [options]

Commands:
  start                       Start the sidecar in the background (no window)
  stop                        Stop all sidecar processes
  restart                     Stop then start
  status                      Show running state, pid, uptime, recent log
  logs [-Tail N] [-Follow]    Show last N lines of latest log (default 30); -Follow tails live
  help                        Show this help
'@ | Write-Host
}

function Cmd-Start {
    $existing = Get-SidecarProcs
    if ($existing) {
        Write-Host "* hs - already running (pid $($existing[0].ProcessId))"
        return
    }
    if (-not (Test-Path $Py)) {
        Write-Host "x pythonw.exe not found at $Py" -ForegroundColor Red
        return
    }
    Start-Process -FilePath $Py `
        -ArgumentList @('-m','src.main') `
        -WorkingDirectory $Root `
        -WindowStyle Hidden | Out-Null
    Start-Sleep -Milliseconds 800
    $procs = Get-SidecarProcs
    if ($procs) {
        Write-Host "* hs - started (pid $($procs[0].ProcessId))"
    } else {
        Write-Host "x hs - failed to start; check logs (hs logs)" -ForegroundColor Red
    }
}

function Cmd-Stop {
    $procs = Get-SidecarProcs
    if (-not $procs) {
        Write-Host "- hs - not running"
        return
    }
    foreach ($p in $procs) {
        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
    }
    Write-Host "* hs - stopped ($($procs.Count) process(es))"
}

function Cmd-Restart {
    Cmd-Stop
    Start-Sleep -Seconds 1
    Cmd-Start
}

function Cmd-Status {
    $procs = Get-SidecarProcs
    if (-not $procs) {
        Write-Host "* hs - inactive (dead)"
        $logFile = Get-LatestLogFile
        if ($logFile) { Write-Host "  Last log:  $($logFile.FullName)" }
        return
    }

    $main = $procs | Where-Object { $_.CommandLine -match '\.venv' } | Select-Object -First 1
    if (-not $main) { $main = $procs | Select-Object -First 1 }
    $proc = Get-Process -Id $main.ProcessId -ErrorAction SilentlyContinue
    $started = if ($proc) { $proc.StartTime } else { $null }
    $uptime  = if ($started) { (Get-Date) - $started } else { $null }
    $logFile = Get-LatestLogFile

    Write-Host "* hs - active (running)"
    Write-Host "  Main PID:  $($main.ProcessId)"
    if ($started) { Write-Host ("  Started:   {0}" -f $started) }
    if ($uptime)  { Write-Host ("  Uptime:    {0:dd\.hh\:mm\:ss}" -f $uptime) }
    Write-Host "  Procs:     $($procs.Count) (incl. venv launcher child)"
    if ($logFile) {
        Write-Host "  Log:       $($logFile.FullName)"
        Write-Host ''
        Write-Host '  Recent log:'
        Get-Content $logFile.FullName -Tail 8 -ErrorAction SilentlyContinue |
            ForEach-Object { "  $_" }
    } else {
        Write-Host '  Log:       (none yet)'
    }
}

function Cmd-Logs {
    $logFile = Get-LatestLogFile
    if (-not $logFile) {
        Write-Host "no log files in $LogDir"
        return
    }
    Write-Host "=== $($logFile.FullName) ==="
    if ($Follow) {
        Get-Content $logFile.FullName -Tail $Tail -Wait
    } else {
        Get-Content $logFile.FullName -Tail $Tail
    }
}

switch ($Command) {
    'start'   { Cmd-Start }
    'stop'    { Cmd-Stop }
    'restart' { Cmd-Restart }
    'status'  { Cmd-Status }
    'logs'    { Cmd-Logs }
    'help'    { Show-Help }
    default   { Show-Help }
}
