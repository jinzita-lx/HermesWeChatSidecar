$ErrorActionPreference = 'Continue'

$sshPath = 'C:\WINDOWS\System32\OpenSSH\ssh.exe'
$logDir  = 'C:\HermesWeChatSidecar\logs'
$logPath = Join-Path $logDir 'tunnel-watchdog.log'

if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

function Write-TunnelLog($msg) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $logPath -Value "$ts $msg" -Encoding utf8
}

$sshArgs = @(
    '-N',
    '-L', '8787:127.0.0.1:8787',
    '-o', 'ServerAliveInterval=30',
    '-o', 'ServerAliveCountMax=3',
    '-o', 'ExitOnForwardFailure=yes',
    '-o', 'StrictHostKeyChecking=accept-new',
    'xin'
)

Write-TunnelLog "watchdog started; pid=$PID"

$consecutiveFastFails = 0

while ($true) {
    $startedAt = Get-Date
    Write-TunnelLog "launching: ssh $($sshArgs -join ' ')"
    try {
        $p = Start-Process -FilePath $sshPath -ArgumentList $sshArgs -NoNewWindow -PassThru -Wait
        $code = $p.ExitCode
    } catch {
        Write-TunnelLog "Start-Process threw: $_"
        $code = -1
    }
    $duration = (Get-Date) - $startedAt
    Write-TunnelLog ("ssh exited code={0} after {1:N1}s" -f $code, $duration.TotalSeconds)

    if ($duration.TotalSeconds -lt 10) {
        $consecutiveFastFails++
    } else {
        $consecutiveFastFails = 0
    }

    $sleep = [Math]::Min(5 + ($consecutiveFastFails * 5), 60)
    Write-TunnelLog "sleeping ${sleep}s before reconnect"
    Start-Sleep -Seconds $sleep
}
