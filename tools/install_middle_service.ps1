# install_middle_service.ps1 - resident middle service for VB6MCP.
#
# WHY THIS FILE EXISTS
#   The middle service (bridge/vb6mcp_bridge.py) is the RESIDENT layer between the MCP hub
#   (bridge/vb6mcp_hub.mjs) and every VB6 add-in instance: it owns ONE TCP port, accepts the
#   add-ins' named pipes as they appear, and forwards calls.  It must be running BEFORE any
#   IDE and must stay up with NO IDE at all - otherwise the hub has nobody to talk to and the
#   whole add-in tool face goes silently dumb (the 2026-10-19 outage: hub reported port null).
#
#   Nothing in the repository installed it, so today it only lived while somebody started it
#   by hand.  This script is the missing installer.  It does NOT elevate itself: install and
#   uninstall need an elevated shell; status and smoke do not.
#
# DESIGN
#   install   : write a launcher .cmd (env vars + log redirection), register a scheduled task
#               at STARTUP running as SYSTEM with restart-on-failure, start it, then poll
#               /health until the service answers.
#   uninstall : stop + remove the task (and a -UseService/nssm service if present), stop the
#               middle-service process, KEEP the token file and the logs/audit trail.
#   status    : read-only report of task/service, processes, port listener and the raw
#               /health response (token from -TokenFile when present).  No admin needed.
#   smoke     : run the shipped --smoke acceptance on a temporary port in a child process;
#               no admin needed and it never touches the live port.
#
# USAGE
#   powershell -File tools\install_middle_service.ps1 -Action status
#   powershell -File tools\install_middle_service.ps1 -Action smoke
#   powershell -File tools\install_middle_service.ps1 -Action install -DryRun
#   powershell -File tools\install_middle_service.ps1 -Action install            (elevated)
#   powershell -File tools\install_middle_service.ps1 -Action uninstall          (elevated)
#
#   Defaults: -Port 8797, -ExePath dist\vb6mcp_bridge.exe (falls back to the source script
#   through python), -TokenFile <current user profile>\.vb6mcp\token.txt.
[CmdletBinding()]
param(
    [ValidateSet('install', 'uninstall', 'status', 'smoke')]
    [string]$Action = 'status',

    [int]$Port = 8797,

    [string]$ExePath = 'dist\vb6mcp_bridge.exe',

    # Empty means <current user profile>\.vb6mcp\token.txt (resolved at run time; this file
    # never hardcodes a machine-specific absolute path).
    [string]$TokenFile = '',

    [string]$TaskName = 'VB6MCP-MiddleService',

    # Optional explicit JSONL audit log for remote mode; empty = <log dir>\remote-audit.jsonl.
    [string]$AuditLog = '',

    # Use a real Windows service via nssm instead of a scheduled task (only if nssm is found).
    [switch]$UseService,

    # Print the exact commands and the privileges they need; execute nothing.
    [switch]$DryRun,

    [int]$HealthTimeoutSec = 30,
    [int]$SmokeSeconds = 5
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot

function Write-Info([string]$msg) { Write-Host $msg }

function Emit-Result($obj) {
    Write-Host ('RESULT_JSON ' + ($obj | ConvertTo-Json -Depth 10 -Compress))
}

function Test-Admin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $pr = New-Object Security.Principal.WindowsPrincipal($id)
        return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Resolve-TokenFile {
    if ($TokenFile -and $TokenFile.Trim().Length -gt 0) { return $TokenFile }
    if ($env:USERPROFILE) { return (Join-Path $env:USERPROFILE '.vb6mcp\token.txt') }
    return $null
}

function Get-LogDir {
    if ($env:ProgramData) { return (Join-Path $env:ProgramData 'vb6mcp\logs') }
    if ($env:TEMP) { return (Join-Path $env:TEMP 'vb6mcp-logs') }
    return (Join-Path $Root 'logs')
}

function Resolve-Bridge {
    $exeFull = $ExePath
    if (-not [IO.Path]::IsPathRooted($exeFull)) { $exeFull = Join-Path $Root $exeFull }
    $pyScript = Join-Path $Root 'bridge\vb6mcp_bridge.py'
    if (Test-Path -LiteralPath $exeFull) {
        return [ordered]@{
            kind = 'exe'; executable = $exeFull; prefix = @(); display = $exeFull;
            exeMissing = $false; exeTried = $exeFull; python = $null; pyScript = $pyScript;
            pyScriptExists = (Test-Path -LiteralPath $pyScript)
        }
    }
    $py = Get-Command python -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $py) { $py = Get-Command py -ErrorAction SilentlyContinue | Select-Object -First 1 }
    if (-not $py) {
        return [ordered]@{
            kind = 'none'; executable = $null; prefix = @(); display = '(nothing runnable)';
            exeMissing = $true; exeTried = $exeFull; python = $null; pyScript = $pyScript;
            pyScriptExists = (Test-Path -LiteralPath $pyScript)
        }
    }
    return [ordered]@{
        kind = 'python'; executable = $py.Source; prefix = @($pyScript);
        display = ($py.Source + ' ' + $pyScript);
        exeMissing = $true; exeTried = $exeFull; python = $py.Source; pyScript = $pyScript;
        pyScriptExists = (Test-Path -LiteralPath $pyScript)
    }
}

function Get-FreePort {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
    $listener.Start()
    $p = [int]$listener.LocalEndpoint.Port
    $listener.Stop()
    return $p
}

function Quote-Arg([string]$a) {
    if ($a -match '[\s"]') { return '"' + ($a -replace '"', '\"') + '"' }
    return $a
}

function Get-Health([int]$p, [string]$tf) {
    $url = 'http://127.0.0.1:' + $p + '/health'
    $headers = @{}
    $tok = ''
    if ($tf -and (Test-Path -LiteralPath $tf)) {
        $rawTok = Get-Content -LiteralPath $tf -Raw -ErrorAction SilentlyContinue
        if ($rawTok) { $tok = $rawTok.Trim() }
    }
    if ($tok) { $headers['Authorization'] = 'Bearer ' + $tok }
    $res = [ordered]@{
        url = $url; authUsed = [bool]$tok; httpStatus = $null; reachable = $false;
        bodyOk = $null; addinOnline = $null; raw = $null; error = $null
    }
    try {
        $r = Invoke-WebRequest -Uri $url -Headers $headers -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        $res.httpStatus = [int]$r.StatusCode
        $res.reachable = $true
        $res.raw = $r.Content
        try {
            $body = $r.Content | ConvertFrom-Json
            $res.bodyOk = $body.ok
            $res.addinOnline = $body.ok
        } catch { }
    } catch {
        $resp = $_.Exception.Response
        if ($resp) { try { $res.httpStatus = [int]$resp.StatusCode } catch { } }
        $res.error = $_.Exception.Message
        # A 401 still proves the HTTP layer is alive.
        if ($res.httpStatus -eq 401) { $res.reachable = $true }
    }
    return $res
}

function Get-TaskInfo([string]$name) {
    $info = [ordered]@{ name = $name; exists = $false; state = $null; lastRunTime = $null; lastTaskResult = $null; error = $null }
    try {
        $t = Get-ScheduledTask -TaskName $name -ErrorAction Stop
        $info.exists = $true
        $info.state = [string]$t.State
        try {
            $ti = Get-ScheduledTaskInfo -TaskName $name -ErrorAction Stop
            $info.lastRunTime = [string]$ti.LastRunTime
            $info.lastTaskResult = [int]$ti.LastTaskResult
        } catch { }
    } catch {
        $info.error = $_.Exception.Message
    }
    return $info
}

function Get-ServiceInfo([string]$name) {
    $info = [ordered]@{ name = $name; exists = $false; status = $null; startType = $null; error = $null }
    try {
        $s = Get-Service -Name $name -ErrorAction Stop
        $info.exists = $true
        $info.status = [string]$s.Status
        try { $info.startType = [string]((Get-CimInstance Win32_Service -Filter ("Name='" + $name + "'") -ErrorAction Stop).StartMode) } catch { }
    } catch {
        $info.error = $_.Exception.Message
    }
    return $info
}

function Get-Listener([int]$p) {
    $out = [ordered]@{ listening = $false; pid = $null; address = $null; error = $null }
    try {
        $c = Get-NetTCPConnection -LocalPort $p -State Listen -ErrorAction Stop | Select-Object -First 1
        if ($c) {
            $out.listening = $true
            $out.pid = [int]$c.OwningProcess
            $out.address = [string]$c.LocalAddress
        }
    } catch {
        $out.error = $_.Exception.Message
    }
    return $out
}

function Get-BridgeProcesses {
    $list = @()
    try {
        $all = Get-CimInstance Win32_Process -ErrorAction Stop | Where-Object {
            $_.CommandLine -and ($_.CommandLine -match 'vb6mcp_bridge') -and
            ($_.Name -match '^(python|pythonw|py|vb6mcp_bridge)([.]exe)?$')
        }
        foreach ($p in $all) {
            $list += [ordered]@{ pid = [int]$p.ProcessId; name = [string]$p.Name; commandLine = [string]$p.CommandLine }
        }
    } catch { }
    return $list
}

function Format-BridgeDisplay($bridge) {
    if ($bridge.kind -eq 'none') { return '(nothing runnable: no dist exe and no python)' }
    return $bridge.display
}

# ------------------------------------------------------------------ status --------
function Invoke-Status {
    $bridge = Resolve-Bridge
    $tf = Resolve-TokenFile
    $task = Get-TaskInfo $TaskName
    $svc = $null
    if ($UseService) { $svc = Get-ServiceInfo $TaskName }
    $listener = Get-Listener $Port
    $procs = @(Get-BridgeProcesses)
    $health = Get-Health $Port $tf
    $tokenExists = [bool]($tf -and (Test-Path -LiteralPath $tf))
    $tokenLen = 0
    if ($tokenExists) {
        $rawTok = Get-Content -LiteralPath $tf -Raw -ErrorAction SilentlyContinue
        if ($rawTok) { $tokenLen = $rawTok.Trim().Length }
    }

    $anyRuntime = ($listener.listening -or $procs.Count -gt 0 -or $task.exists -or ($svc -and $svc.exists))
    $verdict = 'down'
    if ($listener.listening) {
        if ($health.reachable -and $health.httpStatus -eq 200) { $verdict = 'up' } else { $verdict = 'up-health-unverified' }
    } elseif ($task.exists -or ($svc -and $svc.exists) -or $procs.Count -gt 0) {
        $verdict = 'degraded'
    }

    $result = [ordered]@{
        action = 'status'; checkedAt = (Get-Date).ToString('o'); port = $Port;
        task = $task;
        service = $svc;
        bridge = [ordered]@{ kind = $bridge.kind; executable = $bridge.executable; prefix = $bridge.prefix;
                             display = (Format-BridgeDisplay $bridge); exeTried = $bridge.exeTried;
                             python = $bridge.python; pyScript = $bridge.pyScript };
        process = [ordered]@{ found = ($procs.Count -gt 0); count = $procs.Count; pids = @($procs | ForEach-Object { $_.pid }); items = @($procs) };
        listener = $listener;
        health = $health;
        token = [ordered]@{ path = $tf; exists = $tokenExists; length = $tokenLen };
        verdict = $verdict;
        serviceHealthy = [bool]($listener.listening -and $health.reachable -and ($health.httpStatus -eq 200));
        resident = [bool]$anyRuntime;
        addinOnline = $health.addinOnline;
        ok = [bool]($verdict -eq 'up')
    }

    Write-Info ('STATUS task      : ' + $TaskName + ' exists=' + $task.exists + ' state=' + $task.state)
    if ($svc) { Write-Info ('STATUS service   : ' + $TaskName + ' exists=' + $svc.exists + ' status=' + $svc.status) }
    Write-Info ('STATUS process   : ' + $procs.Count + ' pid(s) ' + (@($procs | ForEach-Object { $_.pid }) -join ','))
    Write-Info ('STATUS listener  : ' + $listener.address + ':' + $Port + ' listening=' + $listener.listening + ' pid=' + $listener.pid)
    Write-Info ('STATUS health    : http=' + $health.httpStatus + ' reachable=' + $health.reachable + ' bodyOk=' + $health.bodyOk + ' addinOnline=' + $health.addinOnline)
    Write-Info ('STATUS token     : ' + $tf + ' exists=' + $tokenExists + ' length=' + $tokenLen)
    Write-Info ('STATUS bridge    : ' + (Format-BridgeDisplay $bridge))
    if ($verdict -eq 'up' -and $health.addinOnline -eq $false) {
        Write-Info 'STATUS note      : the middle service is UP; no add-in is online yet (body.ok reflects a live IDE, not the service)'
    }
    Write-Info ('STATUS verdict   : ' + $verdict.ToUpper())
    Emit-Result $result
    if ($result.ok) { return 0 }
    if ($verdict -eq 'degraded') { return 1 }
    return 2
}

# ----------------------------------------------------------------- install --------
function Get-InstallPlan {
    $bridge = Resolve-Bridge
    $tf = Resolve-TokenFile
    $logDir = Get-LogDir
    $logFile = Join-Path $logDir 'middle-service.log'
    $audit = $AuditLog
    if (-not $audit) { $audit = Join-Path $logDir 'remote-audit.jsonl' }
    $launcher = Join-Path $logDir 'run-middle-service.cmd'

    $bridgeArgs = @()
    if ($bridge.kind -eq 'python') { $bridgeArgs = @($bridge.prefix) }
    $bridgeArgs += @('--port', [string]$Port)
    $inner = (@((Quote-Arg $bridge.executable)) + @($bridgeArgs | ForEach-Object { Quote-Arg $_ })) -join ' '

    $launcherLines = @(
        '@echo off',
        'rem AUTO-GENERATED by tools\install_middle_service.ps1 - do not edit by hand.',
        ('set "VB6MCP_TOKEN_FILE=' + $tf + '"'),
        ('set "VB6MCP_BRIDGE_LOG=' + $logFile + '"'),
        ('rem audit JSONL (only used when remote mode is configured): ' + $audit),
        ($inner + ' >> "' + $logFile + '" 2>&1')
    )

    $commands = @(
        ('New-Item -ItemType Directory -Force -Path "' + $logDir + '"'),
        ('write launcher: "' + $launcher + '" (' + $launcherLines.Count + ' lines)'),
        ('Register-ScheduledTask -TaskName "' + $TaskName + '" -Action (cmd.exe /c "' + $launcher + '") ' +
         '-Trigger (AtStartup) -Principal (SYSTEM / Highest) -Settings (RestartCount 999, RestartInterval 1m) -Force'),
        ('Start-ScheduledTask -TaskName "' + $TaskName + '"'),
        ('Invoke-WebRequest http://127.0.0.1:' + $Port + '/health  (poll up to ' + $HealthTimeoutSec + ' s)')
    )
    if ($UseService) {
        $commands = @(
            ('New-Item -ItemType Directory -Force -Path "' + $logDir + '"'),
            ('write launcher: "' + $launcher + '" (' + $launcherLines.Count + ' lines)'),
            ('nssm install "' + $TaskName + '" ' + $inner),
            ('nssm set "' + $TaskName + '" AppDirectory "' + $Root + '"'),
            ('nssm set "' + $TaskName + '" AppEnvironmentExtra "VB6MCP_TOKEN_FILE=' + $tf + '" "VB6MCP_BRIDGE_LOG=' + $logFile + '"'),
            ('nssm set "' + $TaskName + '" AppExit Default Restart'),
            ('nssm set "' + $TaskName + '" AppStdout "' + $logFile + '"'),
            ('nssm set "' + $TaskName + '" AppStderr "' + $logFile + '"'),
            ('nssm start "' + $TaskName + '"'),
            ('Invoke-WebRequest http://127.0.0.1:' + $Port + '/health  (poll up to ' + $HealthTimeoutSec + ' s)')
        )
    }

    return [ordered]@{
        action = 'install'; dryRun = [bool]$DryRun; elevated = (Test-Admin);
        taskName = $TaskName; port = $Port; useService = [bool]$UseService;
        logDir = $logDir; logFile = $logFile; launcher = $launcher; auditLog = $audit;
        tokenFile = $tf; tokenExists = [bool]($tf -and (Test-Path -LiteralPath $tf)); tokenLength = 0;
        bridge = [ordered]@{ kind = $bridge.kind; executable = $bridge.executable; prefix = $bridge.prefix;
                             display = (Format-BridgeDisplay $bridge); exeMissing = $bridge.exeMissing;
                             exeTried = $bridge.exeTried; python = $bridge.python; pyScript = $bridge.pyScript };
        launcherLines = $launcherLines;
        requiredPrivileges = @('Administrator (elevated) shell: create a scheduled task that runs as SYSTEM / start a service',
                               ('write access to ' + $logDir),
                               'read access to the token file (SYSTEM can read it)');
        commands = $commands
    }
}

function Invoke-Install {
    $plan = Get-InstallPlan
    $tf = $plan.tokenFile
    if ($tf -and (Test-Path -LiteralPath $tf)) {
        $rawTok = Get-Content -LiteralPath $tf -Raw -ErrorAction SilentlyContinue
        if ($rawTok) { $plan.tokenLength = $rawTok.Trim().Length }
    }

    Write-Info '=== INSTALL PLAN (resident middle service) ==='
    Write-Info ('action=install task=' + $plan.taskName + ' port=' + $plan.port + ' useService=' + $plan.useService)
    Write-Info ('bridge=' + (Format-BridgeDisplay $plan.bridge) + ' (exe missing=' + $plan.bridge.exeMissing + '; tried=' + $plan.bridge.exeTried + ')')
    Write-Info ('tokenFile=' + $plan.tokenFile + ' exists=' + $plan.tokenExists + ' length=' + $plan.tokenLength)
    Write-Info ('logDir=' + $plan.logDir + ' logFile=' + $plan.logFile + ' auditLog=' + $plan.auditLog)
    Write-Info '--- launcher ---'
    foreach ($line in $plan.launcherLines) { Write-Info ('  ' + $line) }
    Write-Info '--- required privileges ---'
    foreach ($p in $plan.requiredPrivileges) { Write-Info ('  * ' + $p) }
    Write-Info '--- commands ---'
    foreach ($c in $plan.commands) { Write-Info ('  $ ' + $c) }

    if ($plan.bridge.kind -eq 'none') {
        $plan.ok = $false
        $plan.reason = 'not_runnable'
        $plan.message = 'Neither the dist EXE nor python is available, so there is nothing to install.'
        Emit-Result $plan
        return 4
    }
    if ($DryRun) {
        $plan.ok = $true
        $plan.reason = 'dry_run'
        $plan.message = 'Nothing was changed.  Re-run without -DryRun from an elevated shell to install.'
        Emit-Result $plan
        return 0
    }
    if (-not $plan.elevated) {
        $plan.ok = $false
        $plan.reason = 'needs_elevation'
        $plan.message = 'install needs an elevated PowerShell; this script never tries to elevate itself.  Nothing was changed.'
        Emit-Result $plan
        return 3
    }

    # ---- real install (elevated) ----
    New-Item -ItemType Directory -Force -Path $plan.logDir | Out-Null
    Set-Content -LiteralPath $plan.launcher -Value $plan.launcherLines -Encoding ASCII
    if ($UseService) {
        $nssm = Get-Command nssm -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $nssm) {
            $plan.ok = $false
            $plan.reason = 'nssm_not_found'
            $plan.message = '-UseService needs nssm on PATH; nothing was changed.'
            Emit-Result $plan
            return 5
        }
        $svcArgs = @()
        if ($bridge.kind -eq 'python') { $svcArgs += @($bridge.prefix) }
        $svcArgs += @('--port', [string]$Port)
        & $nssm.Source install $TaskName $bridge.executable (($svcArgs | ForEach-Object { Quote-Arg $_ }) -join ' ') | Out-Null
        & $nssm.Source set $TaskName AppDirectory $Root | Out-Null
        & $nssm.Source set $TaskName AppEnvironmentExtra ('VB6MCP_TOKEN_FILE=' + $tf) ('VB6MCP_BRIDGE_LOG=' + $plan.logFile) | Out-Null
        & $nssm.Source set $TaskName AppExit Default Restart | Out-Null
        & $nssm.Source set $TaskName AppStdout $plan.logFile | Out-Null
        & $nssm.Source set $TaskName AppStderr $plan.logFile | Out-Null
        & $nssm.Source start $TaskName | Out-Null
    } else {
        $action = New-ScheduledTaskAction -Execute (Join-Path $env:SystemRoot 'System32\cmd.exe') -Argument ('/c "' + $plan.launcher + '"') -WorkingDirectory $Root
        $trigger = New-ScheduledTaskTrigger -AtStartup
        $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Days 0) -MultipleInstances IgnoreNew -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
        Start-ScheduledTask -TaskName $TaskName
    }

    $deadline = (Get-Date).AddSeconds($HealthTimeoutSec)
    $health = $null
    while ((Get-Date) -lt $deadline) {
        $health = Get-Health $Port $tf
        if ($health.reachable -and $health.httpStatus -eq 200) { break }
        Start-Sleep -Milliseconds 500
    }
    $plan.ok = [bool]($health -and $health.reachable -and $health.httpStatus -eq 200)
    $plan.health = $health
    $plan.installedAt = (Get-Date).ToString('o')
    $plan.message = if ($plan.ok) { 'middle service installed and /health answered 200.' } else { 'task created, but /health did not answer 200 within the timeout; inspect the log file.' }
    Emit-Result $plan
    if ($plan.ok) { return 0 }
    return 6
}

# --------------------------------------------------------------- uninstall --------
function Get-UninstallPlan {
    $bridge = Resolve-Bridge
    $tf = Resolve-TokenFile
    $logDir = Get-LogDir
    $audit = $AuditLog
    if (-not $audit) { $audit = Join-Path $logDir 'remote-audit.jsonl' }
    $commands = @(
        ('Stop-ScheduledTask -TaskName "' + $TaskName + '"'),
        ('Unregister-ScheduledTask -TaskName "' + $TaskName + '" -Confirm:$false'),
        ('stop any process whose command line matches vb6mcp_bridge'),
        ('KEEP "' + $tf + '" (token)'),
        ('KEEP "' + $logDir + '" (service log + audit JSONL)')
    )
    if ($UseService) {
        $commands = @(
            ('nssm stop "' + $TaskName + '"'),
            ('nssm remove "' + $TaskName + '" confirm'),
            ('stop any leftover vb6mcp_bridge process'),
            ('KEEP "' + $tf + '" (token)'),
            ('KEEP "' + $logDir + '" (service log + audit JSONL)')
        )
    }
    return [ordered]@{
        action = 'uninstall'; dryRun = [bool]$DryRun; elevated = (Test-Admin);
        taskName = $TaskName; port = $Port; useService = [bool]$UseService;
        logDir = $logDir; tokenFile = $tf; auditLog = $audit;
        bridge = [ordered]@{ kind = $bridge.kind; display = (Format-BridgeDisplay $bridge) };
        requiredPrivileges = @('Administrator (elevated) shell: stop/remove a scheduled task or service');
        commands = $commands; preserved = @($tf, $logDir, $audit);
        processes = @(Get-BridgeProcesses); listener = Get-Listener $Port
    }
}

function Invoke-Uninstall {
    $plan = Get-UninstallPlan
    Write-Info '=== UNINSTALL PLAN (resident middle service) ==='
    Write-Info ('action=uninstall task=' + $plan.taskName + ' port=' + $plan.port + ' useService=' + $plan.useService)
    Write-Info '--- required privileges ---'
    foreach ($p in $plan.requiredPrivileges) { Write-Info ('  * ' + $p) }
    Write-Info '--- commands ---'
    foreach ($c in $plan.commands) { Write-Info ('  $ ' + $c) }
    Write-Info ('--- preserved --- token=' + $plan.tokenFile + ' logs=' + $plan.logDir + ' audit=' + $plan.auditLog)
    if ($plan.processes.Count -gt 0) {
        Write-Info ('running middle-service process(es): ' + (@($plan.processes | ForEach-Object { $_.pid }) -join ','))
    }

    if ($DryRun) {
        $plan.ok = $true; $plan.reason = 'dry_run'
        $plan.message = 'Nothing was changed.  Re-run without -DryRun from an elevated shell to uninstall.'
        Emit-Result $plan
        return 0
    }
    if (-not $plan.elevated) {
        $plan.ok = $false; $plan.reason = 'needs_elevation'
        $plan.message = 'uninstall needs an elevated PowerShell; this script never tries to elevate itself.  Nothing was changed.'
        Emit-Result $plan
        return 3
    }

    $stopped = @()
    try { Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue } catch { }
    if ($UseService) {
        $nssm = Get-Command nssm -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($nssm) {
            try { & $nssm.Source stop $TaskName | Out-Null } catch { }
            try { & $nssm.Source remove $TaskName confirm | Out-Null } catch { }
        }
    } else {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    }
    $lp = Get-Listener $Port
    $candidates = @()
    foreach ($p in (Get-BridgeProcesses)) {
        if ($lp.listening -and ($lp.pid -eq $p.pid)) { $candidates += $p; continue }
        if ($p.commandLine -match ('--port\s+' + [regex]::Escape([string]$Port) + '(\s|$)')) { $candidates += $p }
    }
    foreach ($p in $candidates) {
        try { Stop-Process -Id $p.pid -Force -ErrorAction Stop; $stopped += $p.pid } catch { }
    }
    $plan.stoppedPids = $stopped
    $plan.ok = $true
    $plan.message = 'task/service removed; token and logs were kept.'
    Emit-Result $plan
    return 0
}

# ------------------------------------------------------------------- smoke --------
function Invoke-Smoke {
    $bridge = Resolve-Bridge
    if ($bridge.kind -eq 'none') {
        Emit-Result ([ordered]@{ action = 'smoke'; ok = $false; reason = 'not_runnable'; bridge = $bridge })
        return 4
    }
    $p = Get-FreePort
    $argv = @()
    if ($bridge.kind -eq 'python') { $argv += $bridge.prefix }
    $argv += @('--smoke', '--port', [string]$p, '--smoke-seconds', [string]$SmokeSeconds)

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $bridge.executable
    $psi.Arguments = (@($argv | ForEach-Object { Quote-Arg $_ }) -join ' ')
    $psi.WorkingDirectory = $Root
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $psi.EnvironmentVariables['VB6MCP_TOKEN'] = ''
    $psi.EnvironmentVariables['VB6MCP_TOKEN_FILE'] = ''

    $started = Get-Date
    Write-Info ('SMOKE start: ' + $bridge.display + ' --smoke --port ' + $p + ' --smoke-seconds ' + $SmokeSeconds)
    $proc = [System.Diagnostics.Process]::Start($psi)
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    $errTask = $proc.StandardError.ReadToEndAsync()

    # While the smoke server runs, read /health from the same temporary port (no token).
    $health = $null
    $deadline = (Get-Date).AddSeconds(12)
    while ((Get-Date) -lt $deadline) {
        $health = Get-Health $p $null
        if ($health.reachable) { break }
        Start-Sleep -Milliseconds 250
    }

    $timeoutMs = [int](($SmokeSeconds + 45) * 1000)
    $exited = $proc.WaitForExit($timeoutMs)
    if (-not $exited) { try { $proc.Kill() } catch { } }
    $stdout = ''
    $stderr = ''
    try { $stdout = $outTask.Result } catch { }
    try { $stderr = $errTask.Result } catch { }
    $elapsed = [int](((Get-Date) - $started).TotalMilliseconds)

    $parsed = @()
    foreach ($line in ($stdout -split "[\r\n]+")) {
        if (-not $line.Trim()) { continue }
        try { $parsed += ($line | ConvertFrom-Json) } catch { }
    }
    $ready = $parsed | Where-Object { $_.smoke -eq 'ready' } | Select-Object -Last 1
    $roundtrip = $parsed | Where-Object { $_.smoke -eq 'pipe_roundtrip' } | Select-Object -Last 1
    $fatal = $parsed | Where-Object { $_.smoke -eq 'fatal' } | Select-Object -Last 1
    if (-not $fatal) { $fatal = $null }
    $stopped = $parsed | Where-Object { $_.smoke -eq 'stopped' } | Select-Object -Last 1
    $exitCode = $null
    if ($exited) { $exitCode = [int]$proc.ExitCode }
    $faceMatches = $false
    if ($roundtrip -and $roundtrip.faceMatchesSnapshot) { $faceMatches = $true }
    $ok = ($exitCode -eq 0) -and $faceMatches -and ($ready -ne $null) -and ($fatal -eq $null)

    $result = [ordered]@{
        action = 'smoke'; ok = [bool]$ok; exitCode = $exitCode; timedOut = (-not $exited);
        port = $p; elapsedMs = $elapsed; bridge = [ordered]@{ kind = $bridge.kind; display = (Format-BridgeDisplay $bridge) };
        ready = $ready; pipeRoundtrip = $roundtrip; fatal = $fatal; stopped = $stopped;
        health = $health; jsonLineCount = $parsed.Count; stderrTail = ($stderr -split "[\r\n]+" | Select-Object -Last 8)
    }
    $readyPort = $null
    if ($ready) { $readyPort = $ready.port }
    Write-Info ('SMOKE exit=' + $exitCode + ' readyPort=' + $readyPort + ' faceMatchesSnapshot=' + $faceMatches)
    Write-Info ('SMOKE health: http=' + $health.httpStatus + ' bodyOk=' + $health.bodyOk + ' url=' + $health.url)
    Emit-Result $result
    if ($ok) { return 0 }
    return 7
}

# -------------------------------------------------------------------- main --------
Write-Info ('install_middle_service.ps1: action=' + $Action + ' port=' + $Port + ' elevated=' + (Test-Admin) + ' dryRun=' + [bool]$DryRun)
switch ($Action) {
    'status'    { $rc = Invoke-Status; exit $rc }
    'install'   { $rc = Invoke-Install; exit $rc }
    'uninstall' { $rc = Invoke-Uninstall; exit $rc }
    'smoke'     { $rc = Invoke-Smoke; exit $rc }
    default     { exit 64 }
}
