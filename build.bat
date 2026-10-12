@echo off
rem ---------------------------------------------------------------------------
rem Build VB6MCP.dll.  Card TASK-P1-012.
rem
rem This script used to report a FALSE SUCCESS, for two reasons:
rem   1. the verdict was only `if not exist VB6MCP.dll`, which passes while a
rem      STALE dll is still on disk. A running VB6 IDE auto-loads the add-in and
rem      therefore LOCKS VB6MCP.dll, so /make compiles and then cannot write the
rem      output (log: 'permission denied'); the stale file satisfied the check
rem      and the script still exited 0.
rem   2. every message went to stdout, so a caller doing `build.bat > build.log`
rem      wrote into the very file VB6's /out writes, destroying the compiler
rem      diagnostics.
rem
rem Now: the artifact must have been FRESHLY WRITTEN AND VB6 must have exited 0.
rem All diagnostics go to STDERR, so a stdout redirect can no longer clobber
rem build.log, and a machine-readable build.status is written on both paths.
rem
rem TASK-P1-072 - the "freshly written" test is now the artifact's SHA-256 CHANGE
rem (certutil), with seconds-resolution mtime and size as auxiliary evidence, and
rem both shas are printed on BOTH the OK and the FAILED path. The old test compared
rem the DOS timestamp (MINUTE resolution) plus the file size; a same-minute rebuild
rem whose output kept its size was therefore reported as FAILED even though build.log
rem said the compile succeeded, rc was 0 and the on-disk sha HAD changed (TASK-P1-069).
rem A rebuild that changes nothing is now NOCHANGE (exit 0), not FAILED: a stale DLL
rem that cannot be rewritten is still caught, because VB6 then logs 'permission
rem denied' and exits non-zero.
rem ---------------------------------------------------------------------------
setlocal EnableExtensions
cd /d "%~dp0"

set "VB6=C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist "%VB6%" set "VB6=C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
if not exist "%VB6%" (
    >&2 echo VB6.EXE not found. Edit build.bat and point VB6 at your VB98\VB6.EXE.
    exit /b 1
)

set "DLL=%~dp0VB6MCP.dll"
set "LOG=%~dp0build.log"
set "STATUS=%~dp0build.status"

rem --- TASK-P1-103: WHICH regsvr32 -------------------------------------------------
rem This build registers a 32-BIT COM server, so it must use the SysWOW64 copy of THIS
rem Windows installation. On this host SystemRoot is B:\Windows while C:\Windows is a
rem stale second installation, so a bare "regsvr32" resolves through PATH and can hit the
rem wrong copy, where every 32-bit registration fails with 0xC0000139 (and fails
rem SILENTLY: /s prints nothing and a PowerShell caller keeps a STALE $LASTEXITCODE of 0).
rem So: never a bare regsvr32, never a hardcoded C:\Windows path - and the verdict is
rem never rc, it is the registry (see tools\verify_registration.bat).
set "REGSVR=%SystemRoot%\SysWOW64\regsvr32.exe"
if not exist "%REGSVR%" goto no_regsvr

rem --- freshness baseline: the artifact as it is BEFORE the build -------------
set "B_TIME="
set "B_TIME_S="
set "B_SIZE=0"
set "B_SHA="
if exist "%DLL%" (
    for %%A in ("%DLL%") do (
        set "B_TIME=%%~tA"
        set "B_SIZE=%%~zA"
    )
    call :sha256 "%DLL%" B_SHA
    call :mtime_s "%DLL%" B_TIME_S
)

>&2 echo Checking Add-In Designer (MSADDNDR.DLL) registration ...
set "ADDNDR=%CommonProgramFiles(x86)%\designer\MSADDNDR.DLL"
rem TASK-P1-103: label-based instead of one parenthesised block - inside a block cmd.exe
rem expands %ERRORLEVEL% when the whole block is PARSED, so the rc read would be stale.
if not exist "%ADDNDR%" goto addndr_missing
"%REGSVR%" /s "%ADDNDR%"
set "ADDNDRC=%ERRORLEVEL%"
if "%ADDNDRC%"=="0" goto addndr_done
>&2 echo WARNING: registering MSADDNDR.DLL failed ^(regsvr32 rc=%ADDNDRC%^) - not fatal for
>&2 echo          this build, but the "Microsoft Add-In Designer" reference may be missing.
goto addndr_done
:addndr_missing
>&2 echo WARNING: MSADDNDR.DLL not found at "%ADDNDR%" - the project needs the
>&2 echo "Microsoft Add-In Designer" reference. Install it from your VB6/VS media.
:addndr_done

>&2 echo Building VB6MCP.dll with VB6 ...
if exist "%LOG%" del /q "%LOG%"
"%VB6%" /make "%~dp0MCP_VB6.vbp" /out "%LOG%"
set "VB6RC=%ERRORLEVEL%"

rem --- freshness verdict: exists, non-empty, and its bytes actually changed -----
set "A_TIME="
set "A_TIME_S="
set "A_SIZE=0"
set "A_SHA="
if exist "%DLL%" (
    for %%A in ("%DLL%") do (
        set "A_TIME=%%~tA"
        set "A_SIZE=%%~zA"
    )
    call :sha256 "%DLL%" A_SHA
    call :mtime_s "%DLL%" A_TIME_S
)

set "FRESH=0"
set "NOCHANGE=0"
if not exist "%DLL%" goto failed
if "%A_SIZE%"=="" goto failed
if "%A_SIZE%"=="0" goto failed
if not "%VB6RC%"=="0" goto failed
if "%A_SHA%"=="" goto legacy_verdict
if /i "%A_SHA%"=="%B_SHA%" (set "NOCHANGE=1") else (set "FRESH=1")
goto built

:legacy_verdict
rem Degraded environment (no usable certutil): keep the old minute+size test, but say so.
>&2 echo WARNING: certutil could not hash the artifact - falling back to the legacy
>&2 echo          minute-resolution timestamp + size test (TASK-P1-072).
if "%A_TIME%%A_SIZE%"=="%B_TIME%%B_SIZE%" goto failed
set "FRESH=1"

:built
rem --- register, then report refreshed-or-unchanged --------------------------
rem NOTE: this block is deliberately written with LABELS instead of one big
rem "if (...) else (...)" block: an unescaped ')' anywhere inside a parenthesised
rem block (e.g. in "mtime (seconds: ...)") silently closes the block early and
rem BOTH branches then execute with the second one overwriting build.status.
>&2 echo Registering COM server ...
"%REGSVR%" /s "%DLL%"
set "REGRC=%ERRORLEVEL%"
>&2 echo   regsvr32 rc=%REGRC% ^(informational only - the verdict is read from the registry^)
>&2 echo Registering VB6 add-in ...
reg import "%~dp0register.reg" >nul
rem TASK-P1-103: rc is NOT the verdict. Read the registry and refuse to report success if
rem the ProgID does not point at THIS file or the add-in is not set to auto-load.
>&2 echo Verifying registration by READING THE REGISTRY ...
call "%~dp0tools\verify_registration.bat" "%DLL%"
if errorlevel 1 goto regfailed
if "%FRESH%"=="1" goto refreshed

rem --- NOCHANGE: byte-identical output; exit 0, this is NOT a failure ---------
> "%STATUS%" echo NOCHANGE rc=%VB6RC% size=%A_SIZE% mtime=%A_TIME% mtimeS=%A_TIME_S% sha=%A_SHA% shaBefore=%B_SHA%
>&2 echo.
>&2 echo Build OK - NOCHANGE: the compiler produced byte-identical output.
>&2 echo   size        : %A_SIZE%
>&2 echo   sha256      : %B_SHA% == %A_SHA%   - identical, not a failure
>&2 echo   mtime       : %B_TIME% -^> %A_TIME%   seconds: %B_TIME_S% -^> %A_TIME_S%
>&2 echo   note        : a locked or stale DLL is still caught by the rc!=0 check
>&2 echo                 VB6 logs 'permission denied' and exits non-zero.
goto done

:refreshed
rem --- OK: the artifact's bytes changed ---------------------------------------
> "%STATUS%" echo OK rc=%VB6RC% size=%A_SIZE% mtime=%A_TIME% mtimeS=%A_TIME_S% sha=%A_SHA% shaBefore=%B_SHA%
>&2 echo.
>&2 echo Build OK - artifact refreshed.
>&2 echo   size        : %B_SIZE% -^> %A_SIZE%
>&2 echo   sha256      : %B_SHA%
>&2 echo               -^> %A_SHA%
>&2 echo   mtime       : %B_TIME% -^> %A_TIME%   seconds: %B_TIME_S% -^> %A_TIME_S%

:done
>&2 echo Done. Start the VB6 IDE; "VB6 MCP Server" should appear in
>&2 echo the Add-In Manager (Add-Ins menu). It is set to auto-load.
>&2 echo The MCP server listens on http://127.0.0.1:8766 (see README.md).
endlocal & exit /b 0

:no_regsvr
>&2 echo.
>&2 echo Build FAILED - the 32-bit regsvr32 of THIS Windows installation was not found.
>&2 echo   expected : %REGSVR%
>&2 echo   SystemRoot: %SystemRoot%
>&2 echo   cause    : this build registers a 32-bit COM server; the SysWOW64 regsvr32 is
>&2 echo              required and a bare "regsvr32" is deliberately NOT used (TASK-P1-103).
endlocal & exit /b 1

:regfailed
>&2 echo.
>&2 echo Build FAILED - the artifact was built but its REGISTRATION could not be verified.
>&2 echo   artifact      : %DLL%
>&2 echo   size / sha256 : %A_SIZE% / %A_SHA%
>&2 echo   cause         : the registry does not point at this DLL, or the add-in
>&2 echo                   LoadBehavior is not 3 ^(see the REGVERIFY lines above^).
>&2 echo   fix           : "%REGSVR%" /s "%DLL%"
>&2 echo                   reg import "%~dp0register.reg" /reg:32
>"%STATUS%" 2>nul echo FAILED-REG rc=%VB6RC% size=%A_SIZE% sha=%A_SHA%
endlocal & exit /b 1

:failed
>&2 echo.
>&2 echo Build FAILED - the artifact was NOT refreshed.
>&2 echo   vb6 exit code : %VB6RC%
>&2 echo   artifact      : %DLL%
>&2 echo   before        : time=%B_TIME% (%B_TIME_S%) size=%B_SIZE%
>&2 echo                   sha256=%B_SHA%
>&2 echo   after         : time=%A_TIME% (%A_TIME_S%) size=%A_SIZE%
>&2 echo                   sha256=%A_SHA%
>&2 echo   sha verdict   : %B_SHA% -^> %A_SHA%  (compare with re_work/p1_072 evidence)
>&2 echo   common cause  : a running VB6 IDE holds VB6MCP.dll (the add-in is
>&2 echo                   auto-loaded), so /make compiles but cannot write it.
>&2 echo                   Close every VB6 IDE and rebuild.
>&2 echo   compiler log  : %LOG%
> "%STATUS%" 2>nul echo FAILED rc=%VB6RC% before=%B_TIME%/%B_SIZE%/%B_SHA% after=%A_TIME%/%A_SIZE%/%A_SHA%
if exist "%LOG%" type "%LOG%" 1>&2
endlocal & exit /b 1

rem --- helpers (TASK-P1-072) --------------------------------------------------
:sha256
rem %1 = file, %2 = variable name that receives the hex digest.
rem certutil is present on every supported Windows; findstr drops the "SHA256 hash
rem of ..." and "CertUtil: ..." lines (both contain ':'), leaving the digest.
set "%~2="
if not exist "%~1" exit /b 0
set "MCPHASH="
for /f "delims=" %%H in ('certutil -hashfile "%~1" SHA256 2^>nul ^| findstr /v ":"') do (
    if not defined MCPHASH set "MCPHASH=%%H"
)
if defined MCPHASH set "%~2=%MCPHASH: =%"
exit /b 0

:mtime_s
rem %1 = file, %2 = variable name that receives the last-write time with SECONDS.
set "%~2="
if not exist "%~1" exit /b 0
for /f "delims=" %%T in ('powershell -NoProfile -NonInteractive -Command "(Get-Item -LiteralPath '%~1').LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')" 2^>nul') do set "%~2=%%T"
exit /b 0
