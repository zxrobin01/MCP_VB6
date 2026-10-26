@echo off
rem ---------------------------------------------------------------------------
rem Build VB6MCP.dll.
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
rem the "freshly written" test is now the artifact's SHA-256 CHANGE
rem (certutil), with seconds-resolution mtime and size as auxiliary evidence, and
rem both shas are printed on BOTH the OK and the FAILED path. The old test compared
rem the DOS timestamp (MINUTE resolution) plus the file size; a same-minute rebuild
rem whose output kept its size was therefore reported as FAILED even though build.log
rem said the compile succeeded, rc was 0 and the on-disk sha HAD changed.
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

rem --- BINARY-COMPATIBILITY REFERENCE GUARD --------------------------
rem .vbp carries CompatibleMode="2" + CompatibleEXE32=<reference>. If that file is
rem MISSING or has been REPLACED, VB6 does not stop: it silently falls back to No
rem Compatibility and mints NEW class IDs, so every registration on every machine
rem goes stale - the exact failure this mode exists to prevent. Fail LOUDLY.
set "COMPAT_REF=%~dp0dist\compat\VB6MCP_ref.dll"
set "COMPAT_SHA=%~dp0dist\compat\VB6MCP_ref.sha256"
rem PUBLISHED (source-only) trees declare CompatibleMode="0" and ship NO binary
rem compatibility reference, so this guard must not fatal there. Read the mode from
rem the project file and only enforce the reference checks when it is 2.
set "COMPAT_MODE=0"
for /f "tokens=2 delims==" %%M in ('findstr /b /c:"CompatibleMode=" "%~dp0MCP_VB6.vbp"') do set "COMPAT_MODE=%%~M"
if not "%COMPAT_MODE%"=="2" goto compat_ref_skipped

:compat_ref_skipped
>&2 echo Binary-compatibility guard SKIPPED: MCP_VB6.vbp declares CompatibleMode="%COMPAT_MODE%"
>&2 echo        (a source-only tree mints its own class IDs on every build; no reference is shipped).
goto compat_ref_done

if not exist "%COMPAT_REF%" goto compat_ref_missing
if not exist "%COMPAT_SHA%" goto compat_ref_nosha
set "WANT_SHA="
set /p WANT_SHA=<"%COMPAT_SHA%"
set "GOT_SHA="
call :sha256 "%COMPAT_REF%" GOT_SHA
if /i not "%GOT_SHA%"=="%WANT_SHA%" goto compat_ref_swapped
>&2 echo Binary-compatibility reference OK: dist\compat\VB6MCP_ref.dll sha256=%GOT_SHA%
goto compat_ref_done
:compat_ref_missing
>&2 echo FATAL: dist\compat\VB6MCP_ref.dll is MISSING, but MCP_VB6.vbp says
>&2 echo        CompatibleMode="2". Building now would mint NEW class IDs and make
>&2 echo        every existing registration stale. Restore the reference (see
>&2 echo        dist\compat\README.md) or deliberately set CompatibleMode="0".
exit /b 1
:compat_ref_nosha
>&2 echo FATAL: dist\compat\VB6MCP_ref.sha256 is MISSING; the reference cannot be verified.
exit /b 1
:compat_ref_swapped
>&2 echo FATAL: the binary-compatibility reference was REPLACED.
>&2 echo        expected sha256=%WANT_SHA%
>&2 echo        actual   sha256=%GOT_SHA%
>&2 echo        A different reference re-defines the class IDs of this project.
>&2 echo        Do NOT build until that is deliberate (see dist\compat\README.md).
exit /b 1
:compat_ref_done

rem --- WHICH regsvr32 -------------------------------------------------
rem This build registers a 32-BIT COM server, so it must use the SysWOW64 copy of THIS
rem Windows installation. On a host whose SystemRoot is not C:, C:\Windows may be a
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
rem label-based instead of one parenthesised block - inside a block cmd.exe
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

rem --- the SOURCE STAMP must exist BEFORE the compiler runs ----------------
rem BuildStamp.bas carries I = sha256 over the .vbp's component inputs (strict mode,
rem docs/SOURCE_STAMP.md) and is EXCLUDED from that input set, so writing it cannot
rem change I. A build must never ship a stale or absent stamp, so a missing or failing
rem generator is FATAL here - never a skipped step (spec section 6: no stamp => FAIL).
>&2 echo Emitting the source stamp (tools\srcstamp.py emit) ...
if not exist "%~dp0tools\srcstamp.py" goto stamp_nogen
rem Exact CLI agreed with the generator author (B3/B4): emit writes BOTH
rem BuildStamp.bas and the .vbp Comments= line, so the two read points cannot drift.
rem Relative paths are safe here: line 30 already did cd /d "%~dp0".
python "tools\srcstamp.py" emit . --vbp "MCP_VB6.vbp" --out "BuildStamp.bas"
set "STAMPRC=%ERRORLEVEL%"
if not "%STAMPRC%"=="0" goto stamp_failed
if not exist "%~dp0BuildStamp.bas" goto stamp_nogen

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
>&2 echo minute-resolution timestamp + size test .
if "%A_TIME%%A_SIZE%"=="%B_TIME%%B_SIZE%" goto failed
set "FRESH=1"

:built
rem --- the artifact must carry the SAME I as this source tree ----------
rem check reads I back from the ARTIFACT (mcp_info when an IDE is running, otherwise
rem the version resource Comments=algo=..;src=..) and recomputes it from the tree; the
rem two read points must agree (spec section 7). A mismatch is FATAL, not a warning.
>&2 echo Checking the artifact's source stamp (tools\srcstamp.py check) ...
python "tools\srcstamp.py" check "%DLL%" "."
set "CHECKRC=%ERRORLEVEL%"
if not "%CHECKRC%"=="0" goto stamp_check_failed

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
rem rc is NOT the verdict. Read the registry and refuse to report success if
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

:stamp_nogen
>&2 echo.
>&2 echo Build FAILED - tools\srcstamp.py or BuildStamp.bas is missing, so the SOURCE
>&2 echo   STAMP could not be emitted. A build must never ship a stale/absent stamp.
>&2 echo   fix : restore tools\srcstamp.py, then run it once (see docs\SOURCE_STAMP.md).
>"%STATUS%" 2>nul echo FAILED-STAMP-NOGEN
endlocal & exit /b 1

:stamp_failed
>&2 echo.
>&2 echo Build FAILED - tools\srcstamp.py emit returned rc=%STAMPRC%.
>&2 echo   fix : python tools\srcstamp.py emit .   ^(read its diagnostics^)
>"%STATUS%" 2>nul echo FAILED-STAMP-EMIT rc=%STAMPRC%
endlocal & exit /b 1

:stamp_check_failed
>&2 echo.
>&2 echo Build FAILED - the artifact's SOURCE STAMP does not match this source tree.
>&2 echo   artifact : %DLL%
>&2 echo   sha256   : %A_SHA%
>&2 echo   reason   : tools\srcstamp.py check returned rc=%CHECKRC% ^(see its per-file diff^)
>&2 echo   note     : an artifact with a stale stamp must NOT be released (spec section 6).
>"%STATUS%" 2>nul echo FAILED-STAMP-CHECK rc=%CHECKRC% size=%A_SIZE% sha=%A_SHA%
endlocal & exit /b 1

:no_regsvr
>&2 echo.
>&2 echo Build FAILED - the 32-bit regsvr32 of THIS Windows installation was not found.
>&2 echo   expected : %REGSVR%
>&2 echo   SystemRoot: %SystemRoot%
>&2 echo   cause    : this build registers a 32-bit COM server; the SysWOW64 regsvr32 is
>&2 echo required and a bare "regsvr32" is deliberately NOT used .
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

rem --- helpers --------------------------------------------------
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
