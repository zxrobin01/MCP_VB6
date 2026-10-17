@echo off
rem The registration verdict for the VB6MCP add-in - decided by READING, never by
rem regsvr32's exit code.  rc is not evidence here: on this host a bare
rem `regsvr32` (or a hardcoded C:\Windows\...) can resolve to the STALE second Windows
rem installation and fail with 0xC0000139 while the caller still sees a stale 0.
rem
rem TWO independent gates are read, and BOTH must be green.
rem
rem   (A) %WINDIR%\VBADDIN.INI  [Add-Ins32] <ProgID>   <-- DECISIVE
rem       This is the switch VB6 consults BEFORE it even tries CoCreateInstance.
rem         0   = DISABLED.  A single failed instantiation makes VB6 write 0 here and
rem               VB6 then NEVER RETRIES - not for the rest of this Windows install.
rem               Rebuilding, regsvr32 and `reg import` cannot clear it, and LoadBehavior
rem               is left untouched at 0x3, so every other check stays green while the
rem               add-in is dead.
rem           1 or 3 = ENABLED.  3 is the steady state: after a clean IDE exit on a
rem               successfully loaded add-in VB6 rewrites the value to 3 itself.
rem           MISSING entry = not an error: VB6 falls back to the registry and loads.
rem               (A later failure then CREATES the entry as 0 and the add-in dies.)
rem
rem   (B) the registry chain ProgID -> CLSID -> InprocServer32 + AddIns LoadBehavior.
rem       SECONDARY signal: a stale/mismatched registration is exactly what makes VB6
rem       write LoadBehavior=0 as a CONSEQUENCE.
rem
rem   usage: verify_registration.bat <path-to-the-DLL-that-must-be-registered>
rem   exit 0 = gate A enabled-or-absent AND the registry points at THAT file AND LoadBehavior=0x3
rem   exit 1 = gate A is 0, or not registered / pointed elsewhere / LoadBehavior != 0x3
setlocal EnableExtensions
set "VG_DLL=%~f1"
set "VG_PROGID=VB6MCP.Connect"
set "VG_ADDIN=HKCU\Software\Microsoft\Visual Basic\6.0\Addins\VB6MCP.Connect"
set "VG_OK=1"
set "VG_CLSID="
set "VG_INPROC="
set "VG_LB="
set "VG_INI="
set "VG_INISEC="
set "VG_INIVAL="
set "VG_INISTATE="

if "%VG_DLL%"=="" goto vg_nodll

rem --- (A) the enable gate VB6 itself consults ---------------------------------------
set "VG_INI=%SystemRoot%\VBADDIN.INI"
if not exist "%VG_INI%" goto vg_ini_absent
for /f "usebackq tokens=1,* delims==" %%A in ("%VG_INI%") do call :vg_ini_line "%%A" "%%B"
if not defined VG_INIVAL goto vg_ini_missing
if /i "%VG_INIVAL%"=="0x0" set "VG_INIVAL=0"
if /i "%VG_INIVAL%"=="0x1" set "VG_INIVAL=1"
if /i "%VG_INIVAL%"=="0x3" set "VG_INIVAL=3"
if "%VG_INIVAL%"=="0" goto vg_ini_disabled
set "VG_INISTATE=enabled"
goto vg_registry
:vg_ini_absent
set "VG_INISTATE=absent"
goto vg_registry
:vg_ini_missing
set "VG_INISTATE=missing_entry"
goto vg_registry
:vg_ini_disabled
set "VG_INISTATE=disabled"
set "VG_OK=0"
goto vg_registry

:vg_registry
rem --- ProgID -> CLSID (32-bit view; the 64-bit view has no such key) ---------------
for /f "tokens=2,*" %%A in ('reg query "HKCR\%VG_PROGID%\CLSID" /reg:32 2^>nul ^| findstr /i "REG_SZ"') do set "VG_CLSID=%%B"
if not defined VG_CLSID goto vg_report

rem --- that CLSID -> InprocServer32 (must be THIS file) -----------------------------
for /f "tokens=2,*" %%A in ('reg query "HKCR\CLSID\%VG_CLSID%\InprocServer32" /reg:32 2^>nul ^| findstr /i "REG_SZ"') do set "VG_INPROC=%%B"

rem --- the add-in entry must still auto-load ----------------------------------------
for /f "tokens=2,*" %%A in ('reg query "%VG_ADDIN%" /v LoadBehavior /reg:32 2^>nul ^| findstr /i "REG_DWORD"') do set "VG_LB=%%B"
if /i "%VG_LB%"=="0x3" goto vg_inproc
set "VG_OK=0"
goto vg_inproc

:vg_inproc
if "%VG_INPROC%"=="" set "VG_OK=0"
if "%VG_INPROC%"=="" goto vg_report
if /i "%VG_INPROC%"=="%VG_DLL%" goto vg_report
set "VG_OK=0"
goto vg_report

:vg_report
>"&2" echo REGVERIFY gate A (VB6 enable gate, DECISIVE): %VG_INI%
if /i "%VG_INISTATE%"=="enabled" goto vg_note_enabled
if /i "%VG_INISTATE%"=="disabled" goto vg_note_disabled
if /i "%VG_INISTATE%"=="absent" goto vg_note_absent
if /i "%VG_INISTATE%"=="missing_entry" goto vg_note_absent
goto vg_gate_b
:vg_note_enabled
>"&2" echo   %VG_PROGID% = %VG_INIVAL%   PASS - enabled (1 = enabled, 3 = the steady state VB6 writes itself)
goto vg_gate_b
:vg_note_disabled
>"&2" echo   %VG_PROGID% = %VG_INIVAL%   FAIL - this add-in is DISABLED at the enable gate.
>"&2" echo   VB6 will NOT load it: it never calls DllGetClassObject and it does NOT
>"&2" echo   touch LoadBehavior (which can still read 0x3, i.e. every other check is green).
>"&2" echo   A value of 0 is written after ONE failed instantiation and VB6 NEVER RETRIES:
>"&2" echo   no rebuild, no regsvr32 and no `reg import` clears it.
>"&2" echo   fix: set %VG_PROGID%=1 in the [Add-Ins32] section of that INI and start the
>"&2" echo        IDE once - on a successful load VB6 rewrites the value to 3 by itself.
goto vg_gate_b
:vg_note_absent
>"&2" echo   no [Add-Ins32] entry for %VG_PROGID% (%VG_INISTATE%): VB6 falls back to the
>"&2" echo   registry and can still load. This is NOT an error, but it is not protection
>"&2" echo   either - one failed instantiation creates the entry as 0 and the add-in then
>"&2" echo   stays dead until somebody sets it back to 1.
goto vg_gate_b
:vg_gate_b
>"&2" echo REGVERIFY gate B (registry chain, SECONDARY):
>"&2" echo   expected DLL       : %VG_DLL%
>"&2" echo   ProgID             : %VG_PROGID%
>"&2" echo   ProgID CLSID       : %VG_CLSID%
>"&2" echo   InprocServer32     : %VG_INPROC%
>"&2" echo   AddIns LoadBehavior: %VG_LB%   ^(must be 0x3^)
if "%VG_OK%"=="1" goto vg_pass
>"&2" echo REGVERIFY FAILED - see the gate above that printed FAIL or a mismatched path.
>"&2" echo   note : rc from regsvr32 is NOT a verdict; this check reads the INI and the registry.
>"&2" echo   fix  : "%SystemRoot%\SysWOW64\regsvr32.exe" /s "%VG_DLL%"
>"&2" echo          reg import "%~dp0..\register.reg" /reg:32
>"&2" echo          then start the IDE once and re-run this check (gate A must read 1 or 3).
endlocal & exit /b 1

:vg_pass
>"&2" echo REGVERIFY OK: gate A %VG_INISTATE% ^(%VG_PROGID%=%VG_INIVAL%^) ^; %VG_PROGID% -^> %VG_CLSID% ^; InprocServer32 = %VG_INPROC% ^; LoadBehavior = %VG_LB%
endlocal & exit /b 0

:vg_nodll
>"&2" echo REGVERIFY ERROR: no DLL path given.
>"&2" echo   usage: verify_registration.bat ^<path-to-the-DLL-that-must-be-registered^>
endlocal & exit /b 1

:vg_ini_line
rem %%1 = the text left of the first '=', %%2 = the rest of the line.
rem Only the [Add-Ins32] section counts: that is the section VB6 reads.
set "VG_L1=%~1"
set "VG_L2=%~2"
for /f "tokens=*" %%S in ("%VG_L1%") do set "VG_L1=%%S"
if /i "%VG_L1%"=="[Add-Ins32]" set "VG_INISEC=1"
if /i "%VG_L1%"=="[Add-Ins]" set "VG_INISEC="
if not defined VG_INISEC goto :eof
if /i not "%VG_L1%"=="%VG_PROGID%" goto :eof
set "VG_L2=%VG_L2: =%"
set "VG_INIVAL=%VG_L2%"
goto :eof
