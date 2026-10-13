@echo off
rem The registration verdict for the VB6MCP add-in - decided by READING THE
rem REGISTRY, never by regsvr32's exit code.  rc is not evidence here: on this host a bare
rem `regsvr32` (or a hardcoded C:\Windows\...) can resolve to the STALE second Windows
rem installation and fail with 0xC0000139 while the caller still sees a stale 0.
rem
rem   usage: verify_registration.bat <path-to-the-DLL-that-must-be-registered>
rem   exit 0 = the ProgID points at THAT file AND the add-in is set to auto-load (0x3)
rem   exit 1 = not registered / pointed elsewhere / LoadBehavior != 0x3 (stderr has detail)
setlocal EnableExtensions
set "VG_DLL=%~f1"
set "VG_PROGID=VB6MCP.Connect"
set "VG_ADDIN=HKCU\Software\Microsoft\Visual Basic\6.0\Addins\VB6MCP.Connect"
set "VG_OK=1"
set "VG_CLSID="
set "VG_INPROC="
set "VG_LB="

if "%VG_DLL%"=="" goto vg_nodll

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
if "%VG_OK%"=="1" goto vg_pass
>&2 echo REGVERIFY FAILED - the registry does NOT point at the DLL that was just registered.
>&2 echo   expected DLL       : %VG_DLL%
>&2 echo   ProgID             : %VG_PROGID%
>&2 echo   ProgID CLSID       : %VG_CLSID%
>&2 echo   InprocServer32     : %VG_INPROC%
>&2 echo   AddIns LoadBehavior: %VG_LB%   ^(must be 0x3^)
>&2 echo   note               : rc from regsvr32 is NOT a verdict; this check reads the registry.
>&2 echo   fix                : "%SystemRoot%\SysWOW64\regsvr32.exe" /s "%VG_DLL%"
>&2 echo                        reg import "%~dp0..\register.reg" /reg:32
endlocal & exit /b 1

:vg_pass
>&2 echo REGVERIFY OK: %VG_PROGID% -^> %VG_CLSID% ^; InprocServer32 = %VG_INPROC% ^; LoadBehavior = %VG_LB%
endlocal & exit /b 0

:vg_nodll
>&2 echo REGVERIFY ERROR: no DLL path given.
>&2 echo   usage: verify_registration.bat ^<path-to-the-DLL-that-must-be-registered^>
endlocal & exit /b 1
