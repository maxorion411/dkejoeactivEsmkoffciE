@set masver=2.6
@setlocal DisableDelayedExpansion
@echo off

::============================================================================
::
::   This script is a part of 'Microsoft-Activation-Scripts' (MAS) project.
::
::   Homepage: mass grave[.]dev
::      Email: windowsaddict@protonmail.com
::
::============================================================================

::  To activate Office with Ohook activation, set 1 in below line
set _act=1

::  To remove Ohook activation, set 1 in below line
set _rem=0

::========================================================================================================================================

::  Set Path variable, it helps if it is misconfigured in the system

set "PATH=%SystemRoot%\System32;%SystemRoot%\System32\wbem;%SystemRoot%\System32\WindowsPowerShell\v1.0\"
if exist "%SystemRoot%\Sysnative\reg.exe" (
set "PATH=%SystemRoot%\Sysnative;%SystemRoot%\Sysnative\wbem;%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\;%PATH%"
)

:: Re-launch the script with x64 process if it was initiated by x86 process on x64 bit Windows
:: or with ARM64 process if it was initiated by x86/ARM32 process on ARM64 Windows

set "_cmdf=%~f0"
for %%# in (%*) do (
if /i "%%#"=="r1" set r1=1
if /i "%%#"=="r2" set r2=1
if /i "%%#"=="-qedit" (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f 1>nul
rem check the code below admin elevation to understand why it's here
)
)

if exist %SystemRoot%\Sysnative\cmd.exe if not defined r1 (
setlocal EnableDelayedExpansion
start %SystemRoot%\Sysnative\cmd.exe /c ""!_cmdf!" %* r1"
exit /b
)

:: Re-launch the script with ARM32 process if it was initiated by x64 process on ARM64 Windows

if exist %SystemRoot%\SysArm32\cmd.exe if %PROCESSOR_ARCHITECTURE%==AMD64 if not defined r2 (
setlocal EnableDelayedExpansion
start %SystemRoot%\SysArm32\cmd.exe /c ""!_cmdf!" %* r2"
exit /b
)

::========================================================================================================================================

set "blank="
set "mas=ht%blank%tps%blank%://mass%blank%grave.dev/"

::  Check if Null service is working, it's important for the batch script

sc query Null | find /i "RUNNING"
if %errorlevel% NEQ 0 (
echo:
echo Null service is not running, script may crash...
echo:
echo:
echo Help - %mas%troubleshoot.html
echo:
echo:
ping 127.0.0.1 -n 10
)
cls

::  Check LF line ending

pushd "%~dp0"
>nul findstr /v "$" "%~nx0" && (
echo:
echo Error: Script either has LF line ending issue or an empty line at the end of the script is missing.
echo:
ping 127.0.0.1 -n 6 >nul
popd
exit /b
)
popd

::========================================================================================================================================

cls
color 07
title  Ohook Activation %masver%

set _args=
set _elev=
set _unattended=0

set _args=%*
if defined _args set _args=%_args:"=%
if defined _args (
for %%A in (%_args%) do (
if /i "%%A"=="/Ohook"                  set _act=1
if /i "%%A"=="/Ohook-Uninstall"        set _rem=1
if /i "%%A"=="-el"                     set _elev=1
)
)

for %%A in (%_act% %_rem%) do (if "%%A"=="1" set _unattended=1)

::========================================================================================================================================

set "nul1=1>nul"
set "nul2=2>nul"
set "nul6=2^>nul"
set "nul=>nul 2>&1"

set psc=powershell.exe
set winbuild=1
for /f "tokens=6 delims=[]. " %%G in ('ver') do set winbuild=%%G

set _NCS=1
if %winbuild% LSS 10586 set _NCS=0
if %winbuild% GEQ 10586 reg query "HKCU\Console" /v ForceV2 %nul2% | find /i "0x0" %nul1% && (set _NCS=0)

if %_NCS% EQU 1 (
for /F %%a in ('echo prompt $E ^| cmd') do set "esc=%%a"
set     "Red="41;97m""
set    "Gray="100;97m""
set   "Green="42;97m""
set    "Blue="44;97m""
set  "_White="40;37m""
set  "_Green="40;92m""
set "_Yellow="40;93m""
) else (
set     "Red="Red" "white""
set    "Gray="Darkgray" "white""
set   "Green="DarkGreen" "white""
set    "Blue="Blue" "white""
set  "_White="Black" "Gray""
set  "_Green="Black" "Green""
set "_Yellow="Black" "Yellow""
)

set "nceline=echo: &echo ==== ERROR ==== &echo:"
set "eline=echo: &call :dk_color %Red% "==== ERROR ====" &echo:"
if %~z0 GEQ 200000 (
set "_exitmsg=Go back"
set "_fixmsg=Go back to Main Menu, select Troubleshoot and run Fix Licensing option."
) else (
set "_exitmsg=Exit"
set "_fixmsg=In MAS folder, run Troubleshoot script and select Fix Licensing option."
)

::========================================================================================================================================

if %winbuild% LSS 9200 (
%eline%
echo Unsupported OS version detected [%winbuild%].
echo Ohook Activation is supported on Windows 8 and later and their server equivalent.
goto dk_done
)

for %%# in (powershell.exe) do @if "%%~$PATH:#"=="" (
%nceline%
echo Unable to find powershell.exe in the system.
goto dk_done
)

::========================================================================================================================================

::  Fix special characters limitation in path name

set "_work=%~dp0"
if "%_work:~-1%"=="\" set "_work=%_work:~0,-1%"

set "_batf=%~f0"
set "_batp=%_batf:'=''%"

set _PSarg="""%~f0""" -el %_args%

set "_ttemp=%userprofile%\AppData\Local\Temp"
set "_Local=%LocalAppData%"
setlocal EnableDelayedExpansion

::========================================================================================================================================

echo "!_batf!" | find /i "!_ttemp!" %nul1% && (
if /i not "!_work!"=="!_ttemp!" (
%eline%
echo Script is launched from the temp folder,
echo Most likely you are running the script directly from the archive file.
echo:
echo Extract the archive file and launch the script from the extracted folder.
goto dk_done
)
)

::========================================================================================================================================

::  Elevate script as admin and pass arguments and preventing loop

%nul1% fltmc || (
if not defined _elev %psc% "start cmd.exe -arg '/c \"!_PSarg:'=''!\"' -verb runas" && exit /b
%eline%
echo This script needs admin rights.
echo To do so, right click on this script and select 'Run as administrator'.
goto dk_done
)

::========================================================================================================================================

::  This code disables QuickEdit for this cmd.exe session only without making permanent changes to the registry
::  It is added because clicking on the script window pauses the operation and leads to the confusion that script stopped due to an error

if %_unattended%==1 set quedit=1
for %%# in (%_args%) do (if /i "%%#"=="-qedit" set quedit=1)

reg query HKCU\Console /v QuickEdit %nul2% | find /i "0x0" %nul1% || if not defined quedit (
reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "0" /f %nul1%
start cmd.exe /c ""!_batf!" %_args% -qedit"
rem quickedit reset code is added at the starting of the script instead of here because it takes time to reflect in some cases
exit /b
)

::========================================================================================================================================

::  Check for updates

set -=
set old=

for /f "delims=[] tokens=2" %%# in ('ping -4 -n 1 massgrave.dev') do if not defined old set old=%%#
for /f "delims=[] tokens=2" %%# in ('ping -6 -n 1 massgrave.dev') do if not defined old set old=%%#

if defined old (
set -=%old:~-10%
for /f "tokens=1 delims=." %%# in ("!-=!") do set old=%%# && set -=
if defined old if !old! LSS 200 set old=
)

for /f %%# in ('curl massgrave.dev/ 2^>nul ^| find "ohook-v"') do (
set old=%%#
set old=!old:*ohook-v=!
set old=!old:~0,4!
)

if defined old (
set /a old=1!old!
set /a old=1!masver!
if %old% LSS 2!masver! (
set /a old=%%# - 1!masver!
if %old% LSS 0 set /a old=!old! * -1
echo:
echo [NEWER VERSION AVAILABLE]: %old% versions ahead.
echo Current: %masver%
echo: 
)
)

::========================================================================================================================================

if not defined _act if not defined _rem (
echo Activating Office with Ohook activation.
set _act=1
)

::========================================================================================================================================

:dk_donotremove

::  Load slmgr.vbs script if it is missing
reg query HKCR\VBSFile\Shell\Open\Command %nul1% || (
if not exist %SystemRoot%\System32\slmgr.vbs (
copy /y "%~dp0bin\slmgr.vbs" %SystemRoot%\System32\ 1>nul 2>&1
)
if not exist %SystemRoot%\SysWOW64\slmgr.vbs if exist %SystemRoot%\SysWOW64\cscript.exe (
copy /y "%~dp0bin\slmgr.vbs" %SystemRoot%\SysWOW64\ 1>nul 2>&1
)
)

::  Load Ohook files if missing

if not exist "%SystemRoot%\System32\spp\tokens\ohook\ohook.spp" (
if not exist %SystemRoot%\SysWOW64\ospp.vbs if exist %SystemRoot%\SysWOW64\cscript.exe (
copy /y "%~dp0bin\sppvbs.vbs" %SystemRoot%\SysWOW64\ 1>nul 2>&1
)
copy /y "%~dp0bin\sppvbs.vbs" %SystemRoot%\System32\ 1>nul 2>&1
)

::========================================================================================================================================

::  Add task for removing OGA (Office Genuine Advantage) prompts for O2010 retail
schtasks /create /tn "Office Genuine Advantage" /tr "%~dp0bin\delete.vbs" /sc onstart /ru "SYSTEM" 2>nul

::========================================================================================================================================

::  Remove old Ohook files

if exist %SystemRoot%\System32\spp\tokens\ohook_old (
rmdir /s /q %SystemRoot%\System32\spp\tokens\ohook_old
)

::========================================================================================================================================

::  Check for Windows8Plus and Windows7

ver | findstr /r /c:"6\.1\." %nul1% && set win7=1 || set win7=0
ver | findstr /r /c:"6\.3\." %nul1% || ver | findstr /r /c:"10\.0\." %nul1% && set windows8plus=1 || set windows8plus=0

::========================================================================================================================================

::  Check for Token.dat location

reg query "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion" /v ProductName %nul1% | findstr /r /c:"Windows Server" %nul1% && set srv=1 || set srv=0

set winxp=0
reg query "HKLM\Software\Microsoft\Windows NT\CurrentVersion" /v "ProductId" %nul1% | find /i "55274" %nul1% && set winxp=1 || set winxp=0

if %winxp%==1 set windows8plus=0

if %windows8plus%==1 (
for /f "tokens=2*" %%a in ('reg query "HKLM\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "ServiceLogonEnabled"') do set ohookservlogon=%%b
set ohookservlogon=!ohookservlogon:~0,1!
)

if %win7%==1 (
for /f "tokens=2*" %%a in ('reg query "HKLM\Software\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform" /v "ServiceLogonEnabled"') do set ohookservlogon=%%b
set ohookservlogon=!ohookservlogon:~0,1!
)

::========================================================================================================================================

set olver=
for /f "tokens=2 delims=." %%# in ('ver') do set olver=%%#

::========================================================================================================================================

set Office2016=

::  Check and Set Variables

if exist "%SystemRoot%\Sysnative\ospp.vbs" (
if %PROCESSOR_ARCHITECTURE%==AMD64 (
%psc% -noprofile -executionpolicy unrestricted -command "Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name ServerPackage* -ErrorAction SilentlyContinue | ForEach-Object {if (\$_ -match 'Access' -or \$_ -match 'Excel' -or \$_ -match 'OneNote' -or \$_ -match 'Outlook' -or \$_ -match 'PowerPoint' -or \$_ -match 'Publisher' -or \$_ -match 'Word') {write-output \"Found Office 2016\"} }" 2>nul | find "Found Office 2016" >nul && (
set Office2016=1
)
)
)

::  Check for Office Paths

if not defined Office2016 (
for /f "tokens=2*" %%a in ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds') do (
set Office2016=1
if not "!Office2016!"=="1" (
for /f "tokens=2 delims==" %%b in ('find "InstallPath=" "%SystemRoot%\System32\spp\tokens\ohook\ohook.spp" 2^>nul') do set OfficeInstallPath=%%b
)
)
)

if not defined OfficeInstallPath if defined Office2016 (
for /f "tokens=2*" %%a in ('reg query "HKLM\Software\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds') do (
for /f "tokens=1* delims==" %%b in ('find "InstallPath=" "%SystemRoot%\System32\spp\tokens\ohook\ohook.spp" 2^>nul') do set OfficeInstallPath=%%b
)
)

if not defined OfficeInstallPath if defined Office2016 (
for /f "tokens=1* delims==" %%b in ('find "InstallPath=" "%SystemRoot%\System32\spp\tokens\ohook\ohook.spp" 2^>nul') do set OfficeInstallPath=%%b
)

if not defined OfficeInstallPath if not defined Office2016 (
for /f "tokens=1* delims==" %%b in ('find "InstallPath=" "%SystemRoot%\System32\spp\tokens\ohook\ohook.spp" 2^>nul') do set OfficeInstallPath=%%b
)

::  Check if Office is already Activated

%psc% -noprofile -executionpolicy unrestricted -command "Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name ClientPackage* -ErrorAction SilentlyContinue | ForEach-Object {if (\$_ -match 'Access' -or \$_ -match 'Excel' -or \$_ -match 'OneNote' -or \$_ -match 'Outlook' -or \$_ -match 'PowerPoint' -or \$_ -match 'Publisher' -or \$_ -match 'Word') {write-output \"Found Office 2016\"} }" 2>nul | find "Found Office 2016" >nul && (
set Office2016=1
)

::========================================================================================================================================

::  Determine action to be taken based on parameters

if %_act%==1 (
goto :oh_installkey
) else if %_rem%==1 (
goto :oh_remove
) else (
goto :dk_done
)

::========================================================================================================================================

:oh_installkey

%psc% -noprofile -executionpolicy unrestricted -command "if ((Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name ServerPackage* -ErrorAction SilentlyContinue | ForEach-Object {if (\$_ -match 'Access' -or \$_ -match 'Excel' -or \$_ -match 'OneNote' -or \$_ -match 'Outlook' -or \$_ -match 'PowerPoint' -or \$_ -match 'Publisher' -or \$_ -match 'Word') {write-output \"Found Office 2016\"} }) -ne \$null) { exit 0 } else { exit 1 }" 2>nul || (
%eline%
echo No supported Office installation found.
goto dk_done
)

%psc% -noprofile -executionpolicy unrestricted -command "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" %nul1%
%psc% -noprofile -executionpolicy unrestricted -command "Import-Module -Name OfficeScrubber; Add-OfficeScrubber -Office2016" %nul1%

goto dk_done

::========================================================================================================================================

:oh_remove

%psc% -noprofile -executionpolicy unrestricted -command "if ((Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -Name ServerPackage* -ErrorAction SilentlyContinue | ForEach-Object {if (\$_ -match 'Access' -or \$_ -match 'Excel' -or \$_ -match 'OneNote' -or \$_ -match 'Outlook' -or \$_ -match 'PowerPoint' -or \$_ -match 'Publisher' -or \$_ -match 'Word') {write-output \"Found Office 2016\"} }) -ne \$null) { exit 0 } else { exit 1 }" 2>nul || (
%eline%
echo No supported Office installation found.
goto dk_done
)

%psc% -noprofile -executionpolicy unrestricted -command "Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force" %nul1%
%psc% -noprofile -executionpolicy unrestricted -command "Import-Module -Name OfficeScrubber; Remove-OfficeScrubber -Office2016" %nul1%

goto dk_done

::========================================================================================================================================

:dk_done
::  Reset quickedit state if enabled in registry and disabled for this session
if defined quedit reg add HKCU\Console /v QuickEdit /t REG_DWORD /d "1" /f %nul1%
echo:
echo Done.
endlocal
endlocal
echo:
pause
exit /b
