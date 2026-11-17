@echo off
setlocal

rem Build DocxGen with the Delphi/RAD Studio available on this machine.
rem Set BDS to use a specific installation, otherwise the newest one is used.

set "ROOT=%~dp0"
set "PROJECT=%ROOT%DocxGen.dproj"

if defined BDS (
    set "BDS_ROOT=%BDS%\"
) else (
    call :FindNewestDelphi
)
if not defined BDS_ROOT (
    echo [build] No RAD Studio found in the registry.
    exit /b 1
)
if not exist "%BDS_ROOT%bin\rsvars.bat" (
    echo [build] rsvars.bat not found in "%BDS_ROOT%bin".
    exit /b 1
)

call "%BDS_ROOT%bin\rsvars.bat"
echo [build] Using "%BDS%"

msbuild "%PROJECT%" /t:Build /p:Config=Release /p:Platform=Win32 /v:m
if errorlevel 1 (
    echo [build] Build failed.
    exit /b 1
)

echo [build] Output: %ROOT%build\bin\Win32\Release\DocxGen.exe
exit /b 0

:FindNewestDelphi
set BEST_MAJOR=0
for /f "delims=" %%K in ('reg query "HKLM\SOFTWARE\WOW6432Node\Embarcadero\BDS" 2^>nul ^| findstr /r "\\[0-9][0-9]*\.[0-9]*$"') do call :ConsiderVersion "%%K"
if not defined BDS_KEY exit /b 0
for /f "tokens=2,*" %%A in ('reg query "%BDS_KEY%" /v RootDir 2^>nul ^| findstr /i "RootDir"') do set "BDS_ROOT=%%B"
exit /b 0

:ConsiderVersion
for /f "delims=." %%M in ("%~nx1") do set /a MAJOR=%%M
if %MAJOR% gtr %BEST_MAJOR% (
    set /a BEST_MAJOR=%MAJOR%
    set "BDS_KEY=%~1"
)
exit /b 0
