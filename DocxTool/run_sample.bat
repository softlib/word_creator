@echo off
setlocal
set "ROOT=%~dp0"
set "EXE=%ROOT%build\bin\Win32\Release\DocxGen.exe"

if not exist "%EXE%" (
    echo [run] DocxGen.exe not found. Please run build.bat first.
    exit /b 1
)

"%EXE%" -c "%ROOT%config.sample.json" -o "%ROOT%output\sample.docx"
if errorlevel 1 exit /b 1

echo [run] Sample document: %ROOT%output\sample.docx
exit /b 0
