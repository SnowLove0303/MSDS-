@echo off
rem ============================================================
rem  Msds Editor - unified command entry
rem  GUI:    double-click or run  Msds-Editor.bat
rem  CLI:    Msds-Editor.bat cli <docx-path>
rem  TEST:   Msds-Editor.bat test
rem  DOCTOR: Msds-Editor.bat doctor
rem ============================================================
cd /d "%~dp0"
where pwsh >nul 2>&1
if %ERRORLEVEL% equ 0 (
    pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0Msds-Editor.ps1" %*
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Msds-Editor.ps1" %*
)
if %ERRORLEVEL% neq 0 pause