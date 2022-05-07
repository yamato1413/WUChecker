@echo off

C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe -target:winexe  %~dp0\WUChecker.cs -win32icon:%~dp0\wmploc_132_9.ico
rem C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe %~dp0\WUChecker.cs -win32icon:%~dp0\wmploc_132_9.ico

if %ERRORLEVEL% equ 0 (
    .\WUChecker.exe
)
