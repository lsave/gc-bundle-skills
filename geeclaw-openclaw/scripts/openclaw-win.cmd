@echo off
setlocal

set "PREFIX=[geeclaw-openclaw]"
set "WRAPPER="

if defined GEECLAW_OPENCLAW_WRAPPER (
  if exist "%GEECLAW_OPENCLAW_WRAPPER%" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "$p='%GEECLAW_OPENCLAW_WRAPPER%'; if (Test-Path $p -PathType Leaf) { if (Select-String -Path $p -Pattern 'GeeClaw' -Quiet) { exit 0 } }; exit 1" >nul 2>&1
    if not errorlevel 1 set "WRAPPER=%GEECLAW_OPENCLAW_WRAPPER%"
  )
)

if not defined WRAPPER (
  for /f "usebackq delims=" %%i in (`powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='SilentlyContinue';" ^
    "function Test-GeeClawWrapper([string]$p) { if (-not $p) { return $false }; if (-not (Test-Path $p -PathType Leaf)) { return $false }; return [bool](Select-String -Path $p -Pattern 'GeeClaw' -Quiet) }" ^
    "$candidates = New-Object System.Collections.Generic.List[string];" ^
    "$cmd = Get-Command openclaw -ErrorAction SilentlyContinue;" ^
    "if ($cmd -and (Test-GeeClawWrapper $cmd.Source)) { Write-Output $cmd.Source; exit 0 }" ^
    "$geeclaw = Get-Command GeeClaw -ErrorAction SilentlyContinue;" ^
    "if ($geeclaw) { $candidates.Add((Join-Path (Split-Path $geeclaw.Source -Parent) 'resources\managed-bin\openclaw.cmd')) }" ^
    "if ($env:LOCALAPPDATA) { $candidates.Add((Join-Path $env:LOCALAPPDATA 'Programs\GeeClaw\resources\managed-bin\openclaw.cmd')) }" ^
    "if ($env:ProgramFiles) { $candidates.Add((Join-Path $env:ProgramFiles 'GeeClaw\resources\managed-bin\openclaw.cmd')) }" ^
    "if ($env:'ProgramFiles(x86)') { $candidates.Add((Join-Path $env:'ProgramFiles(x86)' 'GeeClaw\resources\managed-bin\openclaw.cmd')) }" ^
    "foreach ($candidate in $candidates) { if (Test-GeeClawWrapper $candidate) { Write-Output $candidate; exit 0 } }" ^
    "exit 1"`) do set "WRAPPER=%%i"
)

if not defined WRAPPER (
  echo %PREFIX% 错误: 未找到 GeeClaw wrapper。
  echo %PREFIX% 不会回退到未知的 system openclaw。
  exit /b 1
)

call "%WRAPPER%" %*
set "EXIT_CODE=%ERRORLEVEL%"
endlocal & exit /b %EXIT_CODE%
