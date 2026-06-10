@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PORT=8763"
set "APP_URL=http://127.0.0.1:%PORT%/app/index.html"
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $r = Invoke-WebRequest -UseBasicParsing -Uri '%APP_URL%' -TimeoutSec 1; if ($r.StatusCode -ge 200) { exit 0 } } catch {}; exit 1" >nul 2>nul
if not errorlevel 1 goto open_app

if not exist "%~dp0wordlab-server.ps1" (
  echo [WORD LAB] wordlab-server.ps1 파일을 찾을 수 없습니다.
  echo ZIP 압축을 완전히 해제한 뒤 START.bat을 다시 실행해 주세요.
  pause
  exit /b 1
)

start "WORD LAB Server" /min powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0wordlab-server.ps1" -Port %PORT% -Root "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$deadline = (Get-Date).AddSeconds(10); do { try { $r = Invoke-WebRequest -UseBasicParsing -Uri '%APP_URL%' -TimeoutSec 1; if ($r.StatusCode -ge 200) { exit 0 } } catch {}; Start-Sleep -Milliseconds 500 } while ((Get-Date) -lt $deadline); exit 1" >nul 2>nul
if errorlevel 1 (
  echo [WORD LAB] 로컬 서버를 시작하지 못했습니다.
  echo 서버 창에 표시된 오류를 확인한 뒤 다시 실행해 주세요.
  echo 직접 열 주소: %APP_URL%
  pause
  exit /b 1
)

:open_app
start "" "%APP_URL%"
exit /b 0
