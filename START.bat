@echo off
setlocal EnableExtensions
chcp 65001 >nul
set "PORT=8763"
set "APP_URL=http://127.0.0.1:%PORT%/app/index.html"
set "WORDLAB_ROOT=%~dp0."
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$expected = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($env:WORDLAB_ROOT).TrimEnd([IO.Path]::DirectorySeparatorChar))); try { $r = Invoke-WebRequest -UseBasicParsing -Uri '%APP_URL%' -TimeoutSec 1; if ($r.StatusCode -eq 200 -and $r.Headers['X-Word-Lab-Root'] -eq $expected) { exit 0 }; exit 2 } catch { if ($_.Exception.Response) { exit 2 } }; exit 1" >nul 2>nul
if errorlevel 2 goto port_conflict
if not errorlevel 1 goto open_app

if not exist "%~dp0wordlab-server.ps1" (
  echo [WORD LAB] wordlab-server.ps1 파일을 찾을 수 없습니다.
  echo ZIP 압축을 완전히 해제한 뒤 START.bat을 다시 실행해 주세요.
  pause
  exit /b 1
)

start "WORD LAB Server" /min powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0wordlab-server.ps1" -Port %PORT% -Root "%~dp0."

powershell -NoProfile -ExecutionPolicy Bypass -Command "$expected = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([IO.Path]::GetFullPath($env:WORDLAB_ROOT).TrimEnd([IO.Path]::DirectorySeparatorChar))); $deadline = (Get-Date).AddSeconds(10); do { try { $r = Invoke-WebRequest -UseBasicParsing -Uri '%APP_URL%' -TimeoutSec 1; if ($r.StatusCode -eq 200 -and $r.Headers['X-Word-Lab-Root'] -eq $expected) { exit 0 } } catch {}; Start-Sleep -Milliseconds 500 } while ((Get-Date) -lt $deadline); exit 1" >nul 2>nul
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

:port_conflict
echo [WORD LAB] 8763 포트에서 현재 폴더의 WORD LAB 서버를 확인할 수 없습니다.
echo 이전 버전 WORD LAB 서버라면 해당 창을 닫은 뒤 이 폴더의 START.bat을 다시 실행해 주세요.
echo 다른 프로그램이 사용 중이라면 임의로 종료하지 마세요.
pause
exit /b 1
