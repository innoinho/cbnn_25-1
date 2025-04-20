@echo off
REM 근무출장명령서 생성기 EXE 빌드를 위한 배치 스크립트 (콘솔 창 숨김 옵션 포함)
REM 이 파일을 스크립트(.py)와 동일한 폴더에 저장 후 실행하세요.

:: 현재 디렉터리로 이동
cd /d "%~dp0"

:: PyInstaller로 단일 EXE 생성 (폰트 리소스 포함, 콘솔 창 숨김)
pyinstaller --onefile --windowed --add-data "C:/Windows/Fonts/malgun.ttf;." --name "근무출장명령서_생성기" autowork_v2.py

:: 빌드 완료 메시지 표시 후 잠시 대기
echo 빌드가 완료되었습니다. dist 폴더를 확인하세요.
pause
