@echo off
chcp 65001 > nul
echo ====================================================
echo      오전시황 대시보드 자동화 파이프라인 구동 (엑셀 전용)
echo ====================================================
echo.

echo [1/2] 엑셀 데이터 기반 - HTML 대시보드 빌드 중...
python generate_dashboard_v4_2.py "★ 시장_모니터링.xlsx" dashboard_v4_2.html
if %errorlevel% neq 0 (
    echo [오류] HTML 대시보드 빌드 실패!
    pause
    exit /b %errorlevel%
)
echo.

echo [2/2] 깃허브 저장소 동기화 중 (Push)...
git add "★ 시장_모니터링.xlsx" dashboard_v4_2.html app.py requirements.txt
git commit -m "Auto Update: 시황 데이터 업데이트"
git push origin main
if %errorlevel% neq 0 (
    echo [오류] 깃허브 푸시에 실패했습니다. 네트워크 상태 또는 로컬 Git 설정을 확인하세요.
    pause
    exit /b %errorlevel%
)
echo.

echo ====================================================
echo   동기화 완료! 이제 대시보드가 실시간 업데이트됩니다.
echo ====================================================
pause
