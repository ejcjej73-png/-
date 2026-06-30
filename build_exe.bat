@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  카드·외환 뉴스 스크랩 - EXE 빌드
echo ============================================
echo PyInstaller 및 라이브러리를 설치합니다...
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
echo.
echo EXE 파일을 만듭니다...
pyinstaller --onefile --name news_scrap news_scrap.py
echo.
echo ============================================
echo  완료! dist 폴더 안의 news_scrap.exe 확인
echo  (exe 와 같은 폴더에 '설정.ini'를 두면
echo   키워드/설정을 바꿀 수 있습니다)
echo ============================================
pause
