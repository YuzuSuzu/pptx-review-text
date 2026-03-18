@echo off
chcp 65001 > nul
:: ============================================================
:: combine_agents.bat
::
:: ~/.codex/skills/ 内の全 .md ファイルを結合して
:: ~/.codex/AGENTS.md を自動生成するスクリプト。
::
:: 使い方:
::   1. このファイルを %USERPROFILE%\.codex\ に置く
::   2. スキルファイルを %USERPROFILE%\.codex\skills\ に置く
::   3. このファイルをダブルクリック or コマンドプロンプトで実行
::
:: 実行するたびに AGENTS.md が上書き再生成されます。
:: ============================================================

set SKILLS_DIR=%USERPROFILE%\.codex\skills
set OUTPUT=%USERPROFILE%\.codex\AGENTS.md

:: skills フォルダが存在しない場合は作成
if not exist "%SKILLS_DIR%" (
    mkdir "%SKILLS_DIR%"
    echo [INFO] skills フォルダを作成しました: %SKILLS_DIR%
)

:: スキルファイルが存在するか確認
set FILE_COUNT=0
for %%f in ("%SKILLS_DIR%\*.md") do set /a FILE_COUNT+=1

if %FILE_COUNT% == 0 (
    echo [WARNING] %SKILLS_DIR% に .md ファイルが見つかりません。
    echo           スキルファイルを配置してから再実行してください。
    pause
    exit /b 1
)

:: AGENTS.md を生成（既存ファイルは上書き）
echo. > "%OUTPUT%"

for %%f in ("%SKILLS_DIR%\*.md") do (
    echo # =========================================>> "%OUTPUT%"
    echo # %%~nf>> "%OUTPUT%"
    echo # =========================================>> "%OUTPUT%"
    type "%%f" >> "%OUTPUT%"
    echo.>> "%OUTPUT%"
    echo.>> "%OUTPUT%"
    echo   追加: %%~nxf
)

echo.
echo [完了] AGENTS.md を生成しました: %OUTPUT%
echo        スキル数: %FILE_COUNT% 個
echo.
pause
