@chcp 65001
@echo off
setlocal

echo "=========================================="
echo "Synapsen 統合インストーラー (Poetry版)"
echo "=========================================="

:CheckPoetry
echo "[1/3] 環境チェック中..."
where poetry >nul 2>nul
IF %ERRORLEVEL% NEQ 0 (
    echo "Poetryが見つかりません。pipでインストールします..."
    pip install poetry
)

REM --- オプション選択: AI機能 ---
echo.
set "EXTRAS_CMD="
echo "[2/3] 機能選択"
echo.
echo  "AI機能 (Ollama連携) は、ローカルLLMを使用して画像を自動タグ付けする機能です。"
set /p install_ai=" AI機能をインストールしますか? (y/n): "
IF /I "%install_ai%"=="y" (
    echo  "AI機能を追加します。"
    set "EXTRAS_CMD=--extras ai"
)

REM Web機能の選択肢は削除しました（上級者向け手動インストール）

REM --- ライブラリインストール実行 ---
echo.
echo "------------------------------------------"
echo "Pythonライブラリのインストールを開始します..."
echo "------------------------------------------"

if "%EXTRAS_CMD%"=="" (
    call poetry install --no-root --without dev
) else (
    call poetry install --no-root --without dev %EXTRAS_CMD%
)

REM --- 外部ツール (旧 install_option.bat) ---
echo.
echo "[3/3] 外部ツールのセットアップ"
echo.
echo " 以下のツールは Synapsen の一部機能 (Webクリップ、Markdown変換) に必要です。"
echo "  1. Playwright ブラウザ (Chromium) - Webページ保存に使用"
echo "  2. Pandoc - 文書形式の変換に使用"
echo.
set /p install_tools=" これらの外部ツールをインストールしてよいですか? (y/n): "

IF /I "%install_tools%"=="y" (
    GOTO :InstallTools
) ELSE (
    echo "外部ツールのインストールをスキップしました。"
    GOTO :End
)

:InstallTools
echo.
echo "--- Playwright (Chromium) のインストール ---"
echo "数百MBのダウンロードが発生します..."
call poetry run playwright install chromium

echo.
echo "--- Pandoc のインストール ---"
where /q winget
IF %ERRORLEVEL% EQU 0 (
    echo "Winget経由でPandocをインストールします..."
    winget install -e --id JohnMacFarlane.Pandoc
) ELSE (
    echo "[警告] Wingetが見つかりません。"
    echo "Pandocが必要な場合は、公式サイトから手動でインストールしてください。"
    echo "https://pandoc.org/installing.html"
)

:End
echo.
echo "=========================================="
echo " すべてのセットアップが完了しました"
echo "=========================================="
echo.
echo  "起動するには 'Run_Launcher.bat' を実行してください。"
echo.
pause
endlocal
exit