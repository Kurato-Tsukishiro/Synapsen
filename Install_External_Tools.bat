@chcp 65001
@echo off
setlocal

echo "=================================================="
echo "Synapsen 外部ツールセットアップ (EXE版ユーザー向け)"
echo "=================================================="
echo.
echo " このスクリプトは、Synapsenの拡張機能に必要な以下のツールをセットアップします。"
echo.
echo "  1. Pandoc (必須: Markdown変換機能に使用)"
echo "  2. Playwright ブラウザ (Python環境がある場合のみ: Webクリップ機能に使用)"
echo.

set /p yn_check="インストールを開始してよいですか? (y/n): "
IF /I "%yn_check%"=="y" (
    GOTO :StartInstall
) ELSE (
    echo "キャンセルしました。"
    pause
    exit
)

:StartInstall

REM --- Pandoc Installation ---
echo.
echo "[1/2] Pandoc のインストール確認..."
where pandoc >nul 2>nul
IF %ERRORLEVEL% EQU 0 (
    echo  "Pandoc は既にインストールされています。スキップします。"
) ELSE (
    echo  "Pandoc が見つかりません。Winget経由でインストールを試みます..."
    where winget >nul 2>nul
    IF %ERRORLEVEL% EQU 0 (
        winget install -e --id JohnMacFarlane.Pandoc
    ) ELSE (
        echo  "[警告] Winget が見つかりませんでした。"
        echo  "Pandoc を自動インストールできません。以下から手動でインストールしてください。"
        echo  "https://pandoc.org/installing.html"
    )
)

REM --- Playwright Installation ---
echo.
echo "[2/2] Playwright ブラウザ (Chromium) のインストール..."
echo " ※ Python環境がインストールされていない場合、この手順はスキップされます。"

where playwright >nul 2>nul
IF %ERRORLEVEL% EQU 0 (
    echo "Playwrightコマンドが見つかりました。ブラウザをインストールします..."
    playwright install chromium
) ELSE (
    echo  "[Info] Playwrightコマンドが見つかりませんでした (Python未検出)。"
    echo  "Webクリップ機能を使用するには、別途Python環境とライブラリの導入が必要です。"
    echo  "(通常のPDF管理機能には影響ありません)"
)

echo.
echo "=================================================="
echo " セットアップ処理が完了しました。"
echo "=================================================="
pause
endlocal
exit