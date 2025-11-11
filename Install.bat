@chcp 65001
@echo off

set /p yn_check="必要ライブラリ (Python), ブラウザ (Playwright), Pandoc (MD変換) をインストールしてよいですか? (y/n): "

REM /I オプションで大文字/小文字を区別せず比較する
IF /I "%yn_check%"=="y" (
    GOTO :Install
) ELSE (
    echo "インストールをキャンセルしました。"
    pause
    exit
)

:Install
echo "pipでPythonライブラリのインストールを開始します..."
pip install -r requirements.txt

echo "---"
echo "Playwright (WebClip機能) に必要な Chromium ブラウザをインストールします..."
echo "(これには数分かかり、数百MBのファイルをダウンロードします)"
playwright install chromium

echo "---"
echo "Pandoc (Markdown変換機能) をインストールします..."
echo "(Wingetを使用してインストールを試みます)"
echo "  (ユーザーアカウント制御(UAC)のポップアップが表示された場合は「はい」を押してください)"

REM Winget が利用可能か確認し、Pandocをインストール
where /q winget
IF %ERRORLEVEL% EQU 0 (
    echo "Wingetが見つかりました。Pandocのインストールを開始します..."
    winget install -e --id JohnMacFarlane.Pandoc
    echo "Pandocのインストールが完了しました。"
) ELSE (
    echo "警告: Winget が見つかりませんでした。"
    echo "Pandoc の自動インストールをスキップします。"
    echo "Markdown連携機能を使用する場合は、手動で Pandoc をインストールしてください。"
    echo "https://pandoc.org/installing.html"
)

echo "---"
echo "インストールが完了しました。"
echo "このバッチファイルは手動で削除して構いません。"
pause
exit