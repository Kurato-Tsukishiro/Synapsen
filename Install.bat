@chcp 65001
@echo off

set /p yn_check="必要ライブラリ (Python) 及びブラウザ (Playwright Chromium) をインストールしてよいですか? (y/n): "

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
echo "インストールが完了しました。"
echo "このバッチファイルは手動で削除して構いません。"
pause
exit