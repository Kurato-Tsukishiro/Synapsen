@chcp 65001
@echo off

set /p yn_check="必要ライブラリをインストールしてよいですか? (y/n): "

REM /I オプションで大文字/小文字を区別せず比較する
IF /I "%yn_check%"=="y" (
    GOTO :Install
) ELSE (
    echo "インストールをキャンセルしました。"
    pause
    exit
)

:Install
echo "pipでライブラリのインストールを開始します..."
pip install -r requirements.txt

echo "インストールが完了しました。"
echo "このバッチファイルは手動で削除して構いません。"
pause
exit
