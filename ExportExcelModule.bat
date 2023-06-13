@echo off

rem ---------------------------------------------------------------
rem 機能：VBAのソースをExportする
rem ---------------------------------------------------------------
rem 使い方：
rem     当バッチファイルと下記呼び出しているExportExcelModule.vbsを
rem     VBAソースを記述したExcelファイルの保存フォルダの親フォルダに格納する
rem ---------------------------------------------------------------

rem VBAのソースの保存箇所を設定
set EXPORT_PATH="C:\Users\RightServe\Documents\GitHub\seikyuusyoVBA"

rem このバッチが存在するフォルダをカレントに移動
pushd %0\..

cls


rem --------------------------------------------------------------------
rem カレントフォルダとサブフォルダに含めている全てのEXCEL（xlsmも対象になる）をループし、
rem ExportExcelModule.vbsでソースをエクスポートする
rem (ループはここでしているため、若干性能が悪いが...VBSで再帰処理を書かなくて済む。)
rem --------------------------------------------------------------------
for /F "usebackq" %%i in (`dir /s /b *.xls `) do ( 
    echo %%i 
    CScript ExportExcelModule.vbs %%i %EXPORT_PATH%
    rem pause
)
pause
exit

rem --------------------------------------------------------------------
rem 勉強時メモ
rem メモ１：カレントディレクトリの拡張子がxlsのファイルを出力
rem    for %%i in (*.xls) do ( echo %%i )
rem メモ２(正規表現が利かなかった)：
rem for /F "usebackq" %%i in (`dir /s /b *.xls ^| findstr /V ".*\.xls$" `) do ()
rem --------------------------------------------------------------------
