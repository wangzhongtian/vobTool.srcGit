echo off
set DataPath=
set curdir=%cd%\
echo %curdir%
pushd %curdir%

rem cd ..\
set absDatapath=%cd%\%DataPath%
echo %absDatapath%
set libpath=%cd%\libs\
set libpath=f:\ipy\libs\
popd

set BasicRoot=SSDBIWIn

set BasicRoot=wangzht

set Vobfoldername=\%BasicRoot%\TestVOb_2\nsbdvob

rem 
set sourceFoldername=%BasicRoot%\_a\nsbd
rem set sourceFoldername=_a\nsbd
rem set releaseID=00169


set rootÎÄ¼þ¼Ð=%absDatapath%
cmd.exe /k
pause