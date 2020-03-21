echo off
set DataPath=
set curdir=%cd%\
echo %curdir%
pushd %curdir%

rem cd ..\
set absDatapath=%cd%\%DataPath%
echo %absDatapath%
rem set libpath=%cd%\libs\
popd

set Vobfoldername="\TestVOb_2\nsbdvob"
set sourceFoldername="_a\nsbd"
rem set releaseID=


set rootÎÄ¼þ¼Ð=%absDatapath%
cmd.exe /k
pause