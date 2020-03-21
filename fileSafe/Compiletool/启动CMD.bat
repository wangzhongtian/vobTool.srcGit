set DataPath=定时调测数据
set curdir=%cd%\
echo %curdir%
pushd %curdir%

cd ..\..
set absDatapath=%cd%\%DataPath%
echo %absDatapath
REM set LibPAth= %cd%\tools\compiled\libs
path=%path%;%LibPAth%
echo %path%

set root文件夹=%curdir%
set libpath=%cd%\libs
popd

set IRONPYTHONSTARTUP=E:\ipy\LibSRC\Shell.py

popd
cmd /Q
pause

