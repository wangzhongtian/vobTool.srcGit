set DataPath=��ʱ��������
set curdir=%cd%\
echo %curdir%
pushd %curdir%

cd ..\..
set absDatapath=%cd%\%DataPath%
echo %absDatapath
REM set LibPAth= %cd%\tools\compiled\libs
path=%path%;%LibPAth%
echo %path%

set root�ļ���=%curdir%
set libpath=%cd%\libs
popd

set IRONPYTHONSTARTUP=E:\ipy\LibSRC\Shell.py

popd
cmd /Q
pause

