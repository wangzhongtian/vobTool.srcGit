set DataPath=定时调测数据
set curdir=%cd%\
echo %curdir%
pushd %curdir%

cd ..\..
set absDatapath=%cd%\%DataPath%
echo %absDatapath
set LibPAth= %cd%\tools\compiled\libs
set path="C:\Program Files\IronPython 2.7";%path%;%LibPAth%
echo %path%
set IRONPYTHONSTARTUP=E:\ipy\LibSRC\Shell.py

popd

rem 
ipy64 compileIPy2DLL.py
copy AppAssembly.dll ..\libs\

copy StdLibALLs.dll ..\libs\
ipy compileIPy2DLL.py
pause