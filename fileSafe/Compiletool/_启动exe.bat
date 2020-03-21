echo off
set DataPath=tools\定时调测数据
set curdir=%cd%\
echo %curdir%
pushd %curdir%

cd ..\

set absDatapath=%cd%\%DataPath%
echo %absDatapath%
set libpath=%cd%\libs
popd


set 主机名号=易县004号
rem set root文件夹=C:\Cur\2016-10-25\tools\定时调测数据
set root文件夹=%absDatapath%
set deltaMax=0 
set 现场时间调校=0


prompt %主机名号%:
cmd.exe /k
pause