echo off
set DataPath=tools\��ʱ��������
set curdir=%cd%\
echo %curdir%
pushd %curdir%

cd ..\

set absDatapath=%cd%\%DataPath%
echo %absDatapath%
set libpath=%cd%\libs
popd


set ��������=����004��
rem set root�ļ���=C:\Cur\2016-10-25\tools\��ʱ��������
set root�ļ���=%absDatapath%
set deltaMax=0 
set �ֳ�ʱ���У=0


prompt %��������%:
cmd.exe /k
pause