
@echo off
echo ****************************************************************
echo **************�Զ���jar����װ��maven���ؿ�***************
echo ****************************************************************
echo ****************************************************************

::���ñ���
set p_Length = 4
set Cfile=%cd%\config.bcfg
echo �����ļ���ַ��%Cfile%

::�������ļ��л�ȡ����
setlocal enabledelayedexpansion
for /f "tokens=1* delims==" %%a  in (%Cfile%) do (
    ::set p[!count!]=%%a
     set s1=%%a
     if !s1! equ groupId (set p[0]=%%b)
     if !s1! equ artifactId (set p[1]=%%b)
     if !s1! equ Dversion (set p[2]=%%b)
)

::��ȡjar������·��
for %%j in (*.jar) do ( set p[3]=%cd%\%%j)

::�ȴ��û�ȷ�ϲ���
echo.
echo ****************************���ò�����Ϣ����******************************
echo *jar��λ�ã�%p[3]%
echo *groupId ��    %p[0]%    
echo *artifactId :  %p[1]%   
echo *Dversion :    %p[2]%
echo ***************************************************************************
set /p ok=��ȷ�����ò����Ƿ���ȷ(y/n)?:
echo.
if %ok%==y (
echo ��ʼ��������Ժ�....
call mvn install:install-file -Dfile=%p[3]% -DgroupId=%p[0]% -DartifactId=%p[1]% -Dversion=%p[2]% -Dpackaging=jar
) else (
   exit
)
::���maven�ο�����
echo *******��������������ӵ������Ŀpom�ļ���*******
echo.
echo   ^<dependency^>
echo        ^<groupId^>%p[0]%^</groupId^>
echo        ^<artifactId^>%p[1]%^</artifactId^>
echo        ^<version^>%p[2]%^</version^>
echo   ^</dependency^>
echo.

cmd