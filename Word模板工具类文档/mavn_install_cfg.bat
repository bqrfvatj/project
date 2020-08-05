
@echo off
echo ****************************************************************
echo **************自动将jar包安装到maven本地库***************
echo ****************************************************************
echo ****************************************************************

::设置变量
set p_Length = 4
set Cfile=%cd%\config.bcfg
echo 配置文件地址：%Cfile%

::从配置文件中获取参数
setlocal enabledelayedexpansion
for /f "tokens=1* delims==" %%a  in (%Cfile%) do (
    ::set p[!count!]=%%a
     set s1=%%a
     if !s1! equ groupId (set p[0]=%%b)
     if !s1! equ artifactId (set p[1]=%%b)
     if !s1! equ Dversion (set p[2]=%%b)
)

::获取jar包所在路径
for %%j in (*.jar) do ( set p[3]=%cd%\%%j)

::等待用户确认参数
echo.
echo ****************************配置参数信息如下******************************
echo *jar包位置：%p[3]%
echo *groupId ：    %p[0]%    
echo *artifactId :  %p[1]%   
echo *Dversion :    %p[2]%
echo ***************************************************************************
set /p ok=请确认配置参数是否正确(y/n)?:
echo.
if %ok%==y (
echo 开始打包，请稍候....
call mvn install:install-file -Dfile=%p[3]% -DgroupId=%p[0]% -DartifactId=%p[1]% -Dversion=%p[2]% -Dpackaging=jar
) else (
   exit
)
::输出maven参库配置
echo *******复制以下内容添加到你的项目pom文件中*******
echo.
echo   ^<dependency^>
echo        ^<groupId^>%p[0]%^</groupId^>
echo        ^<artifactId^>%p[1]%^</artifactId^>
echo        ^<version^>%p[2]%^</version^>
echo   ^</dependency^>
echo.

cmd