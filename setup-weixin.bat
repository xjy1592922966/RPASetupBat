@echo off
setlocal enabledelayedexpansion
chcp 65001

@REM 判断是否存在D盘或E盘，选择其中一个硬盘作为项目目录
IF EXIST D:\ (
    set project_dir=D:\wxrpa
) ELSE IF EXIST E:\ (
    set project_dir=E:\wxrpa
) ELSE IF EXIST C:\ (
    set project_dir=C:\wxrpa
) ELSE (
    echo 未找到可用的硬盘，请手动指定项目目录。
    pause
    exit
)

@REM 创建项目目录和相关文件夹
mkdir %project_dir%
mkdir %project_dir%\download
mkdir %project_dir%\lib
mkdir %project_dir%\lib\python39
mkdir %project_dir%\lib\javaopenjdk
mkdir %project_dir%\lib\code

@REM 下载文件到指定目录
set files[0]=https://oss-cdn.hcolor.pro/autojs/project/wechat/wexin.zip,wechat.zip
set files[1]=https://oss-cdn.hcolor.pro/autojs/project/wechat/python.exe,python3964.exe
set files[2]=https://download.java.net/java/GA/jdk19.0.2/fdb695a9d9064ad6b064dc6df578380c/7/GPL/openjdk-19.0.2_windows-x64_bin.zip,java.zip
set files[3]=https://download.microsoft.com/download/9/C/D/9CD480DC-0301-41B0-AAAB-FE9AC1F60237/VSU4/vcredist_x86.exe,vcredist_x86.exe
set files[4]=https://oss-cdn.hcolor.pro/autojs/project/wechat/TagUIWindows.zip,TagUIWindows.zip


for /l %%i in (0, 1, 4) do (
    set "line=!files[%%i]!"
    for /f "tokens=1,2 delims=," %%a in ("!line!") do (
        set "@REMote=%%a"
        set "localfile=%%b"
        set "localpath=%project_dir%\download\!localfile!"
        echo.
        echo ----------------------------
        echo 下载文件: !@REMote!
        echo 本地路径: !localpath!
        if exist !localpath! (
            echo 文件已存在，跳过下载。
        ) else (
            echo ----------------------------
            echo.
            curl.exe --url "!@REMote!" --output "!localpath!"
        )
    )
)

echo 解压核心代码中,请等待
PowerShell -Command Expand-Archive -Force %project_dir%\download\wechat.zip %project_dir%\lib
echo 解压核心代码完成

echo ----------------------------
echo 静默安装python中,请等待
start /wait %project_dir%\download\python3964.exe /quiet InstallAllUsers=1 PrependPath=1 TargetDir=%project_dir%\lib\python39
echo 静默安装python完成

echo ----------------------------
echo 添加python环境变量中,请等待
setx /M PATH "%project_dir%\lib\python39;%project_dir%\lib\python39\Scripts;%PATH%"
echo 添加python环境变量完成

echo ----------------------------


echo 解压java的openjdk中,请等待
PowerShell -Command Expand-Archive -Force %project_dir%\download\java.zip %project_dir%\lib\javaopenjdk
echo 解压java的openjdk完成

echo ----------------------------

echo 添加java的openjdk环境变量中,请等待
setx /M JAVA_HOME "%project_dir%\lib\javaopenjdk\jdk-19.0.2"
setx /M PATH "%project_dir%\lib\javaopenjdk\jdk-19.0.2\bin;%PATH%"

echo 添加java的openjdk环境变量完成

echo ----------------------------
echo 静默安装vcredist中,请等待
start /wait %project_dir%\download\vcredist_x86.exe /q
echo 静默安装vcredist完成
echo ----------------------------

echo 复制TagUI中,请等待
copy /y %project_dir%\download\TagUIWindows.zip %project_dir%\lib\code
mkdir %APPDATA%\tagui
copy /y %project_dir%\download\TagUIWindows.zip %APPDATA%\tagui

echo 复制TagUI完成
echo ----------------------------

echo "创建桌面快捷方式..."
set "SRC_FILE=%project_dir%\lib\python39\python.exe"
set "DST_FILE=%userprofile%\Desktop\添加企业微信RPA.lnk"
set "PARAMS=%project_dir%\lib\main.py"
set "WORKING_DIR=%project_dir%\lib"
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%DST_FILE%'); $Shortcut.TargetPath = 'python'; $Shortcut.Arguments= '%PARAMS%'; $Shortcut.WorkingDirectory = '%WORKING_DIR%'; $Shortcut.Save()"

echo 创建桌面快捷方式完成

echo.
echo ----------------------------
echo 下载和安装完成
echo ----------------------------
echo.
pause