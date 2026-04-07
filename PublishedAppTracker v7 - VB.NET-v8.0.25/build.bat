:: Build requires installation of dotnet SDK https://dotnet.microsoft.com/en-us/download/dotnet/8.0

@echo off
:: Set Variables:
set EXENAME=PublishedAppTracker.exe
set DESTDIR=D:\Documents\- TEST -\_Projects\PublishedAppTracker
set PROJNAME=PublishedAppTracker

:: Set the working directory:
cd /d "%~dp0"

:: Close any running instances:
echo Closing any running instances...
taskkill /IM %EXENAME% /F 2>nul

:: Restore NuGet packages (needed for WebView2):
echo Restoring packages...
dotnet restore

:: Compile the project:
echo Compiling...
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeAllContentForSelfExtract=true -o "bin\Publish"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Compilation successful! Copying exe...
    if not exist "%DESTDIR%" mkdir "%DESTDIR%"
    copy /Y "bin\Publish\%EXENAME%" "%DESTDIR%\%EXENAME%"
    echo Starting app...
    timeout /t 2
    start "" "%DESTDIR%\%EXENAME%"
    exit
) else (
    echo.
    echo Compilation failed!
    pause
)