@echo off
set version=%1
set nuggetexe=C:\Program Files\Nuget\nuget.exe
set EnableNugetPackageRestore=true
SET CMD="%nuggetexe%"
ECHO %CMD%
%CMD%
SET CMD="%nuggetexe%" pack plugin.nuspec -ExcludeEmptyDirectories
ECHO %CMD%
%CMD%