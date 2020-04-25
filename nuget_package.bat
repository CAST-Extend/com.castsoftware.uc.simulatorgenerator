@echo off
set version=%1
set nugetexe=C:\Program Files\Nuget\nuget.exe
set EnableNugetPackageRestore=true
SET CMD="%nugetexe%"
ECHO %CMD%
%CMD%
SET CMD="%nugetexe%" pack plugin.nuspec -ExcludeEmptyDirectories
ECHO %CMD%
%CMD%