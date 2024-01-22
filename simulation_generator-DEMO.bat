@echo off

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
REM configure python path, not required if python is on the path
SET PYTHONPATH=
REM SET PYTHONPATH=C:\Python\Python312\
SET PYTHONCMD=python
IF NOT "%PYTHONPATH%" == "" SET PYTHONCMD=%PYTHONPATH%\python

ECHO =================================
"%PYTHONCMD%" -V
ECHO =================================

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
REM install the additional python lib required
REM IF NOT "%PYTHONPATH%" == "" "%PYTHONPATH%\Scripts\pip" install pandas
REM IF NOT "%PYTHONPATH%" == "" "%PYTHONPATH%\Scripts\pip" install pyarrow
REM IF NOT "%PYTHONPATH%" == "" "%PYTHONPATH%\Scripts\pip" install requests 
REM IF NOT "%PYTHONPATH%" == "" "%PYTHONPATH%\Scripts\pip" install xlsxwriter

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:: REST API URL : http|https://host(:port)(/WarName)/rest
SET RESTAPIURL=https://demo-eu.castsoftware.com/Engineering/rest

::When the Engineering dashboard URL and Rest API don't have the same root, fill the below parameter
::if empty will take above URL without /rest 
:: Engineering dahsboard URL  : http|https://host(:port)(/EngineeringWarName) 
REM SET EDURL=https://demo-eu.castsoftware.com/Engineering

REM SET APIKEY=N/A
SET USER=CIO
SET PASSWORD=cast

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:: [Optional] Application name regexp filter, if not defined all application will be exported
::SET APPFILTER=Webgoat^|eComm.*
SET APPFILTER=shopizer
::HR Management with JEE
::Webgoat

:: [Optional] Load modules ? default = false
SET LOADMODULES=false

:: Aggregation mode (FullApplication/ByNumberOfArtifacts) default = FullApplication
::SET AGGREGATIONMODE=FullApplication

:: [Optional] Inputs CSV file containing the quality rules efforts (default is CAST_QualityRulesEffort.csv)
::SET EFFORTFILEPATH=C:/Temp/CAST_QualityRulesEffort.csv

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: This section contains only parameters related to the Violations tab

:: [Optional] Load violations in a Violations tab ? default = false
SET LOADVIOLATIONS=false

:: [Optional] Quality rule id regexp filter
::SET QRIDFILTER=7802|7804

:: [Optional] Quality rule name regexp  filter
::SET QRNAMEFILTER=

:: [Optional] Critical rules violations filter: true|false
::SET CRITICALONLYFILTER=true

:: [Optional] Business criterion filter : 60017 (Total Quality Index)|60016 (Security)|60014 (Efficiency)|60013 (Robustness)|60011 (Transferability)|60012 (Changeability)
:: to filter the violations and retrieve the PRI for this business criterion (if only one is selected)
::SET BCFILTER=60016,60014

:: [Optional] Technology list filter
::SET TECHNOFILTER=JEE,SQL

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
REM Build the command line
SET CMD="%PYTHONCMD%" "%~dp0simulator_generator.py" 

IF DEFINED RESTAPIURL 				SET CMD=%CMD% -restapiurl "%RESTAPIURL%"
IF DEFINED EDURL					SET CMD=%CMD% -edurl "%EDURL%"

IF NOT DEFINED USER 				SET USER=N/A
IF NOT DEFINED PASSWORD 			SET PASSWORD=N/A
IF NOT DEFINED APIKEY 				SET APIKEY=N/A
SET CMD=%CMD% -user "%USER%" -password "%PASSWORD%" -apikey "%APIKEY%"

SET CURRENTFOLDER=%~dp0
:: remove trailing \
SET CURRENTFOLDER=%CURRENTFOLDER:~0,-1%

SET OUTPUTFOLDER=%CURRENTFOLDER%

SET LOGFILE=%CURRENTFOLDER%\simulation_generator.log
IF DEFINED LOGFILE					SET CMD=%CMD% -log "%LOGFILE%"
IF DEFINED OUTPUTFOLDER 			SET CMD=%CMD% -of "%OUTPUTFOLDER%"

SET EXTENSIONINSTALLATIONFOLDER=%CURRENTFOLDER%
SET CMD=%CMD% -extensioninstallationfolder "%EXTENSIONINSTALLATIONFOLDER%"

ECHO APPFILTER=%APPFILTER%
IF DEFINED LOADMODULES				SET CMD=%CMD% -loadmodules "%LOADMODULES%"
IF DEFINED AGGREGATIONMODE			SET CMD=%CMD% -aggregationmode "%AGGREGATIONMODE%"
IF DEFINED LOADVIOLATIONS			SET CMD=%CMD% -loadviolations "%LOADVIOLATIONS%"
IF DEFINED APPFILTER 				SET CMD=%CMD% -applicationfilter "%APPFILTER%"
IF DEFINED EFFORTFILEPATH			SET CMD=%CMD% -effortcsvfilepath "%EFFORTFILEPATH%"
IF DEFINED QRIDFILTER				SET CMD=%CMD% -qridfilter %QRIDFILTER%
IF DEFINED QRNAMEFILTER				SET CMD=%CMD% -qrnamefilter "%QRNAMEFILTER%"
IF DEFINED CRITICALONLYFILTER		SET CMD=%CMD% -criticalrulesonlyfilter "%CRITICALONLYFILTER%"
IF DEFINED BCFILTER					SET CMD=%CMD% -businesscriterionfilter "%BCFILTER%"
IF DEFINED TECHNOFILTER				SET CMD=%CMD% -technofilter "%TECHNOFILTER%"

:: Max nbRows for the Rest API calls
::SET NBROWS=100000000
IF DEFINED NBROWS					SET CMD=%CMD% -nbrows "%NBROWS%"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

ECHO Running the command line 
ECHO %CMD%
%CMD%
SET RETURNCODE=%ERRORLEVEL%
ECHO RETURNCODE %RETURNCODE% 

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


PAUSE