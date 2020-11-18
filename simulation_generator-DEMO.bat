@echo off

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
REM configure python path if not defined
REM min version to use is Python 3.6
SET PYTHON_PATH_IF_NOT_DEF_IN_ENV=C:\Python\Python37
IF "%PYTHONPATH%"=="" SET PYTHONPATH=%PYTHON_PATH_IF_NOT_DEF_IN_ENV%
"%PYTHONPATH%\python" -V
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
REM install the additional python lib required
"%PYTHONPATH%\Scripts\pip" install pandas
"%PYTHONPATH%\Scripts\pip" install requests 
"%PYTHONPATH%\Scripts\pip" install xlsxwriter
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
SET APPFILTER=Webgoat

:: [Optional] Inputs CSV file containing the quality rules efforts (default is CAST_QualityRulesEffort.csv)
::SET EFFORTFILEPATH=C:/Temp/CAST_QualityRulesEffort.csv

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: This section contains only parameters related to the Violations tab

:: [Optional] Load violations in a Violations tab ? default = false
::SET LOADVIOLATIONS=true

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
::SET CMD="%PYTHONPATH%\python" "%~dp0simulator_generator.py" %CMD_URL% %CMD_USER% %CMD_PASSWORD% %CMD_APIKEY% %CMD_LOGFILE% %CMD_OUTPUTFOLDER% %CMD_APPFILTER% %CMD_NBROWS% %CMD_EFFORTFILEPATH% %CMD_LOADVIOLATIONS% %CMD_QRIDFILTER% %CMD_QRNAMEFILTER% %CMD_CRITICALONLYFILTER% %CMD_BCFILTER% %CMD_TECHNOFILTER% %CMD_EXTENSIONINSTALLATIONFOLDER%
SET CMD="%PYTHONPATH%\python" "%~dp0simulator_generator.py" 

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
IF DEFINED APPFILTER 				SET CMD=%CMD% -applicationfilter "%APPFILTER%"
IF DEFINED EFFORTFILEPATH			SET CMD=%CMD% -effortcsvfilepath "%EFFORTFILEPATH%"
IF DEFINED QRIDFILTER				SET CMD=%CMD% -qridfilter %QRIDFILTER%
IF DEFINED QRNAMEFILTER				SET CMD=%CMD% -qrnamefilter "%QRNAMEFILTER%"
IF DEFINED CRITICALONLYFILTER		SET CMD=%CMD% -criticalrulesonlyfilter "%CRITICALONLYFILTER%"
IF DEFINED BCFILTER					SET CMD=%CMD% -businesscriterionfilter "%BCFILTER%"
IF DEFINED CMD_TECHNOFILTER			SET CMD=%CMD% -technofilter "%TECHNOFILTER%"

:: Max nbRows for the Rest API calls
::SET NBROWS=100000000
IF DEFINED CMD_NBROWS				SET CMD=%CMD% -nbrows "%CMD_NBROWS%"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

ECHO Running the command line 
ECHO %CMD%
%CMD%
SET RETURNCODE=%ERRORLEVEL%
ECHO RETURNCODE %RETURNCODE% 

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


PAUSE