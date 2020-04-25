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
:: WAR params : http|https://host(:port)(/WarName)
SET RESTAPIURL=https://demo-eu.castsoftware.com/Engineering
SET CMD_URL=-restapiurl "%RESTAPIURL%"

::When the Engineering dashboard URL and Rest API are different, fill the below parameter 
:: Engineering dahsboard URL  : http|https://host(:port)(/EngineeringWarName) 
::SET EDURL=https://demo-eu.castsoftware.com/Engineering
SET CMD_EDURL=
::SET CMD_EDURL=-edurl "%EDURL%"

::SET USER=N/A
SET USER=CIO
SET CMD_USER=-user "%USER%"
::SET PASSWORD=N/A
SET PASSWORD=cast
SET CMD_PASSWORD=-password "%PASSWORD%"
SET APIKEY=N/A
SET CMD_APIKEY=-apikey "%APIKEY%"

:: Output folder
SET OUTPUTFOLDER=C:\Users\mmr\workspace\com.castsoftware.uc.simulatorgenerator
SET CMD_OUTPUTFOLDER=-of "%OUTPUTFOLDER%"

SET LOGFILE=%OUTPUTFOLDER%\simulation_generator.log
SET CMD_LOGFILE=-log "%LOGFILE%"
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

:: Application name regexp filter
::SET APPFILTER=Webgoat^|eComm.*
SET APPFILTER=CRM Moder
SET CMD_APPFILTER=
SET CMD_APPFILTER=-applicationfilter "%APPFILTER%"

:: Inputs CSV file containing the quality rules efforts (default is CAST_QualityRulesEffort.csv)
::SET EFFORTFILEPATH=C:/Temp/CAST_QualityRulesEffort.csv
SET CMD_EFFORTFILEPATH=
::SET CMD_EFFORTFILEPATH=-effortcsvfilepath "%EFFORTFILEPATH%"

:: Load violations ? default = false
::SET LOADVIOLATIONS=true
SET CMD_LOADVIOLATIONS=
::SET CMD_LOADVIOLATIONS=-loadviolations %LOADVIOLATIONS%

:: Quality rule id regexp filter
::SET QRIDFILTER=7802|7804
SET CMD_QRIDFILTER=
::SET CMD_QRIDFILTER=-qridfilter %QRIDFILTER%

:: Quality rule name regexp  filter
::SET QRNAMEFILTER=
SET CMD_QRNAMEFILTER=
::SET CMD_QRIDFILTER=-qrnamefilter "%QRNAMEFILTER%"

:: Critical rules violations filter: true|false
::SET CRITICALONLYFILTER=true
SET CMD_CRITICALONLYFILTER=
::SET CMD_CRITICALONLYFILTER=-criticalrulesonlyfilter "%CRITICALONLYFILTER%"

:: Business criterion filter : 60017 (Total Quality Index)|60016 (Security)|60014 (Efficiency)|60013 (Robustness)|60011 (Transferability)|60012 (Changeability)
:: to filter the violations and retrieve the PRI for this business criterion (if only one is selected)
::SET BCFILTER=60016,60014
::SET BCFILTER=60016
SET CMD_BCFILTER=
::SET CMD_BCFILTER=-businesscriterionfilter "%BCFILTER%"

:: Technology list filter
::SET TECHNOFILTER=JEE,SQL
SET CMD_TECHNOFILTER=
::SET CMD_TECHNOFILTER=-technofilter "%TECHNOFILTER%"

:: Max nbRows for the Rest API calls
::SET NBROWS=100000000
SET CMD_NBROWS=
::SET CMD_NBROWS=-nbrows "%CMD_NBROWS%"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ECHO Running the command line 
SET CMD="%PYTHONPATH%\python" "%~dp0simulator_generator.py" %CMD_URL% %CMD_USER% %CMD_PASSWORD% %CMD_APIKEY% %CMD_LOGFILE% %CMD_OUTPUTFOLDER% %CMD_APPFILTER%  %CMD_NBROWS% %CMD_EFFORTFILEPATH% %CMD_LOADVIOLATIONS% %CMD_QRIDFILTER% %CMD_QRNAMEFILTER% %CMD_CRITICALONLYFILTER% %CMD_BCFILTER% %CMD_TECHNOFILTER%
ECHO %CMD%
%CMD%
SET RETURNCODE=%ERRORLEVEL%
ECHO RETURNCODE %RETURNCODE% 

PAUSE