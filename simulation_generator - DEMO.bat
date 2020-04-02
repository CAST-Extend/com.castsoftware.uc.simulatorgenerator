@echo off

SET PYTHONPATH=C:\Python\Python37
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
"%PYTHONPATH%\Scripts\pip" install pandas
"%PYTHONPATH%\Scripts\pip" install requests 
"%PYTHONPATH%\Scripts\pip" install xlsxwriter
:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
:: WAR params : http|https://host(:port)(/WarName)
SET URL=https://demo-eu.castsoftware.com/Engineering
SET CMD_URL=-url "%URL%"

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

SET LOGFILE=%OUTPUTFOLDER%\sim_gene_crmmoder.log
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

:: Max nbRows for the Rest API calls
::SET NBROWS=100000000
SET CMD_NBROWS=
::SET CMD_NBROWS=-nbrows "%CMD_NBROWS%"

:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
ECHO Running the command line 
SET CMD="%PYTHONPATH%\python" "%~dp0simulator_generator.py" %CMD_URL% %CMD_USER% %CMD_PASSWORD% %CMD_APIKEY% %CMD_LOGFILE% %CMD_OUTPUTFOLDER% %CMD_APPFILTER%  %CMD_NBROWS% %CMD_EFFORTFILEPATH%
ECHO %CMD%
%CMD%
SET RETURNCODE=%ERRORLEVEL%
ECHO RETURNCODE %RETURNCODE% 

PAUSE