@echo on

REM generate date format
for /f "tokens=1-4 delims=/ " %%i in ("%date%") do (
     set dow=%%i
     set month=%%j
     set day=%%k
     set year=%%l
   )

REM set log files
set LOGFILE=log\download_log.txt


REM download file to sFTP server
"C:\Program Files (x86)\WinSCP\"WinSCP.exe /ini=nul /script=C:\inetpub\wwwroot\sftp\batch\FTPScript\download.txt /log=C:\inetpub\wwwroot\sftp\batch\log\download_log.txt

"%ProgramFiles%\WinRAR\WinRAR.exe" a -afzip -agYYYY-MM-DD -cfg- -ep1 -ibck -inul -m5 -tn1d -x*.zip -y -- "C:\inetpub\wwwroot\sftp\report\Report_.zip" "C:\inetpub\wwwroot\sftp\report\*"


python message.py


If %ERRORLEVEL% NEQ 0 python errorMessage.py

exit /B
