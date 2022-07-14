@echo on

REM generate date format
for /f "tokens=1-4 delims=/ " %%i in ("%date%") do (
     set dow=%%i
     set month=%%j
     set day=%%k
     set year=%%l
   )

for /f "tokens=1,2,3,4 delims=:." %%m in ("%time%") do (
    
     set hh=%%m
     set mm=%%n
     set ss=%%o
   )

REM set log files
set LOGFILE=log\log.txt
set ErrorLOGFILE=log\errorlog.txt

REM add timestamp to source file and copy it to submission folder
If exist "Source\store_HOHK.dat" copy Source\store_HOHK.dat Submission\store_HOHK_%year%%month%%day%%hh%%mm%%ss%.dat  

REM add timestamp to voucher files and move them to submission folder
REM set result=false

REM add timestamp to voucher files and move them to submission folder
REM If exist "Source\voucher_HOHK_20_*.dat" set result=true
REM If exist "Source\voucher_HOHK_50_*.dat" set result=true
REM If exist "Source\voucher_HOHK_100_*.dat" set result=true

REM If "%result%" == "true" (type Source\voucher_HOHK_20_*.dat Source\voucher_HOHK_50_*.dat Source\voucher_HOHK_100_*.dat > Submission\voucher_HOHK_%year%%month%%day%%hh%%mm%%ss%.dat)

Rem move vouche files to submission folder
If exist "Source\voucher_HOHK_20_*.dat" copy Source\voucher_HOHK_20_*.dat Submission\voucher_HOHK_20_%year%%month%%day%%hh%%mm%%ss%.dat
If exist "Source\voucher_HOHK_50_*.dat" copy Source\voucher_HOHK_50_*.dat Submission\voucher_HOHK_50_%year%%month%%day%%hh%%mm%%ss%.dat
If exist "Source\voucher_HOHK_100_*.dat" copy Source\voucher_HOHK_100_*.dat Submission\voucher_HOHK_100_%year%%month%%day%%hh%%mm%%ss%.dat
If exist "Source\voucher_HOHK_120_*.dat" copy Source\voucher_HOHK_120_*.dat Submission\voucher_HOHK_120_%year%%month%%day%%hh%%mm%%ss%.dat
REM If exist "Source\voucher*.dat" del Source\voucher*.dat

REM delete all the voucher files 
If exist "Source\voucher*.dat" del Source\voucher*.dat


REM upload file to UAT sFTP server
"C:\Program Files (x86)\WinSCP\"WinSCP.exe /ini=nul /script=C:\inetpub\wwwroot\sftp\batch\FTPScript\Upload.txt /log=C:\inetpub\wwwroot\sftp\batch\log\log.txt

If %ERRORLEVEL% NEQ 0 echo.Error was found at uploading file.>>%ErrorLOGFILE%
If %ERRORLEVEL% NEQ 0 python message.py

exit /B
