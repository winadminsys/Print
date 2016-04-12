#@echo off

set CustomFolder="C:\Temp\CustomFiles\"
set PrinterPath="\\myserver\myprinter"
set AcrobatReaderPath="C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe"
set NbrOfLoop=500
set SimpleTexteFilePath="C:\Temp\CustomFiles\helloworld.txt"
cls

IF EXIST %CustomFolder% (
	pushd %CustomFolder%

	For /L %%n in (1, 1, %NbrOfLoop%) do (
		echo loop -----  %%n
		REM PRINT %SimpleTexteFilePath% /D:%PrinterPath%
		REM NOTEPAD /P %SimpleTexteFilePath%
		REM %AcrobatReaderPath% /t %SimpleTexteFilePath%
		
		For %%f in (*.*) do (
			PRINT "%%f" /D:%PrinterPath%
			REM NOTEPAD /P "%%f"
			REM %AcrobatReaderPath% /t %SimpleTexteFilePath%
		)
	)
	popd
) Else (
echo %CustomFolder% does not exist !!!
)
