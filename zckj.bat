@echo off

del %SystemRoot%\system32\smsocx80.ocx
COPY Mscomctl.ocx %SystemRoot%\system32\Mscomctl.ocx
COPY Mscomm32.ocx %SystemRoot%\system32\Mscomm32.ocx
COPY smsocx80.ocx %SystemRoot%\system32\smsocx80.ocx
	
regsvr32 %SystemRoot%\system32\smsocx80.ocx



