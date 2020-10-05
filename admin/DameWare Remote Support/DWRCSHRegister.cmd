
@IF "%2"=="" GOTO Syntax
@IF "%1"=="" GOTO Syntax


@REM Echo "%1"
@REM Echo "%2"
@REM Pause

@IF "%3"=="" GOTO CopyMRCSh
@IF NOT EXIST "%3" mkdir "%3"

:CopyMRCSh

Copy %1 %2
@IF ERRORLEVEL 1 GOTO ERROR
regsvr32 /s %2
@IF ERRORLEVEL 1 GOTO ERROR
@REM Pause
GOTO End

@Pause

:Syntax

@Echo ***********************************************************************
@Echo *                                                                     *
@Echo *  Usage:                                                             * 
@Echo *  DWRCSHRegister.cmd source_file destination_file                    *
@Echo *  DWRCSHRegister.cmd source_file destination_file destination_dir    *
@Echo *                                                                     *
@Echo *  Example:                                                           *
@Echo *  DWRCSHRegister C:\DWRCSh.DLL.Win32 C:\Windows\dwrcs\DWRCSh.DLL     *
@Echo *                                                                     *
@Echo *********************************************************************** 


:ERROR
@PAUSE

:End