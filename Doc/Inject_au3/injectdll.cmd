@prompt -$G
  
  @set TargetExe=%~nx1
  @set TargetAu3=%~n1.au3
  
  @set AutoIt3_dll=%~dp0AutoIt3.dll
  @set SecondsToWait=1


  @call :CheckUsage
  %FinalCommand%

  
  start "my injectdll Launcher" "%TargetExe%" "%TargetAu3%"


Call :GetPID "%TargetExe%"

::@Pause "Press any key if target is loaded.
@timeout /T %SecondsToWait%
::@ping -n %SecondsToWait% 127.0.0.1 > NUL
 
::Set /P PID="Please PID of target: "

@cd /D %~dp0
injectdll.exe %PID% "%AutoIt3_dll%"
::@pause
@goto :EOF

:GetPID

  @set CMD_Tasklist=^
  Tasklist /FI "IMAGENAME eq %~1" /FO CSV /NH

  @FOR /F "tokens=1,2 delims=," %%i in ('%CMD_Tasklist%') do set PID=%%j
 
  ::%CMD_Tasklist%

@goto :EOF


:CheckUsage

@Set FinalCommand=

@echo inject Dll
@echo ==========
@echo.

@if not exist "%TargetExe%" (
  @echo Usage: injectdll ^<SomeCompile.Exe^>
  @echo.
  
  
  @Set FinalCommand=@goto :quit
)  

@if not exist "%TargetAu3%" (
  @echo required #1: ^<SomeCompile.au3^> ^(same name as *.exe but *.au3 extension ^)
  @echo              that you like to run inside ^<SomeCompile.Exe^> 
  @echo.
  @Set FinalCommand=@goto :quit

)

@if not exist "%AutoIt3_dll%" (
  @echo required #2: Make a AutoIt3.dll !!! 
  @echo              Create it via 
  @echo                'ExeToDll.au3 ^<some AutoIt3.exe^>'
  @echo.
  @Set FinalCommand=@goto :quit
)

@goto :EOF


:quit

  @pause >nul