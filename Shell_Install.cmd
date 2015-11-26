
::FOR /F "tokens=1,2 delims==" %%i in ('assoc .exe') do @Echo %%i %%j
::if %%j=="exefile" (
@set RegPath=HKCR\exefile\shell\Decompile with myAutToExe\command
::reg query "%RegPath%" /ve
::if errorlevel 1 ()
reg add "%RegPath%" /ve /t REG_SZ /d "%~DP0myAutToExe.exe "%%1""
@echo Installed. (add 'Decompile with myAutToExe' when you rightclick on *.exe)
@pause