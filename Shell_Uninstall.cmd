@set RegPath=HKCR\exefile\shell\Decompile with myAutToExe\command
reg delete "%RegPath%" /ve /f
@echo Uninstalled
@pause