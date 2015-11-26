#NoTrayIcon
ProcessSetPriority(@AutoItPID, 0)

Global $gVarCmdline_1, $gVarRetLoadSkinCrafter
If $cmdline[0] Then
	If $cmdline[0] <= 1 Then $gVarCmdline_1 = $cmdline[1]
EndIf
	
	
Global Const $gVar0167 = 7
Global Const $gVarRndFlag = 1
Global Const $gVarRndMaxA = 9
Global Const $gVar016A = 2
Global $gUseOfEnCodeItNotPermitted = Fn0081("You are not permitted to use EnCodeIt")
Global $gDataUserInfo = Fn0081("You have been verified and authorized to use EnCodeIt")
Global Const $gVarRndMinB = 8
Global Const $gVarRndD1 = 5
Global Const $gVarRndMinA = 3
Global Const $gVar0170 = 6
Global Const $gVarRndC = 4
Global Const $gVarRndC2 = 0
Global Const $AutoItPID = @AutoItPID
Global Const $ScriptDir = @ScriptDir
Global Const $AppDataDir = @AppDataDir
Global Const $DocumentsCommonDir = @DocumentsCommonDir
Global Const $ScriptFullPath = @ScriptFullPath
Global Const $UserName = @UserName
Global Const $HomeDrive = @HomeDrive
Global Const $UserProfileDir = @UserProfileDir
Global Const $CRLF = @CRLF
Global Const $CR = @CR
Global Const $lf = @LF


Global Const $gVarRetDummy1 = InitFiles()
Global Const $gVarRetDummy2 = FnFilesCleanUp()
;Global Const $gVarRetDummy3 = FnInetCheckBlackList1()
Global $EnCodeItInfo_Dat

HotKeySet('{ESC}', 'Unload_EnCodeIt')
Opt("GUIOnEventMode", 1)
Opt("GUIResizeMode", 802)
Opt("WinTitleMatchMode", 4)
Func LoadSkinCrafter($SkinEngineDll, $SkinDll)
	Local $hDll = DllOpen($SkinEngineDll)
	DllCall($hDll, "none","AboutSkinCrafter")
	CloseWnd()
	DllCall($hDll, "int", "InitLicenKeys", "int", MakeUniCodeStr("0"), "int", MakeUniCodeStr("SKINCRAFTER"), "int", MakeUniCodeStr("SKINCRAFTER.COM"), "int", MakeUniCodeStr("support@skincrafter.com"), "int", MakeUniCodeStr("DEMOSKINCRAFTERLICENCE"))
	CloseWnd()
	DllCall($hDll, "int", "DefineLanguage", "int", 0)
	CloseWnd()
	DllCall($hDll, "int", "InitDecoration", "int", 1)
	CloseWnd()
	DllCall($hDll, "int", "LoadSkinFromFile", "int", MakeUniCodeStr($SkinDll))
	CloseWnd()
	DllCall($hDll, "int", "ApplySkin")
	CloseWnd()
	Return $hDll
EndFunc

Func MakeUniCodeStr($Arg00)
	Local $StrLen = StringLen($Arg00)
	Local $Var0002 = DllCall("oleaut32.dll", "int", "SysAllocStringLen", "int", 0, "int", $StrLen)
	
	DllCall("kernel32.dll", "int", "MultiByteToWideChar", "int", 0, "int", 0, "str", $Arg00, "int", $StrLen, "ptr", $Var0002[0], "int", $StrLen)
	
	Return $Var0002[0]
EndFunc

Func CloseWnd($WndClassName = 'classname=#32770')
	$MatchMode = Opt("WinTitleMatchMode", 4)
	
	Local $hWnd = WinGetHandle($WndClassName)
	ControlHide($hWnd, '', '')
	
	WinKill($hWnd)
	Opt("WinTitleMatchMode", $MatchMode)
EndFunc

Func UnLoadSkinCrafter()
	DllCall($gVarRetLoadSkinCrafter, "int", "DeInitDecoration")
	DllCall($gVarRetLoadSkinCrafter, "int", "RemoveSkin")
	DllClose($gVarRetLoadSkinCrafter)
EndFunc



Global $gDivKey2 = Random($gVarRndMinA & $gVarRndMinB, $gVarRndMinB & $gVarRndMaxA, $gVarRndFlag), $gObfuFileData
Global $gVar0184 = Random($gVarRndC & $gVarRndC2, $gVarRndD1 & $gVarRndC2, 1), $gSleepTimeMs, $gOutPFileName, $gVar0187
Global $gEncodeStartStop, $gOutpFileNDrive, $gOutpFileNPath, $gOutpFileNName, $gOutpFileNExt, $gSleepTime = IniRead($ScriptDir & "\EnCodeIt.ini", "SleepTime", "Sleep", 10)
Global $gRndCharPrefix = Chr(Random(65, 90, 1)), $g4ByteRandHex = Fn4ByteRandHex()
Global $gLogFile_EnCodeIt = ";~===========================================================================================" & $CRLF & ";~                                          EnCodeIt" & $CRLF & ";~                                             By" & $CRLF & ";~                                          SmOke_N" & $CRLF & ";~===========================================================================================" & $CRLF
Global $gLogFile_Functions = $CRLF & ";~===========================================================================================" & $CRLF & ";~                                          Functions" & $CRLF & ";~                                       " & FnGetMonthName(@MON) & " " & @MDAY & ", " & @YEAR & $CRLF & ";~                                           " & FnGetTime() & $CRLF & ";~===========================================================================================" & $CRLF & ";     No.     |              Function             |            EnCoded Function" & $CRLF & ";~===========================================================================================" & $CRLF & $CRLF
Global $gLogFile_Vars = $CRLF & ";~===========================================================================================" & $CRLF & ";~                                          Variables" & $CRLF & ";~                                       " & FnGetMonthName(@MON) & " " & @MDAY & ", " & @YEAR & $CRLF & ";~                                           " & FnGetTime() & $CRLF & ";~===========================================================================================" & $CRLF & ";     No.     |              Variables            |            EnCoded Variables" & $CRLF & ";~===========================================================================================" & $CRLF & $CRLF
$GUIhWnd = GUICreate("EnCodeIt v2.0 By SmOke_N", 355, 150)
$gVar0187 = $GUIhWnd

$GUIhMenu = GUICtrlCreateMenu("&File")
$MenuItem_Open = GUICtrlCreateMenuitem("O&pen", $GUIhMenu, 1)
GUICtrlCreateMenuitem('', $GUIhMenu, 2)
$MenuItem_Exit = GUICtrlCreateMenuitem("&Exit", $GUIhMenu, 3)
$MenuItem_Options = GUICtrlCreateMenu("&Options")
$MenuItem_Compiler_Location = GUICtrlCreateMenu("Compiler Location", $MenuItem_Options, 1)
GUICtrlCreateMenuitem('', $MenuItem_Options, 2)
$M8D5A722A6C5B4FC9 = GUICtrlCreateMenuitem(IniRead($ScriptDir & "\EnCodeIt.ini", "Compiler", "Aut2Exe", ''), $MenuItem_Compiler_Location)
GUICtrlCreateMenuitem('', $MenuItem_Compiler_Location, 2)
$MenuItem_Change = GUICtrlCreateMenuitem("Change", $MenuItem_Compiler_Location, 3)
$MenuItem_DecompilerLocation = GUICtrlCreateMenu("Decompiler Location", $MenuItem_Options, 3)
$M8D5A722F6C5A4F79 = GUICtrlCreateMenuitem(IniRead($ScriptDir & "\EnCodeIt.ini", "Decompiler", "Exe2Aut", ''), $MenuItem_DecompilerLocation)
GUICtrlCreateMenuitem('', $MenuItem_DecompilerLocation, 2)
$M8D5A7223685B4F29 = GUICtrlCreateMenuitem("Change", $MenuItem_DecompilerLocation, 3)
GUICtrlCreateMenuitem('', $MenuItem_Options, 4)
$M8D5A72296C504F29 = GUICtrlCreateMenu("PC Relief", $MenuItem_Options, 5)
$M8D5A728F4C5B4F29 = GUICtrlCreateMenuitem("Current Sleep Average Is: " & $gSleepTime & " Milliseconds", $M8D5A72296C504F29, 1)
GUICtrlCreateMenuitem('', $M8D5A72296C504F29, 2)
$M8D5A722F6C5B5F29 = GUICtrlCreateMenuitem("Change", $M8D5A72296C504F29, 3)
$M8D5A72266C5B4A29 = GUICtrlCreateLabel("Input the complete path for the file to Encode...", 10, 15, 300, 20, 0x0001)
$GUI_Txt_Input = GUICtrlCreateInput('', 10, 35, 300, 20)

If $cmdline[0] Then GUICtrlSetData($GUI_Txt_Input, $gVarCmdline_1)
$GUILabelStatus = GUICtrlCreateLabel('', 20, 90, 150, 20)
$lblEncoding = GUICtrlCreateLabel('', 160, 90, 80, 20)
$GUILabelProcentComplett = GUICtrlCreateLabel('', 290, 90, 50, 20, 0x0001)
$M8D5A792FFC5B4F29 = GUICtrlCreateButton("www", 310, 35, 30, 20)
GUICtrlSetFont($M8D5A792FFC5B4F29, 10, 400, -1, "Wingdings")
$GUIButton_StartEncoding = GUICtrlCreateButton("Start Encoding", 130, 60, 80, 30)
$GUIProgressBar = GUICtrlCreateProgress(20, 110, 310, 15)
ControlHide($GUIhWnd, '', $GUIProgressBar)
GUISetOnEvent(-3, 'FnQuit_EnCodeIt', $GUIhWnd)
GUICtrlSetOnEvent($MenuItem_Open, 'FnWndHandler')
GUICtrlSetOnEvent($MenuItem_Exit, 'FnWndHandler')
GUICtrlSetOnEvent($MenuItem_Change, 'FnWndHandler')
GUICtrlSetOnEvent($M8D5A7223685B4F29, 'FnWndHandler')
GUICtrlSetOnEvent($M8D5A792FFC5B4F29, 'FnWndHandler')
GUICtrlSetOnEvent($GUIButton_StartEncoding, 'FnWndHandler')
GUICtrlSetOnEvent($M8D5A722F6C5B5F29, 'FnWndHandler')
FnWndShowHide($GUIhWnd)
GUISetState()
CloseWnd()
Fn0078()

If UBound($cmdline) - 1 > 1 Then
	FileInstall('C:\Documents and Settings\EnCodeItInfo\Restart_EnCoded1.au3', $ScriptDir & "\EnCodeItInfo\Backup.bak", 1)
	
	If $gSleepTime > 0 And $gSleepTime < 11 Then
		$gSleepTimeMs = $gSleepTime * 100
	ElseIf $gSleepTime = 0 Or $gSleepTime = '' Then
		$gSleepTimeMs = 10000
	Else
		$gSleepTimeMs = 1000
	EndIf

	Local $Var0004, $Var0005 = TimerInit(), $Var0006, $Var0007, $Var0008
	Local $Var0009 = UBound($cmdline) - 1, $Var000A
	FnWndShowHide($GUIhWnd, 1)
	ControlShow($GUIhWnd, '', $GUIProgressBar)
	AdlibEnable('FnShowEncodingLabel', $gSleepTimeMs)
	For $Var00B1 = 1 To $Var0009
		Local $Var000B, $Var000C, $Var000D, $Var000E
		Local $Var000F = SplitFileName($cmdline[$Var00B1], $Var000B, $Var000C, $Var000D, $Var000E)
		GUICtrlSetData($GUILabelStatus, $Var000D & $Var000E)
		GUICtrlSetData($GUI_Txt_Input, '')
		GUICtrlSetData($GUI_Txt_Input, $cmdline[$Var00B1])
		FnCompileDecompile($cmdline[$Var00B1])
		Sleep($gSleepTime)
		
		$gObfuFileData = FnRemoveComments($ScriptDir & "\EnCodeItInfo\Compiled.au3")
		Sleep($gSleepTime)
		
		$gObfuFileData = FnRemoveLineBreaks($gObfuFileData)
		Sleep($gSleepTime)
		
		FnScanFuncNames($gObfuFileData)
		Sleep($gSleepTime)
		
		$gObfuFileData = FnAddEnCodeItFuncs($gObfuFileData, $gVar0184)
		Sleep($gSleepTime)
		
		FnGetVarName($gObfuFileData)
		Sleep($gSleepTime)
		
		$gObfuFileData = FnReplaceVarStrFuncName($gObfuFileData)
		$gObfuFileData = FnJoin($gObfuFileData)
		Sleep($gSleepTime)
		
		$Var000A += UBound(StringSplit(StringStripCR($gObfuFileData), $lf)) - 1
		$gOutPFileName &= FnOutpAndCleanup($cmdline[$Var00B1], $gObfuFileData) & @CR
		FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "File " & $Var00B1 & " of " & $Var0009 & " Done", (($Var00B1 / $Var0009) * 100))
		Sleep(1000)
	Next
	AdlibDisable()
	$Var0006 = TimerDiff($Var0005) / 1000
	MsgBox(64, "Info:", "You're Encoded files are here:" & $CR & $gOutPFileName & $CR & "It took EnCodeIt " & StringFormat("%i minute(s) and %i second(s)", $Var0006 / 60, Mod($Var0006, 60)) & " to obfuscate " & $Var0009 & " Files and " & $Var000A & " lines.")
	Local $Var0010 = Run(FileGetShortName($ScriptFullPath) & " /AutoIt3ExecuteScript " & FileGetShortName($ScriptDir & "\EnCodeItInfo\Backup.bak") & " " & $AutoItPID & " " & FileGetShortName($DocumentsCommonDir), '', @SW_HIDE)
	ProcessWait($Var0010)
	FileDelete($ScriptDir & "\EnCodeItInfo\Backup.bak")
	Exit
EndIf
While 1
	Sleep(100000)
WEnd
Func FnWndHandler()
	Switch @GUI_CtrlId
		
		Case $MenuItem_Open, $M8D5A792FFC5B4F29
			Local $Var0011 = FileOpenDialog("Choose a file to encode", @WorkingDir, "All Files (*.*)")
			If Not @error Then GUICtrlSetData($GUI_Txt_Input, $Var0011)
		
		Case $MenuItem_Exit
			FnQuit_EnCodeIt()
		
		Case $MenuItem_Change
			Local $Var0012 = FileOpenDialog("Change Compiler Aut2Exe", @WorkingDir, "Exe Files (*.exe)", "Aut2Exe.exe")
			If Not @error Then
				IniWrite($ScriptDir & "\EnCodeIt.ini", "Compiler", "Aut2Exe", $Var0012)
				GUICtrlSetData($M8D5A722A6C5B4FC9, $Var0012)
			EndIf
		
		Case $M8D5A7223685B4F29
			Local $Var0013 = FileOpenDialog("Change Decompiler Aut2Exe", @WorkingDir, "Exe Files (*.exe)", "Exe2Aut.exe")
			If Not @error Then
				IniWrite($ScriptDir & "\EnCodeIt.ini", "Decompiler", "Exe2Aut", $Var0013)
				GUICtrlSetData($M8D5A722F6C5A4F79, $Var0012)
			EndIf
		
		Case $M8D5A722F6C5B5F29
			Local $Var0014 = InputBox("Change Sleep Time", " ", "<Change Sleep Here>", '', 210, 120)
			If StringIsInt($Var0014) Then
				IniWrite($ScriptDir & "\EnCodeIt.ini", "SleepTime", "Sleep", Int($Var0014))
				$gSleepTime = Int($Var0014)
				GUICtrlSetData($M8D5A728F4C5B4F29, "Current Sleep Average Is: " & Int($Var0014) & " Milliseconds")
			EndIf
		
	Case $GUIButton_StartEncoding
		
			If UBound($cmdline) - 1 <= 1 Then
				
				;Set SleepTime
				If $gSleepTime > 0 And $gSleepTime < 11 Then
					$gSleepTimeMs = $gSleepTime * 100
				ElseIf $gSleepTime = 0 Or $gSleepTime = '' Then
					$gSleepTimeMs = 10000
				Else
					$gSleepTimeMs = 1000
				EndIf
				
				;Check 
				$gVarCmdline_1 = GUICtrlRead($GUI_Txt_Input)
				If Not FileExists($gVarCmdline_1) Then
					MsgBox(16, "Error", "Please check to ensure that the file path is correct to the file you wish to encode.")
					Return ''
				EndIf
				
				Local $Var0004, $TimeStamp = TimerInit(), $VarTimerDiff, $VarMsgBoxRet, $Var0008
				
				FnWndShowHide($GUIhWnd, 1)
				
				ControlShow($GUIhWnd, '', $GUIProgressBar)
				AdlibEnable('FnShowEncodingLabel', $gSleepTimeMs)

;Comment out because ExeToAut Decompiler don't support CommandlineArguments
;~ 				FnCompileDecompile($gVarCmdline_1)
;~ 				Sleep($gSleepTime)
;~ 				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 1 of 9 Done", 5)
				
;				$gObfuFileData = FnRemoveComments($ScriptDir & "\EnCodeItInfo\Compiled.au3")
				$gObfuFileData = FnRemoveComments($gVarCmdline_1)
				Sleep($gSleepTime)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 2 of 9 Done", 10)
				MsgBox(0,"FilteredFileData",$gObfuFileData)
				
				$gObfuFileData = FnRemoveLineBreaks($gObfuFileData)
				Sleep($gSleepTime)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 3 of 9 Done", 20)
				MsgBox(0,"FnRemoveLineBreaks",$gObfuFileData)
				
				FnScanFuncNames($gObfuFileData)
				Sleep($gSleepTime)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 4 of 9 Done", 30)
				MsgBox(0,"FnScanFuncNames",$gObfuFileData)
				
				
				$gObfuFileData = FnAddEnCodeItFuncs($gObfuFileData, $gVar0184)
				Sleep($gSleepTime)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 5 of 9 Done", 40)
				MsgBox(0,"FnAddEnCodeItFuncs " & $gVar0184,$gObfuFileData)
				
				FnGetVarName($gObfuFileData)
				Sleep($gSleepTime)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 6 of 9 Done", 60)
				MsgBox(0,"FnGetVarName " ,$gObfuFileData)
				
				
				$gObfuFileData = FnReplaceVarStrFuncName($gObfuFileData)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 7 of 9 Done", 85)
				MsgBox(0,"FnReplaceVarStrFuncName " ,$gObfuFileData)
				
				
				$gObfuFileData = FnJoin($gObfuFileData)
				Sleep($gSleepTime)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Step 8 of 9 Done", 90)
				MsgBox(0,"FnJoin " ,$gObfuFileData)				
				
				$VarOutPLines = StringSplit(StringStripCR($gObfuFileData), $lf)
				$gOutPFileName = FnOutpAndCleanup($gVarCmdline_1, $gObfuFileData)
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, "Finished Encoding", 100)
				
				
				AdlibDisable()
				
				$VarTimerDiff = TimerDiff($TimeStamp) / 1000
				$VarMsgBoxRet = MsgBox(68, "Info:", "You're Encoded file is here:" & $CR & $gOutPFileName & $CR & $CR & "It took EnCodeIt " & StringFormat("%i minute(s) and %i second(s)", $VarTimerDiff / 60, Mod($VarTimerDiff, 60)) & " to obfuscate " & $VarOutPLines[0] & " lines." & $CR & $CR & "Would you like to open the file now?")
				If $VarMsgBoxRet = 6 Then
					Run(@ComSpec & " /c """ & $gOutPFileName & """", '', @SW_HIDE)
				EndIf
				FnWndShowHide($GUIhWnd)
				ControlHide($GUIhWnd, '', $GUIProgressBar)
				GUICtrlSetData($lblEncoding, '')
				$gEncodeStartStop = False
				FnSetStatus($GUIProgressBar, $GUILabelStatus, $GUILabelProcentComplett, '', '')
				
				$gRndCharPrefix = Chr(Random(65, 90, 1))
				
				
			EndIf
	EndSwitch
EndFunc

Func FnWndShowHide($Wnd, $ArgOpt01 = 0)
	Local $VarOldPos = WinGetPos($Wnd)
	If Not $ArgOpt01 And IsArray($VarOldPos) Then
	   ;                  X position   , Y position   , Width        , Height 
		WinMove($Wnd, '', $VarOldPos[0], $VarOldPos[1], $VarOldPos[2], $VarOldPos[3] - 35)
		
	ElseIf IsArray($VarOldPos) Then
		WinMove($Wnd, '', $VarOldPos[0], $VarOldPos[1], $VarOldPos[2], $VarOldPos[3] + 35)
	EndIf
EndFunc
Func FnShowEncodingLabel()
	$gEncodeStartStop = Not $gEncodeStartStop
	If $gEncodeStartStop Then
		GUICtrlSetData($lblEncoding, "...Encoding...")
	Else
		GUICtrlSetData($lblEncoding, '')
	EndIf
	
	If IsArray($EnCodeItInfo_Dat) Then
		For $Var00B1 = 1 To UBound($EnCodeItInfo_Dat) - 1
			If $EnCodeItInfo_Dat[$Var00B1] = "PausingEnCodeItUsage" Then
				MsgBox(16, "Warning!", "All usage has been temporarily suspened, this could be due to an update.", 5)
				Exit
			EndIf
		Next
	EndIf
EndFunc
Func FnSetStatus($Arg00, $Arg01, $Arg02, $Arg03, $Arg04)
	GUICtrlSetData($Arg00, $Arg04 + 10)
	GUICtrlSetData($Arg01, $Arg03)
	If $Arg04 <> '' Then
		GUICtrlSetData($Arg02, Round($Arg04, 2) & "%")
	Else
		GUICtrlSetData($Arg02, '')
	EndIf
	Return 1
EndFunc
Func FnOutpAndCleanup($Arg00, $Arg01)
	Local $Var001B, $Var001C, $Var001D, $SeqNum = 1
	$Var001B = SplitFileName($Arg00, $gOutpFileNDrive, $gOutpFileNPath, $gOutpFileNName, $gOutpFileNExt)
	
	;Scan for exisiting File
	While FileExists($gOutpFileNDrive & $gOutpFileNPath & $gOutpFileNName & "_EnCoded" & $SeqNum & ".au3")
		$SeqNum += 1
	WEnd
	
	Local $VarNewFileName = $gOutpFileNDrive & $gOutpFileNPath & $gOutpFileNName & "_EnCoded" & $SeqNum & ".au3"
	Local $VarhFile = FileOpen($VarNewFileName, 2)
	Local $VarhFileLog = FileOpen($gOutpFileNDrive & $gOutpFileNPath & $gOutpFileNName & "_EnCoded" & $SeqNum & ".log", 2)
	FileWrite($VarNewFileName, $Arg01)
	
	FileWrite($VarhFileLog, $gLogFile_EnCodeIt)
	FileWrite($VarhFileLog, $gLogFile_Functions)
	FileWrite($VarhFileLog, $gLogFile_Vars)
	
	FileClose($VarhFile)
	FileClose($VarhFileLog)
	
	FileDelete($ScriptDir & "\EnCodeItInfo\Compiled.au3")
	FileDelete($ScriptDir & "\EnCodeItInfo\Text.ini")
	FileDelete($ScriptDir & "\EnCodeItInfo\Var.ini")
	FileDelete($ScriptDir & "\EnCodeItInfo\Func.ini")
	FileDelete($ScriptDir & "\EnCodeItInfo\FuncReplace.ini")
	
	Return $VarNewFileName
EndFunc
Func FnJoin($Arg00)
	Local $Lines = StringSplit(StringStripCR($Arg00), $lf), $Var001D
	For $i = 1 To UBound($Lines) - 1
		$Var001D &= StringStripWS($Lines[$i ], 7) & $CRLF
	Next
	Return StringTrimRight($Var001D, StringLen($CRLF))
EndFunc

Func FnGetVarName($Arg00)
	Local $VarFuncNames, $Lines, $Var0026 = $gRndCharPrefix
	Local $Var0027 = StringRegExp ($Arg00, "(?i:\$)([a-zA-Z0-9/_]+)", 3)
	If Not IsArray($Var0027) Then Return SetError(1, 0, 0)
	For $Var00B1 = 0 To UBound($Var0027) - 1
		$VarFuncNames &= $Var0027[$Var00B1] & Chr(01)
	Next
	
	$VarFuncNames = StringSplit(StringTrimRight($VarFuncNames, 1), Chr(01))
	$VarFuncNames = Fn007C($VarFuncNames)
	Sleep($gSleepTime)
	$Lines = Fn0069($VarFuncNames)
	Sleep($gSleepTime)
	
	For $i = 1 To UBound($Lines) - 1
		If $Lines[$i] = "cmdline" Or $Lines[$i] = "cmdlineraw" Then ContinueLoop
		$Lines[$i] = $gRndCharPrefix & $Lines[$i]
	Next
	
	Sleep($gSleepTime)
	
	FnVarNamesToVar_ini($VarFuncNames, $Lines)
	
	Sleep($gSleepTime)
	
	Return 1
EndFunc
Func Fn0069(ByRef $ArgRef00)
	Local $Var0028, $Var0029, $Var002A = $g4ByteRandHex
	Local $Arr002B[UBound($ArgRef00) ]
	$Arr002B[0] = UBound($ArgRef00) - 1
	For $Var004C = 1 To $Arr002B[0]
		If $ArgRef00[$Var004C] = "cmdline" Then
			$Arr002B[$Var004C] = "cmdline"
			ContinueLoop
		EndIf
		If $ArgRef00[$Var004C] = "cmdlineraw" Then
			$Arr002B[$Var004C] = "cmdlineraw"
			ContinueLoop
		EndIf
		$Arr002B[$Var004C] = $Var002A
		Do
			Local $Var002C = 0
			Local $Var002D = Random(2, StringLen($Var002A) - 1, 1)
			$Arr002B[$Var004C] = StringLeft($Arr002B[$Var004C], $Var002D - 1) & Hex(Random(0, 15, 1), 1) & StringTrimLeft($Arr002B[$Var004C], $Var002D)
			For $Var0075 = 1 To $Var004C - 1
				If $Arr002B[$Var0075] = $Arr002B[$Var004C] Then
					$Arr002B[$Var0075] = $Arr002B[$Var004C]
					$Var002C = 1
					ExitLoop
				EndIf
			Next
		Until Not $Var002C
		Sleep($gSleepTime)
	Next
	Return $Arr002B
EndFunc
Func FnScanFuncNames($Arg00)
	Local $VarFuncNames, $Lines
	Local $VarFuncNames = GetFuncNames($Arg00)
	If Not IsArray($VarFuncNames) Then Return 1 ;Doesn't contain Functions
		
;	Add internal EnCodeIt FunctionNames to $intEnCodeIt_FuncNames
	Local $intEnCodeIt_FuncNames = StringSplit("_EnCodeIt_UStr,_EnCodeIt_UStr2,_EnCodeIt_OStr", ",")
	For $i = 1 To UBound($intEnCodeIt_FuncNames) - 1
		FnAddElementToArray($VarFuncNames, $intEnCodeIt_FuncNames[$i])
	Next
	
	$VarFuncNames = Fn007C($VarFuncNames)
	Sleep($gSleepTime)
	
	$Lines = Fn006B($VarFuncNames)
	Sleep($gSleepTime)
	For $i = 1 To UBound($Lines) - 1
		If $Lines[$i] = 'Fn009F' Or $Lines[$i] = "OnAutoItStart" Then ContinueLoop
		$Lines[$i] = "_" & $gRndCharPrefix & $Lines[$i]
	Next
	FnWriteFunc_ini($VarFuncNames, $Lines)
	Sleep($gSleepTime)
	Return 1
EndFunc
Func Fn006B(ByRef $ArgRef00)
	Local $Var0028, $Var0029, $Var002A = $g4ByteRandHex
	Local $Arr002B[UBound($ArgRef00) ]
	$Arr002B[0] = UBound($ArgRef00) - 1
	For $Var004C = 1 To $Arr002B[0]
		If $ArgRef00[$Var004C] = 'Fn009F' Then
			$Arr002B[$Var004C] = 'Fn009F'
			ContinueLoop
		EndIf
		If $ArgRef00[$Var004C] = "OnAutoItStart" Then
			$Arr002B[$Var004C] = "OnAutoItStart"
			ContinueLoop
		EndIf
		$Arr002B[$Var004C] = $Var002A
		Do
			Local $Var002C = 0
			Local $Var002D = Random(2, StringLen($Var002A) - 1, 1)
			$Arr002B[$Var004C] = StringLeft($Arr002B[$Var004C], $Var002D - 1) & Hex(Random(0, 15, 1), 1) & StringTrimLeft($Arr002B[$Var004C], $Var002D)
			For $Var0075 = 1 To $Var004C - 1
				If $Arr002B[$Var0075] = $Arr002B[$Var004C] Then
					$Var002C = 1
					ExitLoop
				EndIf
			Next
		Until Not $Var002C
		Sleep($gSleepTime)
	Next
	Return $Arr002B
EndFunc
Func GetFuncNames($Arg00)
	Local $Lines = StringSplit(StringStripCR($Arg00), @LF)

	Local $ArrFuncNames[$Lines[0] + 1], $VarFuncCount, $Var003B
	For $i = 1 To $Lines[0]
		
		;Is Line a Function?
		If StringLeft(StringStripWS($Lines[$i], 8), 4) = "func" Then
			;Does it have '(' ?
			If Not StringInStr($Lines[$i], "(") Then ContinueLoop
			$VarFuncCount += 1
			$VarFuncSplit = StringSplit($Lines[$i], "(")
			$ArrFuncNames[$VarFuncCount] = StringTrimLeft(StringStripWS($VarFuncSplit[1], 8), 4)
		EndIf
		
	Next
	
	ReDim $ArrFuncNames[$VarFuncCount + 1]
	Return $ArrFuncNames
EndFunc

Func FnAddEnCodeItFuncs($ScriptData, $DivKey)
	Local $Var003C = Random(30, 190, 1), $Var003D
	Local $VariEnCodeItLevel5Call = FnMake_EnCodeIt_UStr2_Call($DivKey)
	
	Local $Var001D = Fn006E($ScriptData), $Var0040
	Fn007F($Var001D)
	Sleep($gSleepTime)
	Local $Arr0041[UBound($Var001D) ], $Var0040[UBound($Var001D) ]
	
	For $Var004C = 1 To UBound($Var001D) - 1
		If StringLen($Var001D[$Var004C]) > 2 And StringLeft($Var001D[$Var004C], 1) = """" And StringRight($Var001D[$Var004C], 1) = """" Then
			If StringInStr($Var001D[$Var004C], """""") Then
				$Var0040[$Var004C] = StringReplace($Var001D[$Var004C], """""", """")
			Else
				$Var0040[$Var004C] = $Var001D[$Var004C]
			EndIf
		ElseIf StringLen($Var001D[$Var004C]) > 2 And StringLeft($Var001D[$Var004C], 1) = "'" And StringRight($Var001D[$Var004C], 1) = "'" Then
			If StringInStr($Var001D[$Var004C], "''") Then
				$Var0040[$Var004C] = StringReplace($Var001D[$Var004C], "''", "'")
			Else
				$Var0040[$Var004C] = $Var001D[$Var004C]
			EndIf
		Else
			$Var0040[$Var004C] = $Var001D[$Var004C]
		EndIf
		$Arr0041[$Var004C] = "iEnCodeItEnS" & $Var004C
		$Var003D &= "Global Const $iEnCodeItEnS" & $Var004C & " = " & "_EnCodeIt_UStr(" & "'" & Fn007D($Var0040[$Var004C], $DivKey) & "'" & ", $iEnCodeItLevel5)" & $CRLF
	Next
	
	$ScriptData = StringTrimRight($Var003D, StringLen($CRLF)) & $CRLF & $ScriptData
	Sleep($gSleepTime)
	
	Fn008B($Var001D, $Arr0041)
	Sleep($gSleepTime)
	
	$ScriptData = "Global Const $iEnCodeItLevel5 = " & $VariEnCodeItLevel5Call & $CRLF & $ScriptData
	$ScriptData = "Global Const $EnCodeItConstVar = Int(99/3+15*100/4-13^2+81/3-17-245+99/3+15*100/4-13^2+81/3-17)" & $CRLF & $ScriptData
	Sleep($gSleepTime)
	
	$ScriptData &= $CRLF & Fn_EnCodeIt_Funcs_Body()
	
	Return $ScriptData
EndFunc

Func Fn006E($Arg00)
	Local $Var0043, $Arr0044[2] = ['', False], $Var0045 = False, $Var0046, $Var0047
	Local $Var0048 = Fn00AE($ScriptDir & "\EnCodeItInfo\Func.ini", "Funcs")
	$Var0043 = Fn00A1($Arg00)
	If Not @error Then
		$Arg00 = Fn0071($Var0043, $Arg00)
		$Arr0044[1] = True
	EndIf
	For $M845A72266C5B4F29 = 1 To UBound($Var0043, 1) - 1
		IniWrite($ScriptDir & "\EnCodeItInfo\FuncReplace.ini", "FunctionReplace", $Var0043[$M845A72266C5B4F29][1], $Var0043[$M845A72266C5B4F29][0])
	Next
	Local $Var0049, $Var004A
	Local $Var004B = StringReplace($Arg00, "''", "ss")
	$Var004B = StringReplace($Var004B, """""", "dd")
	Local $Var004C = 0, $Var003A = 0, $Var004E, $Var0045
	While StringLen($Arg00) > 0
		$Var004A = StringInStr($Var004B, "'")
		$Var0049 = StringInStr($Var004B, """")
		If Not $Var004A And Not $Var0049 Then ExitLoop
		If $Var0049 > 0 And ($Var004A > $Var0049 Or $Var004A = 0) Then
			$M8DAA72276C5B4F29 = """"
			$Arg00 = StringTrimLeft($Arg00, $Var0049 - 1)
			$Var004B = StringTrimLeft($Var004B, $Var0049 - 1)
		Else
			$M8DAA72276C5B4F29 = "'"
			$Arg00 = StringTrimLeft($Arg00, $Var004A - 1)
			$Var004B = StringTrimLeft($Var004B, $Var004A - 1)
		EndIf
		$M8D8A722F6C5B5F29 = StringInStr($Var004B, $M8DAA72276C5B4F29, 0, 2)
		If $M8D8A722F6C5B5F29 > 0 Then
			$Var0046 = StringLeft($Arg00, $M8D8A722F6C5B5F29)
			If Not StringInStr(Chr(01) & $Var0047, Chr(01) & $Var0046 & Chr(01), 1) Then
				$Var0047 &= $Var0046 & Chr(01)
				For $M8D5A722F0CDB4F29 = 1 To $Var0048[0][0]
					If $Var0048[$M8D5A722F0CDB4F29][0] = StringTrimRight(StringTrimLeft($Var0046, 1), 1) Then
						$Var0045 = True
						$Arg00 = StringTrimLeft($Arg00, $M8D8A722F6C5B5F29)
						$Var004B = StringTrimLeft($Var004B, $M8D8A722F6C5B5F29)
						ExitLoop
					EndIf
				Next
				If Not $Var0045 Then
					$Var004E &= $Var0046 & $lf
					$Arg00 = StringTrimLeft($Arg00, $M8D8A722F6C5B5F29)
					$Var004B = StringTrimLeft($Var004B, $M8D8A722F6C5B5F29)
				EndIf
				$Var0045 = False
			Else
				$Arg00 = StringTrimLeft($Arg00, $M8D8A722F6C5B5F29)
				$Var004B = StringTrimLeft($Var004B, $M8D8A722F6C5B5F29)
			EndIf
		Else
			ExitLoop
		EndIf
	WEnd
	If $Arr0044[1] Then $Var004E = Fn006F($Var0043, $Var004E)
	Return StringSplit(StringTrimRight($Var004E, 1), $lf)
EndFunc
Func Fn006F(ByRef $ArgRef00, $Arg01)
	For $Var004C = UBound($ArgRef00) - 1 To 1 Step - 1
		$Arg01 = StringReplace($Arg01, $ArgRef00[$Var004C][1], $ArgRef00[$Var004C][0], 0, 1)
	Next
	Return $Arg01
EndFunc
Func Fn0070($Arg00, $Arg01, $Arg02)
	Local $Var0050 = Fn0088($Arg00, "\s" & $Arg01, $Arg02)
	If Not @extended And Not IsArray($Var0050) Then Return SetError(1, 0, 0)
	Local $Var0051 = "Somethingreallylongandobnoxiousforreplacementof" & StringStripWS($Arg01, 8)
	$Var0050 = Fn007C($Var0050, 0)
	Local $Arr0052[$Var0050[0] + 1][2]
	For $Var004C = 1 To $Var0050[0]
		$Arr0052[$Var004C][0] = $Var0050[$Var004C]
		$Arr0052[$Var004C][1] = $Var0051 & $Var004C
	Next
	Return $Arr0052
EndFunc
Func Fn0071(ByRef $ArgRef00, $Arg01)
	For $Var004C = 1 To UBound($ArgRef00) - 1
		$Arg01 = StringReplace($Arg01, $ArgRef00[$Var004C][0], $ArgRef00[$Var004C][1], 0, 1)
	Next
	Return $Arg01
EndFunc
Func FnRemoveComments($FileName)
	
	Local $VarLinesFromFile = StringSplit(StringStripCR(FileRead($FileName)), @LF)
	
	Local $I, $Var001D, $VarAposPos, $VarStrPos, $VarCommentPos, $LineFiltered, $VarChar
	
	;Go through all lines in the File
	For $I = 1 To UBound($VarLinesFromFile) - 1
		If StringLeft(StringStripWS($VarLinesFromFile[$I], 7), 3) = "#cs" Then
			Do
				$I += 1
			Until StringLeft(StringStripWS($VarLinesFromFile[$I], 7), 3) = "#ce" Or $I = $VarLinesFromFile[0]
			ContinueLoop
		EndIf
		If StringLeft(StringStripWS($VarLinesFromFile[$I], 7), 7) = "#region" Or StringLeft(StringStripWS($VarLinesFromFile[$I], 7), 10) = "#endregion" Then ContinueLoop
		
		;Replace "" => dd   and   '' =>ss
		$LineFiltered = StringReplace($VarLinesFromFile[$I], "''", "ss")

;Bug?	$LineFiltered = StringReplace($VarLinesFromFile[$I], """""", "dd")
		$LineFiltered = StringReplace($LineFiltered, """""", "dd")
		$Var0080 = 0
		$VarSomePos = 0
		
		While 1
			$VarAposPos = StringInStr($LineFiltered, "'")
			$VarStrPos = StringInStr($LineFiltered, """")
			$VarCommentPos = StringInStr($LineFiltered, ";")
			
			If Not $VarCommentPos Then
				$Var0080 = StringLen($VarLinesFromFile[$I]) + 1
				ExitLoop
			EndIf
			
			$VarSomePos = $VarCommentPos
			If $VarAposPos <> 0 And $VarAposPos < $VarSomePos Then $VarSomePos = $VarAposPos
			If $VarStrPos <> 0 And $VarStrPos < $VarSomePos Then $VarSomePos = $VarStrPos
			
			$VarChar = StringMid($LineFiltered, $VarSomePos, 1)
			$Var0080 += $VarSomePos
			If $VarChar = ";" Then ExitLoop
			
			$LineFiltered = StringTrimLeft($LineFiltered, $VarSomePos)
			$VarSomePos = StringInStr($LineFiltered, $VarChar)
			$LineFiltered = StringTrimLeft($LineFiltered, $VarSomePos)
			$Var0080 += $VarSomePos
		WEnd
		
		If StringStripWS(StringLeft($VarLinesFromFile[$I], $Var0080 - 1), 7) <> '' Then
			$Var001D &= StringLeft($VarLinesFromFile[$I], $Var0080 - 1) & $CRLF
		EndIf
	Next
	
	Return StringTrimRight($Var001D, StringLen($CRLF))
	
EndFunc


Func FnRemoveLineBreaks($Arg00)
	
	Local $Lines = StringSplit(StringStripCR($Arg00), $lf)
	
	Local $RetStr, $i, $Var005E, $Var005F, $Var0045, $Var0059
	
	For $i = 1 To UBound($Lines) - 1
		If StringMid(StringStripWS($Lines[$i], 8), StringLen(StringStripWS($Lines[$i], 8)), 1) = "_" Then
			$Var005F = StringLen($Lines[$i])
			While $Var005F > 1
				If StringMid($Lines[$i], $Var005F, 1) = "_" And StringMid($Lines[$i], $Var005F - 1, 1) = " " Then
					$Var0059 = StringLeft($Lines[$i], $Var005F - 1)
					If StringLeft($Var0059, 1) = Chr(32) Or StringLeft($Var0059, 1) = Chr(09) Then
						While StringLeft($Var0059, 1) = Chr(32) Or StringLeft($Var0059, 1) = Chr(09)
							If StringLeft($Var0059, 1) = Chr(32) Then
								$Var0059 = StringTrimLeft($Var0059, 1)
							ElseIf StringLeft($Var0059, 1) = Chr(09) Then
								$Var0059 = StringTrimLeft($Var0059, StringLen(Chr(09)))
							EndIf
						WEnd
					EndIf
					$RetStr &= $Var0059
					$Var0045 = True
					ExitLoop
				EndIf
				$Var005F -= 1
			WEnd
		EndIf
		
		If $Var0045 Then
			If StringMid(StringStripWS($Lines[$i + 1], 8), StringLen(StringStripWS($Lines[$i + 1], 8)), 1) = "_" Then ContinueLoop
			$i += 1
			If StringLeft($Lines[$i], 1) = Chr(32) Or StringLeft($Lines[$i], 1) = Chr(09) Then
				While StringLeft($Lines[$i], 1) = Chr(32) Or StringLeft($Lines[$i], 1) = Chr(09)
					If StringLeft($Lines[$i], 1) = Chr(32) Then
						$Lines[$i] = StringTrimLeft($Lines[$i], 1)
					ElseIf StringLeft($Lines[$i], 1) = Chr(09) Then
						$Lines[$i] = StringTrimLeft($Lines[$i], StringLen(Chr(09)))
					EndIf
				WEnd
			EndIf
			$RetStr &= $Lines[$i] & $CRLF
		Else
			$RetStr &= $Lines[$i] & $CRLF
		EndIf
		$Var0045 = False
	Next
	
	Return StringTrimRight($RetStr, StringLen($CRLF))
	
EndFunc
Func Fn0074($Arg00, ByRef $ArgRef01)
	$ArgRef01 = StringSplit(StringStripCR(FileRead($Arg00)), $lf)
	Return 1
EndFunc
Func FnCompileDecompile($Arg00)
	$Arg00 = $Arg00
	Local $VarDecompiler = IniRead($ScriptDir & "\EnCodeIt.ini", "Decompiler", "Exe2Aut", '')
	Local $VarCompiler = IniRead($ScriptDir & "\EnCodeIt.ini", "Compiler", "Aut2Exe", '')
	Local $VarCompiled = $ScriptDir & "\EnCodeItInfo\Compiled."
	
MsgBox(0,"Executing Compiler with",$VarCompiler & " /in """ & $Arg00 & """ /out """ & $VarCompiled & "exe" & """ /pass ""A""")
	RunWait($VarCompiler & " /in """ & $Arg00 & """ /out """ & $VarCompiled & "exe" & """ /pass ""A""")
MsgBox(0,"Executing DeCompiler with",$VarDecompiler & " /in """ & $VarCompiled & "exe" & """ /out """ & $VarCompiled & "au3" & """ /pass ""A""")
	RunWait($VarDecompiler & " /in """ & $VarCompiled & "exe" & """ /out """ & $VarCompiled & "au3" & """ /pass ""A""")
	FileDelete($VarCompiled & "exe")
EndFunc

Func InitFiles()
  ; Delete	Note: ScriptDir: Directory containing the running script.
	If FileExists($ScriptDir & "\EnCodeItInfo\_Revamp191.dll") Then FileDelete($ScriptDir & "\EnCodeItInfo\_Revamp191.dll")
	If FileExists($ScriptDir & "\EnCodeItInfo\Tranquill.skf") Then FileDelete($ScriptDir & "\EnCodeItInfo\Tranquill.skf")
		
  ; Make 'EnCodeItInfo\' Dir 
	If Not FileExists($ScriptDir & "\EnCodeItInfo") Then
 	 ;  mkdir & Move "c:\EnCodeItInfo" -> "E:\test\EnCodeItInfo"	
		DirCreate($HomeDrive & "\EnCodeItInfo")
		FileSetAttrib($HomeDrive & "\EnCodeItInfo", "+H", 1)
		DirMove($HomeDrive & "\EnCodeItInfo", $ScriptDir & "\EnCodeItInfo", 1)
	EndIf
	
	
	If Not FileExists($ScriptDir & "\EnCodeItInfo\SkinCrafter.dll") Then FileInstall('C:\Documents and Settings\EnCodeItInfo\SkinCrafter.dll', $ScriptDir & "\EnCodeItInfo\SkinCrafter.dll", 1)
	If Not FileExists($ScriptDir & "\EnCodeItInfo\200699.dll") Then FileInstall('C:\Documents and Settings\EnCodeItInfo\200699.dll', $ScriptDir & "\EnCodeItInfo\200699.dll", 1)
	
   ; Start Skincrafter	
	$gVarRetLoadSkinCrafter = LoadSkinCrafter ($ScriptDir & "\EnCodeItInfo\SkinCrafter.dll", $ScriptDir & "\EnCodeItInfo\200699.dll")
	
   ; Run Config
;	If Not FileExists($ScriptDir & "\EnCodeIt.ini") Then FirstTimeConfig()
   
   ;Stupid path name check
;	If Not StringInStr($ScriptFullPath, "EnCodeIt 2.0") Then
;		MsgBox(16, "Error", "EnCodeIt.exe must be ran from within the EnCodeIt 2.0 folder.")
;		Exit
;		EndIf
	
EndFunc

Func FirstTimeConfig()
	Local $Var0065, $Var0063, $Var0062, $Okay
	$Var0065 = MsgBox(64, "Info:", "Welcome to EnCodeIt for AutoIt.au3 Scripts." & $CR & "Let's just take a few seconds to setup file locations.")
	If $Var0065 = 7 Then MsgBox(16, "Error", "GoodBye")
	While 1
		$Var0063 = FileOpenDialog("Find Aut2Exe", $HomeDrive, "Exe Files (*.*)", 3, "Aut2Exe.exe")
		If Not @error Then
			MsgBox(64, "Exe2Aut", "Now find the decompiler.")
			While 1
				$Var0062 = FileOpenDialog("Find Exe2Au3", $HomeDrive, "Exe Files (*.*)", 3, "Exe2Aut.exe")
				If @error Then
					$Var0065 = MsgBox(20, "Error", "There was an error retrieving the file, Would you like to quit?")
					If $Var0065 = 6 Then
						MsgBox(64, "Quitting", "GoodBye")
						Exit
					EndIf
				Else
					$Okay = 1
					ExitLoop
				EndIf
			WEnd
		Else
			$Var0065 = MsgBox(20, "Error", "There was an error retrieving the file, Would you like to quit?")
			If $Var0065 = 6 Then
				MsgBox(64, "Quitting", "GoodBye")
				Exit
			EndIf
		EndIf
		If $Okay Then ExitLoop
	WEnd
		
	IniWrite($ScriptDir & "\EnCodeIt.ini", "Compiler", "Aut2Exe", FileGetShortName($Var0063))
	IniWrite($ScriptDir & "\EnCodeIt.ini", "Decompiler", "Exe2Aut", FileGetShortName($Var0062))
	IniWrite($ScriptDir & "\EnCodeIt.ini", "SleepTime", "Sleep", 10)
EndFunc

Func Fn0078($ArgOpt00 = -1)
	If $ArgOpt00 <> -1 Then
		Local $Var0069 = DllCall("kernel32.dll", "int", "OpenProcess", "int", 0x1F0FFF, "int", False, "int", $ArgOpt00)
		Local $Var006A = DllCall($ScriptDir & "\psapi.dll", "int", "EmptyWorkingSet", "long", $Var0069[0])
		DllCall("kernel32.dll", "int", "CloseHandle", "int", $Var0069[0])
	Else
		Local $Var006A = DllCall("psapi.dll", "int", "EmptyWorkingSet", "long", -1)
	EndIf
	Return $Var006A[0]
EndFunc
Func FnMake_EnCodeIt_UStr2_Call($DivKey)
	Local $VarChar
	For $i = 1 To StringLen($DivKey)
		$VarChar = $VarChar & Chr(Asc(StringMid($DivKey, $i, 1)) + ($gDivKey2 - 11))
	Next
	Return "_EnCodeIt_UStr2(" & "'" & Fn0081($VarChar) & "'" & ", $EnCodeItConstVar)"
EndFunc
Func FnAddElementToArray(ByRef $ArgRef00, $Arg01)
	If Not IsArray($ArgRef00) Then Return SetError(1, 0, 0)
		
	ReDim $ArgRef00[UBound($ArgRef00) + 1]
	$ArgRef00[UBound($ArgRef00) - 1] = $Arg01
	
	Return 1
EndFunc
Func Fn007B($Arg00, $Arg01)
	If IsArray($Arg00) Then Return Fn007C($Arg00)
	If StringRight($Arg00, StringLen($Arg01)) = $Arg01 Then $Arg00 = StringTrimRight($Arg00, StringLen($Arg01))
	Local $Var005E = StringSplit(StringStripCR($Arg00), $Arg01), $Var001D, $Lines
	For $Var004C = 1 To $Var005E[0]
		If StringInStr($Arg01 & $Var001D, $Arg01 & $Var005E[$Var004C] & $Arg01, 1) Or $Var005E[$Var004C] = '' Then ContinueLoop
		$Var001D &= $Var005E[$Var004C] & $Arg01
	Next
	$Lines = StringSplit(StringTrimRight($Var001D, StringLen($Arg01)), $Arg01)
	Fn007F($Lines)
	Return $Lines
EndFunc
Func Fn007C($Arg00, $ArgOpt01 = 1, $ArgOpt02 = '')
	If IsString($Arg00) Then Return Fn007B($Arg00, $ArgOpt02)
		
	If $ArgOpt02 = '' Then $ArgOpt02 = Chr(01)
	
	Local $Var001D, $Lines
	
	For $i = $ArgOpt01 To UBound($Arg00) - 1
		If StringInStr($ArgOpt02 & $Var001D, $ArgOpt02 & $Arg00[$i] & $ArgOpt02, 1) Then ContinueLoop
		If $Arg00[$i] <> '' Then $Var001D &= $Arg00[$i] & $ArgOpt02
	Next
	
	$Lines = StringSplit(StringTrimRight($Var001D, StringLen($ArgOpt02)), $ArgOpt02)
	Fn007F($Lines, $ArgOpt01)
	Return $Lines
EndFunc
Func Fn007D($Arg00, $Arg01)
	Local $Var006C
	$Arg00 = StringTrimLeft(StringTrimRight($Arg00, 1), 1)
	For $M8D5A722F635B4FB9 = 1 To StringLen($Arg00)
		$Var006C = $Var006C & Chr(Asc(StringMid($Arg00, $M8D5A722F635B4FB9, 1)) + $Arg01)
	Next
	Return Fn0081($Var006C)
EndFunc
Func Fn007E(ByRef $ArgRef00, $Arg01, $ArgOpt02 = 1)
	If Not IsArray($ArgRef00) Then Return SetError(1, 0, 0)
	Local $Var0073 = UBound($ArgRef00), $Var004C, $Var0075
	Local $VarFuncNames[$Var0073], $Var0077 = ($Var0073 - $Arg01)
	For $Var004C = $Var0073 - 1 To ($Var0077 + $ArgOpt02) Step - 1
		$VarFuncNames[$Var004C] = $ArgRef00[$Var004C - $Var0077]
	Next
	For $Var0075 = $ArgOpt02 To $Var0077
		$VarFuncNames[$Var0075] = $ArgRef00[ ($Var0073 - 1) - ($Var0077 - $Var0075) ]
	Next
	$ArgRef00 = $VarFuncNames
	Return True
EndFunc
Func Fn007F(ByRef $ArgRef00, $ArgOpt01 = 1)
	For $Var004C = $ArgOpt01 To UBound($ArgRef00) - 2
		Local $Var0078 = $Var004C, $Var0079
		For $Var0075 = $Var004C + 1 To UBound($ArgRef00) - 1
			If StringLen($ArgRef00[$Var0075]) > StringLen($ArgRef00[$Var0078]) Then $Var0078 = $Var0075
			If @error Then SetError(1, 0, False)
		Next
		If $Var0078 <> $Var004C Then
			Local $Var0079 = $ArgRef00[$Var004C]
			$ArgRef00[$Var004C] = $ArgRef00[$Var0078]
			$ArgRef00[$Var0078] = $Var0079
		EndIf
	Next
	Return True
EndFunc
Func Fn0080($Arg00)
;	Return BinaryString("0x" & $Arg00)
	Return Binary("0x" & $Arg00)
EndFunc
Func Fn0081($Arg00)
;	Return Hex(BinaryString($Arg00))
	Return Hex(Binary($Arg00))
EndFunc
Func FnQuit_EnCodeIt()
	GUIDelete(HWnd($gVar0187))
	DllCall($gVarRetLoadSkinCrafter, "int", "DeInitDecoration")
	DllCall($gVarRetLoadSkinCrafter, "int", "RemoveSkin")
	DllClose($gVarRetLoadSkinCrafter)
	Exit 0
EndFunc
Func Unload_EnCodeIt()
	If MsgBox(36, "Exit", "Are you sure you would like to exit EnCodeIt?") = 6 Then
		GUIDelete(HWnd($gVar0187))
		DllCall($gVarRetLoadSkinCrafter, "int", "DeInitDecoration")
		DllCall($gVarRetLoadSkinCrafter, "int", "RemoveSkin")
		DllClose($gVarRetLoadSkinCrafter)
		
		FileDelete($ScriptDir & "\EnCodeItInfo\Compiled.au3")
		FileDelete($ScriptDir & "\EnCodeItInfo\Text.ini")
		FileDelete($ScriptDir & "\EnCodeItInfo\Var.ini")
		FileDelete($ScriptDir & "\EnCodeItInfo\Func.ini")
		FileDelete($ScriptDir & "\EnCodeItInfo\FuncReplace.ini")
		FileDelete($ScriptDir & "\EnCodeItInfo\Backup.bak")
		
		Exit
	EndIf
EndFunc
Func Fn0084($Arg00, $ArgOpt01 = 0, $ArgOpt02 = 0)
	$ArgOpt01 = StringInStr($Arg00, $ArgOpt01) + StringLen($ArgOpt01)
	Return StringMid($Arg00, $ArgOpt01, StringInStr($Arg00, $ArgOpt02) - $ArgOpt01)
EndFunc
Func SplitFileName($FinishedPath, ByRef $ArgRef01, ByRef $ArgRef02, ByRef $ArgRef03, ByRef $ArgRef04)
	Local $VarDrive, $Path, $Filename, $Ext, $i, $Var0080, $ArrFileName[5]
	
	;AbsolutePath?
	If StringMid($FinishedPath, 2, 1) = ":" Then
		$VarDrive = StringLeft($FinishedPath, 2)
		$FinishedPath = StringTrimLeft($FinishedPath, 2)
	;NetworkPath?
	ElseIf StringLeft($FinishedPath, 2) = "\\" Then
		$FinishedPath = StringTrimLeft($FinishedPath, 2)
		$Var0080 = StringInStr($FinishedPath, "\")
		If $Var0080 = 0 Then $Var0080 = StringInStr($FinishedPath, "/")
		If $Var0080 = 0 Then
			$VarDrive = "\\" & $FinishedPath
			$FinishedPath = ""
		Else
			$VarDrive = "\\" & StringLeft($FinishedPath, $Var0080 - 1)
			$FinishedPath = StringTrimLeft($FinishedPath, $Var0080 - 1)
		EndIf
	EndIf
	
	;split Path & Name.ext
	For $i = StringLen($FinishedPath) To 0 Step - 1
		If StringMid($FinishedPath, $i, 1) = "\" Or StringMid($FinishedPath, $i, 1) = "/" Then
			$Path = StringLeft($FinishedPath, $i)
			$Filename = StringRight($FinishedPath, StringLen($FinishedPath) - $i)
			ExitLoop
		EndIf
	Next
	
	If StringLen($Path) = 0 Then $Filename = $FinishedPath
	
	For $i = StringLen($Filename) To 0 Step - 1
		If StringMid($Filename, $i, 1) = "." Then
			$Ext = StringRight($Filename, StringLen($Filename) - ($i - 1))
			$Filename = StringLeft($Filename, $i - 1)
			ExitLoop
		EndIf
	Next
	
	$ArgRef01 = $VarDrive
	$ArgRef02 = $Path
	$ArgRef03 = $Filename
	$ArgRef04 = $Ext
	$ArrFileName[1] = $VarDrive
	$ArrFileName[2] = $Path
	$ArrFileName[3] = $Filename
	$ArrFileName[4] = $Ext
	Return $ArrFileName
EndFunc
Func FnReplaceVarStrFuncName($Arg00, $ArgOpt01 = 1)
	Local $Var0082 = $ScriptDir & "\EnCodeItInfo\", $Var0083
	$Arg00 = Fn00AC(Fn00AB(Fn00A5($Arg00)))
	Sleep($gSleepTime)
	Return $Arg00
EndFunc
Func Fn0087($Arg00, $ArgOpt01 = 1)
	Local $Var0082 = $ScriptDir & "\EnCodeItInfo\", $Var0083
	FileClose(FileOpen($Var0082 & "1.properties", 2))
	$Arg00 = Fn00A5($Arg00)
	FileWrite($Var0082 & "1.properties", $Arg00)
	Sleep(100)
	$Arg00 = Fn00AB(FileRead($Var0082 & "1.properties"))
	FileClose(FileOpen($Var0082 & "1.properties", 2))
	FileWrite($Var0082 & "1.properties", $Arg00)
	Sleep(100)
	$Arg00 = Fn00AC(FileRead($Var0082 & "1.properties"))
	FileClose(FileOpen($Var0082 & "1.properties", 2))
	FileWrite($Var0082 & "1.properties", $Arg00)
	$Var0083 &= FileRead($Var0082 & "1.properties")
	FileDelete($Var0082 & "1.properties")
	Sleep($gSleepTime)
	FileClose(FileOpen($Var0082 & "Compiled.au3", 2))
	FileWrite($Var0082 & "Compiled.au3", $Var0083)
	Return 1
EndFunc
Func Fn0088($Arg00, $Arg01, $Arg02, $ArgOpt03 = 'i')
	If $ArgOpt03 <> "i" Then $ArgOpt03 = ''
	$Var0027 = StringRegExp ($Arg00, "(?" & $ArgOpt03 & ":" & $Arg01 & ")(.*?)(?" & $ArgOpt03 & ":" & $Arg02 & ")", 3)
	If @extended & IsArray($Var0027) Then Return $Var0027
	Return SetError(1, 0, 0)
EndFunc
Func FnVarNamesToVar_ini(ByRef $ArgRef00, ByRef $ArgRef01)
	Local $Var0086 = Fn00AE($ScriptDir & "\EnCodeItInfo\Var.ini", "Vars")
	
	Local $Var003A, $Var0088 = Fn008C($ArgRef00)
	If IsArray($Var0086) And $Var0086[0][0] > 0 Then
		For $Var004C = 1 To UBound($ArgRef00) - 1
			If Not StringInStr($ArgRef00[$Var004C], "EnCodeIt") Then
				$Var003A += 1
				$gLogFile_Vars &= "    " & StringFormat("% 6d", $Var003A) & ".    " & "$" & $ArgRef00[$Var004C] & $Var0088[$Var004C] & StringFormat("% 6d", $Var003A) & ".    " & "$" & $ArgRef01[$Var004C] & $CRLF
			EndIf
			IniWrite($ScriptDir & "\EnCodeItInfo\Var.ini", "Vars", $ArgRef00[$Var004C], $ArgRef01[$Var004C])
			For $Var0075 = 1 To $Var0086[0][0]
				If $ArgRef00[$Var004C] = $Var0086[$Var0075][1] Then
					IniWrite($ScriptDir & "\EnCodeItInfo\Var.ini", "Vars", $Var0086[$Var0075][0], "$" & $ArgRef01[$Var004C])
					ExitLoop
				EndIf
			Next
		Next
	Else
		For $Var004C = 1 To UBound($ArgRef00) - 1
			IniWrite($ScriptDir & "\EnCodeItInfo\Var.ini", "Vars", $ArgRef00[$Var004C], $ArgRef01[$Var004C])
			If Not StringInStr($ArgRef00[$Var004C], "EnCodeIt") Then
				$Var003A += 1
				$gLogFile_Vars &= "    " & StringFormat("% 6d", $Var003A) & ".    " & "$" & $ArgRef00[$Var004C] & $Var0088[$Var004C] & StringFormat("% 6d", $Var003A) & ".    " & "$" & $ArgRef01[$Var004C] & $CRLF
			EndIf
		Next
	EndIf
	Return 1
EndFunc


Func FnWriteFunc_ini(ByRef $ArgRef00, ByRef $ArgRef01)
	Local $Var003A, $Var0088 = Fn008C($ArgRef00)
	For $Var004C = 1 To UBound($ArgRef00) - 1
		IniWrite($ScriptDir & "\EnCodeItInfo\Func.ini", "Funcs", $ArgRef00[$Var004C], $ArgRef01[$Var004C])
		If Not StringInStr($ArgRef00[$Var004C], "_EnCodeIt_") Then
			$Var003A += 1
			$gLogFile_Functions &= "    " & StringFormat("% 6d", $Var003A) & ".    " & $ArgRef00[$Var004C] & $Var0088[$Var004C] & StringFormat("% 6d", $Var003A) & ".    " & $ArgRef01[$Var004C] & $CRLF
		EndIf
	Next
	Return 1
EndFunc
Func Fn008B(ByRef $ArgRef00, ByRef $ArgRef01)
	For $Var004C = 1 To UBound($ArgRef00) - 1
		IniWrite($ScriptDir & "\EnCodeItInfo\Text.ini", "Text", $ArgRef01[$Var004C], $ArgRef00[$Var004C])
	Next
	Return 1
EndFunc
Func Fn008C(ByRef $ArgRef00)
	Local $Arr008B[UBound($ArgRef00) ], $Var003A
	For $Var004C = 1 To UBound($ArgRef00) - 1
		For $Var0075 = 1 To ((12 + StringLen($ArgRef00[1])) - StringLen($ArgRef00[$Var004C]))
			$Arr008B[$Var004C] &= Chr(32)
		Next
	Next
	Return $Arr008B
EndFunc
Func FnGetMonthName($Arg00)
	Local $ArrMonthNames[13] = ['', "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
	Return $ArrMonthNames[$Arg00]
EndFunc
Func FnGetTime()
	Local $VarHOUR = @HOUR, $VarMIN = @MIN, $VarAM_PM
	If $VarHOUR < 12 Then
		If $VarHOUR = 00 Then $VarHOUR = 12
		$VarAM_PM = " AM"
	Else
		If $VarHOUR > 12 Then $VarHOUR -= 12
		$VarAM_PM = " PM"
	EndIf
	Return $VarHOUR & ":" & $VarMIN & $VarAM_PM
EndFunc
Func Fn4ByteRandHex()
	Local $Var002A
	For $Var004C = 1 To 4
		$Var002A &= Hex(Random(0, 2 ^ 16 - 1, 1), 4)
	Next
	Return $Var002A
EndFunc
Func Fn_EnCodeIt_Funcs_Body()
	Return "Func _EnCodeIt_UStr($sEnCodeItFunction, $iEnCodeItLevel3)" & $CRLF & "Local $iEncodeItOutKey" & $CRLF & "$sEnCodeItFunction = _EnCodeIt_OStr($sEnCodeItFunction)" & $CRLF & "For $xEnCodeItCount = 1 to StringLen($sEnCodeItFunction)" & $CRLF & "$iEncodeItOutKey = $iEncodeItOutKey & Chr(Asc(StringMid($sEnCodeItFunction,$xEnCodeItCount,1))-" & "$iEnCodeItLevel3" & ")" & $CRLF & "Next" & $CRLF & "Return $iEncodeItOutKey" & $CRLF & "EndFunc" & $CRLF & "Func _EnCodeIt_OStr($sEnCodeItStrHex)" & $CRLF & "Local $sEnCodeItChar" & $CRLF & "$aEnCodeIt = StringSplit($sEnCodeItStrHex, """")" & $CRLF & "If Mod($aEnCodeIt[0], 2) <> 0 Then Return SetError(1, 0, -1)" & $CRLF & "For $iXEnCodeItCount = 1 To $aEnCodeIt[0]" & $CRLF & "$iEnCodeItOne = $aEnCodeIt[$iXEnCodeItCount]" & $CRLF & "$iXEnCodeItCount += 1" & $CRLF & "$iEnCodeItTwo = $aEnCodeIt[$iXEnCodeItCount]" & $CRLF & "$iEnCodeItDec = Dec($iEnCodeItOne & $iEnCodeItTwo)" & $CRLF & "If @error <> 0 Then Return SetError(1, 0, -1)" & $CRLF & "$iEnCodeItChar = Chr($iEnCodeItDec)" & $CRLF & "$sEnCodeItChar &= $iEnCodeItChar" & $CRLF & "Next" & $CRLF & "Return $sEnCodeItChar" & $CRLF & "EndFunc" & $CRLF & "Func _EnCodeIt_UStr2($sEnCodeItFunction, $EnCodeItConstVar)" & $CRLF & "Local $iEncodeItOutKey" & $CRLF & "$iEnCodeItDiv = " & ($gDivKey2 - 11) & $CRLF & "$sEnCodeItFunction = _EnCodeIt_OStr($sEnCodeItFunction)" & $CRLF & "For $xEnCodeItCount = 1 to StringLen($sEnCodeItFunction)" & $CRLF & "$iEncodeItOutKey &= Chr(Asc(StringMid($sEnCodeItFunction,$xEnCodeItCount,1))-" & "Int($iEnCodeItDiv)" & ")" & $CRLF & "Next" & $CRLF & "Return $iEncodeItOutKey" & $CRLF & "EndFunc"
EndFunc
Func Fn0091($Arg00, $Arg01, $Arg02, $ArgOpt03 = 0)
	If $Arg01 <= 0 Then Return SetError(4, 0, 0)
	If Not IsString($Arg02) Then Return SetError(6, 0, 0)
	If $ArgOpt03 <> 0 And $ArgOpt03 <> 1 Then Return SetError(5, 0, 0)
	If Not FileExists($Arg00) Then Return SetError(2, 0, 0)
	Local $Var0092 = FileRead($Arg00)
	$Var0092 = StringSplit($Var0092, $CRLF, 1)
	If UBound($Var0092, 1) < $Arg01 Then Return SetError(1, 0, 0)
	Local $Var0093 = FileOpen($Arg00, 2)
	If $Var0093 = -1 Then Return SetError(3, 0, 0)
	For $M8D5A722F600B4F29 = 1 To UBound($Var0092) - 1
		If $M8D5A722F600B4F29 = $Arg01 Then
			If $ArgOpt03 = 1 Then
				If $Arg02 <> '' Then
					FileWrite($Var0093, $Arg02 & $CRLF)
				Else
					FileWrite($Var0093, $Arg02)
				EndIf
			EndIf
			If $ArgOpt03 = 0 Then
				FileWrite($Var0093, $Arg02 & $CRLF)
				FileWrite($Var0093, $Var0092[$M8D5A722F600B4F29] & $CRLF)
			EndIf
		ElseIf $M8D5A722F600B4F29 < UBound($Var0092, 1) - 1 Then
			FileWrite($Var0093, $Var0092[$M8D5A722F600B4F29] & $CRLF)
		ElseIf $M8D5A722F600B4F29 = UBound($Var0092, 1) - 1 Then
			FileWrite($Var0093, $Var0092[$M8D5A722F600B4F29])
		EndIf
	Next
	FileClose($Var0093)
	Return 1
EndFunc
Func FnInetCheckBlackList1()
	If FileExists($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak") Then
		If Not CheckUserInfo() Then Exit
		FnBlackListCheck()
	Else
		FnBlackListCheck()
	EndIf
	Return 1
EndFunc
Func FnBlackListCheck()
;	Get http://www.autoitscript.com/fileman/users/SmOke_N/EnCodeIt_BlackList.dat
	Local $Var0094, $Lines, $Var0096, $Var0097, $InetData = FnInetGet4()
	If Not IsArray($InetData) Then Exit

;	Get .autoitscript.com {...}member_id   >25174< from firefox/cookie.dat
	Local $CookieDataIE = FnGetCookieDataIE(), $CookieDataFireFox = FnGetFireFoxCookie_Member_ids()
	FileDelete($ScriptDir & "\EnCodeItInfo\456E436F646549745F426C61636B4C697374.dat")
	
	If IsArray($CookieDataIE) Or IsArray($CookieDataFireFox) Then

		If IsArray($CookieDataIE) Then
			For $i_Cookie = 1 To $CookieDataIE[0]
				For $i_BlacklistDat = 2 To $InetData[0]
					If $InetData[$i_BlacklistDat] = "PausingEnCodeItUsage" Then
						MsgBox(16, "Warning!", "All usage has been temporarily suspened, this could be due to an update.", 5)
						Exit
					EndIf
					
					If Number($CookieDataIE[$i_Cookie]) = Number($InetData[$i_BlacklistDat]) Then
						RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\AutoIt v3\EnCodeIt", "UserInfo", "REG_SZ", $gUseOfEnCodeItNotPermitted)
						If Not FileExists($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak") Then
							DirCreate($AppDataDir & "\Microsoft\456E436F64654974")
							FileClose(FileOpen($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak", 2))
							FileWriteLine($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak", $gUseOfEnCodeItNotPermitted)
							FileSetAttrib($AppDataDir & "\Microsoft\456E436F64654974", "+H", 1)
						Else
							DirRemove($AppDataDir & "\Microsoft\456E436F64654974", 1)
						EndIf
						Exit
					EndIf
					
				Next
			Next
		EndIf

		If IsArray($CookieDataFireFox) Then
			For $i_Cookie = 1 To $CookieDataFireFox[0]
				For $i_BlacklistDat = 2 To $InetData[0]
					If $InetData[$i_BlacklistDat] = "PausingEnCodeItUsage" Then Exit
					
					If Number($CookieDataFireFox[$i_Cookie]) = Number($InetData[$i_BlacklistDat]) Then
						RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\AutoIt v3\EnCodeIt", "UserInfo", "REG_SZ", $gUseOfEnCodeItNotPermitted)
						If Not FileExists($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak") Then
							DirCreate($AppDataDir & "\Microsoft\456E436F64654974")
							FileClose(FileOpen($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak", 2))
							FileWriteLine($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak", $gUseOfEnCodeItNotPermitted)
							FileSetAttrib($AppDataDir & "\Microsoft\456E436F64654974", "+H", 1)
						Else
							DirRemove($AppDataDir & "\Microsoft\456E436F64654974", 1)
						EndIf
						Exit
					EndIf
					
				Next
			Next
		EndIf

	Else
		MsgBox(16, "Error", "You must be a registered member with AutoIt to use EnCodeIt.")
		Exit
	EndIf
	
	RegWrite("HKEY_LOCAL_MACHINE\SOFTWARE\AutoIt v3\EnCodeIt", "UserInfo", "REG_SZ", $gDataUserInfo)
	If Not FileExists($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak") Then
		DirCreate($AppDataDir & "\Microsoft\456E436F64654974")
		FileClose(FileOpen($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak", 2))
		FileWriteLine($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak", $gDataUserInfo)
		FileSetAttrib($AppDataDir & "\Microsoft\456E436F64654974", "+H", 1)
	EndIf
	Return 1
EndFunc
Func FnToNumber($Arg00)
	Return Number($Arg00)
EndFunc
Func CheckUserInfo()
	If RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\AutoIt v3\EnCodeIt", "UserInfo") == $gDataUserInfo And _
		StringStripWS(FileRead($AppDataDir & "\Microsoft\456E436F64654974\456E436F64654974.bak"), 8) == $gDataUserInfo Then _
			Return 1
	Return SetError(1, 0, 0)
EndFunc

Func FnGetCookieDataIE()
	Local $Var0094, $Autoitscript_txt_Data, $Var009D, $Var0096, $Var009F, $Var001D, $bCookieFileFound = False
	
	Local $VarCookieDir = $UserProfileDir & "\Cookies\" & $UserName & "@autoitscript"
	If Not FileExists($VarCookieDir & ".txt") Then
		For $i = 1 To 100
			If FileExists($VarCookieDir & "[" & $i & "].txt") Then
				$bCookieFileFound = True
				$Var009F &= $Var0094 = $VarCookieDir & "[" & $i & "].txt" & Chr(01)
			EndIf
		Next
		
		If $bCookieFileFound Then
			$Var009D = StringSplit(StringTrimRight($Var001D, 1), Chr(01))
			For $Var00D4 = 1 To $Var009D[0]
				$Autoitscript_txt_Data = StringSplit(StringStripCR(FileRead($Var009D[$Var00D4])), @LF)
				If $Autoitscript_txt_Data[0] >= 2 Then $Var009F &= $Autoitscript_txt_Data[2] & Chr(01)
			Next
			Return StringSplit(StringTrimRight($Var009F, 1), Chr(01))
		EndIf
	
	Else	;"\Cookies\cw2k@autoitscript.txt present
		$Autoitscript_txt_Data = StringSplit(StringStripCR(FileRead($VarCookieDir & ".txt")), @LF)
		If $Autoitscript_txt_Data[0] >= 2 Then _
			Return $Autoitscript_txt_Data[2]
	
	EndIf
		
	Return False
EndFunc
Func FnDecryFileData($Arg00)
	Local $DecryData
	For $Var00B1 = 1 To StringLen($Arg00)
		Switch Asc(StringMid($Arg00, $Var00B1, 1))
			Case FnToNumber("127") To FnToNumber("135")
				$DecryData &= "A"
			Case FnToNumber("136") To FnToNumber("144")
				$DecryData &= "B"
			Case FnToNumber("149") To FnToNumber("153")
				$DecryData &= "C"
			Case FnToNumber("154") To FnToNumber("162")
				$DecryData &= "D"
			Case FnToNumber("163") To FnToNumber("171")
				$DecryData &= "E"
			Case FnToNumber("172") To FnToNumber("180")
				$DecryData &= "F"
			Case FnToNumber("181") To FnToNumber("189")
				$DecryData &= FnToNumber("0")
			Case FnToNumber("190") To FnToNumber("198")
				$DecryData &= FnToNumber("2")
			Case FnToNumber("199") To FnToNumber("207")
				$DecryData &= FnToNumber("4")
			Case FnToNumber("208") To FnToNumber("216")
				$DecryData &= FnToNumber("3")
			Case FnToNumber("217") To FnToNumber("225")
				$DecryData &= FnToNumber("5")
			Case FnToNumber("226") To FnToNumber("234")
				$DecryData &= FnToNumber("6")
			Case FnToNumber("235") To FnToNumber("243")
				$DecryData &= FnToNumber("8")
			Case FnToNumber("244") To FnToNumber("252")
				$DecryData &= FnToNumber("9")
			Case FnToNumber("253") To FnToNumber("254")
				$DecryData &= FnToNumber("7")
			Case FnToNumber("255")
				$DecryData &= FnToNumber("1")
		EndSwitch
	Next
	Return BinaryString("0x" & $DecryData)
EndFunc
Func FnGetFireFoxCookie_Member_ids()
	If Not FileExists($AppDataDir & "\Mozilla\Firefox\Profiles") Then Return SetError(1, 0, 0)
	Local $member_ids, $FilesInProfiles, $FileData, $RegExpData
	$FilesInProfiles = FnDir($AppDataDir & "\Mozilla\Firefox\Profiles")
	If IsArray($FilesInProfiles) Then
		;For all Files in "\Mozilla\Firefox\Profiles" Do
		For $i = 0 To UBound($FilesInProfiles) - 1
			
			If StringInStr(FileGetAttrib($FilesInProfiles[$i]), "D") And _
				Not StringInStr(StringTrimLeft($FilesInProfiles[$i], StringLen($AppDataDir & "\Mozilla\Firefox\Profiles") + 1), "\") Then
				$FileData = FileRead($FilesInProfiles[$i])
;Some line from cookies.txt
;.autoitscript.com   TRUE   /forum/   FALSE   1215984826   member_id   21435
				$RegExpData = StringRegExp($FileData, ".autoitscript.com.+?member_id\s*(\d*)\n", 1)
				For $index = 0 To UBound($RegExpData) - 1
					$member_ids &= $RegExpData[$index] & Chr(01)
				Next
			EndIf
			
		Next
		
		Return StringSplit(StringTrimRight($member_ids, 1), Chr(01))
		
	EndIf
EndFunc
Func FnInetGet4()
   ;If FnInetGet3("http://www.autoitscript.com//SmOke_N/EnCodeIt_BlackList.dat", $ScriptDir & "\EnCodeItInfo\456E436F646549745F426C61636B4C697374.dat", 1, 0) Then
	If FnInetGet3("http://www.autoitscript.com/fileman/index.php?op=get&target=SmOke_N/EnCodeIt_BlackList.dat.txt", $ScriptDir & "\EnCodeItInfo\456E436F646549745F426C61636B4C697374.dat", 1, 0) Then
		$EnCodeItInfo_Dat = StringSplit(StringStripCR(FnDecryFileData(FileRead($ScriptDir & "\EnCodeItInfo\456E436F646549745F426C61636B4C697374.dat"))), Chr(01))
		Return $EnCodeItInfo_Dat
	Else
		MsgBox(16, "Error", "EnCodeIt must have access to the internet upon start up.")
		Exit
	EndIf
	Return 0
EndFunc
Func FnDir($Arg00)
	Local $search , $FoundFileName, $FileList
	$search  = FileFindFirstFile("*.*")
	If $search  = -1 Then Return SetError(1, 0, 0)
	
	While 1
		$FoundFileName = FileFindNextFile($search )
		If @error Then ExitLoop
			
		$FileList &= $Arg00 & "\" & $FoundFileName & Chr(01)
		
	WEnd
	
	FileClose($search )
	
	Return StringSplit(StringTrimRight($FileList, 1), Chr(01))
EndFunc
Func FnInetGet($URL, $DownloadToFileName, $reload, $DownloadInBackground)
	Return InetGet($URL, $DownloadToFileName, $reload, $DownloadInBackground)
EndFunc
Func FnInetGet3($URL, $DownloadToFileName, $reload, $DownloadInBackground, $ArgOpt04 = 1, $RetValFnInetGet = 8, $DoDownload = False)
	$ArgOpt04 = FnInetGet2($URL, $DownloadToFileName, $reload, $DownloadInBackground, $DoDownload)
	$RetValFnInetGet = FnInetGet($URL, $DownloadToFileName, $reload, $DownloadInBackground)
	Return $RetValFnInetGet
EndFunc
Func FnInetGet2($URL, $DownloadToFileName, $reload, $DownloadInBackground, $DoDownload)

	If Not $DoDownload Then Return 1
	;Success: Returns 1
	Return InetGet($URL, $DownloadToFileName, $reload, $DownloadInBackground)
EndFunc


Func FnFilesCleanUp()
	If FileExists($ScriptDir & "\EnCodeItInfo\Compiled.au3") Then FileDelete($ScriptDir & "\EnCodeItInfo\Compiled.au3")
	If FileExists($ScriptDir & "\EnCodeItInfo\Text.ini") Then FileDelete($ScriptDir & "\EnCodeItInfo\Text.ini")
	If FileExists($ScriptDir & "\EnCodeItInfo\Var.ini") Then FileDelete($ScriptDir & "\EnCodeItInfo\Var.ini")
	If FileExists($ScriptDir & "\EnCodeItInfo\Func.ini") Then FileDelete($ScriptDir & "\EnCodeItInfo\Func.ini")
	If FileExists($ScriptDir & "\EnCodeItInfo\FuncReplace.ini") Then FileDelete($ScriptDir & "\EnCodeItInfo\FuncReplace.ini")
	If FileExists($ScriptDir & "\EnCodeItInfo\Backup.back") Then FileDelete($ScriptDir & "\EnCodeItInfo\Backup.bak")
EndFunc
Func Fn009F()
	DirRemove($ScriptDir & "\EnCodeItInfo", 1)
EndFunc
Func Fn00A0($Arg00, $Arg01, $ArgOpt02 = True)
	Local $Var00AB = FileRead($Arg00)
	FileClose(FileOpen($Arg00, 2))
	If $ArgOpt02 Then
		FileWrite($Arg00, $Arg01 & $CRLF & $Var00AB)
	Else
		FileWrite($Arg00, $Var00AB & $CRLF & $Arg01)
	EndIf
	Return 1
EndFunc
Func Fn00A1($Arg00)
	Local $Lines = StringSplit(StringStripCR($Arg00), $lf)
	Local $Var0051 = "Somethingreallylongandobnoxiousforreplacementoffunc"
	Local $Arr0052[$Lines[0] + 1][2], $Var003A
	For $Var004C = 1 To $Lines[0]
		If StringLeft(StringStripWS($Lines[$Var004C], 8), 4) = "func" Then
			$Var003A += 1
			$Arr0052[$Var003A][0] = $Lines[$Var004C]
			$Arr0052[$Var003A][1] = $Var0051 & $Var003A
		EndIf
	Next
	ReDim $Arr0052[$Var003A + 1][2]
	Return $Arr0052
EndFunc
Func Fn00A2($Arg00)
	Local $Arr00B0[4], $Var00B1, $Var005A
	If (StringIsAlpha(StringLeft($Arg00, 1)) Or StringLeft($Arg00, 1) = "_") And (StringIsAlpha(StringRight($Arg00, 1)) Or StringRight($Arg00, 1) = "_") Then
		$Arr00B0[2] = $Arg00
		Return $Arr00B0
	EndIf
	For $Var00B1 = 1 To StringLen($Arg00)
		$Var005A = StringMid($Arg00, $Var00B1, 1)
		If StringIsAlpha($Var005A) Or $Var005A = "_" Then
			$Arr00B0[1] = StringMid($Arg00, $Var00B1 - 1, $Var00B1 - 1)
			$Arg00 = StringTrimLeft($Arg00, $Var00B1 - 1)
			For $Var00D4 = StringLen($Arg00) To 1 Step - 1
				$Var005A = StringMid($Arg00, $Var00D4, 1)
				If StringIsAlpha($Var005A) Or $Var005A = "_" Then
					$Arr00B0[2] = StringTrimRight($Arg00, StringLen($Arg00) - ($Var00D4))
					$Arr00B0[3] = StringReplace($Arg00, $Arr00B0[2], '')
					ExitLoop
				EndIf
			Next
			If $Arr00B0[2] = '' Then $Arr00B0[2] = $Arg00
			ExitLoop
		EndIf
	Next
	Return $Arr00B0
EndFunc
Func Fn00A3(ByRef $ArgRef00, $Arg01)
	For $Var004C = UBound($ArgRef00) - 1 To 1 Step - 1
		$Arg01 = StringReplace($Arg01, $ArgRef00[$Var004C][0], $ArgRef00[$Var004C][1], 0, 1)
	Next
	Return $Arg01
EndFunc
Func Fn00A4(ByRef $ArgRef00, $Arg01)
	Local $Lines = StringSplit(StringStripCR($Arg01), @LF)
	Local $Var001D
	$Arg01 = ''
	For $Var00D4 = 1 To $Lines[0]
		For $Var004C = UBound($ArgRef00) - 1 To 1 Step - 1
			If StringLeft($Lines[$Var00D4], StringLen($Lines[$Var00D4])) == $ArgRef00[$Var004C][1] Then
				$Lines[$Var00D4] = StringReplace($Lines[$Var00D4], $ArgRef00[$Var004C][1], $ArgRef00[$Var004C][0], 1)
				ExitLoop
			EndIf
		Next
		$Arg01 &= $Lines[$Var00D4] & $CRLF
	Next
	Return StringTrimRight($Arg01, StringLen($CRLF))
EndFunc
Func Fn00A5($Arg00)
	Local $Lines, $Var001D, $Var0059, $Var00B8
	Local $Var00B9, $Var0050, $Var00BB, $Var00BC, $Arr0044[5] = ['', False, False, False, False]
	Local $Var00BE = Fn00AE($ScriptDir & "\EnCodeItInfo\Text.ini", "Text")
	If Not IsArray($Var00BE) Then Return $Arg00
	If $Var00BE[1][0] = '' Then Return $Arg00
	$Var0043 = Fn00AE($ScriptDir & "\EnCodeItInfo\FuncReplace.ini", "FunctionReplace")
	If IsArray($Var0043) Then
		$Arg00 = Fn00A4($Var0043, $Arg00)
		$Arr0044[1] = True
	EndIf
	$Var00B9 = Fn00AA($Arg00, "hotkeyset", "\)")
	If Not @error Then
		$Arg00 = Fn00A7($Var00B9, $Arg00)
		$Arr0044[2] = True
	EndIf
	$Var0050 = Fn00AA($Arg00, "fileinstall", ",")
	If Not @error Then
		$Arg00 = Fn00A7($Var0050, $Arg00)
		$Arr0044[3] = True
	EndIf
	$Var00BC = Fn00AA($Arg00, "adlibenable", "\)")
	If Not @error Then
		$Arg00 = Fn00A7($Var00BC, $Arg00)
		$Arr0044[4] = True
	EndIf
	$Arg00 = Fn00AD($Arg00)
	If $Arr0044[1] Then $Arg00 = Fn00A3($Var0043, $Arg00)
	If $Arr0044[2] Then $Arg00 = Fn00A6($Var00B9, $Arg00)
	If $Arr0044[3] Then $Arg00 = Fn00A6($Var0050, $Arg00)
	If $Arr0044[4] Then $Arg00 = Fn00A6($Var00BC, $Arg00)
	Return $Arg00
EndFunc
Func Fn00A6(ByRef $ArgRef00, $Arg01)
	For $Var004C = UBound($ArgRef00) - 1 To 1 Step - 1
		$Arg01 = StringReplace($Arg01, $ArgRef00[$Var004C][1], $ArgRef00[$Var004C][0], 0, 1)
	Next
	Return $Arg01
EndFunc
Func Fn00A7(ByRef $ArgRef00, $Arg01)
	For $Var004C = UBound($ArgRef00) - 1 To 1 Step - 1
		$Arg01 = StringReplace($Arg01, $ArgRef00[$Var004C][0], $ArgRef00[$Var004C][1], 0, 1)
	Next
	Return $Arg01
EndFunc
Func Fn00A8($Arg00, $ArgOpt01 = 1, $ArgOpt02 = '')
	If IsString($Arg00) Then Return Fn00A9($Arg00, $ArgOpt02)
	If $ArgOpt02 = '' Then $ArgOpt02 = Chr(01)
	Local $Var001D
	Fn007F($Arg00, $ArgOpt01)
	For $Var004C = $ArgOpt01 To UBound($Arg00) - 1
		If StringInStr($ArgOpt02 & $Var001D, $ArgOpt02 & $Arg00[$Var004C] & $ArgOpt02, 1) Then ContinueLoop
		If $Arg00[$Var004C] <> '' Then $Var001D &= $Arg00[$Var004C] & $ArgOpt02
	Next
	Return StringSplit(StringTrimRight($Var001D, StringLen($ArgOpt02)), $ArgOpt02)
EndFunc
Func Fn00A9($Arg00, $Arg01)
	If IsArray($Arg00) Then Return Fn00A8($Arg00)
	If StringRight($Arg00, StringLen($Arg01)) = $Arg01 Then $Arg00 = StringTrimRight($Arg00, StringLen($Arg01))
	Local $Var005E = StringSplit(StringStripCR($Arg00), $Arg01), $Var001D
	Fn007F($Var005E)
	For $Var004C = 1 To $Var005E[0]
		If StringInStr($Arg01 & $Var001D, $Arg01 & $Var005E[$Var004C] & $Arg01, 1) Then ContinueLoop
		If $Var005E[$Var004C] <> '' Then $Var001D &= $Var005E[$Var004C] & $Arg01
	Next
	Return StringSplit(StringTrimRight($Var001D, StringLen($Arg01)), $Arg01)
EndFunc
Func Fn00AA($Arg00, $Arg01, $Arg02)
	Local $Var0050 = Fn0088($Arg00, "\s" & $Arg01, $Arg02)
	If Not @extended And Not IsArray($Var0050) Then Return SetError(1, 0, 0)
	Local $Var0051 = "Somethingreallylongandobnoxiousforreplacementof" & StringStripWS($Arg01, 8)
	$Var0050 = Fn00A8($Var0050, 0)
	Local $Arr0052[$Var0050[0] + 1][2]
	For $Var004C = 1 To $Var0050[0]
		$Arr0052[$Var004C][0] = $Var0050[$Var004C]
		$Arr0052[$Var004C][1] = $Var0051 & $Var004C
	Next
	Return $Arr0052
EndFunc
Func Fn00AB($Arg00)
	Local $Var00C5 = Fn00AE($ScriptDir & "\EnCodeItInfo\Func.ini", "Funcs")
	If Not IsArray($Var00C5) Then Return $Arg00
	If $Var00C5[1][0] = '' Then Return $Arg00
	For $Var00B1 = 1 To $Var00C5[0][0]
		$Arr00B0 = Fn00A2($Var00C5[$Var00B1][0])
		$Arg00 = StringReplace($Arg00, "'" & $Var00C5[$Var00B1][0] & "'", "'_!_" & $Var00C5[$Var00B1][1] & "'")
		$Arg00 = StringReplace($Arg00, """" & $Var00C5[$Var00B1][0] & """", """_!_" & $Var00C5[$Var00B1][1] & """")
		$Arg00 = StringRegExpReplace($Arg00, $Arr00B0[1] & "(\<" & $Arr00B0[2] & "\>)" & $Arr00B0[3], "_!_" & $Var00C5[$Var00B1][1])
		Sleep($gSleepTime)
	Next
	Return StringReplace($Arg00, "_!_", '', 0, 1)
EndFunc
Func Fn00AC($Arg00)
	Local $Var0086 = Fn00AE($ScriptDir & "\EnCodeItInfo\Var.ini", "Vars")
	If Not IsArray($Var0086) Or $Var0086[1][0] = '' Then Return SetError(1, 0, $Arg00)
	For $Var00B1 = 1 To $Var0086[0][0]
		$Arg00 = StringReplace($Arg00, "$" & $Var0086[$Var00B1][0], "$@!~@`" & $Var0086[$Var00B1][1])
		Sleep($gSleepTime)
	Next
	Return StringReplace($Arg00, "@!~@`", '')
EndFunc
Func Fn00AD($Arg00)
	Local $Var0043, $Var0045 = False, $Var0046, $Var00CA, $Var0049, $Var004A
	Local $Var00BE = Fn00AE($ScriptDir & "\EnCodeItInfo\Text.ini", "Text")
	If Not IsArray($Var00BE) Then Return $Arg00
	If $Var00BE[1][0] = '' Then Return $Arg00
	Local $Var004B = StringReplace($Arg00, "'" & "'", ";" & ";")
	$Var004B = StringReplace($Var004B, """" & """", "~" & "~")
	$Var00CA = $Var004B
	Local $Var003A = 0, $Var004E
	While StringLen($Arg00) > 0
		$Var004A = StringInStr($Var004B, "'")
		$Var0049 = StringInStr($Var004B, """")
		If Not $Var004A And Not $Var0049 Then ExitLoop
		If $Var0049 > 0 And ($Var004A > $Var0049 Or $Var004A = 0) Then
			$M8DAA72276C5B4F29 = """"
			$Arg00 = StringTrimLeft($Arg00, $Var0049 - 1)
			$Var004B = StringTrimLeft($Var004B, $Var0049 - 1)
		Else
			$M8DAA72276C5B4F29 = "'"
			$Arg00 = StringTrimLeft($Arg00, $Var004A - 1)
			$Var004B = StringTrimLeft($Var004B, $Var004A - 1)
		EndIf
		$M8D8A722F6C5B5F29 = StringInStr($Var004B, $M8DAA72276C5B4F29, 0, 2)
		If $M8D8A722F6C5B5F29 > 0 Then
			$Var0046 = StringLeft($Arg00, $M8D8A722F6C5B5F29)
			For $Var00B1 = 1 To $Var00BE[0][0]
				If $Var0046 == $Var00BE[$Var00B1][1] Then
					If StringLen($Var0046) > 2 And StringInStr($Var0046, """" & """") Or StringInStr($Var0046, "'" & "'") Then
						$Var00CA = StringReplace($Var00CA, ";" & ";", "'" & "'")
						$Var00CA = StringReplace($Var00CA, "~" & "~", """" & """")
						$Var00CA = StringReplace($Var00CA, $Var0046, "$" & $Var00BE[$Var00B1][0], 1)
						$Var00CA = StringReplace($Var00CA, "'" & "'", ";" & ";")
						$Var00CA = StringReplace($Var00CA, """" & """", "~" & "~")
					Else
						$Var00CA = StringReplace($Var00CA, $Var0046, "$" & $Var00BE[$Var00B1][0], 1)
					EndIf
					ExitLoop
				EndIf
			Next
			$Arg00 = StringTrimLeft($Arg00, $M8D8A722F6C5B5F29)
			$Var004B = StringTrimLeft($Var004B, $M8D8A722F6C5B5F29)
		Else
			ExitLoop
		EndIf
		Sleep($gSleepTime)
	WEnd
	$Var00CA = StringReplace($Var00CA, ";" & ";", "'" & "'")
	$Var00CA = StringReplace($Var00CA, "~" & "~", """" & """")
	Return $Var00CA
EndFunc
Func Fn00AE($Arg00, $Arg01)
	Local $Var00D1 = FileGetSize($Arg00) / 1024
	If $Var00D1 <= 31 Then Return IniReadSection($Arg00, $Arg01)
	Local $Var00A6 = FileRead($Arg00), $Var001D, $Var00D4, $Var00D5 = 1
	Local $Lines = StringSplit(StringStripCR($Var00A6), @LF)
	For $Var00D4 = 1 To $Lines[0]
		If StringLeft(StringStripWS($Lines[$Var00D4], 7), StringLen(StringStripWS("[" & $Arg01 & "]", 7))) = StringStripWS("[" & $Arg01 & "]", 7) Then
			For $M8D5A722F6C5B1F29 = $Var00D4 + 1 To $Lines[0]
				If StringLeft(StringStripWS($Lines[$M8D5A722F6C5B1F29], 8), 1) = "[" Then ExitLoop
				$Var001D &= $Lines[$M8D5A722F6C5B1F29] & Chr(01)
			Next
			ExitLoop
		EndIf
	Next
	If $Var001D = '' Then Return SetError(1, 0, 0)
	$Var001D = StringSplit(StringTrimRight($Var001D, 1), Chr(01))
	FileWriteLine(@TempDir & "\IniReadSectionTemp" & $Var00D5 & ".ini", "[" & $Arg01 & "]")
	For $Var00B1 = 1 To $Var001D[0]
		If FileGetSize(@TempDir & "\IniReadSectionTemp" & $Var00D5 & ".ini") / 1024 > 25 Then
			$Var00D5 += 1
			FileWriteLine(@TempDir & "\IniReadSectionTemp" & $Var00D5 & ".ini", "[" & $Arg01 & "]")
		EndIf
		FileWriteLine(@TempDir & "\IniReadSectionTemp" & $Var00D5 & ".ini", $Var001D[$Var00B1])
	Next
	Local $Var00D7, $Var00D8, $Var00D9
	For $M8D5A722B695B4F29 = 1 To $Var00D5
		$Var00D7 = IniReadSection(@TempDir & "\IniReadSectionTemp" & $M8D5A722B695B4F29 & ".ini", $Arg01)
		If $M8D5A722B695B4F29 = 1 Then
			$Var00D8 = $Var00D7
			FileDelete(@TempDir & "\IniReadSectionTemp" & $M8D5A722B695B4F29 & ".ini")
		Else
			$Var00D9 = UBound($Var00D8, 1) - 1
			ReDim $Var00D8[($Var00D9 + 1) + $Var00D7[0][0]][2]
			For $M8D5A722F68544F29 = 1 To $Var00D7[0][0]
				$Var00D8[$M8D5A722F68544F29 + $Var00D9][0] = $Var00D7[$M8D5A722F68544F29][0]
				$Var00D8[$M8D5A722F68544F29 + $Var00D9][1] = $Var00D7[$M8D5A722F68544F29][1]
			Next
			FileDelete(@TempDir & "\IniReadSectionTemp" & $M8D5A722B695B4F29 & ".ini")
		EndIf
	Next
	$Var00D8[0][0] = UBound($Var00D8, 1) - 1
	Return $Var00D8
EndFunc