#NoTrayIcon
;----------------------------------------------------
; TestFile - for AutoIt3 tokens expansions (Decomiler)
;----------------------------------------------------

;PreProcessor

;AutoItFunction  UserString  Macro
FileInstall('>>>AUTOIT SCRIPT<<<', @ScriptDir & '\TokenTestFile_Extracted.au3')
Exit

;operators
$Dummy1 = (1+2-3/4)^2
$Dummy2 += 1>2<3<>4>=5<=6=7&8
$Dummy3 -= true==true=-true
$Dummy4 /= 0x1
$Dummy5 /= (-0x1)
$Dummy6 *= 1.123					;Float Number
$Dummy7 &= -1.53			    ;Float Number
$Dummy8 = 1234567887654321		;int64 Number
$Dummy9 = -1234567887654321		;int64 Number

;UserFunction
myfunc($Dummy1,$Dummy2)

;user Varible
$oShell = ObjCreate("shell.application")    ; Get the Windows Shell Object

;Properties
$oShellWindows=$oShell.windows


func myfunc($value,$value2)
	dim $Array1[4]
	$Array1[2]=1
	$Dummy1 = " "" "
	$Dummy2 = ' " '
	$Dummy3 = " ' ' """"  "

endfunc

exit
