 #NoTrayIcon
FileInstall(">>>AUTOIT SCRIPT<<<", @ScriptDir & "\TokenTestFile_Extracted.au3")
Exit
$DUMMY = (1 + 2 - 3 / 4) ^ 2
$DUMMY += 1 > 2 < 3 <> 4 >= 5 <= 6 = 7 & 8
$DUMMY -= True == True = True
$DUMMY /= 1
$DUMMY *= 1.123
$DUMMY &= 1500
$DUMMY = 1234567887654321
MYFUNC($DUMMY, $DUMMY)
$OSHELL = ObjCreate("shell.application")
$OSHELLWINDOWS = $OSHELL.windows

Func MYFUNC($VALUE, $VALUE2)
	Dim $ARRAY1[4]
	$ARRAY1[2] = 1
	$DUMMY = ' " '
	$DUMMY = ' " '
	$DUMMY = " ' ' """"  "
EndFunc

Exit