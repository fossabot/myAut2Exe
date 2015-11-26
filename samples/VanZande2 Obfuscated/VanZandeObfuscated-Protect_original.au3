; <AUT2EXE VERSION: 3.1.1.128>

; ----------------------------------------------------------------------------
; <AUT2EXE INCLUDE-START: I:\!Cracks & Projects\AutoIt3\AutoIt 3-Decompiler Improved Version\w0uter\prot.au3>
; ----------------------------------------------------------------------------

Opt("GuiOnEventMode", 1)

GUICreate("Protect", 393, 200)
GUISetFont(8, 0 , 0, 'Arial')
GUISetOnEvent ( -3, "_Exit")

;A3484BBE986C4AA9
;. H K . . l J .

;994C530A86D6487D
;. L S . . . H }

;41553321
;A U 3 !

;45413035
;E A 0 5

Global $av_AU3[24] = ['A3', '48', '4B', 'BE', '98', '6C', '4A', 'A9', '99', '4C', '53', '0A', '86', 'D6', '48', '7D', '41', '55', '33', '21', '45', '41', '30', '35']
$av_AU3 = _CI($av_AU3)

Global $h_executable = GUICtrlCreateInput('', 8, 176, 250, 16)

GUICtrlCreateButton('Browse', 263, 175, 59, 18)
GUICtrlSetOnEvent(-1, '_BR')

GUICtrlCreateButton('Patch', 327, 175, 59, 18)
GUICtrlSetOnEvent(-1, '_Patch')

GUISetState(@SW_SHOW)

While 1
    Sleep(10)
WEnd

Exit

Func _Patch()

    ;general
    Local $s_installdir = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\AutoIt v3\AutoIt', 'InstallDir')
    Local $s_executable = GUICtrlRead($h_executable)
    FileCopy($s_executable, StringTrimRight($s_executable, 3) & 'hack.exe', 1)
    $s_executable = StringTrimRight($s_executable, 3) & 'hack.exe'

    ;de-upx
    RunWait('"' & $s_installdir & '\aut2exe\upx.exe" -d "' & $s_executable & '"')

    ;binary read
    Local $v_executable = String(Binary(FileRead($s_executable)))

    ;modify
    _Replace($v_executable, 'A3484BBE986C4AA9', _Read(0, 07))
    _Replace($v_executable, '994C530A86D6487D', _Read(8, 15))
    _Replace($v_executable, '41553321', _Read(16, 19))
    _Replace($v_executable, '45413035', _Read(20, 23))

    ;write
    Local $h_Open = FileOpen($s_executable, 2)
    FileWrite($h_Open, Binary($v_executable))
    FileClose($h_Open)

    ;upx
    RunWait('"' & $s_installdir & '\aut2exe\upx.exe" --best "' & $s_executable & '"')

    MsgBox(0, 'Protect', 'The Patch has been applied on:' & @CRLF & @CRLF & $s_executable)

EndFunc

Func _Replace(byref $v_byref, $v_hex, $v_hax)
    $v_byref = StringReplace($v_byref, $v_hex, $v_hax)
EndFunc

Func _Read($j, $k)
    Local $v_tmp = ''
    For $i = $j to $k
        $v_tmp &= Hex('0x' & GUICtrlRead($av_AU3[$i][1]), 2)
    Next
    return $v_tmp
EndFunc

Func _BR()
    Local $v_tmp = @ScriptDir
    If GUICtrlRead($h_executable) <> '' Then $v_tmp = GUICtrlRead($h_executable)
    $v_tmp = FileOpenDialog('Protect', $v_tmp, 'Compiled AutoIt Scripts (*.exe)|All (*.*)', 1, '*.exe')
    If Not @error Then GUICtrlSetData($h_executable, $v_tmp)
EndFunc

Func _CI($ai_code)

    Local $i_max = 23
    Local $v_return[24][2]

    for $i = 0 to 7
        For $j = 0 to 2
            GUICtrlCreateInput(Chr('0x' & $ai_code[$i+8*$j]),   8+128*$j,      8+21*$i, 25, 16, 2049)
            GUICtrlCreateInput($ai_code[$i+8*$j],         8+32+128*$j, 8+21*$i, 25, 16, 2049)

            $v_return[$i+8*$j][0] = GUICtrlCreateInput(Chr('0x' & $ai_code[$i+8*$j]), 64+8+128*$j, 8+21*$i, 25, 16, 1)
            GUICtrlSetLimit(-1, 1)
            GUICtrlSetOnEvent(-1, "_Change")
            $v_return[$i+8*$j][1] = GUICtrlCreateInput($ai_code[$i+8*$j], 64+8+32+128*$j, 8+21*$i, 25, 16, 1)
            GUICtrlSetLimit(-1, 2)
            GUICtrlSetOnEvent(-1, "_Change")

        Next
    Next

    return $v_return
EndFunc

Func _Change()
    For $i = 0 To 23
        If $av_AU3[$i][0] = @GUI_CtrlId Then Return GUICtrlSetData($av_AU3[$i][1], Hex(Asc(GUICtrlRead($av_AU3[$i][0])), 2))
        If $av_AU3[$i][1] = @GUI_CtrlId Then Return GUICtrlSetData($av_AU3[$i][0], Chr('0x' & GUICtrlRead($av_AU3[$i][1])))
    Next
EndFunc

Func _Exit()
    Exit
EndFunc

; ----------------------------------------------------------------------------
; <AUT2EXE INCLUDE-END: I:\!Cracks & Projects\AutoIt3\AutoIt 3-Decompiler Improved Version\w0uter\prot.au3>
; ----------------------------------------------------------------------------

