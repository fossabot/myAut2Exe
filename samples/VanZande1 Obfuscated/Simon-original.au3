
#include <GUIConstants.au3>
Opt('GuiOnEventMode',1)

HotKeySet("{UP}", "HClick")
HotKeySet("{DOWN}", "HClick")
HotKeySet("{LEFT}", "HClick")
HotKeySet("{RIGHT}","HClick")

Func HClick()   
    IF @HotKeyPressed = "{LEFT}" Then $Num = 1
    IF @HotKeyPressed = "{UP}" Then $Num = 2
    IF @HotKeyPressed = "{DOWN}" Then $Num = 3
    IF @HotKeyPressed = "{RIGHT}" Then $Num = 4
    Switch $Num
        Case 1
            ControlClick("Simon", "", Execute($Red))
        Case 2
            ControlClick("Simon", "", Execute($Green))
        Case 3
            ControlClick("Simon", "", Execute($Blue))
        Case 4
            ControlClick("Simon", "", Execute($Yellow))
    EndSwitch
EndFunc


$TimePeriod = 2000 ;miliseconds
$RedFreq = 1000
$GreenFreq = 750
$BlueFreq = 500
$YellowFreq = 250
$Score = 0
$Count = 0
$Clicked = 0
$Duration = 300;miliseconds
$Last = 0
$Max = 100;game ends after
Dim $Pattern[$Max+1]
For $i = 1 to $Max
    $Pattern[$i] = Random(1,4,1)
Next


$Simon = GUICreate ("Simon", 450, 350, -1,-1, $WS_POPUP)
GuiRoundCorners($Simon, 0,0,450,350)
$Red = GUICtrlCreateLabel("", 0,0, 225, 175)
GUICtrlSetBkColor(-1, 0xAA0000)
GUICtrlSetOnEvent(-1, "Clicked")
$Green = GUICtrlCreateLabel("", 225,0, 225, 175)
GUICtrlSetBkColor(-1, 0x00AA00)
GUICtrlSetOnEvent(-1, "Clicked")
$Blue = GUICtrlCreateLabel("", 0,175, 225, 175)
GUICtrlSetBkColor(-1, 0x0000AA)
GUICtrlSetOnEvent(-1, "Clicked")
$Yellow = GUICtrlCreateLabel("", 225, 175, 225, 175)
GUICtrlSetBkColor(-1, 0xAAAA00)
GUICtrlSetOnEvent(-1, "Clicked")
GUISetState()


While $Count < $Max
    ToolTip("My Turn"&@CRLF&$Score, @DesktopWidth/2-5, 250)
    $Count +=1
    GUICtrlSetOnEvent($Red, "Filler")
    GUICtrlSetOnEvent($Green, "Filler")
    GUICtrlSetOnEvent($Blue, "Filler")
    GUICtrlSetOnEvent($Yellow, "Filler")
    For $x = 1 to $Count
        Demo($Pattern[$x])
    Next
    GUICtrlSetOnEvent($Red, "Clicked")
    GUICtrlSetOnEvent($Green, "Clicked")
    GUICtrlSetOnEvent($Blue, "Clicked")
    GUICtrlSetOnEvent($Yellow, "Clicked")
    ToolTip("Your Turn"&@CRLF&$Score, @DesktopWidth/2-5, 250)
    For $x = 1 to $Count
        $Timer = TimerInit()
        Do
            Sleep(10)
        Until TimerDiff($Timer)>=$TimePeriod or $Clicked=1
        If TimerDiff($Timer)>$TimePeriod then Lose()
        If $Clicked = 1 Then
            If $Last = $Pattern[$x] Then
                If $x = $Count then $Score+=1
            Else
                Lose()
            EndIf
            $Clicked = 0
        EndIf
    Next
WEnd

Func Lose()
    MsgBox(0,"Lose","You Lose."&@CRLF&"Score: "&$Score)
    Exit
EndFunc

Func Filler()
EndFunc 
    
Func Clicked()
    $Color = @GUI_CtrlID
    If $Color = $Red Then $Num = 1
    If $Color = $Green Then $Num = 2
    If $Color = $Blue Then $Num = 3
    If $Color = $Yellow Then $Num = 4
    Switch $Num
        Case 1
            GUICtrlSetBkColor($Red, 0xFF0000)
            Beep($RedFreq, $Duration)
            GUICtrlSetBkColor($Red, 0xAA0000)
            Global $Last =1
            Global $Clicked = 1
            Return 1
        Case 2
            GUICtrlSetBkColor($Green,0x00FF00)
            Beep($GreenFreq, $Duration)
            GUICtrlSetBkColor($Green, 0x00AA00)
            Global $Last =2
            Global $Clicked = 1
            Return 2
        Case 3
            GUICtrlSetBkColor($Blue, 0x0000FF)
            Beep($BlueFreq, $Duration)
            GUICtrlSetBkColor($Blue, 0x0000AA)
            Global $Last =3
            Global $Clicked = 1
            Return 3
        Case 4
            GUICtrlSetBkColor($Yellow, 0xFFFF00)
            Beep($YellowFreq,$Duration)
            GUICtrlSetBkColor($Yellow, 0xAAAA00)
            Global $Last =4
            Global $Clicked = 1
            Return 4
    EndSwitch
EndFunc

Func Demo($Color)
    Sleep(500)
    Switch $Color
        Case 1
            GUICtrlSetBkColor($Red, 0xFF0000)
            Beep($RedFreq, $Duration)
            GUICtrlSetBkColor($Red, 0xAA0000)
        Case 2
            GUICtrlSetBkColor($Green,0x00FF00)
            Beep($GreenFreq, $Duration)
            GUICtrlSetBkColor($Green, 0x00AA00)
        Case 3
            GUICtrlSetBkColor($Blue, 0x0000FF)
            Beep($BlueFreq, $Duration)
            GUICtrlSetBkColor($Blue, 0x0000AA)
        Case 4
            GUICtrlSetBkColor($Yellow, 0xFFFF00)
            Beep($YellowFreq,$Duration)
            GUICtrlSetBkColor($Yellow, 0xAAAA00)
    EndSwitch
EndFunc



;===============================================================================
;
; Function Name: GuiRoundCorners()
; Description: Rounds the corners of a window
; Parameter(s): $h_win
; $i_x1
; $i_y1
; $i_x3
; $i_y3
; Requirement(s): AutoIt3
; Return Value(s):
; Author(s): gaFrost
;
;===============================================================================
Func GuiRoundCorners($h_win, $i_x1, $i_y1, $i_x3, $i_y3)
    Dim $pos, $ret, $ret2
    $pos = WinGetPos($h_win)
    $ret = DllCall("gdi32.dll", "long", "CreateRoundRectRgn", "long", $i_x1, "long", $i_y1, "long", $pos[2], "long", $pos[3], "long", $i_x3, "long", $i_y3)
    If $ret[0] Then
        $ret2 = DllCall("user32.dll", "long", "SetWindowRgn", "hwnd", $h_win, "long", $ret[0], "int", 1)
        If $ret2[0] Then
            Return 1
        Else
            Return 0
        EndIf
    Else
        Return 0
    EndIf
EndFunc;==>_GuiRoundCorners