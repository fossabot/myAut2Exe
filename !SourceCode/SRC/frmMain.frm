VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "myAut2Exe >The Open Source AutoIT/AutoHotKey script decompiler<"
   ClientHeight    =   9675
   ClientLeft      =   2670
   ClientTop       =   1005
   ClientWidth     =   9300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9675
   ScaleWidth      =   9300
   Begin VB.ListBox List_Positions 
      Height          =   2010
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Right click for close"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Skip 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Skip >>>"
      Height          =   260
      Left            =   8400
      TabIndex        =   8
      ToolTipText     =   "Skip current step"
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer_TriggerLoad 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7800
      Top             =   120
   End
   Begin VB.Frame Fr_Options 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   8520
      Width           =   9135
      Begin VB.TextBox txt_OffAdjust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   16
         Text            =   "2C"
         ToolTipText     =   $"frmMain.frx":628A
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd_scan 
         Caption         =   "<<"
         Height          =   255
         Left            =   4245
         TabIndex        =   12
         ToolTipText     =   "Finds possible scriptstarts. ( Requires valid options for 'SrcFile_FileInst' and 'CompiledPathName' in options.)"
         Top             =   255
         Width           =   375
      End
      Begin VB.TextBox Txt_Scriptstart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         ToolTipText     =   $"frmMain.frx":631C
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmd_options 
         Caption         =   "More Options >>"
         Default         =   -1  'True
         Height          =   375
         Left            =   7560
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox Chk_TmpFile 
         Caption         =   "Don't delete temp files (for ex. compressed scriptdata)"
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.CheckBox Chk_verbose 
         Caption         =   "Verbose LogOutput"
         Height          =   195
         Left            =   5760
         MaskColor       =   &H8000000F&
         TabIndex        =   6
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label lbl_Adjustment 
         Caption         =   "Off_Adjust"
         Height          =   255
         Left            =   2550
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "StartOffset"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmd_MD5_pwd_Lookup 
      Caption         =   "Lookup Passwordhash"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      ToolTipText     =   "Copies hash to clipboard and does an online query."
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer_TriggerLoad_OLEDrag 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   480
   End
   Begin VB.CommandButton Cmd_About 
      Caption         =   "About"
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox ListLog 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "Double click to see more !"
      Top             =   6615
      Width           =   9135
   End
   Begin VB.ListBox List_Source 
      Appearance      =   0  'Flat
      Height          =   5685
      ItemData        =   "frmMain.frx":63B3
      Left            =   120
      List            =   "frmMain.frx":63B5
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.TextBox Txt_Script 
      Appearance      =   0  'Flat
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   600
      Width           =   9135
   End
   Begin VB.ComboBox Combo_Filename 
      Height          =   315
      ItemData        =   "frmMain.frx":63B7
      Left            =   120
      List            =   "frmMain.frx":63B9
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Text            =   "Drag the compiled AutoItExe / AutoHotKeyExe or obfucated script in here, or enter/paste path+filename."
      Top             =   120
      Width           =   9135
   End
   Begin VB.Shape Sh_ProgressBar 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   120
      Top             =   6360
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Shape Sh_ProgressBar 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   120
      Top             =   6480
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.Menu mu_Tools 
      Caption         =   "&Tools"
      Begin VB.Menu RegExp_Renamer 
         Caption         =   "&RegExp_Renamer"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mi_FunctionRenamer 
         Caption         =   "&FunctionRenamer"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mi_HexToBinTool 
         Caption         =   "&HexToBin_Binary() parser"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mi_CustomDecrypt 
         Caption         =   "&Custom_Decrypt() parser"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mi_GetAutoItVersion 
         Caption         =   "&GetAutoItVersion(Attention this executes the current exe)"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mi_SeperateIncludes 
         Caption         =   "&Seperate includes of *.au3"
      End
   End
   Begin VB.Menu mu_BugFix 
      Caption         =   "&BugFix"
      Begin VB.Menu mi_LocalID 
         Caption         =   "SetLocalID"
      End
   End
   Begin VB.Menu mu_Info 
      Caption         =   "&Info"
      Begin VB.Menu mi_About 
         Caption         =   "About"
         Visible         =   0   'False
      End
      Begin VB.Menu mi_Update 
         Caption         =   "&Update"
      End
      Begin VB.Menu mi_Forum 
         Caption         =   "&Forum"
      End
   End
   Begin VB.Menu mi_MD5_pwd_Lookup 
      Caption         =   "Lookup Passwordhash"
      Visible         =   0   'False
   End
   Begin VB.Menu mi_Reload 
      Caption         =   "&Reload"
      Enabled         =   0   'False
   End
   Begin VB.Menu mi_cancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'for mt_MT_Init to do a multiplation without 'overflow error'
'Private Declare Function Mul Lib "MSVBVM60.DLL" Alias "_allmul" (ByVal dw1 As Long, ByVal dw2 As Long, ByVal dw3 As Long, ByVal dw4 As Long) As Long

'Mersenne Twister
'Private Declare Function MT_Init Lib "MT.DLL" (ByVal initSeed As Long) As Long
'Private Declare Function MT_GetI8 Lib "MT.DLL" () As Long

'Private Declare Function Uncompress Lib "LZSS.DLL" (ByVal CompressedData$, ByVal CompressedDataSize&, ByVal OutData$, ByVal OutDataSize&) As Long
'Private Declare Function GetUncompressedSize Lib "LZSS.DLL" (ByVal CompressedData$, ByRef nUncompressedSize&) As Long

'Dim PE As New PE_info
Dim DeObfuscate As New ClsDeobfuscator
   
Dim myRegExp As New RegExp

Dim FilePath_for_Txt$

Const Combo_Filename_ClearList$ = "<Clear List>"
'Const MD5_CRACKER_URL$ = "http://gdataonline.com/qkhash.php?mode=txt&hash="

'Const MD5_CRACKER_URL$ = "http://www.md5cracker.de/crack.php?form=Cracken&md5="
'Const MD5_CRACKER_URL$ = "http://web18.server10.nl.kolido.net/md5cracker/crack.php?form=Cracken&md5="

Const MD5_CRACKER_URL$ = "http://hashkiller.com/api/api.php?md5="

'   http://www.milw0rm.com/cracker/info.php?'

Public ScriptLines

Public WithEvents Console As Console
Attribute Console.VB_VarHelpID = -1

Public WithEvents Console2 As Console
Attribute Console2.VB_VarHelpID = -1
Public Console2Output As clsStrCat
Attribute Console2Output.VB_VarHelpID = -1


'Public LogData As New clsStrCat
Private LogData()

Public bCmd_Skip_HasFocus As Boolean
Dim GUIEvent_InitialWidth(0 To 1) As Long

Dim GUIEvent_ProcessScale(0 To 1) As Double
Dim GUIEvent_Max(0 To 1) As Long

Dim GUIEvent_Width_before(0 To 1) As Long

Private Form_Initial_Height&
Private Form_Initial_Width&
Private Form_ResizeEventDisable As Boolean

Public StartLocations As New Collection

Private ListLogClickEventDisable As Boolean

Private Sub mi_cancel_Click()
   CancelAll = True
End Sub

Private Sub mi_CustomDecrypt_Click()
   CustomDecrypt
End Sub

Private Sub mi_Reload_Click()
   StartProcessing
End Sub

Private Sub mi_Update_Click()
   openURL "http://deioncube.in/files/MyAutToExe/index.html"
End Sub
Private Sub mi_Forum_Click()
   openURL "http://board.deioncube.in/showthread.php?tid=29"
End Sub

Public Sub GUIEvent_ProcessBegin(Target&, Optional BarLevel& = 0, Optional Skipable As Boolean = False)
On Error GoTo ERR_GUIEvent_ProcessBegin
   
   With Sh_ProgressBar(BarLevel)
      .Visible = True
'      .Tag = .Width
      .Width = 0
      
      GUIEvent_Max(BarLevel) = Target
      
    ' Avoid a division by Zero
      If Target > 0 Then
       ' Get stored length from when created the Form
         GUIEvent_ProcessScale(BarLevel) = GUIEvent_InitialWidth(BarLevel) / Target
      Else
         GUIEvent_ProcessScale(BarLevel) = 1
      End If
   End With
   
   GUIEvent_Width_before(BarLevel) = 0
   
   If BarLevel = 0 Then
      
      If Skipable Then
         GUI_SkipEnable
      Else
         GUI_SkipDisable
      End If
      
   End If
   
'   myDoEvents
   
ERR_GUIEvent_ProcessBegin:
End Sub

Public Sub GUIEvent_ProcessUpdate(CurrentValue&, Optional BarLevel& = 0)
On Error GoTo ERR_GUIEvent_ProcessUpdate
   With Sh_ProgressBar(BarLevel)
      
      .Width = CurrentValue * GUIEvent_ProcessScale(BarLevel)
      
      If (.Width - GUIEvent_Width_before(BarLevel)) > 10 Then
         GUIEvent_Width_before(BarLevel) = .Width
         
         On Error GoTo 0
         myDoEvents
         
      End If
   End With
ERR_GUIEvent_ProcessUpdate:
End Sub

Public Sub GUIEvent_Increase(PerCentToIncrease As Double, Optional BarLevel& = 0)
   
   Dim NewValue&
   NewValue = GUIEvent_ProcessScale(BarLevel) * PerCentToIncrease

   With Sh_ProgressBar(BarLevel)
      .Width = .Width + NewValue
   End With
End Sub

Public Sub GUIEvent_ProcessEnd(Optional BarLevel& = 0)
On Error GoTo ERR_GUIEvent_ProcessEnd
   With Sh_ProgressBar(BarLevel)
   
'      .Width = .Tag
      .Visible = False
   End With
   
   If BarLevel = 0 Then
      Cmd_Skip.Visible = False
   End If



'   myDoEvents
   
ERR_GUIEvent_ProcessEnd:
End Sub
Sub FL_verbose(Text)
   log_verbose H32(File.Position) & " -> " & Text
End Sub

Sub log_verbose(TextLine$)
   If Chk_verbose.value = vbChecked Then Log TextLine
End Sub



Sub FL(Text)
   Log H32(File.Position) & " -> " & Text
End Sub

Public Sub LogSub(TextLine$)
   Log "  " & TextLine
End Sub


Public Sub log2(TextLine$)
'   log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Public Sub Log(TextLine$, Optional LinePrefix$ = "")
On Error Resume Next
   
 ' Output Text /split into line and output it line wise
   Dim Line
   For Each Line In Split(TextLine, vbCrLf)
      ListLog.AddItem LinePrefix & Line
      
'      LogData.Concat LinePrefix & Line & vbCrLf
      ArrayAdd LogData, LinePrefix & Line
      
      
   Next
   
'   ListLog.AddItem H32(GetTickCount) & vbTab & TextLine
 
 ' Process windows messages (=Refresh display)
   If RangeCheck(ListLog.ListCount, 10000) Then
      ListLogClickEventDisable = True
       
    ' Scroll to last item ; when there are more than &h7fff items there will be an overflow error
      Dim ListCount&
      ListLog.ListIndex = ListLog.ListCount - 1
      
      ListLogClickEventDisable = False
      
      myDoEvents
      
   ElseIf (Rnd < 0.01) Then
      myDoEvents
      
   End If
End Sub

'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Public Sub Log_Clear()
On Error Resume Next
   ListLog.Clear
End Sub





Private Sub Chk_verbose_Click()
   Static value
   Checkbox_TriStateToggle Chk_verbose, value

End Sub

Private Sub Cmd_About_Click()
   FrmAbout.Show vbModal
End Sub

Private Sub ListLogShowCaption()
   Log Me.Caption
   Log String(80, "=")
End Sub

Private Sub ListLogClear()
   ListLog.Clear
   
'   LogData.Clear
    ArrayDelete LogData
   
End Sub


Private Sub cmd_options_Click()
   Frm_Options.Show
End Sub

Private Sub Cmd_Skip_Click()
   Cmd_Skip.Visible = False
   Skip = True
End Sub

Private Sub Combo_Filename_Additem(Text)
   With Combo_Filename
      
      Dim bAlreadyInList As Boolean
      Dim i
      For i = 0 To .ListCount - 1
         If Text = .List(i) Then
            If bAlreadyInList = True Then
             ' two occurence in list
               .RemoveItem i
            Else
             ' First occurence
               bAlreadyInList = True
            End If
            
         End If
      Next
      
      
      If bAlreadyInList = False Then
         If Text = Combo_Filename_ClearList Then
            Combo_Filename.AddItem Text, 0
        
         Else
            Combo_Filename.AddItem Text, 1
         
         End If
      End If
   End With

End Sub

Private Sub Combo_Filename_Clear()
   With Combo_Filename
      .Clear
      .AddItem Combo_Filename_ClearList
   End With
End Sub

Private Sub Combo_Filename_Click()
   With Combo_Filename
      
      If .Text = Combo_Filename_ClearList Then
         Combo_Filename_Clear
         
      ElseIf FileExists(.Text) Then
         Combo_Filename_Additem Trim(Combo_Filename)
         cmd_scan.Visible = True
         StartProcessing
      
      Else
         .RemoveItem .ListIndex
      
      End If
   End With
End Sub


Private Sub Console_OnInit(ProgramName As String)
    On Error Resume Next
'    GUI_SkipEnable
    GUIEvent_ProcessBegin UBound(ScriptLines), 0, True
    GUIEvent_ProcessBegin UBound(ScriptLines), 1
End Sub

Private Function GetCurLineFromTidyOutput(TextLine As String, MatchKeyWord$) As Long

   With myRegExp
      .Pattern = MatchKeyWord & RE_Group("\d+") & RE_NewLine
      
      Dim Match As Match
      For Each Match In .Execute(TextLine)
         GetCurLineFromTidyOutput = Match.SubMatches(0)
      Next
   End With

End Function
Private Sub Console_OnOutput(TextLine As String, ProgramName As String)
   On Error GoTo Console_OnOutput_err
 ' cut last newline
   Dim NewLinePos&
   NewLinePos = InStrRev(TextLine, vbCrLf)
   If NewLinePos > 0 Then
      Dec NewLinePos
   End If
   
  'TidyOutput - updateProcessBar
   Dim curline&
   curline = GetCurLineFromTidyOutput(TextLine, "Pre-processing record: ")
   If curline Then
      GUIEvent_ProcessUpdate curline, 1
   Else
      curline = GetCurLineFromTidyOutput(TextLine, "Processing record: ")
      If curline Then
         GUIEvent_ProcessUpdate curline, 0
      End If
   End If
 
 ' Show first 100 Lines
   ShowScriptPart ScriptLines, curline
   
 ' Log output
   Log Left(TextLine, NewLinePos), ProgramName & ": "
   
Console_OnOutput_err:
 Exit Sub
Log "ERR: " & Err.Description & "in  FrmMain.Console_OnOutput(TextLine , ProgramName )"
End Sub

 ' Show first 100 Lines
Private Sub ShowScriptPart(ScriptLines, curline&, Optional Lines& = 100)
   Dim ScriptLinesPreview_Start&
   ScriptLinesPreview_Start = Min(curline, UBound(ScriptLines))
   
   Dim ScriptLinesPreview_End&
   ScriptLinesPreview_End = Min(curline + Lines, UBound(ScriptLines))
   
   ReDim ScriptLinesPreview(ScriptLinesPreview_Start To ScriptLinesPreview_End)
   
   Dim i
   For i = ScriptLinesPreview_Start To ScriptLinesPreview_End
      ScriptLinesPreview(i) = ScriptLines(i)
   Next
   ShowScript Join(ScriptLinesPreview, vbCrLf)

End Sub

Private Sub Console_OnDone(ExitCode As Long)
'   GUI_SkipDisable
   GUIEvent_ProcessEnd 0
   GUIEvent_ProcessEnd 1
End Sub



Private Sub Console2_OnInit(ProgramName As String)
   Set Console2Output = New clsStrCat
End Sub

Private Sub Console2_OnOutput(TextLine As String, ProgramName As String)
   Console2Output.Concat TextLine
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
'      Case vbKeyDelete, vbKeyBack
'         ListLogClear
         
      Case vbKeyEscape
         CancelAll = True

   End Select


End Sub



Private Sub Form_Resize()
   If Form_ResizeEventDisable Then Exit Sub
   Form_ResizeEventDisable = True
   
   On Error GoTo Form_Resize_err
      
      If WindowState = vbMaximized Then
         WindowState = vbNormal
         Log "Sorry Form_Resize is not supported so maximize don't makes sense."
     
      End If
      
      If WindowState = vbNormal Then
         If (Me.Height <> Form_Initial_Height) Or _
            (Me.Width <> Form_Initial_Width) Then
               
             Log "Sorry Form_Resize is not supported " & Me.Height
               
             Me.Height = Form_Initial_Height
             Me.Width = Form_Initial_Width
         End If
      End If
      
Form_Resize_err:
   Form_ResizeEventDisable = False
End Sub
Function WH_Open() As Long
   On Error Resume Next
   If Frm_Options.chk_disableWinhex = vbChecked Then Exit Function
   If IsCurrentFileValid = False Then Exit Function
   
   
   Dim FileName$
   
   FileName = Combo_Filename
' As ClsFilename
   Dim retval&
   retval = WHX_Init(1)
   
   Dim msg$
   Select Case retval
      Case 2
         msg = "Success (limited)"
      Case 1
         msg = "Success"
      Case 0
         msg = "General Error"
      Case -1
         msg = "WinHex installation not ready"
      Case -2
         msg = "APIVersion incorrect(should be 1)"
      Case Else
         msg = "Unknown"
   End Select
   If retval <> 1 Then
      Log "Winhex Error: " & retval & " " & msg
   End If
   
 ' File already opened ?
   Dim CurrentFile$
   CurrentFile = Space(256)
   retval = WHX_GetCurObjName(CurrentFile)
   szNullCutProc CurrentFile
   If CurrentFile <> Trim(FileName) Then
  
    ' Open file
      retval = WHX_Open(FileName)
      If retval <> 1 Then
         
         Log "Winhex Error: WHX_Open() returned: " & retval
         
       ' Set Default
         Dim WH_Path As New ClsFilename
         WH_Path = App.Path & "\data\WinHex\"
         
       ' get path from registry
         Console2.ShellExConsole "reg", "query ""HKEY_CLASSES_ROOT\WHSFile\DefaultIcon"""
         WH_Path = Split(Console2Output, """")(1)
         WH_Path = InputBox("Please enter path to Winhex", "Winhex not found", WH_Path.Path)
         WH_Path.NameWithExt = "."
         
         Dim Param$
         Param = "ADD ""HKCU\Software\X-Ways AG\WinHex"" /d " & Quote(WH_Path) & " /f /v Path /t REG_SZ"
         Log "Running: Reg " & Param
         Console2.ShellExConsole "Reg", Param
         'HKEY_CLASSES_ROOT\WHSFile\DefaultIcon
      End If
   End If
   

End Function

Function WH_Goto(ByVal Position As Currency) As Long
   WHX_Goto VB2API(Position)
End Function

Sub WH_close()
   WHX_Done
End Sub



Private Sub List_Positions_Click()
   WH_Open
   WH_Goto HexToInt(List_Positions.Text)
'   WH_close
End Sub

Private Sub List_Positions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = MouseButtonConstants.vbRightButton Then
      List_Positions.Visible = False
      lbl_Adjustment.Visible = False
      txt_OffAdjust.Visible = False
   End If
   
End Sub

Private Sub ListLog_Click()
   On Error Resume Next
   If ListLogClickEventDisable Then Exit Sub
   With ListLog
      If .Text Like "???????? -> *" Then
         Debug.Print .Text
         Dim Offset$
         Offset = Left(.Text, 8)
         WH_Open
         WH_Goto HexToInt(Offset)
      End If
   End With
End Sub

Private Sub ListLog_DblClick()
   frmLogView.txtlog = Replace( _
                        FrmMain.Log_GetData, _
                        vbNullChar, ".")
   frmLogView.Show
End Sub

Private Sub ListLog_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDelete, vbKeyBack
         ListLogClear
         
   End Select
   
   Form_KeyDown KeyCode, Shift

End Sub


Private Sub ListLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = MouseButtonConstants.vbRightButton Then
      ListLogClear
   End If
End Sub

Private Sub mi_GetAutoItVersion_Click()

'3.1.1 ( 7th Apr   , 2005) (Release)
'...
'
'3.2.0 (12th August, 2006) (Release)
' Added: /AutoIt3ExecuteScript command line option.

   
   On Error GoTo ERR_mi_GetAutoItVersion_Click
 
 ' Minimize during execution
   Dim WindowState_old
   WindowState_old = WindowState
   WindowState = vbMinimized
      
 ' Log & execute
   Dim ShellCommandParams$
   ShellCommandParams = " /AutoIt3ExecuteLine ""Exit(999+MsgBox(0x40,'AutoIt version of ' & @ScriptName & ' is', @AutoItVersion,10))"""
   
   Log "GetAutoItVersion executes: " & Quote(Combo_Filename) & " " & ShellCommandParams
   Dim ShellExitCode&
   ShellExitCode = ShellEx(Combo_Filename, ShellCommandParams, vbNormalFocus)
   
   If ShellExitCode = 1000 Then
      Log "Execution finished; ExitCode: " & ShellExitCode
      openURL "http://www.autoitscript.com/autoit3/files/archive/autoit"
   Else
      Log "Execution failed; ExitCode: " & ShellExitCode
   End If
'                   ' Run "LZSS.exe -d *.debug *.au3" to extract the script (...and wait for its execution to finish)
'                     Dim LZSS_Output$, ExitCode&
'                     LZSS_Output = Console.ShellExConsole( _
'                              App.Path & "\" & "data\LZSS.exe", _
'                              "-d " & Quote(.FileName) & " " & Quote(OutFileName.FileName), _
'                              ExitCode)
'
'                     If ExitCode <> 0 Then Log LZSS_Output, "LZSS_Output: "
   
   
   
 ' Restore
   WindowState = WindowState_old
Exit Sub
ERR_mi_GetAutoItVersion_Click:
   Log "ERROR " & Err.Description
End Sub

Private Sub mi_HexToBinTool_Click()
   HexToBinTool
End Sub

Private Sub mi_LocalID_Click()
InputValue:
    On Error GoTo ERR_mi_LocalID_Click
    
    Dim InputboxTmp$
    InputboxTmp = InputBox( _
        "You will need to adjust that value if your Windows is not a german of english one and you are getting errors(checksum fail;modified JB LZSS Signature) when decompiling ya own freshly compiled files or the included examples." _
        & "See '!SourceCode\languages-ids.txt' for your LCID. '0' will tell VB to use the current LCID. Any invalid value will reset this to the default(German).", _
        "Enter your LocalID(as hex) for handling strings", H16(LocaleID))

    If InputboxTmp = "" Then Exit Sub
    
    Dim InputboxTmpVal&
    
    InputboxTmpVal = HexToInt(InputboxTmp)
    If InputboxTmpVal <> 0 Then
        RangeCheck InputboxTmpVal, &H5000, &H400, "LCID is not inside the valid range"
    End If
    
    LocaleID = InputboxTmpVal
    
ERR_mi_LocalID_Click:
Select Case Err
    Case 0
    Case 13
        LocaleID = LocaleID_GER
    
    Case Else
        MsgBox Err.Description
        Resume InputValue
End Select
End Sub

'Copies hash to clipboard and does an online query.
Private Sub mi_MD5_pwd_Lookup_Click()
   Clipboard.Clear
   Clipboard.SetText MD5PassphraseHashText

   openURL MD5_CRACKER_URL$ & LCase$(MD5PassphraseHashText)

End Sub

Private Sub HexToBinTool()

   Dim FileName As New ClsFilename
   FileName.FileName = InputBox("FileName:", "", Combo_Filename)
   
   If FileName.FileName = "" Then Exit Sub
   

   Dim Data$
   Data = FileLoad(FileName.FileName)

   
   
   Dim myRegExp As New RegExp
   With myRegExp

      
      .Pattern = RE_WSpace(RE_Group("\w*?") & "\(", "[""']0x", _
                            "[0-9A-Fa-f]" & "*?", "[""']", ".*?", "\)")
      Dim matches As MatchCollection
      Set matches = .Execute(Data)
      Dim FunctionName$
      If matches.Count < 1 Then
         
         FunctionName = "FnNameOfBinaryToString"
      Else
      
         FunctionName = matches(0).SubMatches(0)
      End If
   

      FunctionName = InputBox("FunctionName:", "", FunctionName)
      

      
      .Global = True
      .Pattern = RE_WSpace(RE_Literal(FunctionName), _
                            "\(", "[""']0x", _
                               RE_Group("[0-9A-Fa-f]" & "*?"), _
                            "[""']", _
                              RE_Group_NonCaptured( _
                                 RE_WSpace( _
                                   ",", _
                                   RE_Group("[1-4]")) _
                                 ) & "?", _
                            "\)")
                            
      Set matches = .Execute(Data)
      Dim Match As Match
      For Each Match In matches
         With Match
         
            Dim IsPrintable As Boolean
            Dim BinData$
            BinData = MakeAutoItString( _
               HexStringToString(.SubMatches(0), IsPrintable, .SubMatches(1)))
            
            If IsPrintable Then
               Log "Replacing: " & BinData & " <= " & .value
               ReplaceDo Data, .value, EncodeUTF8(BinData), .FirstIndex, 1
            Else
               Log "Skipped replace(not printable): " & MakePrintable(BinData) & " <= " & .value
            End If
            
         End With
      Next
      
      
      
   End With
   
   If matches.Count Then
      FileName.Name = FileName.Name & "_HexToBin"
       
    ' Save
      FileSave FileName.FileName, Data

       
       Log matches.Count & " replacements done."
       Log "File save to: " & FileName.FileName
   Else
      Log "Nothing found."
   End If
   
   

End Sub

Private Sub CustomDecrypt()

   Dim FileName As New ClsFilename
   FileName.FileName = InputBox("Note: The CustomDecrypt only makes sense together with the VB6-IDE !" & vbCrLf & _
                     "" & vbCrLf & _
                     "It helps if you encounter stuff like this: 'MsgBox(0, Fn04B6(""dHBKQL LWW~W"", ""FI""),...'" & vbCrLf & _
                     "" & vbCrLf & _
                     "FileName:", "Programmers only!", Combo_Filename)
   
   If FileName.FileName = "" Then Exit Sub
   

   Dim Data$
   Data = FileLoad(FileName.FileName)

   
   
   Dim myRegExp As New RegExp
   With myRegExp

      
      .Pattern = RE_WSpace(RE_Group("\w*?") & "\(", "[""']0x", _
                            "[0-9A-Fa-f]" & "*?", "[""']", ".*?", "\)")
      Dim matches As MatchCollection
      Set matches = .Execute(Data)
      Dim FunctionName$
      If matches.Count < 1 Then
         
         FunctionName = "FnNameOfBinaryToString"
      Else
      
         FunctionName = matches(0).SubMatches(0)
      End If
   
   
FunctionName = "_deCode"

      FunctionName = InputBox("FunctionName:", "", FunctionName)
      

'_deCode("rATNQ7", "BA")

      .Global = True
      
      
    'We'll just care about "doublequoted" Strings
      Const RE_AU3_QUOTE$ = "[""]"
      
      Const RE_AU3_String$ = _
         RE_AU3_QUOTE & "(" & _
             "[^""]*?" & _
         ")" & RE_AU3_QUOTE
      
      .Pattern = RE_WSpace(RE_Literal(FunctionName), _
                            "\(", _
                              RE_AU3_String$, _
                              ",", _
                              RE_AU3_String$, _
                            "\)")
                            
      Set matches = .Execute(Data)
      Dim Match As Match
      For Each Match In matches
         With Match
         
            Dim IsPrintable As Boolean
            Dim BinData$
            

            BinData$ = CryptCall(.SubMatches(0), .SubMatches(1))
            BinData = MakeAutoItString(BinData$)
            
 '           If IsPrintable Then
               Log "Replacing: " & BinData & " <= " & .value
               ReplaceDo Data, .value, EncodeUTF8(BinData), .FirstIndex, 1
  '          Else
  '             Log "Skipped replace(not printable): " & MakePrintable(BinData) & " <= " & .value
  '          End If
            
         End With
      Next
      
      
      
   End With
   
   If matches.Count Then
      FileName.Name = FileName.Name & "_CustomDecrypt"
       
    ' Save
      FileSave FileName.FileName, Data

       
       Log matches.Count & " replacements done."
       Log "File save to: " & FileName.FileName
   Else
      Log "Nothing found."
   End If
   
   

End Sub

Private Sub Form_Load()

 '  CamoGet


 ' Create ConsoleObj
   Set Console = New Console
   
   Set Console2 = New Console
  

   GUIEvent_InitialWidth(0) = Sh_ProgressBar(0).Width
   GUIEvent_InitialWidth(1) = Sh_ProgressBar(1).Width


'
'   Dim myRegExp As New RegExp
'   With myRegExp 'New RegExp
'      .Global = True
'      .MultiLine = True
'
'      myRegExp.Pattern = "(?:SHELLEXECUTE)*(EXECUTE\(\s*\$A\s*\))?"
'
'      Dim test$, Out$
'      test = "SHELLEXECUTE($A) | WHILE EXECUTE($A) | EXECUTE($A606)"
'      Out = myRegExp.Replace(test, "--RP--")
'
'   End With


'   Dim str$, i&
'   Dim leni%
'   Do
'      BenchStart
'      For i = 0 To 5000000
'         Dim a
'         ArrayEnsureBounds a
'
'      Next
'      BenchEnd
'   Loop While True

   
   FrmMain.Caption = FrmMain.Caption & " " & App.Major & "." & App.Minor & " build(" & App.Revision & ")"
   
   LocaleID = LocaleID_GER
   FormSettings_Load Me, "txt_OffAdjust"
   
   
'  'Just for the case of the first run
'   txt_FILE_DecryptionKey_Change
'   txt_FILE_DecryptionKey_Validate True
   Load Frm_Options
   
 ' Ensure combo has "<Clear List>" item
   Combo_Filename_Additem Combo_Filename_ClearList
   
   
   
   'Extent Listbox width
   Listbox_SetHorizontalExtent ListLog, 6000
   
   ListLogClear
   ListLogShowCaption

 
 ' Commandlinesupport   :)
   ProcessCommandline
   

  'Show Form if SilentMode is not Enable
   If IsOpt_RunSilent = False Then
      Form_ResizeEventDisable = True
      
      Me.Show
      
      Form_ResizeEventDisable = False
      
      Form_Initial_Height = Me.Height
      Form_Initial_Width = Me.Width
   End If
  
  'Open the File that was set by the commandline
   If IsCommandlineMode Then
      Combo_Filename = FileName
   Else
    ' try Load file in the 'File textbox'
      Timer_TriggerLoad.Enabled = True
   End If

End Sub
   
   
Private Sub ProcessCommandline()

   Dim CommadLine As New CommandLine
   With CommadLine
   
      If .NumberOfCommandLineArgs Then
      
         Log "Cmdline Args: " & .CommandLine
         
         Dim arg
         For Each arg In .getArgs
            
           'Check for options
            If arg Like "[/-]*" Then

               If arg Like "?[qQ]" Then
                  IsOpt_QuitWhenFinish = True
                  LogSub "Option 'QuitWhenFinish' enabled."
                  
               ElseIf arg Like "?[sS]" Then
                  IsOpt_RunSilent = True
                  LogSub "Option 'RunSilent' enabled."
                  
               Else
                  LogSub "ERR_Unknow option: '" & arg & "'"
                  
               End If
               
          ' Check if CommandArg is a FileName
            Else
           
               If IsCommandlineMode Then
                  LogSub "ERR_Invalid Argument ('" & arg & "') filename already set."
                  
               Else
                  If FileExists(arg) Then
                     IsCommandlineMode = True
                     FileName = arg
                     LogSub "FileName : " & arg
                  Else
                     LogSub "ERR_Invalid Argument. Can't open file '" & arg & "'"
                  End If
               End If
               
            End If
         Next
      End If
   End With

   'Verify
   If IsOpt_RunSilent And Not (IsOpt_QuitWhenFinish) Then
      LogSub "ERR 'RunSilent' only makes sence together with 'QuitWhenFinish'. As long as you don't also enable 'QuitWhenFinish' 'RunSilent' is ignored "
      IsOpt_RunSilent = False
   End If

End Sub


Public Function Log_GetData$()
   Log_GetData = Join(LogData, vbCrLf)
End Function
'   Dim LogData As New clsStrCat
'   LogData.Clear
'   Dim i
'   If (ListLog.ListCount >= 0) Then
'      For i = 0 To ListLog.ListCount
'         LogData.Concat (ListLog.List(i) & vbCrLf)
'      Next
'   Else
'      For i = 0 To &H7FFE
'         LogData.Concat (ListLog.List(i) & vbCrLf)
'      Next
'      LogData.Concat "<Data cut due to VB-listbox.ListCount bug :( >"
'
''   Do While ListLog.ListCount < 0
''      LogData.Concat (ListLog.List(&H7FFF) & vbCrLf)
''      ListLog.RemoveItem &H7FFF
''   Loop
'
'   End If
'
'   Log_GetData = LogData.value
'
'End Function

Private Sub Form_Unload(Cancel As Integer)
   
   FormSettings_Save Me
  
 'Close might be clicked 'inside' some myDoEvents so
 'in case it was do a hard END
 '   End
 
   Dim form_i As Form
   For Each form_i In Forms
     Unload form_i
   Next
 
   APP_REQUEST_UNLOAD = True
   
   WH_close
End Sub



Sub openURL(url$)
   Dim hProc&
   hProc = ShellExecute(0, "open", url, "", "", 1)
End Sub


Private Sub mi_FunctionRenamer_Click()
   Load FrmFuncRename
'   If FileExists(Combo_Filename) Then
'      FrmFuncRename.Txt_Fn_Org_FileName = Combo_Filename
'   Else
'
'
'   End If
   
   FrmFuncRename.Show ' vbModal
'   Unload FrmFuncRename
   
End Sub

Private Sub mi_SeperateIncludes_Click()
   Dim File$
   File = InputBox("Normally seperating includes is done automatically after you decompiled some au3.exe(of old none tokend format)." & vbCrLf & _
          "However that tool is useful in the case you have some decompiled *.au3 with these '; <AUT2EXE INCLUDE-START: C:\ ...' comments you like to process." & vbCrLf & vbCrLf & _
          "Please enter(/paste) full path of the file: (Or drag it into the myAutToExe filebox and then run me again)", "Manually run 'seperate au3 includes' on file", Combo_Filename)
   If File <> "" Then
      FileName.FileName = File
      SeperateIncludes
   End If
End Sub





Private Sub RegExp_Renamer_Click()
   FrmRegExp_Renamer.Show ' vbModal
'   Unload FrmRegExp_Renamer
End Sub

Private Sub Timer_TriggerLoad_OLEDrag_Timer()
   Timer_TriggerLoad_OLEDrag.Enabled = False
   Combo_Filename = FilePath_for_Txt
End Sub


Private Sub Timer_TriggerLoad_Timer()
   Timer_TriggerLoad.Enabled = False
   
   Combo_Filename_Change

End Sub

Private Function IsCurrentFileValid() As Boolean
   IsCurrentFileValid = FileExists(Combo_Filename)
End Function


Private Sub Combo_Filename_Change()
  'Avoid to be triggered during load settings
   If Combo_Filename.Enabled = False Then Exit Sub
  
   On Error GoTo Combo_Filename_err
   
   Dim bFileExists As Boolean
   bFileExists = IsCurrentFileValid
   
   cmd_scan.Visible = bFileExists
   
   If bFileExists Then
      
      Combo_Filename_Additem Trim(Combo_Filename)
      StartProcessing
   
   End If
Combo_Filename_err:
End Sub

Sub StartProcessing()
  On Error GoTo StartProcessing_err
         
  CancelAll = False
  
' Block any new files during DoEvents
  Combo_Filename.Enabled = False
  mi_Reload.Enabled = False
  mi_cancel.Enabled = True
  
  
  
' Reset ProgressBars
  GUIEvent_ProcessEnd 0
  GUIEvent_ProcessEnd 1
  
  
' Clear Log (expect when run via commandline)
  If IsCommandlineMode = False Then
     ListLogClear
     ListLogShowCaption
  End If
  Txt_Script = ""
  
  FileName = Combo_Filename
  

' Log String(80, "=")
' log "           -=  " & Me.Caption & "  =-"
     
  On Error Resume Next
  
  
  Decompile
  
  
  If Err = ERR_CANCEL_ALL Then GoTo StartProcessing_err:
  If Err Then
     Log "ERR: " & Err.Description
  End If
  
  FileName = ExtractedFiles("MainScript")
     
  DeToken
     If Err = ERR_CANCEL_ALL Then GoTo StartProcessing_err:
     If Err Then Log "ERR: " & Err.Description

     Log String(79, "=")
     On Error Resume Next
     
  DeObfuscate.DeObfuscate
     If Err = ERR_CANCEL_ALL Then GoTo StartProcessing_err:
     Select Case Err
     Case 0, ERR_NO_OBFUSCATE_AUT
        If Frm_Options.Chk_RestoreIncludes.value = vbChecked Then _
           SeperateIncludes
           
           
     Case Else
        Log Err.Description
        
     End Select


    If IsTextFile Then CheckScriptFor_COMPILED_Macro
  
 
Err.Clear
GoTo StartProcessing_err
      
' ErrorHandle for resume from Errors
DeToken:
   Log String(79, "=")
   DeToken

DeObfuscate:
   Log String(79, "=")
   DeObfuscate.DeObfuscate
      
StartProcessing_err:

' Add some fileName if it weren't done during decompile()
  If IsAlreadyInCollection(ExtractedFiles, "MainScript") = False Then
     ExtractedFiles.Add File.FileName, "MainScript"
  End If


' Note: Resume is necessary to reenable Errorhandler
'       Else the VB-standard Handler will catch the error -> Exit Programm
  Select Case Err
  Case 0
  
  Case ERR_NO_AUT_EXE
     Log Err.Description
     Resume DeToken
  
  Case NO_AUT_DE_TOKEN_FILE
     Log Err.Description
     Resume DeObfuscate
  
  Case ERR_NO_OBFUSCATE_AUT
'    Log Err.Description
     Resume StartProcessing_err
  
  Case ERR_CANCEL_ALL
     Log "Processing CANCELED!  " & Err.Description
     Resume Finally
     
  Case Else
     Log Err.Description
     Resume StartProcessing_err
  End Select
'-----------------------------------------------
   
Finally:
' Save Log Data
  On Error Resume Next
  Resume Finally
  
  
  FileName = ExtractedFiles("MainScript").FileName
  FileName.NameWithExt = FileName.Name & "_myExeToAut.log"
  
  Log ""
  Log "Saving Logdata to : " & FileName.FileName
  FileSave FileName.FileName, Log_GetData

' process Quit
  If APP_REQUEST_UNLOAD Then End
  
' Allow Reload / Block Cancel
  Combo_Filename.Enabled = True
  mi_Reload.Enabled = True
  mi_cancel.Enabled = False
  
  
  
  IsCommandlineMode = False
  If IsOpt_QuitWhenFinish Then Unload Me
 
End Sub


Private Function OpenFile(Target_FileName As ClsFilename) As Boolean
   
   On Error GoTo Scanfile_err
   Log "------------------------------------------------"

   Log Space(4) & Target_FileName.NameWithExt

   File.Create Target_FileName.mvarFileName, Readonly:=True
   
   Me.Show

Err.Clear
Scanfile_err:
Select Case Err
   Case 0

   Case Else
      Log "-->ERR: " & Err.Description

End Select
   
End Function


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub List_Source_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub ListLog_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub Combo_Filename_KeyDown(KeyCode As Integer, Shift As Integer)
   Form_KeyDown KeyCode, Shift
End Sub

Private Sub txt_OffAdjust_Change()
   updateStartLocations_List
End Sub

Private Sub Txt_Script_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub Combo_Filename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   File_DragDrop Data
End Sub

Private Sub File_DragDrop(Data As DataObject)
   
   On Error GoTo Combo_Filename_OLEDragDrop_err
   
   FilePath_for_Txt = Data.Files(1)
   Timer_TriggerLoad_OLEDrag.Enabled = True
   

Combo_Filename_OLEDragDrop_err:
Select Case Err
Case 0

Case Else
   Log "-->Drop'n'Drag ERR: " & Err.Description

End Select

End Sub


Private Sub Txt_Script_Change()
  If Len(Txt_Script) >= 65535 Then
      Txt_Script.ToolTipText = "Notice: Display limited to 65535 Bytes. File is bigger."
  Else
      Txt_Script.ToolTipText = ""
  End If
End Sub

Private Sub Txt_Script_KeyDown(KeyCode As Integer, Shift As Integer)
   Cancel = KeyCode <> vbKeySpace
   
   If KeyCode = vbKeyEscape Then
      CancelAll = True
   End If
   
End Sub



Public Sub LogSourceCodeLine(TextLine$)
   If Chk_verbose.value = vbChecked Then
   
      On Error Resume Next
      With List_Source
         .AddItem TextLine
       
       ' Process windows messages (=Refresh display)
         If Rnd < 0.01 Then
             ' Scroll to last item
            .ListIndex = .ListCount - 1
         End If
         
      End With
   End If
End Sub




Private Sub Txt_Scriptstart_Change()
   On Error Resume Next
   Dim scriptstart&
   scriptstart = HexToInt(Txt_Scriptstart)
   
   Frm_Options.Chk_NormalSigScan.Enabled = (Err.Number <> 0)
   
   If Txt_Scriptstart.Enabled Then
      WH_Open
      WH_Goto CInt(scriptstart)
   End If
End Sub
Private Sub cmd_scan_Click()
   LongValScan_Init
 
 ' New Script 0xADBC / 0F820 WideChar_Unicode
   LongValScan XORKEY_SrcFile_FileInstSize:=Xorkey_SrcFile_FileInstNEW_Len, _
               XORKEY_CompiledPathNameSize:=Xorkey_CompiledPathNameNEW_Len, _
               CHARSIZE:=2
   
   Log "Testing for old AU3-scripttype"
 ' Old Script 0x29BC / 29AC ACCII
   LongValScan XORKEY_SrcFile_FileInstSize:=Xorkey_SrcFile_FileInst_Len, _
               XORKEY_CompiledPathNameSize:=Xorkey_CompiledPathName_Len, _
               CHARSIZE:=1


End Sub
Private Sub List_Positions_DblClick()
   Txt_Scriptstart = List_Positions.Text
   FrmMain.StartProcessing
End Sub


Private Sub List_Positions_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then List_Positions_DblClick
End Sub


Public Sub updateStartLocations_List()
   On Error Resume Next 'GoTo updateStartLocations_List_err

   With FrmMain.List_Positions
      
      Dim adjustment&
      adjustment = HexToInt(FrmMain.txt_OffAdjust)
      
      .Clear
      
      lbl_Adjustment.Visible = True
      txt_OffAdjust.Visible = True
      .Visible = True
'      .SetFocus
      
      Dim Location
      For Each Location In StartLocations
         Dec Location, adjustment
         .AddItem Right(H32(Location), 6)
      Next
   End With
updateStartLocations_List_err:
End Sub

Private Sub Txt_Scriptstart_Validate(Cancel As Boolean)
   WH_close
End Sub
