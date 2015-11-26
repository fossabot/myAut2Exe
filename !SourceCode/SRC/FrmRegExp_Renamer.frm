VERSION 5.00
Begin VB.Form FrmRegExp_Renamer 
   Caption         =   "RegEx Renamer (alpha!)"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_RegExpSave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmd_RegExpLoad 
      Appearance      =   0  'Flat
      Caption         =   "Load"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   735
   End
   Begin VB.CheckBox chk_Simple 
      Caption         =   "Simple"
      Height          =   375
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Enables Simple Mode - that do not adds 0001,0002... to each match"
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Cmd_Test 
      Appearance      =   0  'Flat
      Caption         =   "&Test"
      Height          =   495
      Left            =   11880
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmd_help 
      Appearance      =   0  'Flat
      Caption         =   "?"
      Height          =   495
      Left            =   12480
      TabIndex        =   10
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Cmd_Save 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   495
      Left            =   11880
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Matches"
      Height          =   5655
      Left            =   8400
      TabIndex        =   6
      Top             =   2400
      Width           =   4335
      Begin VB.ListBox List_Matches 
         Appearance      =   0  'Flat
         Height          =   5880
         ItemData        =   "FrmRegExp_Renamer.frx":0000
         Left            =   120
         List            =   "FrmRegExp_Renamer.frx":0002
         TabIndex        =   8
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txt_Matches 
         BorderStyle     =   0  'None
         Height          =   5055
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   7
         Text            =   "FrmRegExp_Renamer.frx":0004
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.ListBox List_log 
      Appearance      =   0  'Flat
      Height          =   1005
      ItemData        =   "FrmRegExp_Renamer.frx":002C
      Left            =   120
      List            =   "FrmRegExp_Renamer.frx":002E
      TabIndex        =   5
      Top             =   7080
      Width           =   8175
   End
   Begin VB.TextBox txt_FileName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Text            =   "<Drag some au3-file in here>"
      Top             =   120
      Width           =   11655
   End
   Begin VB.TextBox txt_ReplaceString 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "FrmRegExp_Renamer.frx":0030
      Top             =   600
      Width           =   11655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   8175
      Begin VB.TextBox txt_Replace 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Tag             =   ">>>  Dear VB6-Dev: Please note that right behind me is another TextBox  <<<"
         Text            =   "FrmRegExp_Renamer.frx":019A
         Top             =   240
         Width           =   7935
      End
      Begin VB.TextBox txt_Original 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Text            =   "FrmRegExp_Renamer.frx":026C
         Top             =   240
         Width           =   7935
      End
   End
   Begin VB.CommandButton cmd_Quit 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Height          =   255
      Left            =   12120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Lbl_Status 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   8160
      Width           =   8295
   End
   Begin VB.Label Label1 
      Caption         =   """<RegExpSearchPattern(Variable)>"" -> ""<ReplaceString>"" ; Comments"
      Height          =   255
      Left            =   105
      TabIndex        =   13
      Top             =   405
      Width           =   9135
   End
End
Attribute VB_Name = "FrmRegExp_Renamer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScriptData As StringReader
Const LOAD_SAVE_FILENAME$ = "myAutToExe_RegExpRenamerSearchPatter.txt"

'Private Enum TSearchReplacePattern
   Const Pattern_Search& = 0
   Const Pattern_Replace& = 1
   Const Pattern_Comments& = 2
'End Enum

'Private SearchReplaceItem As TSearchReplacePattern
Private SearchReplacePattern As Collection
   
Dim SearchReplace_Matches As MatchCollection
   
   
Private Sub SeperateSearchReplacePattern()
   Dim Line
   Line = txt_ReplaceString
   
   With New RegExp
      .Global = True
      .MultiLine = True
      .IgnoreCase = True
      
      .Pattern = RE_WSpace("""(.*)""", "->", _
                           """(.*)""", _
                           RE_Group_NonCaptured(";(.*)") & "?")
      
      Set SearchReplacePattern = New Collection
      
      Dim SearchReplace_Match As Match
      For Each SearchReplace_Match In .Execute(txt_ReplaceString)
         
      
'         With SearchReplace_Match
'            SearchReplaceItem.Pattern_Search = .SubMatches(1)
'            SearchReplaceItem.Pattern_Replace = .SubMatches(2)
'            SearchReplaceItem.Pattern_Comments = .SubMatches(3)
'         End With
         SearchReplacePattern.Add SearchReplace_Match.SubMatches
      Next
         
      
   End With
End Sub



Private Sub chk_Simple_Click()
   Refresh_Preview
End Sub

Private Sub Refresh_Preview()
' Note: even If txt_Original is passed byref
'       txt_Original.text is not change
   Apply txt_Original
End Sub


Private Sub cmd_help_Click()
   ShellExecute 0, "open", "Doc\regexp.htm", "", "", 0
End Sub

Private Sub cmd_Quit_Click()
   Me.Hide
End Sub

Private Sub cmd_RegExpLoad_Click()
   On Error Resume Next
   
   txt_ReplaceString = FileLoad(LOAD_SAVE_FILENAME)
   If Err Then Log Err.Description
   
End Sub

Private Sub cmd_RegExpSave_Click()
   On Error Resume Next
   
   FileSave LOAD_SAVE_FILENAME, txt_ReplaceString
   If Err Then Log Err.Description
   
End Sub

Private Sub Cmd_Save_Click()

   OpenAndFill FileName.FileName
   
   ScriptData.Data = Apply(ScriptData.Data)
   
   Dim OutputFileName As New ClsFilename
   
   OutputFileName.FileName = FileName.FileName
   OutputFileName.Name = OutputFileName.Name & "_Renamed"
   
   FileSave OutputFileName.FileName, _
            DecodeUTF8(Mid(ScriptData.Data, 4))
            
   txt_Original = ScriptData.Data

   
End Sub

Private Sub Cmd_Test_Click()
   On Error Resume Next
   Apply ScriptData.Data, True
End Sub

Private Sub Form_Load()
   'txt_Filename =
End Sub

Private Sub List_Matches_Click()
 On Error GoTo List_Matches_err

   With SearchReplace_Matches(List_Matches.ListIndex)
   
      OpenAndFill txt_FileName
      txt_Replace = Mid(txt_Original, Max(.FirstIndex - 400, 0))
   
   
      txt_Replace.SelStart = 400
      txt_Replace.SelLength = .Length
      txt_Replace.SetFocus

      
   End With
   
'   List_Matches.SetFocus
Exit Sub

List_Matches_err:
   Resume List_Matches_Load

List_Matches_Load:
On Error GoTo List_Matches_Load_err
   OpenAndFill txt_FileName, SearchReplace_Matches(List_Matches.ListIndex).FirstIndex
   txt_Replace = txt_Original
List_Matches_Load_err:
Exit Sub

End Sub

Private Sub txt_Filename_Change()
   If FileExists(txt_FileName) Then
      FileName = txt_FileName
   
      OpenAndFill FileName.FileName
      Refresh_Preview
   End If

End Sub

Private Sub Log_Clear()
   List_log.Clear
End Sub

Private Sub Log(Text$)
   List_log.AddItem Text
End Sub


Private Sub OpenAndFill(FileName$, Optional StartOffset = 0)
' Open au3 file

   Set ScriptData = New StringReader
   ScriptData.Data = FileLoad(FileName)
   
   txt_Original = Script_RawToText(ScriptData.Data)
   
   Log_Clear
   Log FileName & " loaded."

End Sub

Private Sub FindMatches(Data$, RE_Search$) ' As MatchCollection
   
  
   With New RegExp
      .IgnoreCase = True
      .Global = True
      .MultiLine = False
      
      .Pattern = RE_Search
      
'      Dim SearchReplace_Matches As MatchCollection
      Set SearchReplace_Matches = .Execute(Data)
      
      Log SearchReplace_Matches.Count & " matches found."
'      Set FindMatches = SearchReplace_Matches
    
    ' show Matches
      Dim Match As Match
      For Each Match In SearchReplace_Matches
         With Match
'            txt_Matches = txt_Matches & vbCrLf & .value
            If .SubMatches.Count <= 1 Then
                List_Matches.AddItem Replace(.value, Match.SubMatches(0), "=>" & .SubMatches(0) & "<=")
            
            Else
                List_Matches.AddItem .value
            
            End If
           
            
'            On Error Resume Next
'            txt_Original.SelStart = .FirstIndex
'            txt_Original.SelLength = .Length
'            myDoEvents
            
         End With
         
      Next
  End With

End Sub


Private Sub SimpleSearchReplace(Data$, RE_Search$, RE_Replace$) ' As MatchCollection
   
  
   With New RegExp
      .IgnoreCase = True
      .Global = True
      .MultiLine = False
      
      .Pattern = RE_Search
      
'      Dim SearchReplace_Matches As MatchCollection
       Data = .Replace(Data, RE_Replace$)
  End With

End Sub


'Attentions this function can be unreliable
'   Match: "testme (test)"
'submatch: "test"
'
'To test if this version is reliable it is necessary to check
' if the submatch can be found only one time inside the match

Public Function RE_SubMatch_Offset(Match As Match, Optional SubMatchIndex = 0) As Long
   RE_SubMatch_Offset = InStr(1, Match.value, Match.SubMatches(SubMatchIndex)) - 1
End Function


Public Function RE_SubMatch_FirstIndex(Match As Match, Optional SubMatchIndex = 0) As Long
   RE_SubMatch_FirstIndex = Match.FirstIndex + (InStr(1, Match.value, Match.SubMatches(SubMatchIndex)) - 1)
End Function
Public Function RE_SubMatch_Length(Match As Match, Optional SubMatchIndex = 0) As Long
   RE_SubMatch_Length = Len(Match.SubMatches(SubMatchIndex))
End Function



Public Sub RE_Replace_SplitMatches(Data$, SearchReplace_Matches As MatchCollection, ByRef Replace_FixData)

 ' Dim Array for splited Data
   ReDim Replace_FixData((2 * SearchReplace_Matches.Count))

   Dim StrReader As New StringReader
   With StrReader
      .Data = Data
      .Position = 0

      Dim i&
      i = 0

      Dim Match As Match
      For Each Match In SearchReplace_Matches
        
        Dim MatchStart&, MatchLen&
        If Match.SubMatches.Count = 0 Then
         ' No SubMatches so use normal match data
           MatchStart = Match.FirstIndex
           MatchLen = Match.Length
           
        Else
         ' Okay SubMatches(well care only about the first)
           MatchStart = RE_SubMatch_FirstIndex(Match)
           MatchLen = Len(Match.SubMatches(0))
           
        End If
        
        
       ' Part from Start till the match
         Replace_FixData(i) = .FixedString(MatchStart - .Position)
         Inc i

'         Debug.Assert StrReader.Position = Match.FirstIndex
         
       ' The match part
         Replace_FixData(i) = .FixedString(MatchLen)
         Inc i
         

      Next

    ' Also append all the remaining Data
      Replace_FixData(i) = .FixedString(-1)

   End With
End Sub



Private Sub DoSearchReplace(Data$, RE_Search$, RE_Replace$, Optional Comments = "RegEx Search&Replace", Optional Testonly = False)
   
   Log "Applying '" & Comments & "'"
   
   FindMatches Data, RE_Search
   
   If Testonly Then Exit Sub
   
   If chk_Simple.value = CheckBoxConstants.vbChecked Then
   
      SimpleSearchReplace Data, RE_Search$, RE_Replace$
      
   Else
   
   
'      Dim Replace_FixData
'      RE_Replace_SplitMatches Data, SearchReplace_Matches, Replace_FixData
'
'      Dim VarCount&
'      VarCount = 1
'
'    ' Note Replace_FixData is an Array that contains in
'    ' the first field the match and in
'    ' the next field the data between
'    ' and so on
'      Dim i
'      For i = LBound(Replace_FixData) To UBound(Replace_FixData) - 1 Step 2
'
'
'         Dim StrToReplace_Old$
'         StrToReplace_Old = Replace_FixData(i + 1)
'
'       ' Make New String / Replace
'         Dim StrToReplace_New$
'         StrToReplace_New = RE_Replace & H16(VarCount)
'
'         Dim DupFinder As New Collection
'         On Error Resume Next
'
'       ' Filter duplicates for StrToReplace_Old
'         DupFinder.Add StrToReplace_New, StrToReplace_Old
'         If Err = 0 Then
'          ' Okay VarName is new and unique
'            Inc VarCount
'
'         Else
'          ' VarName already exists load existing
'            StrToReplace_New = DupFinder(StrToReplace_Old)
'         End If
'
'         Replace_FixData(i + 1) = StrToReplace_New
'
'
'      Next
'
'     'Join/Make New String(with replacements applied)
'      Data = Join(Replace_FixData, "")
'
''---------------------------------------------
      
      If SearchReplace_Matches.Count = 0 Then Exit Sub
      If SearchReplace_Matches(0).SubMatches.Count = 0 Then
         Dim ErrText$
         ErrText = "ERROR! - There are SubMatches. Please put the NamePatter of the RegExpSearchPattern into round parentheses()."
         txt_Replace = ErrText
         Log ErrText
         
         Exit Sub
      End If
      
      
      Dim DuplicatesFilter As New clsDuplicateFilter
      DuplicatesFilter.Clear
   
   
      Dim Match As Match
      Dim VarCount&
      VarCount = 0
      GUIEvent_ProcessBegin SearchReplace_Matches.Count
   
      For Each Match In SearchReplace_Matches
         With Match
            Dim SearchValue$
            SearchValue = .SubMatches(0)
            If DuplicatesFilter.IsUnique(SearchValue) Then
               ReplaceDo Data, SearchValue, RE_Replace & H16(VarCount)
               
               GUIEvent_ProcessUpdate VarCount
'               myDoEvents
               
               Inc VarCount
            End If
   
         End With
   
      Next
   
      GUIEvent_ProcessEnd
   
   End If
   txt_Replace = Script_RawToText(Data)
   
'
''    ' Merge lines with _ at the end
''      NewScript = Replace(NewScript, AU_NEWLINE, "")
'
''=== Process Data LineByLine ===
'
'    ' Break into lines and process them
'      Dim ScriptLines
'      ScriptLines = Split(Data, vbCrLf)
'
'      Dim Lines As Long
'      Lines = UBound(ScriptLines)
''      GUI_StatusBar_SetLines Lines
'
'     'Go through all Lines
'      Dim Line_idx
'      For Line_idx = 0 To Lines
'         Dim Line$
'         Line = ScriptLines(Line_idx)
''         GUI_StatusBar_SetLines Line_idx
'
''=== Start Replace ===
'
'
'      With New RegExp
'         .IgnoreCase = True
'         .Global = False
'         .MultiLine = False
'
'         .Pattern = RE_Search
'
'
'         Dim SearchReplace_Matches As MatchCollection
'         Set SearchReplace_Matches = .Execute(ScriptData.Data)
'         Log SearchReplace_Matches.Count & " matches found."
'
'
'            txt_Replace = _
'            .Replace(Txt_Original, RE_Replace & H16(VarCount))
'         Next
'
'   End With
'
  
End Sub

'Do Search'n'Replace
Private Function Apply(ByRef Data$, Optional Testonly = False)
   On Error GoTo Apply_err
   
   Log_Clear
   
   txt_Matches = ""
   List_Matches.Clear
   
   
   txt_Replace = Script_RawToText(Data)
   
 ' get/parse SearchReplacePatterns
   SeperateSearchReplacePattern
   
   
 ' And Execute Search&Replace of each entry
   Dim SearchReplaceItem As SubMatches
   For Each SearchReplaceItem In SearchReplacePattern
            
     'Filter out to short search strings(like "." that are probably incomplete and only slowdown preview)
      If Len(SearchReplaceItem(Pattern_Search)) >= 5 Then
               
         
         DoSearchReplace Data, _
                         SearchReplaceItem(Pattern_Search), _
                         SearchReplaceItem(Pattern_Replace), _
                         SearchReplaceItem(Pattern_Comments), _
                         Testonly
'         List_Matches.AddItem String(50, "_")

      End If
   Next
   
   Apply = Data
   
   Err.Clear
Apply_err:
Select Case Err
   Case 0
   Case Else
      Log "ERROR during DoSearchReplace(): " & Err.Description
End Select

End Function


Private Sub txt_Replace_KeyUp(KeyCode As Integer, Shift As Integer)
   UpdateLabel
End Sub

Private Sub txt_Replace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   UpdateLabel
End Sub


Private Sub txt_Filename_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo txt_Filename_OLEDragDrop_err
   
   txt_FileName = Data.Files(1)
'   Timer_OleDrag.Enabled = True
   

txt_Filename_OLEDragDrop_err:
Select Case Err
Case 0

Case Else
'   log "-->Drop'n'Drag ERR: " & Err.Description

End Select
End Sub

Private Sub UpdateLabel()
   Dim CharsSelected&
   CharsSelected = txt_Replace.SelLength
   If CharsSelected Then
      Lbl_Status.Caption = "Note: " & txt_Replace.SelLength & " chars selected."
   Else
      Lbl_Status.Caption = ""
   End If
End Sub




Private Sub Txt_ReplaceString_Change()
   Refresh_Preview
End Sub
