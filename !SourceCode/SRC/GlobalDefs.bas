Attribute VB_Name = "GlobalDefs"
Option Explicit

Public File As New FileStream
Public FileName As New ClsFilename

Public Const ERR_NO_AUT_EXE& = vbObjectError Or &H10
Public Const ERR_NO_OBFUSCATE_AUT& = vbObjectError Or &H20
Public Const ERR_NO_TEXTFILE& = vbObjectError Or &H30




Public Const StringBody_SingleQuoted As String = "[^']*"
Public Const String_SingleQuoted = "(?:'" & StringBody_SingleQuoted & "')+"

Public Const StringBody_DoubleQuoted As String = "[^""]*"
Public Const String_DoubleQuoted As String = "(?:""" & StringBody_DoubleQuoted & """)+"
   
Public Const StringPattern As String = String_DoubleQuoted & "|" & String_SingleQuoted


Public Const DE_OBFUSC_TYPE_NOT_OBFUSC& = &H0
Public Const DE_OBFUSC_TYPE_VANZANDE& = &H10000
Public Const DE_OBFUSC_TYPE_ENCODEIT& = &H20000
Public Const DE_OBFUSC_TYPE_CHR_ENCODE& = &H10
Public Const DE_OBFUSC_TYPE_CHR_ENCODE_OLD& = &H8


Public Const DE_OBFUSC_VANZANDE_VER14& = &H10014
Public Const DE_OBFUSC_VANZANDE_VER15& = &H10015
Public Const DE_OBFUSC_VANZANDE_VER15_2& = &H100152
Public Const DE_OBFUSC_VANZANDE_VER24& = &H10024


Public Const NO_AUT_DE_TOKEN_FILE& = &H100

Public ExtractedFiles As Collection

Public IsCommandlineMode As Boolean
Public IsOpt_QuitWhenFinish As Boolean
Public IsOpt_RunSilent As Boolean



Public Sub GUIEvent_ProcessBegin(Target&, Optional BarLevel& = 0, Optional Skipable As Boolean = False)
   FrmMain.GUIEvent_ProcessBegin Target, BarLevel, Skipable
End Sub

Public Sub GUIEvent_ProcessUpdate(CurrentValue&, Optional BarLevel& = 0)
   FrmMain.GUIEvent_ProcessUpdate CurrentValue, BarLevel
End Sub
Public Sub GUIEvent_ProcessEnd(Optional BarLevel& = 0)
   FrmMain.GUIEvent_ProcessEnd BarLevel
End Sub

Public Sub GUIEvent_Increase(PerCentToIncrease As Double, Optional BarLevel& = 0)
   FrmMain.GUIEvent_Increase PerCentToIncrease, BarLevel
End Sub

Public Sub GUI_SkipEnable()
   FrmMain.Cmd_Skip.Visible = True
   If FrmMain.bCmd_Skip_HasFocus = False Then
      FrmMain.Cmd_Skip.SetFocus
      FrmMain.bCmd_Skip_HasFocus = True
   End If
End Sub

Public Sub GUI_SkipDisable()
   FrmMain.Cmd_Skip.Visible = False
   FrmMain.bCmd_Skip_HasFocus = False
End Sub



Sub DoEventsSeldom()
   If Rnd < 0.01 Then myDoEvents
End Sub

Sub DoEventsVerySeldom()
   If (GetTickCount() And &H7F) = 1 Then
'   If Rnd < 0.00001 Then
       myDoEvents
   End If
End Sub

Sub ShowScript(ScriptData$)
   
   FrmMain.Txt_Script = Script_RawToText(ScriptData)

End Sub

Function Script_RawToText(ByRef ScriptData$) As String
   If isUTF16(ScriptData) Then
      Script_RawToText = StrConv((Mid(ScriptData, 1 + Len(UTF16_BOM))), vbFromUnicode)
   ElseIf isUTF8(ScriptData) Then
      Script_RawToText = Mid(ScriptData, 1 + Len(UTF8_BOM))
   Else
      Script_RawToText = ScriptData
   End If

End Function

Sub SaveScriptData(ScriptData$, Optional skipTidy As Boolean)

   With FrmMain
      
   ' Not need anymore since Tidy v2.0.24.4 November 30, 2008
'   ' Adding a underscope '_' for lines longer than 2047
'   ' so Tidy will not complain
'      FrmMain.Log "Try to breaks very long lines (about 2000 chars) by adding '_'+<NewLine> ..."
'      ScriptData = AddLineBreakToLongLines(Split(ScriptData, vbCrLf))
      
       ' overwrite script
         If FrmMain.Chk_TmpFile.value = vbChecked Then
            FileName.Name = FileName.Name & "_restore"
            .Log "Saving script to: " & FileName.FileName
         Else
   '         FileDelete FileName.Name
            .Log "Save/overwrite script to: " & FileName.FileName
         End If
   
         FileSave FileName.FileName, ScriptData
      
      End With
      
      RunTidy ScriptData, skipTidy
End Sub

Public Sub RunTidy(ScriptData$, Optional skipTidy As Boolean)
   
   With FrmMain
        
      ShowScript ScriptData
      .Log ""
     
      If skipTidy Then
         .Log "Skipping to run 'data\Tidy\Tidy.exe' on" & FileName.NameWithExt & "' to improve sourcecode readability. (Plz run it manually if you need it.)"
      Else
         
         .Log "Running 'Tidy.exe " & FileName.NameWithExt & "' to improve sourcecode readability."
         
         FrmMain.ScriptLines = Split(ScriptData, vbCrLf)
         
         Dim cmdline$, parameters$, Logfile$
         cmdline = App.Path & "\" & "data\Tidy\Tidy.exe"
         parameters = """" & FileName & """" ' /KeepNVersions=1
         .Log cmdline & " " & parameters
         
         Dim TidyExitCode&
         
         'Dim ConsoleOut$
         'ConsoleOut =
         FrmMain.Console.ShellExConsole cmdline, parameters, TidyExitCode
         
         
         If TidyExitCode = 0 Then
             .Log "=> Okay (ExitCode: " & TidyExitCode & ")."
             Dim TidyBackupFileName As New ClsFilename
             TidyBackupFileName.mvarFileName = FileName.mvarFileName
             TidyBackupFileName.Name = TidyBackupFileName.Name & "_old1"
             
           ' Delete Tidy BackupFile
             If FrmMain.Chk_TmpFile.value = vbUnchecked Then
                .Log "Deleting Tidy BackupFile..." ' & TidyBackupFileName.NameWithExt
                FileDelete TidyBackupFileName.FileName
             End If
            
            
          ' Readin tidy file
            ScriptData = FileLoad(FileName.FileName)
          
            ShowScript ScriptData
            
         Else
            .Log "=> Error (ExitCode: " & TidyExitCode & ")" ' TidyOutput >>>"
'            .Log ConsoleOut, "TIDY OUTPUT: "
'            .Log "<<<"
            .Log "Attention: Tidy.exe failed. Deobfucator will probably also fail because scriptfile is not in proper format."
         End If
         
      End If 'skip tidy
      
   End With
End Sub

