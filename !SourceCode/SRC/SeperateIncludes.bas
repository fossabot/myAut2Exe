Attribute VB_Name = "modSeperateIncludes"
Option Explicit
Private Const INCLUDE_Seperator$ = "; ----------------------------------------------------------------------------" & vbCrLf
Private Const INCLUDE_START$ = INCLUDE_Seperator & "; <AUT2EXE INCLUDE-START: "
Private Const INCLUDE_END$ = INCLUDE_Seperator & "; <AUT2EXE INCLUDE-END: "
Private Const INCLUDE_Close$ = ">" & vbCrLf & INCLUDE_Seperator
   
Private Const INCLUDE_FirstLine = "#include-once" & vbCrLf
Private Const INCLUDE_FirstLine_Len = 13
   
Private Const INCLUDE_REPLACE_START = "#include <"
Private Const INCLUDE_REPLACE_END = ">" & vbCrLf
Private IncludeList As New Collection
Private IncludeListCount&
Private IncludeFileName As New ClsFilename


Const AllButRE_NewLine As String = "[^\r\n]"


Dim str As StringReader
Dim Level&

Public Sub SeperateIncludes()
   
   FrmMain.Log ""
   FrmMain.Log "==============================================================="
   FrmMain.Log "Seperating Includes of : " & FileName.FileName
   
   
  'Read *.au3 into ScriptData
   Dim ScriptData$
   With File
      .Create FileName.FileName
      
      
      'Test for Unicode-Bom
      Dim bUTF16detected As Boolean
      
      Dim UnicodeBomBuff%
      UnicodeBomBuff = .int16
      
      bUTF16detected = UnicodeBomBuff = &HFEFF
      If bUTF16detected Then 'LittleEndian of UTF16
      
      ElseIf UnicodeBomBuff = &HFFFE Then
         Log "ERR: BigEndian of UTF16 detected - Please convert input file manually to 8-bit Accii or LittleEndian UTF16."
      Else
        'Seek to begin
         .Position = 0
      End If

      
      ScriptData = .Data
      .CloseFile
   End With
 
' delete old script
'   Kill FileName.FileName
'   FileRename FileName.FileName, FileName.FileName & "_"
   
   
  
 ' Make DirName with scriptname
   With IncludeFileName
      .mvarFileName = FileName.mvarFileName
      .Name = "_" & .Name & "_Seperated\"
      .NameWithExt = ""
   End With
   
   FrmMain.Log "  " & Len(ScriptData) & IIf(bUTF16detected, "(Unicode)", "") & " bytes loaded."
   
 ' Convert unicode to accii
   If bUTF16detected Then
      ScriptData = StrConv(ScriptData, vbFromUnicode, LocaleID)
   End If
   
   SeperateIncludes2 ScriptData
   
End Sub


Public Sub SeperateIncludes2(ScriptData$)
   
  '
   Set str = New StringReader
   str = ScriptData
   str.DisableAutoMove = True
   
   Level = 0
   IncludeListCount = 0
   SeperateIncludes_Recursiv INCLUDE_END$
   
   If Level <> 0 Then Err.Raise vbObjectError, , "INCLUDE-START/END unembalanced: (" & Level & " too much) in ScriptData: " & str & vbCrLf & "ignored: " & str.FixedString(-1)
   
End Sub


Private Sub SeperateIncludes_Recursiv(ByVal EndSym$)
   
 ' Scan for StartSym until end of String
   Do While str.EOS = False
      
    ' Test for "; <AUT2EXE INCLUDE-START: "
      If INCLUDE_START$ = str.FixedString(Len(INCLUDE_START$)) Then

       ' Set Script Cut Position
         Dim ScriptCutPos_Start&
         ScriptCutPos_Start = str.Position
         
         str.Move Len(INCLUDE_START$)
       
       ' === Cut out include path ===
         Dim pathStartPos&
         pathStartPos = str.Position
         
         Dim pathEndPos&
         pathEndPos& = str.FindString(INCLUDE_Close) ' - Len(INCLUDE_Close)
         
         Dim IncludePath As ClsFilename
         Set IncludePath = New ClsFilename
         IncludePath = str.FixedString(pathEndPos - pathStartPos)
        
      ' === Generate Output path+name for include ===
      ' Copy Original IncludeName to Output IncludeName + Create Output IncludeName Dir
      ' D:\Program Files\AutoIt3\Include\UpDownConstants.au3 -> C:\myscripts\AutoIt3\Include\UpDownConstants.au3
         Dim IncludePathNew As ClsFilename
         Set IncludePathNew = New ClsFilename
         IncludePathNew.NameWithExt = IncludePath.NameWithExt
        
       'CopyName last two path parts
       ' Example: "D:\Program Files\AutoIt3\Include\" ->  "AutoIt3\Include\"
         Dim PathParts, PathPartsCount
         PathParts = Split(IncludePath.Path, "\")
         PathPartsCount = UBound(PathParts)
         If PathPartsCount > 2 Then
            IncludePathNew.Path = PathParts(PathPartsCount - 2) & "\" & _
                               PathParts(PathPartsCount - 1) & "\"
         ElseIf PathPartsCount > 1 Then
            IncludePathNew.Path = PathParts(PathPartsCount - 1) & "\"
            
         Else
            IncludePathNew.Path = "Inc\"
            
         End If
         
        'First Include is the MainScript - Place it in the ScriptDir
         If IncludeListCount = 0 Then IncludePathNew.Path = ""
         
                               
                              
         Inc IncludeListCount
                               
       ' show IncludeFileName
         FrmMain.Log Space(Level) & "#" & IncludeListCount & " " & IncludePath & vbTab & " -> " & IncludePathNew
                               
       ' Make includepath for insert "#include <...>" l
         Dim IncludeLinePath$
         IncludeLinePath = IIf(IncludePathNew.Path Like "AutoIt3\Include\", "", IncludePathNew.Path) & IncludePathNew.NameWithExt
         
        'Add Script path   Example: "AutoIt3\Include\"-> "f:\myscripts\AutoIt3\Include\"
         IncludePathNew.Path = IncludeFileName.Path & IncludePathNew.Path
         IncludePathNew.MakePath
          
          
       ' === Recursiv Call of this function ===
       ' ; ----------------------------------------------------------------------------
       ' ; <AUT2EXE INCLUDE-START: D:\Program Files\AutoIt3\Include\UpDownConstants.au3>
       ' ; ----------------------------------------------------------------------------
         str.Move Len(INCLUDE_Close) + (pathEndPos - pathStartPos)
         
       ' Store Position to cut out Text later
         Dim ScriptTextStartPos&
         ScriptTextStartPos = str.Position
         
         Dim newEndSym$
         newEndSym = INCLUDE_END & IncludePath & INCLUDE_Close
         
      '! Recursiv Call of this function !
         Inc Level
         SeperateIncludes_Recursiv newEndSym
         
       ' now the function returned because some
       ' ; ----------------------------------------------------------------------------
       ' ; <AUT2EXE INCLUDE-END: D:\Program Files\AutoIt3\Include\UpDownConstants.au3>
       ' ; ----------------------------------------------------------------------------
       ' were found.
       ' Note: String position pointer is at the beginning
         Dim ScriptTextEndPos&
         ScriptTextEndPos = str.Position
         
       ' Seek to end of '; <AUT2EXE INCLUDE-END'
         str.Move Len(newEndSym)
         Dim ScriptCutPos_End&
         ScriptCutPos_End& = str.Position
                  
         
         Dim tmpstr2$
         Inc ScriptCutPos_Start
         Inc ScriptCutPos_End
         
         
        'Filter out duplicates
         On Error Resume Next
         IncludeList.Add IncludePathNew.FileName, IncludePathNew.FileName
         If (Err = 0) Then
'        If Len(ScriptData) > (6 + INCLUDE_FirstLine_Len) Then
          
          ' Copy include Text(without '; <AUT2EXE INCLUDE' Comments) to ScriptData
            Dim ScriptData$
            str.Position = ScriptTextStartPos
            ScriptData = INCLUDE_FirstLine & str.FixedString(ScriptTextEndPos - ScriptTextStartPos)
            
   
          ' show ScriptData
            FrmMain.Txt_Script = ScriptData
          
          ' Save ScriptData to file
            FileSave IncludePathNew.FileName, ScriptData

           
         Else
         
           ' Log
            If Err = 457 Then '"Dieser Schlüssel ist bereits einem Element dieser Auflistung zugeordnet"
               FrmMain.Log Space(Level) & "Duplicate Include - Skipped"
            Else
               FrmMain.Log Space(Level) & "Unexp. Err: " & Err.Description
            End If
          
         End If

       ' Delete Include from ScriptFile and Replace it with '#include'
         tmpstr2 = str.mvardata
         Dim IncludeLine$
         IncludeLine = INCLUDE_REPLACE_START & IncludeLinePath & INCLUDE_REPLACE_END
         
         FrmMain.Txt_Script = strCutOut(tmpstr2, ScriptCutPos_Start, ScriptCutPos_End - ScriptCutPos_Start, IncludeLine)
         str.mvardata = tmpstr2
         
        
        'Seek back where deleting of include text started
         str.Position = ScriptCutPos_Start

      
    ' Test for "; <AUT2EXE INCLUDE-END: "
      ElseIf EndSym = str.FixedString(Len(EndSym)) Then ' "; <AUT2EXE INCLUDE-END: "...
           
            Dec Level
            Exit Do

      End If
      
    ' Move to next Position in String to test for '; <AUT2EXE INCLUDE XXX'
      str.Move 1
      
   Loop
   
End Sub

Public Sub AHK_RemoveLineBreaks(ByRef ScriptData As StringReader)
   With New RegExp
      .Pattern = RE_Group_NonCaptured(RE_NewLine) & "?" & _
                 RE_Group("; <COMPILER: v" & AllButRE_NewLine & "*>" & RE_NewLine) & _
                 RE_Group_NonCaptured(RE_NewLine) & "*"
      ScriptData = .Replace(ScriptData, "$1")
   End With

End Sub


'Public Sub AHK_SeperateIncludes_NEW(ByRef ScriptData As StringReader, OutputPath$)
'   Dim myRegExp As New RegExp
'   With myRegExp
'
'    ' Remove RE_NewLines after "; <COMPILER: v1.0.48.3>"
'      .Pattern = "; <COMPILER: v.*" & RE_NewLine
'      .Global = True
'
'    ' Seperate & Save includes
'      ScriptData.Position = 0
'
'      Dim Match As Match
'      For Each Match In myRegExp.Execute(ScriptData.FixedString)
'
'          MainScript.Concat Match.value
'
'          Dim ScriptPos_Start&
'          ScriptPos_Start = Match.FirstIndex + Match.Length
'
'
'      Next
'
'      ScriptData.Position = ScriptPos_Start
'
'      .Pattern = "; " & RE_Group("#include (.*?\.ahk)" & RE_NewLine)
'
'      For Each Match In myRegExp.Execute(ScriptData.FixedString)
'          MainScript.Concat Match.value
'
'       ' Get IncludeFileName
'         Dim IncludeFileName As New ClsFilename
'         With IncludeFileName
'            .FileName = OutputPath
'            .NameWithExt = Match.SubMatches(3)
'            .MakePath
'         End With
'
'
'       ' Get IncludeData
'          ScriptData.Position = ScriptPos_Start
'          ScriptPos_Start = Match.FirstIndex + Match.Length
'
'         Dim IncludeFile As New FileStream
'         With IncludeFile
'            .Create IncludeFileName.FileName, True, False, False
'            .FixedString(-1) = Match.SubMatches(1)
'            .CloseFile
'         End With
'
'
'
'
'      Next
'
'
'
'
'
'      Next
'
'    ' Save mainscript ( with #includes)
'      ' '.Replace' deletes all matches data and inserts there the given data -
'      '     here $1 the CompilerLine and $3 what are the includes
'      ' ... and of course the unmatched data at the end stays too -> what is the main script
'      ScriptData = .Replace(ScriptData, "$1$3")
'
'   End With
'End Sub


Public Sub AHK_SeperateIncludes(ByRef ScriptData As StringReader, OutputPath$)
'   On Error GoTo AHK_SeperateIncludes_err
   
'   If ScriptData.FindString("; #include ") = 0 Then
'      Log "There are no AHK-includes that could be seperated."
'      Exit Sub
'   End If
   
 ' Got Through all Lines of the Script
   Dim ScriptLines
   ScriptLines = Split(ScriptData.Data, vbCrLf)
   
   
   Dim MainFile As New clsStrCat
   Dim IncludeFile As New clsStrCat
   Dim IncludeFileCount&
   
   Dim StoreInIncludeFile As Boolean 'false
   Dim LineWithIncludeDirective As Boolean 'false
   
   
   
   Dim myRegExp As New RegExp
   With myRegExp
   
     .MultiLine = False
     .Global = False

      Dim LineCount& '0
      Dim Line
      For Each Line In ScriptLines
         Inc LineCount
         
         LineWithIncludeDirective = False
      
     
       ' Filter out lines like
       ' "#include %A_ScriptDir%"
       ' ...and use  them as kind of seperator between mainscript and first start of frist include
        .Pattern = "; " & RE_Group("#include (%A_.*\%)")
         Dim Match As Match
         For Each Match In myRegExp.Execute(Line)
            LineWithIncludeDirective = True
            StoreInIncludeFile = True
            
            Line = Match.SubMatches(0)
            log_verbose "AHK-Include directive: " & Match.SubMatches(1) & "@line: " & LineCount
            
   
         Next
      
        
       ' Now get "#include blah.ahk"
        .Pattern = "; " & RE_Group("#include (.*\.ahk)")
         For Each Match In myRegExp.Execute(Line)
            LineWithIncludeDirective = True
            StoreInIncludeFile = True
            Inc IncludeFileCount
            
            
          ' Make IncludeFileName
            Dim IncludeFileName As New ClsFilename
            With IncludeFileName
               .FileName = OutputPath
               .NameWithExt = Replace(Match.SubMatches(1), "/", "\") '<-slash to backslash
               .MakePath
            End With
            
          ' Save include data
            FileSave IncludeFileName.FileName, _
                     IncludeFile.value
          
          ' Clear tmpstorage for Include
            IncludeFile.Clear
            
          ' Line <= " #include blah.ahk"
            Line = Match.SubMatches(0)
            
            
            log_verbose "AHK-Include #" & IncludeFileCount & ": " & _
                        Match.SubMatches(1) & "  @line: " & LineCount
            
            
         Next
            
        
         If StoreInIncludeFile And Not LineWithIncludeDirective Then
         
            IncludeFile.Concat Line & vbCrLf
            
         Else
          ' store first lines and IncludeDirective in mainfile
            MainFile.Concat Line & vbCrLf
            
         End If
   
      Next
   
   End With
   
   ScriptData = MainFile.value
   
   
   
   
   
'   If ScriptData.Length > 10000 Then
'      If vbYes <> MsgBox("Due to some strange RegExp bug this can take some time. (please look at sourcecode and tell me if you found some better solution)" & vbCrLf & _
'         "Do you really like to seperated includes?", vbDefaultButton2 Or vbYesNo Or vbQuestion, "Seperate AHK-includes") Then Exit Sub
'   End If
'
'   Dim myRegExp As New RegExp
'   With myRegExp
'
'      .MultiLine = True
'      .Pattern = RE_Group_NonCaptured( _
'                    RE_Group("; <COMPILER: v.*" & RE_NewLine) & _
'                    RE_Group_NonCaptured(RE_NewLine) & "*" _
'                 ) & "?" & _
' _
'                 RE_Group(RE_Group_NonCaptured(RE_AnyCharNL) & "*?") & _
'                 RE_NewLine & "; " & RE_Group("#include (.*?\.ahk)" & RE_NewLine)
'
'      .Global = True
'
''      BenchStart
''.Execute ScriptData
''      BenchEnd
'
'    ' Seperate & Save includes
'      Dim Match As Match
'      For Each Match In myRegExp.Execute(ScriptData)
'
'       ' Get IncludeFileName
'         Dim IncludeFileName As New ClsFilename
'         With IncludeFileName
'            .FileName = OutputPath
'            .NameWithExt = Match.SubMatches(3)
'            .MakePath
'         End With
'
'       ' Get IncludeData
'         Match.SubMatches(1) = FileLoad(IncludeFileName.FileName)
'
'      Next
'
'    ' Save mainscript ( with #includes)
'      ' '.Replace' deletes all matches data and inserts there the given data -
'      '     here $1 the CompilerLine and $3 what are the includes
'      ' ... and of course the unmatched data at the end stays too -> what is the main script
'      ScriptData = .Replace(ScriptData, "$1$3")
'
'   End With
AHK_SeperateIncludes_err:
End Sub

