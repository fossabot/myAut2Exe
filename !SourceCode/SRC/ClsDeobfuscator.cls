VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDeobfuscator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text

Const AU_NEWLINE$ = " _" & vbCrLf


Dim IsDeTokeniserOutput As Boolean
Dim IsTidyed As Boolean
Dim IsTidyedDeTokeniserOutput As Boolean



Dim myRegExp As New RegExp

Dim DeObfu_FunctionNames As New Collection
Dim NewScript$

Sub Log(Text$)
   FrmMain.Log Text
End Sub

Sub log2(Text$)
   FrmMain.log2 Text
End Sub

Private Sub GUI_Show_Text(Text$)
   FrmMain.Txt_Script = Text
End Sub


Public Function IsObfuscatedFunctionName(FunctionName) As Boolean
   'only rename obfuscate Names that look like these:
   'A45B1002A2B
   'A04B1105530
   'A3AA1801935
   '...
   'a1b00100b0a
   
   
   'IsObfuscatedFunctionName = FunctionName Like "A??????????" '"A[0-9A-F]{10}"
'   With New RegExp
      myRegExp.Pattern = "^[0-9A-Fa-f]{10}"
      IsObfuscatedFunctionName = myRegExp.Test(FunctionName)
'      Debug.Assert IsObfuscatedFunctionName = True
'   End With
End Function

Public Sub TestIsDeTokeniserOutput(Script)
  
  
 ' Check for presence of "; DeTokenise by"
   IsDeTokeniserOutput = InStrRev(Right(Script, 200), DETOKENISE_MAKER, , vbBinaryCompare)
End Sub

Public Function TestIsTidyed(Script$) As Boolean
 
' ' In a decompiled & detokenise file that was not run in tidy KeyWords and VarNames are in UpperCase
'   TestIsTidyed = (0 = InStr(1, Script, "GLOBAL $", vbBinaryCompare)) And _
                  (0 = InStr(1, Script, "  =", vbBinaryCompare))

 ' In a decompiled & detokenise file that there are no identions with space or tab at the beginning of a line
   TestIsTidyed = InStr(1, Script, vbCrLf & vbTab, vbBinaryCompare) Or _
                  InStr(1, Script, vbCrLf & " ", vbBinaryCompare) ' Or _
                  InStr(1, Script, ", ", vbBinaryCompare)
End Function


'
'Sub ReBase_Array(ByRef Arr, Dim_Min, Optional Dim_Max)
'
'   If IsMissing(Dim_Max) Then Dim_Max = UBound(Arr)
'
'   Dim NewArr As New Collection
'
'   Dim i&
'   For i = Dim_Min To Dim_Max
'      NewArr.Add Arr(i)
'   Next
'
'   Arr = CollectionToArray(NewArr)
'End Sub

Private Sub DeObfuscate_Chr_Encode()
   
'   If NewScript Like "*""[!""]*" Then
'      Log "Skipping Replacing 'Consts Globals & Dim' because there are strings in the script. Use the function rename to include these includes first."
'   Else
'     '=== Replace Dim Const and Globals
'      DeObfu_ReplaceGlobals NewScript
'      DeObfu_ReplaceDimsAndConst NewScript
'   End If


' --- Prepare regular expression objects & search pattern ---
   
   Dim Number$
   Number = "(\d{0,4})"
   
   Dim Operator$
   Operator = "([+-])"
   
 
   Dim CHR_Command$
 ' for CHR_Commands for example "CHR(142+8082)"
   CHR_Command = RE_WSpace("CHR", "\(", Number, Operator, Number, "\)")
   
 
   Dim CHR_CommandElements$
 ' One or more CHR_Commands for example "CHR(157-125) & CHR(10+101)"
 ' used to separate & replace ChrStrings like
 ' 'CHR(157-125)' or
 ' 'CHR(157-125) & ""' or
 ' 'CHR(157-125) & CHR(10+101)' or
 ' 'CHR(157-125) & CHR(10+101) & ""

   CHR_CommandElements = RE_WSpace(CHR_Command & _
                      "(", "&", CHR_Command, ")*" & _
                      "(", "&", """""){0,1}")
  
   Dim ChrStrings As New RegExp
   With ChrStrings
      .IgnoreCase = True
      .Global = False
      .Pattern = CHR_CommandElements
      
   End With
 
   Dim Chrs As New RegExp
 ' RegExp for separating Chrs like "CHR( '10' '+' '101' )"
   With Chrs
      .IgnoreCase = True
      .Global = True
      .Pattern = RE_WSpace(CHR_Command)
   End With
 
' 'Break String into lines (but retain "_ "&CRLF as one line)
'   Dim Lines As New RegExp
'   With Lines
'       .IgnoreCase = True
'       .Global = True
'       .MultiLine = True
'       .Pattern = RE_WSpace(".*", "[^_]", "\n")
'   End With
'   Dim Lines_Match As MatchCollection
'   Set Lines_Match = Lines.Execute(NewScript)
'
'


' --- Process Script ---

   Dim ChrStringNew As clsStrCat
   Set ChrStringNew = New clsStrCat
 
 ' Merge lines with _ at the end
   NewScript = Replace(NewScript, AU_NEWLINE, " ")

 ' Break into lines and process them
   Dim ScriptLines
   ScriptLines = Split(NewScript, vbCrLf)
   
   Dim Lines As Long
   Lines = UBound(ScriptLines)
   GUIEvent_ProcessBegin Lines
 
  'Go through all Lines
   Dim Line_idx&
   For Line_idx = 0 To Lines
      Dim Line$
      Line = ScriptLines(Line_idx)
      GUIEvent_ProcessUpdate Line_idx
'      GUI_Show_Text Line_idx & ": " & ScriptLines(Line_idx)
      
    
'    If Line_idx = 13245 Then Stop
    
    
    ' separate ChrString's
      Do While True
         Dim ChrStrings_Matches As MatchCollection
         Set ChrStrings_Matches = ChrStrings.Execute(Line)
    ' Exit (and process nxt line) if there are no ChrStrings
      If ChrStrings_Matches.Count = 0 Then Exit Do

        ChrStringNew.Clear
         
       'separate Chr's
        Dim Chrs_Match As match
        For Each Chrs_Match In Chrs.Execute(ChrStrings_Matches(0))
           
          ' Concat Chrs to a ChrString
            With Chrs_Match
              
             ' "CHR(157-125)" -> ...
               Dim Chrs_op$, Chrs_arg1&, Chrs_arg2&
               Chrs_arg1 = .SubMatches(0)    'Chrs_arg1 = 157
               Chrs_op = .SubMatches(1)      'Chrs_op   = "-"
               Chrs_arg2 = .SubMatches(2)    'Chrs_arg2 = 125
            
            End With
              
          ' calc new Char
            Select Case Chrs_op
            Case "+"
               ChrStringNew.Concat (ChrW(Chrs_arg1 + Chrs_arg2))
            Case "-"
               ChrStringNew.Concat (ChrW(Chrs_arg1 - Chrs_arg2))
            End Select
           
         Next Chrs_Match
       
       ' replace ChrString in line
         ReplaceDo Line, CStr(ChrStrings_Matches(0)), MakeAutoItString(ChrStringNew.value)

       ' save modified Line
         ScriptLines(Line_idx) = Line
         
      Loop ' next ChrString in Line

      GUI_Show_Text Line_idx & ": " & ScriptLines(Line_idx)
      DoEventsSeldom
      
   Next Line_idx
   
   GUIEvent_ProcessEnd

  'Join lines back to a script
   NewScript = Join(ScriptLines, vbCrLf)

End Sub


'Private Function Eval(Expr As String) As String
'   'Dim x  As MSScriptControl.ScriptControl
'   'Set x = New MSScriptControl.ScriptControl
'   Dim x As Object
'   Set x = CreateObject("MSScriptControl.ScriptControl") 'regsvr32.exe <WINXP>\SYSTEM32\msscript.ocx
'   x.Language = "VBScript"
'
'   Eval = x.Eval(Expr)
'
'   Set x = Nothing
'
'
'End Function
'
'Private Sub DeObfuscate_Chr_Encode_Old()
'
'
'
'
' ' Merge lines with _ at the end
' '  NewScript = Replace(NewScript, vbTab, "")
'   NewScript = Replace(NewScript, "&""""&", "&") '
'   NewScript = Replace(NewScript, "_" & vbCrLf, "")
'
' ' Break into lines and process them
'   Dim ScriptLines
'   ScriptLines = Split(NewScript, vbCrLf)
'
'   Dim Line_idx
'   For Line_idx = 0 To UBound(ScriptLines)
'    Dim Line$
'    Line = ScriptLines(Line_idx)
'
'' 'dbg
''  If Line_idx = 33900 Then Stop
'
'     'Seperated ChrStrings of each line
'      Dim ChrStrings
'      ChrStrings = Split(Line, " & """"")
'
'
'     'If contains ChrStrings...
'      If UBound(ChrStrings) > 0 Then
'
'         Dim ChrString_idx&
'         For ChrString_idx = 0 To UBound(ChrStrings)
'          Dim ChrString$
'          ChrString = ChrStrings(ChrString_idx)
'
'            Dim EncodedChars
'            EncodedChars = Split(ChrString, "CHR(")
'
'           'If contains chr's...
'            If UBound(EncodedChars) > 0 Then
'              '... simplify them to a String
'               Dim CHR_PreData$
'               CHR_PreData = EncodedChars(0)
'               EncodedChars(0) = ""
'
'               Dim Clean_ChrString$
'               Clean_ChrString = Join(EncodedChars, "CHR(")
'
'               On Error Resume Next
'               Dim Clean_String$
'               Clean_String = Eval(Clean_ChrString)
'               If Err = 0 Then
'                  ChrStrings(ChrString_idx) = CHR_PreData & Quote(Clean_String)
'               Else
'                  Log "Problem on Line: " & Line_idx & " (" & Err.Description & ") - Line skipped."
'               End If
'            End If
'
'         Next
'        'Rejoin Line
'         ScriptLines(Line_idx) = Join(ChrStrings, " ")
'      End If
'   Next
'
'   NewScript = Join(ScriptLines, vbCrLf)
'
'
'End Sub


Private Sub RE_Replace(ByRef Data$, SearchPattern$, ReplaceText$)
   With New RegExp
      .IgnoreCase = True
      .Global = True
      .Pattern = SearchPattern
      
      Data = .Replace(Data, ReplaceText)
   End With
End Sub

Private Sub ReplaceCRLF_and_Tab_Chrs(ByRef Line$)

         RE_Replace Line, RE_WSpace("Chr\(13\)", "&", "Chr\(10\)"), "@CRLF" & AU_NEWLINE
         RE_Replace Line, "Chr\(13\)", "@CR" & AU_NEWLINE
         RE_Replace Line, "Chr\(10\)", "@LF" & AU_NEWLINE
         RE_Replace Line, "Chr\(9\)", "@TAB" & AU_NEWLINE

End Sub


Private Sub DeObfuscate_Chr_Encode_old()
   

' --- Prepare regular expression objects & search pattern ---
   
   Dim Number$
   Number = "(\d{0,4})"
   
 
   Dim CHR_Command$
 ' for CHR_Commands for example "CHR(142)" or "CHR(8082)"
   CHR_Command = RE_WSpace("CHR", "\(", Number, "\)")
   
 
   Dim CHR_CommandElements$
 ' One or more CHR_Commands for example "CHR(157) & CHR(10)"
 ' used to separate & replace ChrStrings like
 ' 'CHR(157)' or
 ' 'CHR(157) & ""' or
 ' 'CHR(157) & CHR(10)' or
 ' 'CHR(157) & CHR(10) & ""

   CHR_CommandElements = RE_WSpace(CHR_Command & _
                      "(", "&", CHR_Command, ")*" & _
                      "(", "&", """""){0,1}")
  
   Dim ChrStrings As New RegExp
   With ChrStrings
      .IgnoreCase = True
      .Global = False
      .Pattern = CHR_CommandElements
      
   End With
 
   Dim Chrs As New RegExp
 ' RegExp for separating Chrs like "CHR( '10' )"
   With Chrs
      .IgnoreCase = True
      .Global = True
      .Pattern = RE_WSpace(CHR_Command)
   End With
 
' 'Break String into lines (but retain "_ "&CRLF as one line)
'   Dim Lines As New RegExp
'   With Lines
'       .IgnoreCase = True
'       .Global = True
'       .MultiLine = True
'       .Pattern = RE_WSpace(".*", "[^_]", "\n")
'   End With
'   Dim Lines_Match As MatchCollection
'   Set Lines_Match = Lines.Execute(NewScript)
'
'


' --- Process Script ---

   Dim ChrStringNew As clsStrCat
   Set ChrStringNew = New clsStrCat
 
 ' Merge lines with _ at the end
   NewScript = Replace(NewScript, AU_NEWLINE, " ")

 ' Break into lines and process them
   Dim ScriptLines
   ScriptLines = Split(NewScript, vbCrLf)
   
   Dim Lines As Long
   Lines = UBound(ScriptLines)
   GUIEvent_ProcessBegin Lines
 
  'Go through all Lines
   Dim Line_idx&
   For Line_idx = 0 To Lines
      Dim Line$
      Line = ScriptLines(Line_idx)
      GUIEvent_ProcessUpdate Line_idx
'      GUI_Show_Text Line_idx & ": " & ScriptLines(Line_idx)
      
    
'    If Line_idx = 13245 Then Stop
    
  ' Used to detect if all chr-strings in a line were replaced
  ' incase there is something like Chr(0) or Chr(9) this is increased
    Dim ChrStrings_NoMoreMatches_Count&
    ChrStrings_NoMoreMatches_Count = 0
    
    ' separate ChrString's
      Do While True
         Dim ChrStrings_Matches As MatchCollection
         Set ChrStrings_Matches = ChrStrings.Execute(Line)
    ' Exit (and process next line) if there are no ChrStrings
      If ChrStrings_Matches.Count = ChrStrings_NoMoreMatches_Count Then Exit Do

        ChrStringNew.Clear
         
      ' separate Chr's
        Dim bSkipChrStringReplace As Boolean
        bSkipChrStringReplace = False
         
        Dim Chrs_Match As match
        For Each Chrs_Match In Chrs.Execute(ChrStrings_Matches(0))

            Dim ChrCode&
            ChrCode = Chrs_Match.SubMatches(0)
          ' Filter out like  09-Tab or
            Select Case ChrCode
            
            Case 0 To 31
            
'            ' Concat Chrs to a ChrString
'              ChrStringNew.Concat """ & Chr(" & ChrCode & " ) &"""
              
            ' Hope that this is sufficient
              bSkipChrStringReplace = True
              Inc ChrStrings_NoMoreMatches_Count
              Exit For
            
            Case Else
            ' Concat Chrs to a ChrString
              ChrStringNew.Concat ChrW(ChrCode)
            End Select
           
         Next Chrs_Match
       
       ' replace ChrString in line
         If bSkipChrStringReplace = False Then
            ReplaceDo Line, CStr(ChrStrings_Matches(0)), MakeAutoItString(ChrStringNew.value)
         End If
       
       ' save modified Line
         ScriptLines(Line_idx) = Line
         
      Loop ' next ChrString in Line

      ' Apply macro's To Line (execute only if there were really some chr found)
      If ChrStrings_NoMoreMatches_Count >= 1 Then
         ReplaceCRLF_and_Tab_Chrs Line
       
       ' save modified Line
         ScriptLines(Line_idx) = Line

      End If
      
      GUI_Show_Text Line_idx & ": " & ScriptLines(Line_idx)
      DoEventsSeldom
      
   Next Line_idx
   
   GUIEvent_ProcessEnd
   
  'Join lines back to a script
   NewScript = Join(ScriptLines, vbCrLf)

End Sub





Private Sub DeObfuscate_Encodeit()

 ' Merge lines with _ at the end
   NewScript = Replace(NewScript, AU_NEWLINE, " ")


 ' ======== Var Replace
   DeObfu_ReplaceDimsAndConst NewScript
   DeObfu_ReplaceGlobals NewScript
   DeObfu_ReplaceLocals NewScript



'==== Get Vectors
   Dim globs
   globs = Split(NewScript, "Global Const ")
   
   Dim FnDeCryHexToStr$
   FnDeCryHexToStr = Trim(strCrop1(globs(2), "=", "("))
   
   Dim b
   b = Split(globs(2), "=")
   
   Dim FnDeCryKeyVar$
   FnDeCryKeyVar = Trim(b(0))
   
   Dim FnDeCryHexToStrIn$
   FnDeCryHexToStrIn = Trim(strCrop1(globs(2), "'", "'"))
   If FnDeCryHexToStrIn = "" Then
      FnDeCryHexToStrIn = Trim(strCrop1(globs(2), """", """"))
   End If

   
   Dim FnDeCryHexToStrKey, q
   q = Split(NewScript, FnDeCryHexToStr)(2)
   q = Split(q, "=")(1)
   'FnDeCryHexToStrKey = strCrop1(NewScript, FnDeCryHexToStr, "StringLen", , InStrRev(NewScript, FnDeCryHexToStr))
   FnDeCryHexToStrKey = Split(q, vbCrLf)(0)
   
   If FnDeCryHexToStrKey Like "*0x*" Then
      FnDeCryHexToStrKey = CInt(HexToInt(Split(FnDeCryHexToStrKey, "0x")(1)))
   End If
   
   Dim AA&, BB&, Key&
   AA = (HexToInt(Mid(FnDeCryHexToStrIn, 1, 2))) - FnDeCryHexToStrKey
   BB = (HexToInt(Mid(FnDeCryHexToStrIn, 3, 2))) - FnDeCryHexToStrKey
   Key = Chr(AA) & Chr(BB)
   
  
   
  '==== Load String =====
  
   Dim i&, Line$, EncString$
   Dim VarNames As New Collection
   Dim VarValues As New Collection

   For i = 3 To UBound(globs)
      Dim tmpstr$
      tmpstr = globs(i)
      Line = strCropAndDelete(tmpstr, "", vbCrLf)
      globs(i) = tmpstr
   

      
      Dim Words
      Words = Split(Line, "'")
      If UBound(Words) = 0 Then
         Words = Split(Line, """")
      End If
      
      EncString = Words(1)
      
      Dim VarName$
      VarName = Trim(strCrop1(Words(0), "$", "="))
         
      Dim DecString$
      DecString = Space(Len(EncString) \ 2)
      
      Dim pos&
      For pos = 1 To Len(EncString) Step 2
         Mid(DecString, (pos \ 2) + 1) = Chr((HexToInt(Mid(EncString, pos, 2))) - Key)
      Next
      
      VarNames.Add VarName
      VarValues.Add DecString
      
   If tmpstr <> "" Then Exit For
'      log VarName & "   " & DecString
      
   Next
   
'  === Join Script & Remove EncodeIT "Global Const "
   Dim j&
   For j = i To UBound(globs)
      globs(j - i) = globs(j)
   Next
   
   ReDim Preserve globs(j - i - 1)
   NewScript = Join(globs, "Global Const ")
   GUI_Show_Text NewScript
   
   Log "Search'n'replace strings: "
   
'  ==== Replace Strings =====
   
   Dim ItemIndex&
   For ItemIndex = 1 To VarNames.Count
'       log VarNames(ItemIndex) & "   " & VarValues(ItemIndex)
      
      NewScript = Replace(NewScript, CStr("$" & VarNames(ItemIndex)), CStr("""" & Replace(VarValues(ItemIndex), """", """""") & """"))
      
      Log CStr("$" & VarNames(ItemIndex)) & "    " & """" & Replace(VarValues(ItemIndex), """", """""") & """"

   Next
   
   GUI_Show_Text NewScript
   Log "Optimizing strings..."
   'Sample:
   'Blah = "E" & "n" & "C" & "o" & "d" & "e" & "I" & "t"
   '        ^^^^^^
   '        Matchstring
   'Blah = "EnCodeIt"
   
   Const SingleOrDoubleQuotes$ = "[""']"
   Const BackRef2$ = "\2"
   
   myRegExp.Pattern = RE_Group(RE_AnyChar) & RE_WSpace(RE_Group(SingleOrDoubleQuotes), "&", BackRef2)
   
   myRegExp.IgnoreCase = False
   myRegExp.Global = True
   NewScript = myRegExp.Replace(NewScript, "$1")
   
   
   GUI_Show_Text NewScript
   Log "Processing function names and arguments:"
   
' ======== Functions

'  'Securitycheck last  functionname=FnDeCryHexToStr
'   Dim funcName$
'   funcName = Trim(Split(Funcs(UBound(Funcs)), "(")(0))
'   Debug.Assert (FnDeCryHexToStr <> funcName)

  Dim Funcs, FuncIdx
  Funcs = Split(NewScript, vbCrLf & "Func ")
  
' Remove last 3 functions added by EncodeIt
  ReDim Preserve Funcs(UBound(Funcs) - 3)
  
  
  For FuncIdx = LBound(Funcs) + 1 To UBound(Funcs)

   ' getName
   Dim FuncName$
   FuncName = Trim(Split(Funcs(FuncIdx), "(")(0))
   
'   Debug.Assert (FnDeCryHexToStr <> funcName) ' and (FuncIdx=UBound(Funcs)-2)

   ' add name for replace later
   DeObfu_FunctionNames.Add FuncName

   ' seperate Body & rest
   Dim funcBody$, funcRest$, tmp
   tmp = Split(Funcs(FuncIdx), vbCrLf & "EndFunc")
   funcBody = tmp(0)
   funcRest = tmp(1)

   'replace
   DeObfu_ReplaceFnArgs funcBody, FuncName
   ' writeback

   tmp(0) = funcBody
   Funcs(FuncIdx) = Join(tmp, vbCrLf & "EndFunc")


  Next

  NewScript = Join(Funcs, vbCrLf & "Func ")
   
  
  
  DeObfu_ReplaceFunctionNames NewScript
   
End Sub


Public Sub Init()
   Log "DeObfuscating : " & FileName
   
 ' Clear ModulVars
   Set DeObfu_FunctionNames = New Collection
 
 
 '=== Open & Test Obfu-Type ===
   
   File.Create FileName.FileName, False, False, True
   
  'Convert to Accii if it's unicode


   If IsUTF16File Then
      NewScript = StrConv(File.FixedString(-1), vbFromUnicode, LocaleID)
   Else
      If IsUTF8File Then
         NewScript = DecodeUTF8(File.FixedString(-1))
      Else
         File.Position = 0
         NewScript = File.FixedString(-1)
      End If
   End If
   
   File.CloseFile
 
 ' Was file detokenised?
   TestIsDeTokeniserOutput NewScript

End Sub

Public Function DetectScriptObfuscated() As Long
 
GUIEvent_ProcessBegin 6
 
 ' Test for if obfuscated ( find 'func A2C4558BC554')
' Option Compare Text
'   Dim UCaseScript$
'   UCaseScript = UCase(NewScript)

  'Default case
   DetectScriptObfuscated = DE_OBFUSC_TYPE_NOT_OBFUSC
   
   Dim myRegExp As New RegExp
   With myRegExp
   
       .IgnoreCase = True
       
'TIDY uses these strings to recognise the obfuscator:
' func A2C4558BC554
' Execute(Binarytostring('0x46696C6544656C6574652824414631323331323329'))
' Execute(Binarytostring('0x537472696E6753706C69742846696C655265616428244146313233313233292C27
' (^ Note the unclose String("Execute(Binarytostring('0x...2C27") string is no Error)

       .Pattern = RE_WSpace("Func", "A2C4558BC554")
       If .Test(NewScript) Then
          DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE
       Else
       
       GUIEvent_ProcessUpdate 1
       .Pattern = """4074656D70646972"""
       If .Test(NewScript) Then
          DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE
      ' "@TempDir"
       Else
       
     ' Generic VANZANDE (without stringEncryption)
       .Pattern = "Func\s*(A[0-9A-F]{10})\s*\("
       If .Test(NewScript) Then
          DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE
      ' "@TempDir"
       Else
       
       
       
       
       
       GUIEvent_ProcessUpdate 2
       .Pattern = "['""]4054656D70446972['""]"
       If .Test(NewScript) Then
          DetectScriptObfuscated = DE_OBFUSC_VANZANDE_VER24
      
      
      
       'Obfuscator v1.0.24.23  ( find 'func A1D00103002')
       Else
       
       GUIEvent_ProcessUpdate 3
       .Pattern = "A1D00103002"
       If .Test(NewScript) Then
          DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE

       Else
       
      '   Endif GUIEvent_ProcessUpdate
      
       
       GUIEvent_ProcessUpdate 4
      'Tidy            Int(99 / 3 + 15 * 100 / 4 - 13 ^ 2 + 81 / 3 - 17 - 245 + 99 / 3 + 15 * 100 / 4 - 13 ^ 2 + 81 / 3 - 17)
      'Original        Int(99/3+15*100/4-13^2+81/3-17-245+99/3+15*100/4-13^2+81/3-17)
      'RegExp-Encoded  Int\(99/3\+15\*100/4-13\^2\+81/3-17-245\+99/3\+15\*100/4-13\^2\+81/3-17\)
      ' .Pattern = RE_WSpace("Int", "\(", "99", "/", "3", "\+", "15", "\*", "100", "/", "4", _
                  "-", "13", "\^", "2", "\+", "81", "/", "3", "-", "17", "-", "245", "\+", _
                  "99", "/", "3", "\+", "15", "\*", "100", "/", "4", "-", "13", "\^", "2", "\+", _
                  "81", "/", "3", "-", "17", "\)")
       
       Const HexNum = "(?:" & "0x" & ")?" & "[0-9A-Fa-f]" & "{1,8}"
       .Pattern = RE_WSpace("Int", "\(", HexNum, "/", "3", "\+", HexNum, "\*", HexNum, "/", "4", _
                  "-", HexNum, "\^", "2", "\+", HexNum, "/", "3", "-", HexNum, "-", HexNum, "\+", _
                  HexNum, "/", "3", "\+", HexNum, "\*", HexNum, "/", "4", "-", HexNum, "\^", "2", "\+", _
                  HexNum, "/", "3", "-", HexNum, "\)") & RE_NewLine
       If .Test(NewScript) Then
          
          DetectScriptObfuscated = DE_OBFUSC_TYPE_ENCODEIT
          
       Else
       
       Const RE_255DigitNumber$ = "[0-2]?[0-9]?[0-9]"
       
       GUIEvent_ProcessUpdate 5
     ' Example: &CHR(7+108) ''       If NewScript Like "*Chr(##?[+-]?##*" Then
       .Pattern = RE_WSpace("&", "Chr", "\(", RE_255DigitNumber, "[+-]", RE_255DigitNumber, "\)")
       If .Test(NewScript) Then

          DetectScriptObfuscated = DE_OBFUSC_TYPE_CHR_ENCODE
          
       Else
       
       
       GUIEvent_ProcessUpdate 6
       
       GUIEvent_ProcessUpdate 5
       ' Example: &CHR(101)&""
       .Pattern = RE_WSpace("Chr", "\(", RE_255DigitNumber, "\)", "&", """")
       If .Test(NewScript) Then
          DetectScriptObfuscated = DE_OBFUSC_TYPE_CHR_ENCODE_OLD
       End If
       
       
       
       End If '5
       End If '4 DE_OBFUSC_TYPE_ENCODEIT
       End If '3
       End If '2
       End If '11
       End If '1
       End If '0
       
       
      If DetectScriptObfuscated = DE_OBFUSC_TYPE_VANZANDE Then
         If InStr(1, NewScript, "lobal $Os") Then
            DetectScriptObfuscated = DE_OBFUSC_VANZANDE_VER15
         
         ElseIf InStr(1, NewScript, ".tbl") Then
            
            DetectScriptObfuscated = DE_OBFUSC_VANZANDE_VER15_2
         
         Else
            DetectScriptObfuscated = DE_OBFUSC_VANZANDE_VER14
         End If
      End If

       
       
       GUIEvent_ProcessEnd
       
       
   End With
   
End Function

Private Sub WarnIfUntestedObfuType()

         If IsTidyedDeTokeniserOutput = False Then
            Log "Warning - using a detokenise au3-file that was not tidyed!"
            Log "I didn't tested this with that Obfuscation type. If problems occure "
            Log "please run data\Tidy\Tidy.exe manually on this *.au3 and then drag it into myAutToExe."
         End If

End Sub



Public Sub DeObfuscate()
   
   
   Log "Testing for Scripts that were obfuscate by 'Jos van der Zande AutoIt3 Source Obfuscator v1.0.15 [July 1, 2007]' or 'EncodeIt 2.0'"
   Log String(79, "=")
 
   If IsTextFile = False Then
      Err.Raise ERR_NO_TEXTFILE, , "DeObfuscate skipped! " & FileName.NameWithExt & " seems to be no text file. (because it contains 00 Bytes)"
      Exit Sub
   End If
 
 ' Load File into 'newScript'
   Init
   
'   Dim DeobfuRounds%
'   DeobfuRounds = 0
   Do
   
      IsTidyed = TestIsTidyed(NewScript)
      IsTidyedDeTokeniserOutput = IsDeTokeniserOutput And IsTidyed

      ' Test for if obfuscated
      'my dilettantic deObfuscator                                      (okay It still got better and is not that dilettantic it was before :))
      Select Case DetectScriptObfuscated
         Case DE_OBFUSC_VANZANDE_VER24
            Log "Deobfuscating van Zande 1.0.24.22..."
            DeObfuscate_VanZande1_0_24

         Case DE_OBFUSC_VANZANDE_VER15
               Log "Deobfuscating van Zande 1.0.15..."
               DeObfuscate_VanZande1_0_24 'DeObfuscate_VanZande1_0_15

         Case DE_OBFUSC_VANZANDE_VER15_2
               Log "Deobfuscating van Zande 1.0.24..."
               WarnIfUntestedObfuType
               DeObfuscate_VanZande1_0_15


         Case DE_OBFUSC_VANZANDE_VER14
               Log "Deobfuscating van Zande 1.0.14..."
               DeObfuscate_VanZande1_0_14


         Case DE_OBFUSC_TYPE_ENCODEIT

            Log "Deobfuscating Encodeit 2.0..."
            DeObfuscate_Encodeit

         Case DE_OBFUSC_TYPE_CHR_ENCODE

            Log "Deobfuscating Chr_Encode..."
            DeObfuscate_Chr_Encode

         Case DE_OBFUSC_TYPE_CHR_ENCODE_OLD

            Log "Deobfuscating Chr_Encode(old style)..."
            DeObfuscate_Chr_Encode_old

         Case DE_OBFUSC_TYPE_NOT_OBFUSC

          ' Remove DeTokenise marker and run Tidy
           If IsTidyed = False Then
               'NewScript = Split(NewScript, DETOKENISE_MAKER)(0)
               RunTidy NewScript
            End If

            Err.Raise ERR_NO_OBFUSCATE_AUT ', , "Script is NOT obfuscate by 'Jos van der Zande AutoIt3 Source Obfuscator v1.0.15 [July 1, 2007]', v1.0.14[June 16, 2007] or 'EncodeIt 2.0'"
            Exit Sub


      End Select
         
      Log "Deobfucation succeed."
      
      SaveScriptData NewScript, True 'IsTidyed
      
'      Inc DeobfuRounds
   
   Loop While True
   
End Sub

Function StringToHexString$(Txt_String)
   Dim tmpBuff As New StringReader
   tmpBuff = Txt_String
   StringToHexString = ValuesToHexString(tmpBuff, "")
End Function

Private Sub DeObfuscate_VanZande1_0_24()
' In v1.0.24.23 Seperator is wrapped up into "Execute(Binarytostring('0x"- 3 times
' Like this:
'FileInstall('test_org.au3.tbl',$AF123123,1)
'Global $Os = Execute(Binarytostring('0x457865637574652842696E61....
'                                       Execute(Binarytostring('0x457865637574652842696E6...
'                                                                 Execute(Binarytostring('0x537472696E6753706C6974...
'                                                                                           StringSplit(FileRead($AF123123...
   
   
   Dim tmp2$
   tmp2 = NewScript
   Do
      Dim tmp$
      tmp = strCrop(tmp2, _
                  StringToHexString("Execute(Binarytostring('0x"), _
                  StringToHexString("'))"))
      
   If tmp = "" Then Exit Do
      tmp2 = HexStringToString(tmp)
   Loop While True
   
     
   'VanZande1_0_24
   Dim TBLStringSeperator$
   
   '0x537472696E6753706C69742846696C655265616428244146313233313233292C27 | 69354F | 272C3129
   '                                   StringSplit(FileRead($AF123123),' | i5O    | ',1)
   'TBLStringSeperator = HexStringToString(strCrop1(tmp2, StringToHexString("StringSplit(FileRead($AF123123),'"), StringToHexString("',1)"), "4F"))
   'UPDATE                             StringSplit(FileRead($A460000414Esz_),'T5BDO',1)
   TBLStringSeperator = HexStringToString(strCrop1(tmp2, StringToHexString("),'"), StringToHexString("',1)"), "4F"))
   Log "TBL seperator string: " & TBLStringSeperator
   DeObfuscate_VanZande1_0_15 TBLStringSeperator

End Sub


Private Sub DeObfuscate_VanZande1_0_15(Optional TBLStringSeperator = "O")


 ' Merge lines with _ at the end + Add vbCrLf after last EndFunc
   ReplaceDo NewScript, AU_NEWLINE, " " ' & vbCrLf
   
''Correct case (may it's safer to let tidy do)
'   NewScript = Replace(NewScript, "Func ", "Func ", , , vbTextCompare)
'   NewScript = Replace(NewScript, "EndFunc ", "EndFunc ", , , vbTextCompare)
'   NewScript = Replace(NewScript, "IsDeclared", "IsDeclared", , , vbTextCompare)
'   NewScript = Replace(NewScript, "EndIf", "EndIf", , , vbTextCompare)

   
   Dim OSString$
   OSString = strCropWithSeek(NewScript, "($", "[")
   
      
   ReplaceDo NewScript, OSString, "Os", , , vbTextCompare
   
   ReplaceDo NewScript, "$Os [", "$Os[", , , vbTextCompare


'=== Replace Dim Const and Globals

'   GUI_Show_Text NewScript
   
'   DeObfu_ReplaceGlobals NewScript
'   GUI_Show_Text NewScript
'   qw
   
'   DeObfu_ReplaceDimsAndConst NewScript
'   GUI_Show_Text NewScript

   Log "Okay. Obfucated script loaded and displayed."
   
'=== Prepare ===


'    strCropAndDelete NewScript, "Global $", "StringTrimLeft(String(StringToBinary(""\"" & Hex(Random(256, 268435455, 1), 7))), 2)"

  ' Example Data:
  ' global $Os
  ' A07000035222()
'Other Example
'IF NOT ISDECLARED("Os")THEN GLOBAL $OS
'#OnAutoItStartRegister "A5C00002713_"
    Dim FnNameLoadTBL$
 ' Old function
 '  FnNameLoadTBL$ = strCropAndDelete(NewScript, "global $Os" & vbCrLf, "()")
      
  ' Get
    Dim FnNameLoadTBL_Match As match
    FnNameLoadTBL = RE_FindPattern(NewScript, _
                 "[ -~]*?" & RE_Literal("global $Os") & RE_NewLine & _
                 RE_AnyChars & RE_Group("[0-9A-Fa-f_]{12}") & RE_AnyChars, _
                 FnNameLoadTBL_Match)
  'Note: [ -~] goes from 0x20 to 0x7E and is there to retain UTF-8 BomMarker
    
    On Error Resume Next
                 
  ' Delete match
    ReplaceDo NewScript, FnNameLoadTBL_Match.value, ""
  
  ' Example Data:
  ' global $SA3700804E15=A0700003522($Os[0x1]),
  ' ^^^^^^^             ^
    Dim FnNameHexToString$
'    FnNameHexToString = strCropWithSeek(NewScript, "=", "($", vbCrLf & "global ")
   
  ' Get
    FnNameHexToString = RE_FindPattern(NewScript, _
                 RE_Literal("global $") & "[0-9A-Za-z_]*?" & RE_Literal("=") & _
                 RE_Group("[0-9A-Za-z_]*?") & RE_Literal("($"))
   
   
  
  
 Log "Extracting TBLFileName from Script."
  ' GetStringfile(*.tbl) from obfuscate script
    Dim StringTBLFileName As New ClsFilename
    StringTBLFileName.Path = FileName.Path
    
    Dim ExtractedTBLFileName$
    ExtractedTBLFileName = strCrop(NewScript, "(", ".tbl")
    If ExtractedTBLFileName = "" Then
       ExtractedTBLFileName = strCrop1(NewScript, "FileInstall(", ",")
       ExtractedTBLFileName = Mid(ExtractedTBLFileName, 2, Len(ExtractedTBLFileName) - 2)
       StringTBLFileName.NameWithExt = ExtractedTBLFileName
    Else
      
       StringTBLFileName.NameWithExt = Mid(ExtractedTBLFileName, 2) & ".tbl"
    End If
    
    
    
'    StringTBLFileName.NameWithExt = Mid(strCrop(NewScript, "FileInstall(", ".tbl"), 2) & ".tbl"
'    Dim line
'    For Each line In Split(NewScript, vbCrLf)
'    FileInstall('*.tbl",
'      StringTBLFileName.NameWithExt = strCrop(NewScript, "FileInstall('", "'")
'      Exit For
'    Next
    
 
   Log "Loading StringTBLFileName: " & StringTBLFileName.FileName & "..."
    
   Dim TBLFileData As New StringReader
   
   Do
      If Err Then
         StringTBLFileName.FileName = InputBox("Can't open TBL-file - " & Err.Description, "Please try to correct that path manually: ", StringTBLFileName.FileName)
         If StringTBLFileName.FileName = "" Then Exit Do
      End If
      
      On Error Resume Next
      TBLFileData.Data = FileLoad(StringTBLFileName.FileName)
   Loop While Err

' === Find Stringseperator ===
   Dim TBLStringSeperator2$
   TBLStringSeperator2 = TBLFindStringSeperator(TBLFileData)
   
' === Check TBLStringSeperator ===
   If IsValidSeperator(TBLFileData, TBLStringSeperator) Then
      ' okay
   ElseIf IsValidSeperator(TBLFileData, TBLStringSeperator2) Then

         Log "TBLStringSeperator seems to be wrong using heuristic one: " & TBLStringSeperator2
         TBLStringSeperator = TBLStringSeperator2
      
   Else
      Do While True
         TBLStringSeperator = InputBox("Please enter TBLStringSeperator that is used in the *.tbl file to seperate the HexStrings", "Finding TBLStringSeperator failed", TBLStringSeperator)
         Log "Finding TBLStringSeperator failed. Asked user and got: " & TBLStringSeperator
         
         If IsValidSeperator(TBLFileData, TBLStringSeperator) = False Then
            If vbYes = MsgBox("The entered TBLStringSeperator still seems to be wrong. Contiune anyway?", vbYesNo + vbDefaultButton2) Then Exit Do
         Else
            Exit Do
         End If
      Loop
   End If
   
    
   'Read Stringtable
    
    log_verbose "Dump of TBL-File:"
    Dim StringTBL
    ReDim StringTBL(1 To 1)
      Dim Line
      For Each Line In Split(TBLFileData.Data, TBLStringSeperator)
         Dim TBL_ID&
         TBL_ID = UBound(StringTBL)
         
         Dim TBL_Data$
         TBL_Data = HexStringToString(Line)
         
         StringTBL(TBL_ID) = TBL_Data
         ReDim Preserve StringTBL(1 To TBL_ID + 1)
         
         log_verbose "   " & H16(TBL_ID) & " " & TBL_Data
      
      
      Next
     'delete last two elements
     '  the last is always empty because Array is extented in advanced by ReDim Preserve
     '  the prelast is always "" because the tbl ends with the string seperator
     '      split takes this a start of the (last)element, that start after the stringseperator and ends at EOF)
      ReDim Preserve StringTBL(1 To UBound(StringTBL) - 1 - 1)


   
   
Log UBound(StringTBL) & " strings found."
 
  
  
  
Log "Restoring StringNames in ..."
  Dim LocalStringNames
  ReDim LocalStringNames(UBound(StringTBL))
  
  ' get local stringvars like
  'global $SA220010601E=A490000263F($Os[0x1]),$SA3A00204A08=A490000263F($Os[0x2]),$SA200030463A=A490000263F($Os[0x3])
   
   Dim MainIndexes As New Collection

   Dim Stat_NumLocal&, Stat_NumAtAll&
   Stat_NumAtAll = 0

 
   
   Dim NewScript2
   Dim EndToken$
   EndToken = "EndFunc" & vbCrLf
   
   If InStr(1, NewScript, EndToken) = 0 Then
      Log "WARNING - deobfucator needs a clean 'EndFunc' after each function!"
      Log "          Probably it currently looks like this: 'EndFunc   ;==>A27F180133A'. ReRun Tidy with '/Skip_EndFunc_Comment'(/sefc) switch or run it /Remove_EndFun_comment (/refc) on this file."
   End If
   
  
   NewScript2 = Split(NewScript, "Func ")
   Dim i&
   
 ' add new line at the end (to ensure function ends with "ENDFUNC\n" )
   NewScript2(UBound(NewScript2)) = NewScript2(UBound(NewScript2)) & vbCrLf
   
         
   
   GUIEvent_ProcessBegin UBound(NewScript2)
   
   Dim iter: i = 0
   For iter = LBound(NewScript2) To UBound(NewScript2)
      Dim ScriptFunction$
      ScriptFunction = NewScript2(iter)
      
 ' Show ScriptFunction Before
   GUI_Show_Text ScriptFunction
   
   GUIEvent_ProcessUpdate i
      
   
    ' Get Current FunctionName
      Dim FunctionName$
      If i = 0 Then
         FunctionName = "Main"
         Log "FunctionName: " & FunctionName
         
      Else
        
       ' Extract FunctionName, Store it for later
         FunctionName = strCrop1(ScriptFunction, "", "(")
         DeObfu_FunctionNames.Add FunctionName
         Log "FunctionName: " & FunctionName
         
              
       ' === Replace FunctionArguments  ====
         DeObfu_ReplaceFnArgs ScriptFunction, FunctionName
         
         
        'Check for Mainfunction chunks between FUNC and ENDFUNC
         Dim afterEndFunc$
         afterEndFunc = Split(CStr(ScriptFunction), EndToken)(1)
         If Replace(CStr(afterEndFunc), vbCrLf, "") <> "" Then
           '... apply String Restore
           
            Dim ItemCounter&
            ItemCounter = 0
            GUIEvent_ProcessBegin MainIndexes.Count, 1
            
            Dim index
            For Each index In MainIndexes
              DeObfu_ReplaceStrings ScriptFunction, LocalStringNames(index), StringTBL(index)
              
              GUIEvent_ProcessUpdate ItemCounter, 1
              Inc ItemCounter
              
            Next
            
            GUIEvent_ProcessEnd 1
            
         End If


'         GUI_Show_Text ScriptFunction
         

         
         
      End If
      
 ' Exit if FnNameLoadTBL is reached
   If FunctionName = FnNameLoadTBL Then
     'Don't exit if this function is not at the expected place
     '(only 3 other Van Zande functions(that can be deleted) should follow)
      Dim FunctionsCount&
      FunctionsCount = UBound(NewScript2)
      If RangeCheck(FunctionsCount - i, 3, 0) Then
      
       ' Cut of rest after FnNameLoadTBL
         ReDim Preserve NewScript2(i - 1)
  
         Exit For
      End If
   End If
      
'      Debug.Assert "A044140152D" <> FunctionName

     
    'Delete/ Seperate 'Obfu-Header' in ScriptFunction
    'store deleted part in obfuSection for further processing
    'Example for an 'Obfu-Header'
    '   If Not Isdeclared("SSA110020160C") Then
    '      Global $SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])
    '      Global $SSA110020160C=0x01
    '   End If
     Dim obfuSection$
     
      If i = 0 Then
         obfuSection = ScriptFunction
      Else

'         obfuSection = strCropAndDelete(tmpstr, "If Not Isdeclared(", "EndIF" & vbCrLf)
         obfuSection = strCropAndDelete(ScriptFunction, "If Not IsDeclared(", "EndIf" & vbCrLf)
'         'maybe also triggered by some normal 'IsDeclared'
'         Debug.Assert InStr(1, ScriptFunction, "IsDeclared", vbTextCompare) = 0
      End If
         
         
    ' Statistic Counter for String in Current function
      Stat_NumLocal = 0
      
      Do
         
         'Fix for long lines like this
         '  Global $A17CBF14A04 = A460000414E($OS[5102]), ... $A1BDB61490D = A460000414E($OS[5109]),
         '  $A4DFB714513 = A460000414E($OS[5142]), ...., $A3CFBD1063B = A460000414E($OS[5148])
          ReplaceDo obfuSection, "," & vbCrLf, ","
         
         
         'Process Local String Section
          Dim GlobalStringDef, CommandSeperateLines$
          CommandSeperateLines = strCropAndDelete(obfuSection, IIf(i = 0, "global ", "Global "), vbCrLf)
       '  Now CommandSeperateLines contains this:
       ' "$SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])"
      
      
    ' EXIT DO: if CommandSeperateLines dont contains sth like $Os[0x7] it's no obfu definition
      If InStr(1, CommandSeperateLines, "$Os[", vbTextCompare) = 0 Then Exit Do
          
         If i = 0 Then
            strCropAndDelete ScriptFunction, IIf(i = 0, "global ", "Global "), vbCrLf
         End If
          
          GlobalStringDef = Split(CommandSeperateLines, ",")
         'GlobalStringDef[0] = $SA3160B01D3E=A0000006239($Os[0x7])
         'GlobalStringDef[0] = $SA1D60C0011E=A0000006239($Os[0x8])
         '....
         
    'GUI ProgressBar
      Dim GUIEvent_ProcessCounter&
      GUIEvent_ProcessCounter = LBound(GlobalStringDef)
      GUIEvent_ProcessBegin UBound(GlobalStringDef), 1
         
         
        'for all items in GlobalStringDef do...
         Dim LocalStringName
         For Each LocalStringName In GlobalStringDef
             
             'Seperate GlobalStringDef further...
             Dim LocalStringName_Splitted
             LocalStringName_Splitted = Split(LocalStringName, "=")
             'LocalStringName_Splitted    with         $SA3160B01D3E=A0000006239($Os[0x7])
             'LocalStringName_Splitted[0] = $SA3160B01D3E
             'LocalStringName_Splitted[1] = A0000006239($Os[0x7])
             
             If LocalStringName_Splitted(1) Like "*$Os[[]0x*" Then
               index = CLng(HexToInt(strCrop1(LocalStringName_Splitted(1), "$Os[0x", "])")))
             Else
               index = CLng(strCrop1(LocalStringName_Splitted(1), "$Os[", "])"))
             End If
             
             LocalStringNames(index) = Trim(LocalStringName_Splitted(0))
            'LocalStringNames(&h7) = "$SA3160B01D3E"
            
             If i = 0 Then
               ' Store for later search&replace in mainfunction code between functions
                 MainIndexes.Add index
             End If

             'Replace Strings
            ' In ScriptFunction code Replace all "$SA3160B01D3E" with for example "Wrong Name Entered!"
'Debug.Print "Number(" & LocalStringNames(index) & ")", StringTBL(index)
'Debug.Assert index <> 966
              
'Debug.Assert Not (StringTBL(index) Like "gro*")
              DeObfu_ReplaceStrings ScriptFunction, LocalStringNames(index), StringTBL(index)
              
            Inc GUIEvent_ProcessCounter
            GUIEvent_ProcessUpdate GUIEvent_ProcessCounter, 1
         


'GUI_Show_Text CStr(index)
'myDoEvents

              
         Next

   GUIEvent_ProcessEnd 1
   
 ' Unspecific Increase the ProgressBar by 1% - just to show some progress
   GUIEvent_Increase 0.01
   
   
      Stat_NumLocal = Stat_NumLocal + UBound(GlobalStringDef)
      Loop While True
      

'Bench no stringReplace    1219
'Old stringReplace         15172 14813
'
      Log "  Local strings: " & Stat_NumLocal
      Stat_NumAtAll = Stat_NumAtAll + Stat_NumLocal
    
   ' Show ScriptFunction after
     GUI_Show_Text ScriptFunction

    
    
    
    ' Save New Function
      NewScript2(i) = ScriptFunction
   
     
   i = i + 1
   Next
  
 
GUIEvent_ProcessEnd
  
Log "Local strings at all: " & Stat_NumAtAll
Log "strings in Tbl: " & UBound(StringTBL)
  
  
   NewScript = Join(NewScript2, vbCrLf & "Func ")
  
  '=== Replace Functionname
   Log "Renaming FunctionNames..."
   DeObfu_ReplaceFunctionNames NewScript
  
  
   ' === Replace Locals ====
     DeObfu_ReplaceLocals NewScript
   
   ' === Scopeless Locals ====
     z_DeObfu_ReplaceX2 NewScript, "", "Var"
   
  


'  '=== Replace Global Strings

  Log "Restoring global strings..."
  
  GUIEvent_ProcessBegin MainIndexes.Count
  GUIEvent_ProcessCounter = 0
  
  For Each index In MainIndexes
      
      GUIEvent_ProcessUpdate Inc(GUIEvent_ProcessCounter)
      
      DeObfu_ReplaceStrings NewScript, CStr(LocalStringNames(index)), StringTBL(index)
  
  Next
  
  GUIEvent_ProcessEnd


  'unUsed String
'  Dim i&
   For i = LBound(StringTBL) To UBound(StringTBL)
      If LocalStringNames(i) = "" Then
         
         Dim bBlocker_ShowOnlyOnce As Boolean
         If bBlocker_ShowOnlyOnce = False Then
            Log "UnAssigned/Unused strings [Index, Value]: "
            bBlocker_ShowOnlyOnce = True
         End If

         Log "  " & H16(i) & " " & StringTBL(i)
      End If
   Next
    
  
 ' Clean up - delete *.tbl file
   If FrmMain.Chk_TmpFile.value = vbUnchecked Then
'      log "Deleting: " & StringTBLFileName.FileName
      FileDelete StringTBLFileName.FileName
   End If
  
 ' Remove Obfu Maker to avoid double detection
   ReplaceDo NewScript, "(""4054656D70446972"")", "(""Please remove me"")"
  
End Sub



'Function ReplaceAndDetectChange(ByRef Text$, ParamArray Search_Replace()) As Boolean
'   Dim i&
'   For i = 0 To UBound(Search_Replace) Step 2
'
'      Dim LenText
'      LenText = Len(Text)
'
''      Text = Replace(Text, CStr(Search_Replace(i)), CStr(Search_Replace(i + 1)))
'      ReplaceDo Text, CStr(Search_Replace(i)), CStr(Search_Replace(i + 1))
'
''      ReplaceBin Text, Text, Text, CStr(Search_Replace(i)), CStr(Search_Replace(i + 1)), 1, 2147483647
'
'      ReplaceAndDetectChange = (LenText <> Len(Text))
'      If ReplaceAndDetectChange = True Then Exit For
'
'   Next
'
'
'End Function


Private Sub DeObfu_ReplaceFunctionNames(ByRef ScriptCode$)

   On Error GoTo DeObfu_ReplaceFunctionNames_err
   Dim FnNamePure, FnNameCounter&: FnNameCounter = 0

   GUIEvent_ProcessBegin DeObfu_FunctionNames.Count, , True
        
   For Each FnNamePure In DeObfu_FunctionNames
   
    ' only rename obfuscate Names
      If IsObfuscatedFunctionName(FnNamePure) Then
        
        Dim FnNameNew$
        FnNameNew = "Fn" & H16(FnNameCounter)
       
        log_verbose FnNamePure & " -> " & FnNameNew
        
      ' Note: QuickReplace will be it twice as fast, but cause problems with strings containing functionnames like:
      '  for ex. this: TRAYITEMSETONEVENT(-1 ,"Fn0181     ")
      ' Well you could first replace all functionnames in strings - however...
        ReplaceDo ScriptCode, CStr(FnNamePure), FnNameNew
      
      End If
    
      
      Inc FnNameCounter
      GUIEvent_ProcessUpdate FnNameCounter
      
   Next
   
   GUIEvent_ProcessEnd
   
DeObfu_ReplaceFunctionNames_err:
Select Case Err
   Case 0
   Case ERR_SKIP
      Log "User Skipped/cancel Deobfucate->Rename FunctionNames."
      
   Case Else
      ErrThrow
   
End Select

End Sub

Private Sub DeObfu_ReplaceFnArgs(ByRef CodeSnipped$, ByRef FunctionName$)

  Dim ParamList$
  ParamList = strCrop1(CodeSnipped, FunctionName, vbCrLf)
  ParamList = strCrop1(CodeSnipped, "(", ")")
  
  Dim RawParams
  RawParams = Split(ParamList, ",")


   Dim FnArgument, FnNewName$, ArgCounter&
   ArgCounter = 0
   For Each FnArgument In RawParams
     
     FnNewName = "Arg"
     If FnArgument Like "*Const*" Then FnNewName = FnNewName & "C"
     If FnArgument Like "*ByRef*" Then FnNewName = FnNewName & "Ref"
     If FnArgument Like "*=*" Then
        FnNewName = FnNewName & "Opt"
        FnArgument = Split(FnArgument, "=")(0)
     End If
     FnNewName = FnNewName & H8(ArgCounter)
     
     
     Dim FnPureArgName$
     FnPureArgName = Trim(strCrop1(FnArgument, "$", ""))
'            tmpstr = CodeSnipped
     
     ReplaceDo CodeSnipped, CStr(FnPureArgName), FnNewName
'            CodeSnipped = tmpstr
     
     Inc ArgCounter
   Next

End Sub

'old version of z_DeObfu_ReplaceX... (requires Parser.bas)
'Private Sub z_DeObfu_ReplaceX(ByRef CodeSnipped$, KeyWord$, Prefix$, Optional IgnoreLinesWith = "$Os")
'   If IsMissing(IgnoreLinesWith) Then IgnoreLinesWith = "$Os"
'
'   Dim DuplicateFilter As Collection
'   Set DuplicateFilter = New Collection
'
'   Dim LocalDefCommaSep, LocalDefCounter: LocalDefCounter = 0
'   Dim LocalNameCounter&: LocalNameCounter = 0
'   For Each LocalDefCommaSep In Split(CodeSnipped, KeyWord)
'     'Skip first element
'      If LocalDefCounter > 0 Then
'         Dim LocalDefLine$
'         LocalDefLine = Split(LocalDefCommaSep, vbCrLf)(0)
'         If (LocalDefLine Like "*" & IgnoreLinesWith & "*") = False Then
'
'
'           'crop off functions args (Arg1,arg2,arg3...) that also contains ','
'            LocalDefLine = CropParenthesis(LocalDefLine)
'            LocalDefLine = CropParenthesis(LocalDefLine, "[", "]")
'
'            Dim LocalDef
'            For Each LocalDef In Split(LocalDefLine, ",")
'
'             ' Cut out Current Name
'               Dim LocalDefPureName$
'               LocalDefPureName = Split(LocalDef, "$")(1)
'               LocalDefPureName = Split(LocalDefPureName, "=")(0)
'               LocalDefPureName = Split(LocalDefPureName, ")")(0)
'               ReplaceDo LocalDefPureName, vbTab, ""
'               LocalDefPureName = Trim(Split(LocalDefPureName, "[")(0))
'
'               On Error Resume Next
'               DuplicateFilter.Add "", LocalDefPureName
'               If Err = 0 Then
'
'
'
'                ' Create New Name
'                  Dim LocalNameNew$: LocalNameNew = Prefix
'                  If LocalDef Like "*const*" Then LocalNameNew = LocalNameNew & "Const"
'                  If LocalDef Like "*[[]*" Then
'                     LocalNameNew = LocalNameNew & "Arr"
'                  Else
'                     LocalNameNew = LocalNameNew & "Var"
'                  End If
''                  LocalNameNew = BlockAlign_l(LocalNameNew & H16(LocalNameCounter), 11)
'                  LocalNameNew = LocalNameNew & H16(LocalNameCounter)
'
'   ' without myDoEvents
'   ' without Replace 485
'   ' Varsize = 15063
'   ' FixSize = 11031
'   ' dupfilter= 10797
'
'
'                ' Replace current with new name
'                  ReplaceDo CodeSnipped, LocalDefPureName, LocalNameNew
'                  'ReplaceDoMulti CodeSnipped, LocalDefPureName, LocalNameNew, vbTextCompare
'                  'DoEventsSeldom
'
'                  Inc LocalNameCounter
'                  GUI_Show_Text CVar(LocalNameCounter)
'                  myDoEvents
'               End If
'
'            Next
'         Else
'         End If
'      End If
'
'      Inc LocalDefCounter
''      If LocalNameCounter = 1000 Then Exit For
'   Next
'
'End Sub

Private Sub z_DeObfu_Replace_CacheVarNames(CodeSnipped$, ObfuName$, ByRef ObfuVarTree As Collection)
   
 ' Simple heuristic Len Test to see if ObfuVarTree is still valid
   Static CodeSnippedLenOld&
   If CodeSnippedLenOld = Len(CodeSnipped) Then Exit Sub
   
   
 'However there is still the chance that the two CodeSnipped's have same len
 'and matches and are still not equal / we'll miss matches - however chance is low
   
  
     Dim ObfuVar As New RegExp
     With ObfuVar
         .IgnoreCase = True
         .MultiLine = True
      
         .Pattern = "(\s*\$)" & ObfuName
         .Global = True
         
         Set ObfuVarTree = New Collection
         
         Dim Var As match
         For Each Var In ObfuVar.Execute(CodeSnipped)
            
            Dim SubMatchLen&
            SubMatchLen = Len(Var.SubMatches(0)) + 1
            
            Dim VarName$
            VarName = Mid(Var.value, SubMatchLen)
            
            Dim VarNamePos&
            VarNamePos = Var.FirstIndex + SubMatchLen
   
   
          ' Already in the collection?
            On Error Resume Next
            Dim ObfuVarBranch As Collection
            Set ObfuVarBranch = ObfuVarTree(VarName)
            If Err = 5 Then
             
             ' No so add it as new branch (with varName as key)
               Set ObfuVarBranch = New Collection
               ObfuVarTree.Add ObfuVarBranch, VarName
'               Debug.Print VarName
               
            End If
          ' add Var as leaf
            ObfuVarBranch.Add VarNamePos
         Next
         
         On Error GoTo 0
     End With
        
    CodeSnippedLenOld = Len(CodeSnipped)


End Sub

'Replace "<KeyWord> $A4201D12945" with "<KeyWord> $Var001D      "
'http://msdn.microsoft.com/en-us/library/1400241x.aspx
Private Sub z_DeObfu_ReplaceX2(ByRef CodeSnipped$, KeyWord$, Prefix$, Optional IgnoreLinesWith = "$Os")
  
  
  If IsMissing(IgnoreLinesWith) Then IgnoreLinesWith = "$Os"
'                    123456789012345678901234567890
' Example: "Global $A4201D12945"
' Example: "Global $J767027413071708815817

  Dim ObfuName$
  ObfuName = "[A-Z][0123456789ABCDEF]{10,21}"
  
  
  
  Dim GlobalVar As New RegExp
  With GlobalVar
      .IgnoreCase = True
      .MultiLine = True

      If KeyWord = "" Then
        ' Pattern Don't need to start at the beginning
         .Pattern = RE_WSpace( _
                        "\$" & RE_Group(ObfuName) & "\b", _
                        RE_LookHead_negative(RE_AnyCharRepeat() & RE_Literal(IgnoreLinesWith)))
      Else
   
         .Pattern = RE_WSpace( _
                        RE_Anchor_LineBegin, KeyWord, _
                        "\$" & RE_Group(ObfuName), "[[=]", _
                        RE_LookHead_negative(RE_AnyCharRepeat() & RE_Literal(IgnoreLinesWith)))
      End If
      
      .Global = True
  

      Dim VarCounter&
      VarCounter = 0
      
      Dim isUniqueItem As New clsDuplicateFilter
      isUniqueItem.Clear
      
      
'   BenchStart
               
   Dim RE_Searchpattern As New clsStrCat
            
 ' Build/Update Var Search&Replace Cache (used to speed up Search&Replace)
   Static ObfuVarTree As Collection
   z_DeObfu_Replace_CacheVarNames CodeSnipped, ObfuName, ObfuVarTree
  
  
    ' Go through all matches and search'n'replace them with a new name
      Dim GlobalVar_Match As match
      For Each GlobalVar_Match In GlobalVar.Execute(CodeSnipped)
   
         Dim SearchString$
         SearchString = GlobalVar_Match.SubMatches(0)
'         SearchString = MatchesSplitted(i + 1)

        'filter out names where seach&replace was already done
         If isUniqueItem(SearchString) Then

            Dim ReplaceString$
            ReplaceString = Prefix & H16(VarCounter)
           
           'fill up ReplaceString with space to bring it to same size
            If Len(SearchString) > Len(ReplaceString) Then
                ReplaceString = BlockAlign_l(ReplaceString, Len(SearchString))
            Else
               Stop
               'todo handle replacestring is bigger than searchstr
            End If
            
            log_verbose SearchString & " -> " & ReplaceString
            
         ' Replace one Var
         ' Well since the cache build&update function uses the len
         ' of the data were the Search&replace should be applied
         ' to detect if the cache needs to be rebuild is unreliable
         ' Cache maybe outdate and ObfuVarTree($BlahBlah) raise an error
         ' - for this case there is a slower seach&replace function...
Try_DoReplaceWithCache:
            On Error GoTo Catch_DoReplaceWithCache

'            Debug.Print SearchString
'            Debug.Assert "J767027413071708815817" <> SearchString
            
            Dim ReplacePos
            For Each ReplacePos In ObfuVarTree(SearchString)
               If Mid(CodeSnipped, ReplacePos, Len(SearchString)) = SearchString Then
                  Mid(CodeSnipped, ReplacePos) = ReplaceString
               Else
                ' Raise and error to do a normal 'ReplaceDo NewScript, SearchString, ReplaceString'
                  Err.Raise vbObjectError
                  Stop
               End If
            Next
            
   GoTo Finally_DoReplaceWithCache
Catch_DoReplaceWithCache:

   Resume Try_DoReplaceWithoutCache
Try_DoReplaceWithoutCache:

'           'Put a notice on that error in the log
'            Dim bBlocker_ReportErrorOnlyOnce As Boolean
'            If bBlocker_ReportErrorOnlyOnce = False Then
'               bBlocker_ReportErrorOnlyOnce = True
'
'               Log "An expected error situation occured in 'ClsDeobfuscator::z_DeObfu_ReplaceX2' - when replacing '" & SearchString & "'. Well there's an errorhandler that can deal&correct this however it would be nice to see the script that provoked this. So if you like please mail me this script."
'            End If
            
            
            ReplaceDo NewScript, SearchString, ReplaceString
'            ReplaceDoMulti NewScript, SearchString, ReplaceString
            
Finally_DoReplaceWithCache:
         
            Inc VarCounter
            GUI_Show_Text CStr(VarCounter)
            myDoEvents
         End If
            
      Next

Exit Sub

            
'   CodeSnipped = Join(MatchesSplitted, "")
   
' Bench without ReplaceDo              438    469
' Bench with    ReplaceDo    FixLen  10485  10547 ( 8125)
' Bench with    Replace()    FixLen  12719
' Bench with    ReplaceDoMulti xLen  13546  13563
' Bench with    ReplaceDo    VarLen  17625  17313 (15047)
' Bench with    vba.Replace()        44734  45110
' Bench with    regexp.Replace() FixLen   18281  18344
' Bench with    regexp.Replace() FixLen   27218  27172 (.IgnoreCase = True)
   End With
   

End Sub

Private Sub DeObfu_ReplaceGlobals(ByRef CodeSnipped$, Optional IgnoreLinesWith = "$Os")

   Log "Renaming Globals..."
'   z_DeObfu_ReplaceX CodeSnipped, "Global", "g", IgnoreLinesWith
   DeObfu_Replacer CodeSnipped, "Global", "g", IgnoreLinesWith
End Sub

Private Sub DeObfu_Replacer(ByRef CodeSnipped$, KeyWord$, Prefix$, Optional IgnoreLinesWith = "$Os")
   If IsMissing(IgnoreLinesWith) Then IgnoreLinesWith = "$Os"

' 'Global Const $A4A14104939 = $A2014201D1E'
   z_DeObfu_ReplaceX2 CodeSnipped, KeyWord$ & " Const", Prefix & "Const", IgnoreLinesWith

' 'Global $A4201D12945 = '       OR
   z_DeObfu_ReplaceX2 CodeSnipped, KeyWord, Prefix & "Var", IgnoreLinesWith


End Sub



Private Sub DeObfu_ReplaceDimsAndConst(ByRef CodeSnipped$, Optional IgnoreLinesWith)

' Glocal  Const   -> gConst
' Glocal          -> gVar
Log "Renaming Consts..."
   DeObfu_Replacer CodeSnipped, "Global", "g", IgnoreLinesWith
log_verbose String(79, "_")

' Local   Const   -> lConst
' Local           -> lVar
Log "Renaming Locals..."
   DeObfu_Replacer CodeSnipped, "Local", "l", IgnoreLinesWith
log_verbose String(79, "_")

' Dim     Const   -> dConst
' Dim             -> dVar
Log "Renaming Dims..."
   DeObfu_Replacer CodeSnipped, "Dim", "gd", IgnoreLinesWith
log_verbose String(79, "_")

'
'         Const   -> Const
'                 -> Var
Log "Renaming Const..."
   DeObfu_Replacer CodeSnipped, "", "", IgnoreLinesWith
log_verbose String(79, "_")

'Log "Renaming Dims..."
'   z_DeObfu_ReplaceX CodeSnipped, "Dim", "gDim", IgnoreLinesWith
'Log "Renaming Consts..."
'   z_DeObfu_ReplaceX CodeSnipped, vbCrLf & "Const", "gConst", IgnoreLinesWith
'
'   z_DeObfu_ReplaceX CodeSnipped, vbCrLf & "Global Const ", "glConst", IgnoreLinesWith
     
   
End Sub

Private Sub DeObfu_ReplaceLocals(ByRef CodeSnipped$)
' Log "Renaming Locals..."
   z_DeObfu_ReplaceX2 CodeSnipped, "Local ", "Local"
End Sub

Private Function c(Text$)
   If IsDeTokeniserOutput Then
      c = UCase(Text)
   Else
      c = Text
   End If
End Function


Private Sub DeObfu_ReplaceStrings(ByRef Data$, ByRef ReplaceStr, ByRef NewStr)

' ' avoid Replacing ShellExecute
   With myRegExp
      .Global = True
      .IgnoreCase = False
      
' Nice but somehow too slow (due to '|')
'      .Pattern = RE_WSpace("(" & _
'                  RE_Literal(C("Number(")) & ")|(" & _
'                  RE_Literal(C("ShellExecute(")) & ")|(" & _
'                  RE_Literal(C("Execute(")) & _
'                  ")", _
'                     RE_Literal(ReplaceStr), _
'                  RE_Literal(")"))
'
'      Dim match As match
'      For Each match In .Execute(Data)
'         With match
'            If .SubMatches(1) <> Empty Then
'             ' ShellExecute
'              'Do Nothing (ignore this)
'
'            ElseIf .SubMatches(2) <> Empty Then
'             ' Macros/Commands...
'               Stop
'
'            ElseIf .SubMatches(0) <> Empty Then
'             ' Numbers...
'               NewStr = FormatNumber(NewStr)
'               If Len(NewStr) > .Length Then
'                  'Extend String
'                  Stop
'                  ReplaceDo Data, .value, NewStr, .FirstIndex
'               Else
'                  Mid(Data, .FirstIndex + 1, .Length) = BlockAlign_l(FormatNumber(NewStr), .Length)
'               End If
'            Else
'
'               ' Strings...
'
'               myRegExp.Pattern = RE_WSpace(RE_Literal(ReplaceStr), "")
'               If myRegExp.Test(Data) Then
'                  ReplaceDo NewStr, """", """"""
'                  Data = myRegExp.Replace(Data, RE_Replace_Literal(Quote(NewStr)))
'               End If
'
'            End If
'
'         End With
'      Next
      
'-------------------------------

' Macros/Commands...
    ' No ShellExecute(...) but catch all Execute(...)
    ' This is done by 'consuming' ShellExecute and later writing it back via $1
'     .Pattern = RE_WSpace(RE_LookHead_negative( _
                              RE_Literal(C("ShellExecute("))) & _
                           RE_Group(RE_AnyChar & RE_Repeat(5, 5)) & _
                           RE_Literal(C("Execute(")), _
                           RE_Literal(ReplaceStr), _
                           RE_Literal(")"))
'         Data = .Replace(Data, "$1" & RE_Replace_Literal(NewStr)) & " "

'-----Version 1
     .Pattern = RE_WSpace(c("Execute\("), RE_Literal(ReplaceStr), "\)")
      
      Dim bExecuteFound As Boolean
      bExecuteFound = .Test(Data)
      
      If bExecuteFound Then

         Dim bShellExecuteFound As Boolean
         bShellExecuteFound = False

         Dim match As match
         For Each match In .Execute(Data)
            With match
              'Skip ShellExecute
               If Mid(Data, .FirstIndex + 1 - 5, 5) <> c("Shell") Then
                  QuickReplace Data, .value, NewStr, .FirstIndex + 1, 1
               Else
                ' Notify that there's a ShellExecute that might contain strings that need to replaced
                  bShellExecuteFound = True
               End If
            End With
         Next
      End If
         
'-----Version 2 - this is 9x slower than version1 - probably because of '|' Alternation
'     .Pattern = RE_Group_NonCaptured(c("ShellExecute\(")) & "|" & _
'                RE_Group(RE_WSpace(c("Execute\("), RE_Literal(ReplaceStr), "\)"))
'
'      If .test(Data) Then
'
'         Dim Match As Match
'         For Each Match In .Execute(Data)
'            With Match
'              'Skip ShellExecute
'               If .SubMatches(0) <> Empty Then
'
'                 'Do Replacing
'                  QuickReplace Data, NewStr, .FirstIndex + 1, .Length, .value
'               End If
'            End With
'         Next
'-----------------------
      
    ' Continune if there was no Execute() match
    ' ...and when a ShellExecute() was found( <-what is no real match)
      If bShellExecuteFound Or (bExecuteFound = False) Then
' Numbers...
        .Pattern = RE_WSpace(RE_Literal(c("Number(")), RE_Literal(ReplaceStr), RE_Literal(")"))
         If .Test(Data) Then
            
            Data = .Replace(Data, FormatNumber(NewStr) & " ")
            
         Else

' Strings...
'
            QuickReplace Data, ReplaceStr, MakeAutoItString(CStr(NewStr))

' RegExp Version (Speed is nearly the same - slighly slower)
'            .Pattern = RE_Literal(ReplaceStr)
'            If .test(Data) Then
'               Data = .Replace(Data, RE_Replace_Literal(MakeAutoItString(CStr(NewStr))))
'            End If
         
         End If

      End If
   End With
   
   

'If "$H729026783220739177320" = Data Then Stop

'      ReplaceDo Data, "Number(" & ReplaceStr & ")", FormatNumber(NewStr), , , vbTextCompare
'      ReplaceDo Data, ReplaceStr, Quote (Replace(NewStr, """", """""") )

End Sub

Private Sub StringFillUp(ByRef Data, TargetSize&)
   If Len(Data) < TargetSize Then
      Data = BlockAlign_l(Data, TargetSize)
   End If
End Sub

Private Sub QuickReplace(Data$, ReplaceStr, ByVal NewStr$, Optional startPos& = 1, Optional Count& = 2147483647)
   StringFillUp NewStr, Len(ReplaceStr)
   
   ReplaceDo Data, ReplaceStr, NewStr, startPos, Count
'   '-
'   'Do Replacing
'    If Len(NewStr) > NewLen Then
'       'Extend String
'       Stop
'       ReplaceDo Data, oldValue, NewStr, NewPos
'    Else
'       Mid(Data, NewPos, NewLen) = BlockAlign_l(NewStr, NewLen)
'    End If

End Sub

Private Sub DeObfuscate_VanZande1_0_14()

   'Merge lines with _ at the end
    NewScript = Replace(NewScript, AU_NEWLINE, " ")

   Dim badName$
   badName = "$VAR0"
   If NewScript Like "*" & badName & "*" Then
      NewScript = Replace(NewScript, badName, badName & "_")
      Log "WARNING: Found badName " & badName & "* in script and replace it with " & badName & "_"
   End If
   
   badName = "gConst"
   If NewScript Like "*" & badName & "*" Then
      NewScript = Replace(NewScript, badName, badName & "_")
      Log "WARNING: Found badName " & badName & "* in script and replace it with " & badName & "_"
   End If
   
   badName = "FN00"
   If NewScript Like "*" & badName & "*" Then
      NewScript = Replace(NewScript, badName, badName & "_")
      Log "WARNING: Found badName " & badName & "* in script and replace it with " & badName & "_"
   End If

   GUI_Show_Text NewScript

   Log "Okay. Obfucated script loaded and displayed."
  
  
   Dim StringTBL As New Collection
  

  Dim LocalStringNames
  ReDim LocalStringNames(StringTBL.Count)
  
  ' get local stringvars like
  'global $SA220010601E=A490000263F($Os[0x1]),$SA3A00204A08=A490000263F($Os[0x2]),$SA200030463A=A490000263F($Os[0x3])
   
   Dim fnNameDecrypt$
   
   Dim MainIndexes As New Collection

   Dim Stat_NumLocal&, Stat_NumAtAll&
   Stat_NumAtAll = 0

  '=== Replace Dim Const and Globals
   
   
  ' global $SA60D010561B=A0300004831("41415354")
  '                      ^^^^^^^^^^^-> fnStringDecrypt
      Dim NewScript2
      NewScript2 = Split(NewScript, "Func ", , vbTextCompare)
      
      Dim FunctionsCount&
      
      FunctionsCount = UBound(NewScript2)
      
      Dim fnStringDecrypt$
    ' fnStringDecrypt is the Last FunctionName
      fnStringDecrypt = strCrop1(NewScript2(FunctionsCount - 1), "", "(")
   
  'Ignore all lines with "A0300004831()" ( <-fnStringDecrypt ) to speed up replace
  'lines with "A0300004831()" will be deleted later anyway
'!!!   DeObfu_ReplaceGlobals NewScript, fnStringDecrypt


Log "Restoring StringNames in ..."
  
  
   Dim i&: i = 0
'   NewScript2 = Split(NewScript, "Func ", , vbTextCompare)
  
   GUIEvent_ProcessBegin FunctionsCount
   
   Dim iter&: iter = 0
   Do Until iter > FunctionsCount
   
   Dim ScriptFunction$
   ScriptFunction = NewScript2(iter)
 
    ' Get Current FunctionName
      Dim FunctionName$
      
    ' is first function chunk
      If i = 0 Then
         FunctionName = "Main"
         log_verbose "FunctionName: " & FunctionName
      Else
        
         FunctionName = strCrop1(ScriptFunction, "", "(")
'         Debug.Assert FunctionName <> "Stoper"
       
       'if FunctionName was Obfuscated Store it for later
'Debug.Assert "PRINTEOTSWP" <> FunctionName
       
         If IsObfuscatedFunctionName(FunctionName) Then
            DeObfu_FunctionNames.Add FunctionName
            log_verbose "Obfuscated FunctionName: " & FunctionName
            
        ' === Replace Locals ====
          DeObfu_ReplaceLocals ScriptFunction
         
        ' === Replace FunctionArguments  ====
          DeObfu_ReplaceFnArgs ScriptFunction, FunctionName

         Else
            log_verbose "FunctionName: " & FunctionName
         End If
         
         
      End If '//is first function chunk
      
    ' Exit if FnNameLoadTBL is reached
   '   If FunctionName = FnNameLoadTBL Then Exit For
         
         
   'Debug.Assert "PRINTEOTSWP" <> FunctionName
       
      
'    '  Skipping/cut out decrypt function
'    '    Also important to avoid the Au3 to be detected falsely as obfucated
'
'       If (FunctionName Like "A2C4558BC554*") Then
'
'       ' Cut function
'         Dec FunctionsCount
'         ReDim Preserve NewScript2(FunctionsCount)
'
'       Else
   
        
       'Delete/ Seperate 'Obfu-Header' in ScriptFunction
       'store deleted part in obfuSection for further processing
       'Example for an 'Obfu-Header'
       '   If Not Isdeclared("SSA110020160C") Then
       '      Global $SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])
       '      Global $SSA110020160C=0x01
       '   End If
        Dim obfuSection$
        
         If i > 0 Then
         
          ' Find out if it look like this
          ' If Not Isdeclared("SSA1810601658") Then
          ' Global $SA0...
          ' EndIF
          '
          ' OR
          '
          ' If Not Isdeclared("SSA1810601658") Then Global $SA0...
          ' <without EndIF>
            Dim isIf_EndIF_Type As Boolean
            On Error Resume Next
            Dim If_Not_Isdeclared_Chunks
            If_Not_Isdeclared_Chunks = Split(ScriptFunction, "If Not Isdeclared(", , vbTextCompare)(1)
            If Err = 0 Then
               On Error GoTo 0
               If_Not_Isdeclared_Chunks = Split(If_Not_Isdeclared_Chunks, vbCrLf)(0)
               isIf_EndIF_Type = (0 = StrComp(Right(If_Not_Isdeclared_Chunks, 5), " Then", vbTextCompare)) Or _
                                 (0 = StrComp(Right(If_Not_Isdeclared_Chunks, 5), ")Then", vbTextCompare))
               
               
               
               If isIf_EndIF_Type Then
                  obfuSection = strCropAndDelete(ScriptFunction, "If Not Isdeclared(", "EndIF" & vbCrLf)
               Else
                  obfuSection = strCropAndDelete(ScriptFunction, "If Not Isdeclared(", vbCrLf)
               End If
               
               
            Else
               On Error GoTo 0
               obfuSection = ""
            End If
         End If
      
            
       ' Statistic Counter for String in Current function
         Stat_NumLocal = 0
         
         Do
            
            'Process Local String Section
             Dim GlobalStringDef, CommandSeperateLines$
             
             If i = 0 Then
              'Delete first global
               CommandSeperateLines = strCropAndDelete(ScriptFunction, "global ", vbCrLf)
               ReplaceDo CommandSeperateLines, "Const ", "", , , vbTextCompare
             Else
             
               If isIf_EndIF_Type = False Then
                              'If_Not_Isdeclared_Chunks
                  obfuSection = strCropAndDelete(ScriptFunction, "If Not Isdeclared(", vbCrLf) & vbCrLf
                  
                  CommandSeperateLines = strCropAndDelete(obfuSection, "Global Const ", vbCrLf)
               Else
                  CommandSeperateLines = strCropAndDelete(obfuSection, "Global ", vbCrLf)
               End If
   
               
             End If
          
          '  Now CommandSeperateLines contains this:
          ' "$SA3160B01D3E=A0000006239($Os[0x7]),$SA1D60C0011E=A0000006239($Os[0x8]),...,$SA107030544C=A0000006239($Os[0xF])"
         
         
       ' EXIT DO: if CommandSeperateLines dont contains sth like $Os[0x7] it's no obfu definition
   '      If InStr(CommandSeperateLines, "Const ") Then Exit Do
         If InStr(CommandSeperateLines, "=0x01") Then Exit Do
         If CommandSeperateLines = "" Then Exit Do
         If InStr(CommandSeperateLines, fnStringDecrypt) = 0 Then Exit Do
   
   
             
             GlobalStringDef = Split(CommandSeperateLines, ",")
            'GlobalStringDef[0] = $SA3160B01D3E=A0000006239($Os[0x7])
            'GlobalStringDef[0] = $SA1D60C0011E=A0000006239($Os[0x8])
            '....
            
           'for all items in GlobalStringDef do...
            Dim LocalStringName
            For Each LocalStringName In GlobalStringDef
                
                'Sepertate GlobalStringDef further...
                Dim LocalStringName_Splitted
                LocalStringName_Splitted = Split(LocalStringName, "=")
                'LocalStringName_Splitted    with         $SA3160B01D3E=A0000006239($Os[0x7])
                'LocalStringName_Splitted[0] = $SA3160B01D3E
                'LocalStringName_Splitted[1] = A0000006239($Os[0x7])
                
                ReDim Preserve LocalStringNames(1 To UBound(LocalStringNames) + 1)
                
                Dim index
                index = UBound(LocalStringNames)
                'CLng("&h" & strCrop1(LocalStringName_Splitted(1), "$Os[0x", "])"))
                LocalStringNames(index) = Trim(LocalStringName_Splitted(0))
               'LocalStringNames(&h7) = "$SA3160B01D3E"
               
                StringTBL.Add HexStringToString(Trim(strCrop1(LocalStringName_Splitted(1), """", """")))
               
               
               
                If i = 0 Then
                  ' Store for later search&replace in mainfunction code between functions
                    MainIndexes.Add index
   '                 Debug.Print LocalStringNames(index), StringTBL(index), index
                End If
   
                 
              'Replace Local Strings
             ' In ScriptFunction code Replace all "$SA3160B01D3E" with for example "Wrong Name Entered!"
             DeObfu_ReplaceStrings ScriptFunction, CStr(LocalStringNames(index)), StringTBL(index)
   '            tmpstr = Replace(tmpstr, "Execute(" & LocalStringNames(index) & ")", StringTBL(index))
   '            tmpstr = Replace(tmpstr, "Number(" & LocalStringNames(index) & ")", FormatNumber(StringTBL(index)))
   '            tmpstr = Replace(tmpstr, CStr(LocalStringNames(index)), """" & Replace(StringTBL(index), """", """""") & """")
               
               DoEventsSeldom
                 
            Next
            
   
   
         Stat_NumLocal = Stat_NumLocal + UBound(GlobalStringDef)
         Loop While True
         
         If Stat_NumLocal <> 0 Then Log "  Local strings: " & Stat_NumLocal
         Stat_NumAtAll = Stat_NumAtAll + Stat_NumLocal
    
      
         If FrmMain.Chk_verbose.value = vbChecked Then _
         GUI_Show_Text ScriptFunction
      
         GUIEvent_ProcessUpdate i
        
       ' Save New Function
         NewScript2(i) = ScriptFunction
         Inc i



'      End If 'skip crypt functions
      
   Inc iter
   Loop
  
   GUIEvent_ProcessEnd
  
  
  
Log "Strings at all: " & Stat_NumAtAll
Log "strings in Tbl: " & StringTBL.Count
  
' Cut of last Three Functions that where for string load & decrypting
  If StringTBL.Count Then
     ReDim Preserve NewScript2(FunctionsCount - 3)
  End If
  
  
   NewScript = Join(NewScript2, "Func ")
  
  '=== Replace Functionname
   Log DeObfu_FunctionNames.Count & " of " & FunctionsCount & " functions names are obfuscated."
   Log String(79, "_")
   Log "Renaming obfuscated functions..."
   Log "(Enable verbose mode for details)"
   DeObfu_ReplaceFunctionNames NewScript


  '=== Replace Global Strings
  
   GUIEvent_ProcessBegin MainIndexes.Count
  
   Log "Restoring " & MainIndexes.Count & " global strings..."
   For Each index In MainIndexes
'           For index = 1 To StringTBL.Count
'Debug.Assert LocalStringNames(index) <> "$SA5C9AB03457"
      DeObfu_ReplaceStrings NewScript, CStr(LocalStringNames(index)), StringTBL(index)
      
      Dim ReplaceString_counter&
      GUIEvent_ProcessUpdate ReplaceString_counter
      Inc ReplaceString_counter

   Next
   
   GUIEvent_ProcessEnd
  
  'Used String
'  Dim i&
   If StringTBL.Count <> 0 Then Log "UnAssigned/Unused strings [Index, Value]: "
   For i = 1 To StringTBL.Count
      If LocalStringNames(i) = "" Then
         Log "  " & H16(i) & " " & StringTBL(i)
      End If
   Next
   
  
  
  DeObfu_ReplaceDimsAndConst NewScript
  
  
  
End Sub



Private Function FormatNumber(Expr)
   On Error GoTo err_FormatNumber
     'Set default incase of error
      FormatNumber = Expr
    
    ' Exclude decimalNumbers like 1.6 or 234.21
      If Not (Expr Like "*.*") Then
         
        'Beautify Numbers
         If (Expr > 65535) Or (Expr < 0) Then
            FormatNumber = "0x" & H32(Expr)
            
         ElseIf Expr > 10 Then
            FormatNumber = "0x" & H16(Expr)
            
         Else
            '0...10'
            FormatNumber = Expr
            
         End If
         
      Else
         FormatNumber = Expr
      End If
Exit Function

err_FormatNumber:
FormatNumber = Expr
End Function

Private Function IsValidSeperator(TBLFileData As StringReader, TBLStringSeperator) As Boolean
   
'   'Tbl ends with seperator
'   '-> Last element must be empty
'   TBLFileData
'   isTBLStringSeperatorOk = TBLSeperated(UBound(TBLSeperated)) = ""

   
   Dim TBLSeperated
   TBLSeperated = Split(TBLFileData.Data, TBLStringSeperator)
   
 ' Must contain at least 2 HexDigets
   Dim isTBLStringSeperatorOk As Boolean
   isTBLStringSeperatorOk = UBound(TBLSeperated) >= 1
   If isTBLStringSeperatorOk <> False Then
           
     'Remove last element
      ReDim Preserve TBLSeperated(UBound(TBLSeperated) - 1)
      
      Dim item
      For Each item In TBLSeperated
        
       ' Len must be an even number and bigger than 2
         isTBLStringSeperatorOk = isEven(Len(item)) And _
                                   (Len(item) >= 2)
         If isTBLStringSeperatorOk = False Then Exit For
   
      Next
      
      IsValidSeperator = isTBLStringSeperatorOk
   End If

End Function


'Limitation: Will fail on seperators like 'TA1T' -> only returns 'T'
Private Function TBLFindStringSeperator(TBLFileData As StringReader) As String
    With TBLFileData
      
    ' Move to End Of String
      .EOS = True
      .DisableAutoMove = True
      
'      .Move -3
      Dim OccurenceCount&
      OccurenceCount = 0
    ' Blow up TBLStringSeperator2 by 1 byte until the number of its occurence deincreases
      Do
         .Move -1
         Dim TBLStringSeperator$
         TBLStringSeperator = .FixedString(-1)
         
         Dim OccurenceCountNew&
         OccurenceCountNew = UBound(Split(.Data, TBLStringSeperator, , vbBinaryCompare))
      
      If (OccurenceCountNew < OccurenceCount) And _
         IsValidSeperator(TBLFileData, TBLFindStringSeperator) Then
          ' OccurenceCount dropped - TBLStringSeperator is complete
          Exit Do
         
      Else
          ' OccurenceCount has grow(or equal) - set new max
            OccurenceCount = OccurenceCountNew
      End If
       ' store
         TBLFindStringSeperator = TBLStringSeperator
      Loop While True
            
    End With

End Function


