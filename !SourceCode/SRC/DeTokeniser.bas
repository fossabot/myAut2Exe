Attribute VB_Name = "DeTokeniser"
Option Explicit
Public Const DETOKENISE_MAKER$ = "; DeTokenise by "
#Const LineBreak_BeforeAndAfterFunctions = False

Const AUTOIT_SourceCodeLine_MAXLEN& = 4096

Const whiteSpaceTerminal$ = " "
Const ExcludePreWhiteSpaceTerminal$ = "(["
Const ExcludePostWhiteSpaceTerminal$ = ")]."

Const TokenFile_RequiredInputExtensions = ".tok .mem"

Dim bAddWhiteSpace As Boolean


Sub DeToken()
   
   Dim bVerbose As Boolean
   
   bVerbose = FrmMain.Chk_verbose.value = vbGrayed
   With File
    
      Log "DeTokenising: " & FileName.FileName
      
      If InStr(TokenFile_RequiredInputExtensions, FileName.Ext) = 0 Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! Required FileExtension for Tokenfiles: '" & TokenFile_RequiredInputExtensions & "'" & vbCrLf & _
         "Rename this file manually to show that this should be detokenied."
      End If
      
            
'      If Frm_Options.Chk_NoDeTokenise.value = vbChecked Then
'         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! Enable DeTokenise in Options to use it." & FileName.FileName
'
'      End If
      
    ' Since that may be depend on the countrysettings...
      Dim DecimalKomma$
      Const Int64_TestValue As Currency = 1234.1234
    ' ... get it!
      DecimalKomma = Split(Int64_TestValue, "1234")(1)
      
      
      .Create FileName.FileName, False, False, True
      If .Length < 4 Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! File must be at least 4 bytes"
      End If
'   .CloseFile
'   End With
   
'   With New StringReader
'      .Data = FileLoad(FileName.FileName)
   
      
   On Error GoTo DeToken_Err
      .Position = 0
      
      Dim Lines&
      Lines = .int32
      FL "Code Lines: " & Lines & "   0x" & H32(Lines)
      
    ' File shouldn't start with MZ 00 00 -> ExeFile
    ' &HDFEFFF -> Unicodemarker
      If ((Lines And 65535) = &H5A4D) Or (Lines = &HDFEFF) Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "That's no Au3-TokenFile. (MZ-Exe or Dll file)"
      
      ElseIf ((Lines And &H7FFFFFF) > &H3BFEFF) Then
         'It's highly unlikly that there are more that 16 Mio lines in a Sourcefile
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "This seem to be no Au3-TokenFile."
      End If
      
      
            
      
      FrmMain.List_Source.Clear
      FrmMain.List_Source.Visible = True
      
    ' ProgressBarInit
      GUIEvent_ProcessBegin Lines
   
      Dim SourceCodeLine()
      ArrayDelete SourceCodeLine
      
    ' Reset AddWhiteSpace on first item
      Dim bWasLastAnOperator As Boolean
      bWasLastAnOperator = True
      
      
      
      Dim cmd&
      Dim Size&

      Dim SourceCode ' As New Collection
      Dim SourceCodeLineCount&
      ReDim SourceCode(1 To Lines):     SourceCodeLineCount = 1:
      Dim TokenCount&: TokenCount = 0
      
      Dim RawString As StringReader: Set RawString = New StringReader
      Dim DecodeString As StringReader: Set DecodeString = New StringReader

      If bVerbose Then Frm_SrcEdit.Show
      
      Do
   
         Dim TypeName$
         TypeName = ""
   
   
         Dim Atom$
         Atom = ""
         
         If (SourceCodeLineCount > Lines) Then
            Exit Do
         End If
         
         
       ' Default
         bAddWhiteSpace = False
         
         Dim TokenOffset&
         TokenOffset = .Position
         
       ' Read Token
         cmd = .int8
         Inc TokenCount
         
         
         Dim TokenInfo$
         TokenInfo = "Token: " & H8(cmd) & "      (Line: " & SourceCodeLineCount & "  TokenCount: " & TokenCount & ")"
       ' Log it ''" & Chr(Cmd) & "'
         FL_verbose TokenInfo
'         If RangeCheck(SourceCodeLineCount, 3188, 3184) Then
'            Stop
'            If FrmMain.Chk_verbose <> vbChecked Then FrmMain.Chk_verbose = vbChecked
'         Else
'           If FrmMain.Chk_verbose <> vbUnchecked Then FrmMain.Chk_verbose = vbUnchecked
'         End If
'Debug.Assert Not (SourceCodeLine Like "*$NY*")
         
         
         Select Case cmd
         
'------- Numbers -----------
         Case &H0 To &HF
            '&H5
            Dim int32$
            int32 = .int32
            Atom = int32
            
          ' Bugfix for 3.3.8.1 (29th January, 2012)
          ' Tokenoptimisation occure'+-123' -> '-123'
            Dim LastAtom
            If LastAtom = "+" Then
               If Atom <= -1 Then
                  Log " Tokenoptimisation occured '+-' -> '-'  @line: " & SourceCodeLineCount
                  Dim tmp$
                  tmp = ArrayGetLast(SourceCodeLine)
                  tmp = Left2(tmp) ' Cut last char
                  ArraySetLast SourceCodeLine, tmp
               End If
            End If
                         

            
            TypeName = "Int32"
            FL_verbose TypeName & ": 0x" & H32(int32) & "   " & int32
            
          ' So far this value has always been 5
            Debug.Assert cmd = 5
         
         Case &H10 To &H1F
            Dim Int64 As Currency
            Int64 = .int64Value
            'int64 = H32(.int32)
            'int64 = H32(.int32) & int64
            'Replace 123,45 -> 12345
            Atom = Replace(CStr(Int64), DecimalKomma, "")
            TypeName = "Int64"
            FL_verbose TypeName & ": " & Int64
            
            Debug.Assert cmd = &H10
         
         Case &H20 To &H2F
           'Get DoubleValue
            Dim Double_$
            Double_ = .DoubleValue
            'Replace 123,11 -> 123.11
            Atom = Replace(CStr(Double_), DecimalKomma, ".")
            
            TypeName = "64Bit-float"
            FL_verbose TypeName & ": " & Double_
         
            Debug.Assert cmd = &H20
         

'------- Strings -----------
         Case &H30 To &H3F 'Keywords
            
           'Get StrLength and load it
            Size = .int32
            FL_verbose "StringSize: " & H32(Size)
            
            If Size > (.Length - .Position) Then
               Err.Raise vbObjectError, , "Invalid string size(bigger than the file)!"
            End If

            RawString = .FixedStringW(Size)
           
           'XorDecode String
            Dim pos&, XorKey_l As Byte, XorKey_h As Byte
            
            XorKey_l = (Size And &HFF)
            XorKey_h = ((Size \ &H100) And &HFF) ' 2^8 = 256
            
            Dim tmpBuff() As Byte
            tmpBuff = RawString
            
            For pos = LBound(tmpBuff) To UBound(tmpBuff) Step 2
               tmpBuff(pos) = tmpBuff(pos) Xor XorKey_l
               tmpBuff(pos + 1) = tmpBuff(pos + 1) Xor XorKey_h
'               DecodeString = tmpBuff
               
               'If 0 = (pos Mod &H8000) Then myDoEvents
            Next
            
            DecodeString = tmpBuff
            
            
'Comment out due to bad performance
'            RawString.Position = 0
'            DecodeString = Space(RawString.Length \ 2)
'            Do Until RawString.EOS
'               DecodeString.int8 = RawString.int8 Xor Size
'               If Not (RawString.EOS) Then Debug.Assert RawString.int8 = 0
'            Loop
            
            
'------- Commands -----------
            Select Case cmd
            
            Case &H30 'BlockElement (FUNC, IF...) and the Rest of 42 Elements: "AND OR NOT IF THEN ELSE ELSEIF ENDIF WHILE WEND DO UNTIL FOR NEXT TO STEP IN EXITLOOP CONTINUELOOP SELECT CASE ENDSELECT SWITCH ENDSWITCH CONTINUECASE DIM REDIM LOCAL GLOBAL CONST FUNC ENDFUNC RETURN EXIT BYREF WITH ENDWITH TRUE FALSE DEFAULT ENUM NULL"
               TypeName = "BlockElement"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
               
               Atom = DecodeString
               bAddWhiteSpace = True
              
               #If LineBreak_BeforeAndAfterFunctions Then
                  If Atom = "ENDFUNC" Then
                     Atom = Atom & vbCrLf
                  ElseIf Atom = "FUNC" Then
                     Atom = vbCrLf & Atom
                  End If
               #End If

            
            Case &H31 'FunctionCall with params
               Atom = DecodeString
               
               TypeName = "AutoItFunction"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
               
            Case &H32 'Macro
               Atom = "@" & DecodeString
               
               TypeName = "Macro"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H33 'Variable
               Atom = "$" & DecodeString
               
               TypeName = "Variable"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H34 'FunctionCall
               Atom = DecodeString
               
               TypeName = "UserFunction"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H35 'Property
               Atom = "." & DecodeString
               
               TypeName = "Property"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H36 'UserString
               
               Atom = MakeAutoItString(DecodeString.Data)
               
               TypeName = "UserString"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H37 '# PreProcessor
               Atom = DecodeString
               bAddWhiteSpace = True
               
               TypeName = "PreProcessor"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            
            Case Else
               'Unknown StringToken
               If HandleTokenErr("ERROR: Unknown StringToken") Then
               Else
                  Err.Raise vbObjectError Or 1, , "Unknown StringToken"
                  Stop

               End If
               
               
            End Select
            
 '           log String(40, "_")
         
'------- Operators -----------
         Case &H40 To &H56
'            Atom = Choose((Cmd - &H40 + 1), ",", "=", ">", "<", "<>", ">=", "<=", "(", ")", "+", "-", "/", "", "&", "[", "]", "==", "^", "+=", "-=", "/=", "*=", "&=")
         '                     Au3Manual AcciChar
            
            Select Case cmd
               Case &H40: Atom = ","  '        2C
               Case &H41: Atom = "="  ' 1  13  3D
               Case &H42: Atom = ">"  ' 16     3E
               Case &H43: Atom = "<"  ' 18     3C
               Case &H44: Atom = "<>" ' 15     3C
               Case &H45: Atom = ">=" ' 17     3E
               Case &H46: Atom = "<=" ' 19     3C
               Case &H47: Atom = "("  '        28
               Case &H48: Atom = ")"  '        29
               Case &H49: Atom = "+": ' 7      2B
               Case &H4A: Atom = "-": ' 8      2D
               Case &H4B: Atom = "/"  ' 10     2F
               Case &H4C: Atom = "*": ' 9      2A
               Case &H4D: Atom = "&"  ' 11     26
               Case &H4E: Atom = "["  '        5B
               Case &H4F: Atom = "]"  '        5D
               Case &H50: Atom = "==" ' 14     3D
               Case &H51: Atom = "^"  ' 12     5E
               Case &H52: Atom = "+=" '2       2B
               Case &H53: Atom = "-=" '3       2D
               Case &H54: Atom = "/=" '5       2F
               Case &H55: Atom = "*=" '4       2A
               Case &H56: Atom = "&=" '6       26
            End Select
            TypeName = "operator"
            FL_verbose """" & Atom & """   Type: " & TypeName
            
'------- EOL -----------
         Case &H7F
          ' Execute
            
            
            Dim SourceCodeLineFinal$
            SourceCodeLineFinal = Join(SourceCodeLine, whiteSpaceTerminal)
            
            LogSourceCodeLine SourceCodeLineFinal
            
            
            log_verbose ">>>  " & SourceCodeLineFinal
            log_verbose String(80, "_")
            log_verbose ""
 
          ' Test Length
            Dim SourceCodeLine_Len&
            SourceCodeLine_Len = Len(SourceCodeLineFinal)
            If SourceCodeLine_Len >= AUTOIT_SourceCodeLine_MAXLEN Then
               Log "WARNING: SourceCodeLine: " & SourceCodeLineCount & " is " & _
               SourceCodeLine_Len - AUTOIT_SourceCodeLine_MAXLEN & " chars longer than " & _
               AUTOIT_SourceCodeLine_MAXLEN & " - Please remove some spaces manually to make it shorter."
            End If

          ' Processbar update
            GUIEvent_ProcessUpdate SourceCodeLineCount
          
          ' Add SourceCodeLine to SourceCode
            SourceCode(SourceCodeLineCount) = SourceCodeLineFinal
            Inc SourceCodeLineCount
            
          ' del SourceCodeLine
            ArrayDelete SourceCodeLine
            If bVerbose Then Frm_SrcEdit.LineBreak
           
          ' Reset AddWhiteSpace on next item
            bWasLastAnOperator = True
            DelayedReturn False
           

         Case Else
            
           'Unknown Token
            Log "Unknown Token_Command: 0x" & H8(cmd) & " @ " & H32(TokenOffset)
            If HandleTokenErr("ERROR: Unknown Token") Then
            Else
               Err.Raise NO_AUT_DE_TOKEN_FILE, , "Unknown Token"
               'Exit Do
            End If
           'qw
'           Stop
           

         End Select
         
'         Debug.Assert SourceCodeLineCount <> 851

         
         If cmd <> &H7F Then
            
           
          ' Add to SourceLine
            ' Always add a whiteSpace after a command (and preprocessor)
            '    and add a whiteSpace before; except the token before is an operator (Like: [] () = ...)
            If DelayedReturn(bAddWhiteSpace) Or _
               (bAddWhiteSpace And Not (bWasLastAnOperator)) Then
             
             ' Add with whitespace
               ArrayAdd SourceCodeLine, Atom
               If bVerbose Then Frm_SrcEdit.AddItem whiteSpaceTerminal & Atom, cmd, TypeName, TokenInfo & " @ 0x" & H32(TokenOffset)
            Else
              'Append to Last
               
               ArrayAppendLast SourceCodeLine, Atom
               If bVerbose Then Frm_SrcEdit.AddItem Atom, cmd, TypeName, TokenInfo & " @ 0x" & H32(TokenOffset)

            End If
            DoEventsVerySeldom
            
            bWasLastAnOperator = RangeCheck(cmd, &H56, &H40)
'         Else
            
            
         End If
         LastAtom = Atom

      Loop Until .EOS
    
    
    
Err.Clear
DeToken_Err:
Select Case Err
   Case 0
   Case ERR_CANCEL_ALL
      ErrThrowSimple
   
   Case Else
     
     Dim ErrSourceCodeLine$
     ErrSourceCodeLine = Join(SourceCodeLine, whiteSpaceTerminal)
     
     Dim ErrText$
     ErrText = "ERROR: " & Err.Description & vbCrLf & _
      "FileOffset: " & H32(.Position) & vbCrLf & _
      " when de-tokenising script line: " & SourceCodeLineCount & vbCrLf & ErrSourceCodeLine
     Log ErrText
     MsgBox ErrText, vbCritical, "Unexpected Error during detokenising"
     
    'Set incomplete SourceCodeLine
     SourceCode(SourceCodeLineCount) = ErrSourceCodeLine & " <- " & ErrText
     Inc SourceCodeLineCount

    'Cut down SourceCodeArray to Error
     ReDim Preserve SourceCode(SourceCodeLineCount)
     
     Resume DeToken_Finally
End Select

  
  If FrmMain.Chk_TmpFile = vbUnchecked Then
     Log "Keep TmpFile is unchecked => Deleting '" & FileName.NameWithExt & "'"
     FileDelete (FileName)
  End If


DeToken_Finally:
   File.CloseFile
  End With
    
' ProgressBar Finish
  GUIEvent_ProcessEnd
  
  FileName.Ext = ".au3"
  
  
'   If bUnicodeEnable Then
      Dim ScriptData$
      ScriptData = Join(SourceCode, vbCrLf) & vbCrLf & _
                  DETOKENISE_MAKER & FrmMain.Caption & vbCrLf

'      Dim FileName_UTF16 As New ClsFilename
'      FileName_UTF16.FileName = FileName.FileName
'
'      FileName_UTF16.Name = FileName.Name & "_UTF16"
'      FrmMain.Log "Saving UTF16-Script to: " & FileName_UTF16.FileName
'
'      File.Create FileName_UTF16.FileName, True, False, False
'      File.Position = 0
'      File.FixedString(-1) = UTF16_BOM & ScriptData
'      File.setEOF
'      File.CloseFile
'
'   End If
  
  FrmMain.Log "Converting Unicode to UTF8, since Tidy don't support unicode."
  SaveScriptData UTF8_BOM & EncodeUTF8(ScriptData), True
   
  Log "Token expansion succeed."
   
  FrmMain.List_Source.Visible = False

End Sub

Private Function HandleTokenErr(ErrText$) As Boolean

   With File
   
      If vbYes = MsgBox("An Token error occured - possible due to corrupted scriptdata. Contiune?", vbCritical + vbYesNo, ErrText) Then
         HandleTokenErr = True
         
'         Dim Hexdata As New clsStrCat, HexdataLine&
'         Hexdata.Clear
'         For HexdataLine = 0 To &H100 Step &H8
'            Dim Data As New StringReader
'            Data = .FixedString(&H8)
'            Hexdata.Concat H16(HexdataLine) & ":  " & ValuesToHexString(Data) & vbCrLf
'
'         Next
'         .Move -&H100
'         Stop
'         .Move InputBox("The this is the following raw Token data: " & Hexdata.value & "How many bytes should I skip?", "Skip Tokenbytes", "0")
         
      Else
         HandleTokenErr = False
      
      End If
      
   End With
End Function

Private Sub LogSourceCodeLine(TextLine$)
   FrmMain.LogSourceCodeLine TextLine$
End Sub
'Handle UserString with Quotes...
Function MakeAutoItString$(RawString$)
             
   ' HasDoubleQuote ?
     If InStr(RawString, """") <> 0 Then
        
      ' HasSingleQuote ?
        If InStr(RawString, "'") <> 0 Then
         ' Scenario3: " This is a 'Example' on correct "Quoting" String "
           MakeAutoItString = """" & Replace(RawString, """", """""") & """"
        Else
         ' Scenario2: " This is a "Example". "
           MakeAutoItString = "'" & RawString & "'"
        End If
     Else
      ' ' Scenario1: " ExampleString "
        MakeAutoItString = """" & RawString & """"
     End If
     

End Function

' Converts an AutoIt string to a Raw String
' "Test""123""_" -> Test"123"_
Public Function UndoAutoItString$(Au3Str$)
   Dim StringTerminal$
   
  'Get stringchar ( should be " or ')
   StringTerminal$ = Left(Au3Str, 1)
   
  'Cut away Lead&Tailing " or '
  'Is length of Au3Str is smaller than 2 this will give an error
  'since it's no valid Au3String
   Au3Str = Mid(Au3Str, 2, Len(Au3Str) - 2)
   
   
  'Replaces '' -> '  or "" -> "
   UndoAutoItString = Replace(Au3Str, StringTerminal & StringTerminal, StringTerminal)
   
End Function

'   With New RegExp
'      .Global = True
'
'      Const StringTerminal$ = "(['""])"
'      Const StringTerminalBackRef$ = "\1"
'      Const StringBody$ = "(.*?)"
'
'
'      .Pattern = StringTerminal & _
'                   "(?:" & _
'                   StringBody & _
'                     StringTerminalBackRef & StringTerminalBackRef & _
'                        StringBody & _
'                   ")*" & _
'                 StringTerminalBackRef
'      '$2 is the StringBody
'      '$3 is
'      Au3StrToString = .Replace(Au3Str, "$2$3$4")
'   End With
   
'End Function



'
'' Add WhiteSpace Seperator to SourceCodeLine
'Function AddWhiteSpace$()
'
'   'No WhiteSpace at the Beginning
'   If SourceCodeLine = "" Then Exit Function
'
'   Dim LastChar$
'   LastChar = Right(SourceCodeLine, 1)
'
'   Dim NextChar$
'   NextChar = Left(Atom, 1)
'
'   'Don'Append WhiteSpace in cases like this :
'   '"@CMDLIND ["   or   "@CMDLIND [0" <-"].."
'   '         (^-PreCase)                (^-PostCase)
'   If InStr(1, ExcludePreWhiteSpaceTerminal, LastChar) Or _
'      InStr(1, ExcludePostWhiteSpaceTerminal, NextChar) Then
''      Stop
'   ElseIf whiteSpaceTerminal <> LastChar Then
'         AddWhiteSpace = whiteSpaceTerminal
'   End If
'
'End Function





Private Sub FL_verbose(Text)
   FrmMain.FL_verbose Text
End Sub
Private Sub log_verbose(TextLine$)
   FrmMain.log_verbose TextLine$
End Sub

Private Sub FL(Text)
   FrmMain.FL Text
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Private Sub Log(TextLine$)
   FrmMain.Log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Private Sub Log_Clear()
   FrmMain.Log_Clear
End Sub

