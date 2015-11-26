Attribute VB_Name = "Parser"
'_________________________________________________________________
' NOTE: Module Not need anymore since Tidy v2.0.24.4 November 30, 2008 comes along with long scriptLines

Option Explicit
   Dim Str As StringReader
   Dim Level&
   Dim StartSym$, EndSym$
Public Function CropParenthesis$(Expression$, Optional StartSym = "(", Optional EndSym = ")")
   Set Str = New StringReader
   Str = Expression
   
   Parser.StartSym = StartSym
   Parser.EndSym = EndSym
   
   Level = 0
   CropParenthesis = BB()
   
   If Level <> 0 Then Err.Raise vbObjectError, , "Parenthesis unembalanced: (" & Level & " too much) in expression: " & Str & vbCrLf & "ignored: " & Str.FixedString(-1)
   
End Function


Function BB$()
   Dim Text$, char$
   Text = ""
   
   Do While Str.EOS = False
      char = Str.FixedString(1)
      Select Case char
         Case StartSym '"("
            Inc Level
            BB
         
         '  dirty fix:
            Text = Text & char & EndSym

         Case EndSym ' ")"
            Dec Level
            Exit Do

         Case Else
            Text = Text & char
            
      End Select
      
   Loop
   
   BB = Text
'   Debug.Print Space(Level * 2) & text
   
End Function


Function AddLineBreakToLongLines$(ByRef Lines)
' Adding a underscope '_' for lines longer than 2047
' so Tidy will not complain

'  Dim Lines
  Dim Line, NewLine As New clsStrCat
'   Lines = Split(TextLine, vbCrLf)

   'Find place to break line
   '...total "&@CRLF& "fees....
   '                ^-NextAmpersandPos
   'Will be changed to
   '...total "&@CRLF&_
   '"fees....
   '

'   Const MAX_CODE_LINE_LENGHT& = 2000
'   Const MAX_CODE_LINE_LENGHT& = 1897
   Const MAX_CODE_LINE_LENGHT& = 1800


   For Line = 0 To UBound(Lines)

      Dim lineLen&
      lineLen = Len(Lines(Line))


      If lineLen > MAX_CODE_LINE_LENGHT Then

         NewLine.Clear

         Dim linePos&, LastPos&
         linePos = 1
         LastPos = 1

         Do While linePos + MAX_CODE_LINE_LENGHT < lineLen

            Dim CrackAtPos&
            CrackAtPos = InStrRev(Mid(Lines(Line), linePos, MAX_CODE_LINE_LENGHT), "&")
            If (CrackAtPos <> 0) Then

               NewLine.Concat Mid(Lines(Line), linePos, CrackAtPos)
               NewLine.Concat " _" & vbCrLf
             ' Test for special cases
            ElseIf Mid(Lines(Line), linePos, 7) = "GLOBAL " Then
               CrackAtPos = InStrRev(Mid(Lines(Line), linePos, MAX_CODE_LINE_LENGHT), ",")
               If (CrackAtPos <> 0) Then
                  NewLine.Concat Mid(Lines(Line), linePos, CrackAtPos - 1)
                  NewLine.Concat vbCrLf & "GLOBAL "
               Else

                  GoTo notice_user
               End If

            'IF with AND
            ElseIf Mid(Lines(Line), linePos, 3) = "IF " Then
               CrackAtPos = InStrRev(Mid(Lines(Line), linePos, MAX_CODE_LINE_LENGHT), " AND ")
               If (CrackAtPos <> 0) Then
                  NewLine.Concat Mid(Lines(Line), linePos, CrackAtPos)
                  NewLine.Concat " _" & vbCrLf
               Else

                  GoTo notice_user
               End If
            Else
notice_user:
              'notice in the log - user should manually fix this
               CrackAtPos = MAX_CODE_LINE_LENGHT
               NewLine.Concat Mid(Lines(Line), linePos, CrackAtPos)
               Log " PROBLEM: Line " & Line & " is longer than " & MAX_CODE_LINE_LENGHT & " Bytes. Tidy will refuse to work. Fix this manually an then apply Tidy."

            End If

            Inc linePos, CrackAtPos

         Loop

        'add last end
         NewLine.Concat Mid(Lines(Line), linePos)


         Lines(Line) = NewLine.value

      End If
   Next

  AddLineBreakToLongLines = Join(Lines, vbCrLf)

End Function
