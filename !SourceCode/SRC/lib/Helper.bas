Attribute VB_Name = "Helper"
Option Explicit
Option Compare Text

Dim myRegExp As New RegExp

Public Const ERR_CANCEL_ALL& = vbObjectError Or &H1000

Public Const ERR_SKIP& = vbObjectError Or &H2000

'used to quit after doevents
Public APP_REQUEST_UNLOAD As Boolean


Public Cancel As Boolean
Public CancelAll As Boolean

Public Skip As Boolean

'Konstantendeklationen für Registry.cls

'Registrierungsdatentypen
Public Const REG_SZ As Long = 1                         ' String
Public Const REG_BINARY As Long = 3                     ' Binär Zeichenfolge
Public Const REG_DWORD As Long = 4                      ' 32-Bit-Zahl

'Vordefinierte RegistrySchlüssel (hRootKey)
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003

Public Const ERROR_NONE = 0

Public Const LocaleID_ENG = 1033 '0x409 US(Eng)
Public Const LocaleID_GER = 1031 '0x407 German
Public LocaleID&


Public Const ERR_FILESTREAM = &H1000000
Public Const ERR_OPENFILE = vbObjectError Or ERR_FILESTREAM + 1
Private i, j As Integer

Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal src As String, ByVal Length&)
Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal src As String, src As Long, ByVal Length&)
Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal src As Integer, ByVal Length&)


'Public Declare Sub MemCopyAnyToAny Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As Any, src As Any, ByVal Length&)
Public Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As String, ByVal src As Any, ByVal Length&)
Public Declare Sub MemCopyX Lib "kernel32" Alias "RtlMoveMemory" _
() '(Dest As Any, ByVal src As Long, ByVal Length&)
'
Public Declare Sub MemCopyAnyToStr Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal Length&)
'Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As String, src As Long, ByVal Length&)
'
'Public Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (Dest As Long, ByVal src As String, ByVal Length&)
''Public Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As String, src As Long, ByVal Length&)
'Public Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (Dest As Long, ByVal src As Integer, ByVal Length&)
    
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const SM_DBCSENABLED = 42
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer


Private BenchtimeA&, BenchtimeB&

'for mt_MT_Init to do a multiplation without 'overflow error'
Private Declare Function iMul Lib "MSVBVM60.DLL" Alias "_allmul" (ByVal dw1 As Long, ByVal dw2 As Long, ByVal dw3 As Long, ByVal dw4 As Long) As Long


'Ensure that 'myObjRegExp.MultiLine = True' else it will use the beginning of the string!
Public Const RE_Anchor_LineBegin$ = "^"
Public Const RE_Anchor_LineEnd$ = "$"

Public Const RE_Anchor_WordBoarder$ = "\b"
Public Const RE_Anchor_NoWordBoarder$ = "\B"

Public Const RE_AnyChar$ = "."
Public Const RE_AnyChars$ = ".*"

Public Const RE_AnyCharNL$ = "[\S\s]"
Public Const RE_AnyCharsNL$ = "[\S\s]*?"

Public Const RE_NewLine$ = "\r?\n"


Dim ExcludedNames As Collection

Function MulInt32&(a&, b&)
  MulInt32 = iMul(a, 0, b, 0)
End Function

Function AddInt32&(a As Double, b As Double)
  AddInt32 = HexToInt(H32(a + b))
End Function


'Returns whether the user has DBCS enabled
Private Function isDBCSEnabled() As Boolean
   isDBCSEnabled = GetSystemMetrics(SM_DBCSENABLED)
End Function


Function LeftButton() As Boolean
    LeftButton = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function

Function RightButton() As Boolean
    RightButton = (GetAsyncKeyState(vbKeyRButton) And &H8000)
End Function

Function MiddleButton() As Boolean
    MiddleButton = (GetAsyncKeyState(vbKeyMButton) And &H8000)
End Function

Function MouseButton() As Integer
    If GetAsyncKeyState(vbKeyLButton) < 0 Then
        MouseButton = 1
    End If
    If GetAsyncKeyState(vbKeyRButton) < 0 Then
        MouseButton = MouseButton Or 2
    End If
    If GetAsyncKeyState(vbKeyMButton) < 0 Then
        MouseButton = MouseButton Or 4
    End If
End Function

Function KeyPressed(Key) As Boolean
   KeyPressed = GetAsyncKeyState(Key)
End Function

Public Function HexToInt&(ByVal HexString$)
   On Error Resume Next
   HexToInt = "&h" & HexString
End Function

' "414243" -> "ABC"
Public Function HexStringToString$(ByVal HexString$, Optional ByRef IsPrintable As Boolean, Optional flag = 1)
 ' flag = 1 (default), binary data is taken to be ANSI
 ' flag = 2, binary data is taken to be UTF16 Little Endian
 ' flag = 3, binary data is taken to be UTF16 Big Endian
 ' flag = 4, binary data is taken to be UTF8
   
   Dim tmpChar&
   IsPrintable = True
   Select Case flag
   Case 2 ' UTF16 Little Endian
   
      HexStringToString = Space(Len(HexString) \ 4)
      For i = 1 To Len(HexString) Step 4
         tmpChar = HexToInt(Mid$(HexString, i, 2))
         If IsPrintable Then
            IsPrintable = RangeCheck(tmpChar, &HFF, &H20)
         End If
         MidB$(HexStringToString, (i \ 2) + 1) = Chr(tmpChar)
      
         tmpChar = HexToInt(Mid$(HexString, i + 2, 2))
         MidB$(HexStringToString, (i \ 2) + 2) = Chr(tmpChar)
      
      Next
   
   Case 3 ' UTF16 Big Endian

      HexStringToString = Space(Len(HexString) \ 4)
      For i = 1 To Len(HexString) Step 4
         tmpChar = HexToInt(Mid$(HexString, i, 2))
         MidB$(HexStringToString, (i \ 2) + 2) = Chr(tmpChar)
      
         tmpChar = HexToInt(Mid$(HexString, i + 2, 2))
         If IsPrintable Then
            IsPrintable = RangeCheck(tmpChar, &HFF, &H20)
         End If
         MidB$(HexStringToString, (i \ 2) + 1) = Chr(tmpChar)
      
      Next

   Case Else
      HexStringToString = Space(Len(HexString) \ 2)
      For i = 1 To Len(HexString) Step 2
         tmpChar = HexToInt(Mid$(HexString, i, 2))
         If IsPrintable Then
            IsPrintable = RangeCheck(tmpChar, &HFF, &H20)
         End If
         
         Mid$(HexStringToString, (i \ 2) + 1) = Chr(tmpChar)
      Next
   End Select

End Function

' "41 42 43" -> "ABC"
Public Function HexvaluesToString$(Hexvalues$)
   Dim tmpChar
   For Each tmpChar In Split(Hexvalues)
      'HexvaluesToString = HexvaluesToString & ChrB("&h" & tmpchar) & ChrB(0)
      'Note ChrB("&h98") & ChrB(0) is not correct translated
      HexvaluesToString = HexvaluesToString & Chr(HexToInt(tmpChar))
   Next
End Function


' "ABC" -> "41 42 43"
Public Function ValuesToHexString$(Data As StringReader, Optional seperator = " ")
'ValuesToHexString = ""
   With Data
      .EOS = False
      Do Until .EOS
         ValuesToHexString = ValuesToHexString & H8(.int8) & seperator
      Loop
   End With
  
End Function


Function Max(ParamArray values())
   Dim item
   For Each item In values
      Max = IIf(Max < item, item, Max)
   Next
End Function

Function Min(ParamArray values())
   Dim item
   Min = &H7FFFFFFF
   For Each item In values
      Min = IIf(Min > item, item, Min)
   Next
End Function

Function limit(value&, Optional ByVal upperLimit = &H7FFFFFFF, Optional lowerLimit = 0) As Long
   'limit = IIf(Value > upperLimit, upperLimit, IIf(Value < lowerLimit, lowerLimit, Value))

   If (value > upperLimit) Then _
      limit = upperLimit _
   Else _
      If (value < lowerLimit) Then _
         limit = lowerLimit _
      Else _
         limit = value
   
End Function

Function isEven(Number As Long) As Boolean
   isEven = ((Number And 1) = 0)
End Function

Function RangeCheck(ByVal value&, Max&, Optional Min& = 0, Optional ErrText, Optional ErrSource$) As Boolean
   RangeCheck = (Min <= value) And (value <= Max)
   If (RangeCheck = False) And (IsMissing(ErrText) = False) Then _
       Err.Raise vbObjectError, ErrSource, _
           ErrText & " Value must between '" & Min & "'  and '" & Max & "' !"
End Function

Public Function H8(ByVal value As Long)
   H8 = Right(String(1, "0") & Hex(value), 2)
End Function

Public Function H16(ByVal value As Long)
   H16 = Right(String(3, "0") & Hex(value), 4)
End Function

Public Function H32(ByVal value As Double)
   If value <= &H7FFFFFFF Then
      H32 = Hex(value)
   Else
    ' split Number in High a Low part...
      Dim High&, Low&
      High = Int(value / &H10000)
      Low = value - (CDbl(High) * &H10000)
      
      H32 = H16(High) & H16(Low)
   End If
   
   H32 = Right(String(7, "0") & H32, 8)
End Function

Public Function Swap(ByRef a, ByRef b)
   Swap = b
   b = a
   a = Swap
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_r  -  Erzeugt einen rechtsbündigen BlockString
'//
'// Beispiel1:     BlockAlign_r("Summe",7) -> "  Summe"
'// Beispiel2:     BlockAlign_r("Summe",4) -> "umme"
Public Function BlockAlign_r(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Right(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_r = Space(Blocksize - Len(RawString)) & RawString
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_l  -  Erzeugt einen linksbündigen BlockString
'//
'// Beispiel1:     BlockAlign_l("Summe",7) -> "Summe  "
'// Beispiel2:     BlockAlign_l("Summe",4) -> "Summ"
Public Function BlockAlign_l(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Left(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_l = RawString & Space(Blocksize - Len(RawString))
End Function

'used to call from the VB6-debug console to be able to scroll textboxes/Listboxes...
Public Function qw()
   Cancel = True
   Do
      DoEvents
   Loop While Cancel = True
End Function
Public Function szNullCut$(zeroString$)
   Dim nullCharPos&
   nullCharPos = InStr(1, zeroString, Chr(0))
   If nullCharPos Then
      szNullCut = Left(zeroString, nullCharPos - 1)
   Else
      szNullCut = zeroString
   End If
   
End Function
Public Sub szNullCutProc(zeroString$)
   Dim nullCharPos&
   nullCharPos = InStr(1, zeroString, Chr(0))
   If nullCharPos Then
      zeroString = Left(zeroString, nullCharPos - 1)
   End If
   
End Sub



Public Function Inc(ByRef value, Optional Increment& = 1)
   value = value + Increment
   Inc = value
End Function

Public Function Dec(ByRef value, Optional DeIncrement& = 1)
   value = value - DeIncrement
   Dec = value
End Function



Public Function CollectionToArray(Collection As Collection) As Variant
   
   Dim tmp
   ReDim tmp(Collection.Count - 1)
   
   Dim i
   i = LBound(tmp)
   
   Dim item
   For Each item In Collection
      tmp(i) = item
      Inc i
   Next
   
   CollectionToArray = tmp
   
End Function
Public Function isString(StringToCheck) As Boolean
   'isString = False
   Dim i&
   For i = 1 To Len(StringToCheck)
      If RangeCheck(Asc(Mid$(StringToCheck, i, 1)), &H7F, &H20) Then
      
      Else
         Exit Function
      End If
   Next
   
   isString = True
   
End Function



'Searches for some string and then starts there to crop
Function strCropWithSeek$(Text$, LeftString$, RightString$, Optional errorvalue, Optional SeektoStrBeforeSearch$)
   strCropWithSeek = strCrop1(Text$, LeftString$, RightString$, errorvalue, _
            InStr(1, Text, SeektoStrBeforeSearch))
End Function


Function strCrop1$(ByVal Text$, LeftString$, RightString$, Optional errorvalue = "", Optional StartSearchAt = 1)
   
   Dim cutend&, cutstart&
      cutstart = InStr(StartSearchAt, Text, LeftString)
   If cutstart Then
      cutstart = cutstart + Len(LeftString)
      cutend = InStr(cutstart, Text, RightString)
      If cutend > cutstart Then
         strCrop1 = Mid$(Text, cutstart, cutend - cutstart)
      Else
        'is Rightstring empty?
         If RightString = "" Then
            strCrop1 = Mid$(Text, cutstart)
         Else
            strCrop1 = errorvalue
         End If
      End If
   Else
      strCrop1 = errorvalue
   End If

End Function

Function strCropAndDelete(Text$, LeftString$, RightString$, Optional errorvalue = "", Optional StartSearchAt = 1, Optional ReplaceString$ = "")
   strCropAndDelete = strCrop1(Text$, LeftString$, RightString$, errorvalue, StartSearchAt)
   Text = Replace(Text, LeftString & strCropAndDelete & RightString, ReplaceString, , , vbTextCompare)
End Function



Function strCrop$(Text$, LeftString$, RightString$, Optional errorvalue = "", Optional StartSearchAt = 1)
   
   Dim cutend&, cutstart&
      cutend = InStr(StartSearchAt, Text, RightString)
   If cutend Then
      cutstart = InStrRev(Text, LeftString, cutend, vbBinaryCompare) + Len(LeftString)
      strCrop = Mid$(Text, cutstart, cutend - cutstart)
   Else
      strCrop = errorvalue
   End If

End Function

Function MidMbcs(ByVal str As String, Start, Length)
    MidMbcs = StrConv(MidB$(StrConv(str, vbFromUnicode), Start, Length), vbUnicode)
End Function


Function strCutOut$(str$, pos&, Length&, Optional TextToInsert = "")
   strCutOut = Mid(str, pos, Length)
   str$ = Mid(str, 1, pos - 1) & TextToInsert & Mid(str, pos + Length)
End Function


Public Function Int16ToUInt32&(value%)
      Const N_0x8000& = 32767
      If value >= 0 Then
         Int16ToUInt32 = value
      Else
         Int16ToUInt32 = CLng(value And N_0x8000) + N_0x8000
      End If
      
End Function




Public Function BenchStart()

   BenchtimeA = GetTickCount

End Function
Public Function BenchEnd()

   BenchtimeB = GetTickCount
   Debug.Print Time & " - " & BenchtimeB - BenchtimeA

End Function


Public Function FileExists(FileName) As Boolean
   On Error GoTo FileExists_err
   FileExists = FileLen(FileName)

FileExists_err:
End Function

Public Function Quote(ByRef Text) As String
   Quote = """" & Text & """"
End Function

Public Function Brackets(ByRef Text As String) As String
   Brackets = "(" & Text & ")"
End Function

Public Function RE_WSpace(ParamArray Elements()) As String
   Dim WS$ ' WhiteSpace
   WS = "\s*"
   
   RE_WSpace = Join(Elements, WS)
End Function



Public Function RE_LookHead_positive(ExpressionThatShouldBeFound$) As String
   RE_LookHead_positive = "(?=" & ExpressionThatShouldBeFound & ")"
End Function

Public Function RE_LookHead_negative(ExpressionThatShouldNOTBeFound$) As String
   RE_LookHead_negative = "(?!" & ExpressionThatShouldNOTBeFound & ")"
End Function

Public Function RE_Repeat(Optional MinRepeat& = 0, Optional MaxRepeat = "") As String
   If (MinRepeat = MaxRepeat) Then
      RE_Repeat = "{" & MinRepeat & "}"
   Else
      RE_Repeat = "{" & MinRepeat & "," & MaxRepeat & "}"
   End If
   
End Function


Public Function RE_AnyCharRepeat(Optional MinRepeat& = 0, Optional MaxRepeat = "") As String
   RE_AnyCharRepeat = "." & RE_Repeat(MinRepeat, MaxRepeat)
End Function

Public Function RE_Group(RegExpForTheGroup$) As String
   RE_Group = "(" & RegExpForTheGroup & ")"
End Function

Public Function RE_Group_NonCaptured(RegExpForTheNonCapturedGroup$) As String
   RE_Group_NonCaptured = "(?:" & RegExpForTheNonCapturedGroup & ")"
End Function

Public Function RE_Literal(TextWithLiterals) As String
   'Mask metachars
   RE_Literal = RE_Mask(TextWithLiterals, "][{}()*+?.\\^$|")
                                           
End Function


Public Function RE_Replace_Literal(TextWithLiterals) As String
  'Mask Replace metachars
   ' $0-9   Back reference
   ' $+     Last reference
   
   ' $&     MatchText
   
   ' $`     Text left from subject
   ' $'     Text right from subject
   ' $_     Whole subject
   
   RE_Replace_Literal = RE_Mask(TextWithLiterals, "0-9+`'_", "\$", "$$")


End Function
Private Sub RE_Mask_Whitespace(Text)
   ReplaceDo Text, vbCr, "\r"
   ReplaceDo Text, vbLf, "\n"
   ReplaceDo Text, vbTab, "\t"
End Sub

Private Function RE_Mask(Text, CharsToMask$, _
   Optional CharMaskSearch$ = "", _
   Optional CharMaskReplace$ = "\") As String
   With myRegExp
      .Global = True
      
     ' Mask MetaChars like with a preciding '\'
      .Pattern = CharMaskSearch & "[" & CharsToMask & "]"
      
     'Attention Text is passed byref - so don use Text =...!
      RE_Mask = .Replace(Text, CharMaskReplace & "$&")
   
   
   End With

'   RE_Mask_Whitespace Text
   
'   RE_Mask = Text

End Function

Public Function RE_CharCls(Chars$) As String
   ' mask ']' and '-'
   RE_CharCls = "[" & RE_Mask(Chars, "]\\-") & "]"
End Function

Public Function RE_CharCls_Excluded(Chars$) As String
   ' mask ']' and '-'
   RE_CharCls_Excluded = "[^" & RE_Mask(Chars, "]\\-") & "]"

End Function

Public Function IsAlreadyInCollection(CollectionToTest As Collection, Key$) As Boolean
   Dim Description$, Number&, Source$
   Description = Err.Description
   Number = Err.Number
   Source = Err.Source
   
      On Error Resume Next
      CollectionToTest.item Key
      IsAlreadyInCollection = (Err = 0)
      
   Err.Description = Description
   Err.Number = Number
   Err.Source = Source


End Function

'Public Sub ArrayEnsureBounds(Arr)
'
''   Dim tmp_ptr&
''   MemCopy tmp_ptr, VarPtr(Arr) + 8, 4 ' resolve Variant
''   MemCopy tmp_ptr, tmp_ptr, 4               ' get arraypointer
''
''   Dim bIsNullArray As Boolean
''   bIsNullArray = (tmp_ptr = 0)
'' On Error Resume Next
'
'   Dim bIsNullArray As Boolean
'   bIsNullArray = (Not Not Arr) = 0 'use vbBug to get pointer to Arr
'
''   Rnd 1 ' catch Expression too complex error that is cause by the bug
''On Error GoTo 0
'
''   Exit Function
'
'   If bIsNullArray Then
'
'   ElseIf (UBound(Arr) - LBound(Arr)) < 0 Then
'   Else
'      Exit Function
'   End If
'
'   ReDim Arr(0)
'   ArrayEnsureBounds = True
'   Exit Function

Public Sub ArrayEnsureBounds(Arr)

On Error GoTo Array_err
  ' IsArray(Arr)=False        ->  13 - Type Mismatch
  ' [Arr has no Elements]     ->  9 - Subscript out of range
  ' ZombieArray[arr=Array()]  -> GoTo Array_new
   If UBound(Arr) - LBound(Arr) < 0 Then GoTo Array_new
Exit Sub
Array_err:
Select Case Err
    Case 9, 13
Array_new:
      ArrayDelete Arr

'   Case Else
'      Err.Raise Err.Number, "", "Error in ArrayEnsureBounds: " & Err.Description

End Select

End Sub



Public Sub ArrayAdd(Arr, Optional Element = "")
   ArrayEnsureBounds Arr
   ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
   Arr(UBound(Arr)) = Element

End Sub


'Public Sub ArrayAdd(Arr As Variant, Optional element = "")
'' Is that already a Array?
'   If IsArray(Arr) Then
'      ReDim Preserve Arr(LBound(Arr) To UBound(Arr) + 1)
'
' ' VarType(Arr) = vbVariant must be
'   Else 'If VarType(Arr) = vbVariant Then
'      ReDim Arr(0)
'   End If
'
'   Arr(UBound(Arr)) = element
'
'End Sub

Public Sub ArrayRemoveLast(Arr)
   ReDim Preserve Arr(LBound(Arr) To UBound(Arr) - 1)
End Sub

Public Sub ArrayDelete(Arr)
   ReDim Arr(0)
   'Arr = Array()
   'Set Arr = Nothing
End Sub


Public Function ArrayGetLast(Arr)
ArrayEnsureBounds Arr
   ArrayGetLast = Arr(UBound(Arr))
End Function
Public Sub ArraySetLast(Arr, Element)
ArrayEnsureBounds Arr
    Arr(UBound(Arr)) = Element
End Sub
Public Sub ArrayAppendLast(Arr(), Element)
ArrayEnsureBounds Arr
    Arr(UBound(Arr)) = Arr(UBound(Arr)) & Element
End Sub


Public Function ArrayGetFirst(Arr)
ArrayEnsureBounds Arr
   ArrayGetFirst = Arr(LBound(Arr))
End Function
Public Sub ArraySetFirst(Arr, Element)
ArrayEnsureBounds Arr
    Arr(LBound(Arr)) = Element
End Sub
Public Sub ArrayAppendFirst(Arr, Element)
ArrayEnsureBounds Arr
    Arr(LBound(Arr)) = Arr(LBound(Arr)) & Element
End Sub




Function DelayedReturn(Now As Boolean) As Boolean
   Static LastState As Boolean
   
   DelayedReturn = LastState
   
   LastState = Now
   
End Function







'Private Sub QuickSort( _
'                      ByRef ArrayToSort As Variant, _
'                      ByVal Low As Long, _
'                      ByVal High As Long)
'Dim vPartition As Variant, vTemp As Variant
'Dim i As Long, j As Long
'  If Low > High Then Exit Sub  ' Rekursions-Abbruchbedingung
'  ' Ermittlung des Mittenelements zur Aufteilung in zwei Teilfelder:
'  vPartition = ArrayToSort((Low + High) \ 2)
'  ' Indizes i und j initial auf die äußeren Grenzen des Feldes setzen:
'  i = Low: j = High
'  Do
'    ' Von links nach rechts das linke Teilfeld durchsuchen:
'    Do While ArrayToSort(i) < vPartition
'      i = i + 1
'    Loop
'    ' Von rechts nach links das rechte Teilfeld durchsuchen:
'    Do While ArrayToSort(j) > vPartition
'      j = j - 1
'    Loop
'    If i <= j Then
'      ' Die beiden gefundenen, falsch einsortierten Elemente
'austauschen:
'      vTemp = ArrayToSort(j)
'      ArrayToSort(j) = ArrayToSort(i)
'      ArrayToSort(i) = vTemp
'      i = i + 1
'      j = j - 1
'    End If
'  Loop Until i > j  ' Überschneidung der Indizes
'  ' Rekursive Sortierung der ausgewählten Teilfelder. Um die
'  ' Rekursionstiefe zu optimieren, wird (sofern die Teilfelder
'  ' nicht identisch groß sind) zuerst das kleinere
'  ' Teilfeld rekursiv sortiert.
'  If (j - Low) < (High - i) Then
'    QuickSort ArrayToSort, Low, j
'    QuickSort ArrayToSort, i, High
'  Elsea
'    QuickSort ArrayToSort, i, High
'    QuickSort ArrayToSort, Low, j
'  End If
'End Sub
'
'

Public Sub myDoEvents()
   DoEvents
   
   Skip_Test
   CancelAll_Test
   APP_REQUEST_UNLOAD_Test
End Sub

Public Sub Skip_Test()
   If Skip = True Then
      
      Skip = False
      Err.Raise ERR_SKIP, , "User pressed the skip key."
      
   End If
  
End Sub


Public Sub CancelAll_Test()
   If CancelAll = True Then
      
      CancelAll = False
      Err.Raise ERR_CANCEL_ALL, , "User pressed the cancel key."
      
   End If
  
End Sub

Public Sub APP_REQUEST_UNLOAD_Test()
   If APP_REQUEST_UNLOAD = True Then
      
      Err.Raise ERR_CANCEL_ALL, , "Application shutdown."
      
   End If
  
End Sub



Public Function FileLoad$(FileName$)
   Dim File As New FileStream
   With File
      .Create FileName, False, False, True
      FileLoad = .FixedString(-1)
      .CloseFile
   End With
End Function

Public Sub FileSave(FileName$, Data$)
   On Error GoTo err_FileSave
   Dim File As New FileStream
   With File
      .Create FileName, True, False, False
      .FixedString(-1) = Data
      .CloseFile
   End With

Exit Sub
err_FileSave:
   Log "ERROR during FileSave: " & Err.Description
End Sub


Public Function FormatSize$(ByVal SizeValue&)
   On Error GoTo FormatSize_err
   If SizeValue < 0 Then
      FormatSize = "#Error Negative Value: " & SizeValue & "#"
   
   ElseIf SizeValue > &H100000 Then
      Dim SizePostFix$
      Dim tmpSizeValue& 'As Double
      tmpSizeValue = SizeValue \ &H100000 ' clng(&H400) * &H400)
      SizePostFix = "M"
    
   ElseIf SizeValue > &H400 Then
      tmpSizeValue = SizeValue \ &H400
      SizePostFix = "K"
   
   Else
      SizePostFix = ""
   End If

   FormatSize = Format(tmpSizeValue, "##,##0")
 '  If Right(FormatSize, 1) = "," Then
 '     FormatSize = Left(FormatSize, Len(FormatSize) - 1)
 '  End If
   
   FormatSize = FormatSize & " " & SizePostFix & "B"
FormatSize_err:
Select Case Err
   Case 0
   Case Else
      FormatSize = "#Error [" & Err.Description & "]"
End Select


End Function


'///////////////////////////////////////////
'// General Load/Save Configuration Setting
Function ConfigValue_Load(Section$, Key$, Optional DefaultValue)
   ConfigValue_Load = GetSetting(App.Title, Section, Key, DefaultValue)
End Function
Property Let ConfigValue_Save(Section$, Key$, value As Variant)
      SaveSetting App.Title, Section, Key, value
End Property

'///////////////////////////////////////////
'// Load/Save a Form Setting
  'Iterate through all Item on the OptionsFrame
  'incase it's no Checkbox a 'type mismatch error' will occur
  'and due to "On Error Resume Next" it skip the call
Sub FormSettings_Load(Form As Form, Optional ExcludedNames$)
   On Error Resume Next
   ExcludedNamesSet ExcludedNames
   
   Dim controlItem
   For Each controlItem In Form.Controls
      If IsExcludedName(controlItem.Name) = False Then
         Select Case TypeName(controlItem)
         Case "TextBox"
   '         If (controlItem Is Combo_Filename) = False Then
               TextBox_Load Form.Name, controlItem
   '         End If
   
         Case "CheckBox"
            CheckBox_Load Form.Name, controlItem
         
         
         Case "ComboBox"
            ComboBox_Load Form.Name, controlItem
         
         
         End Select
'      Else
'         Debug.Print controlItem.Name
      End If
   Next
   
End Sub
Sub FormSettings_Save(Form As Form, Optional ExcludedNames$)

   On Error Resume Next
   ExcludedNamesSet ExcludedNames
   
   Dim controlItem
   For Each controlItem In Form.Controls
      If IsExcludedName(controlItem.Name) = False Then
         CheckBox_Save Form.Name, controlItem
         TextBox_Save Form.Name, controlItem
         ComboBox_Save Form.Name, controlItem
'      Else
'         Debug.Print "ExcludedName: " & controlItem.Name
      End If
   Next

End Sub


Sub ExcludedNamesSet(ExcludedNamesStr$)
   Set ExcludedNames = New Collection
   Dim item
   For Each item In Split(ExcludedNamesStr)
      ExcludedNames.Add item, item
   Next
End Sub


Function IsExcludedName(controlName) As Boolean
   On Error Resume Next
   ExcludedNames.item controlName
   IsExcludedName = (Err = 0)
End Function



'///////////////////////////////////////////
'// Load/Save a CheckBox State
Sub CheckBox_Load(Section$, ByVal ChkBox As CheckBox)
   ChkBox.value = ConfigValue_Load(Section, ChkBox.Name, ChkBox.value)
End Sub
Sub CheckBox_Save(Section$, ByVal ChkBox As CheckBox)
   ConfigValue_Save(Section, ChkBox.Name) = ChkBox.value
End Sub

'///////////////////////////////////////////
'// Load/Save comboBox States
Sub ComboBox_Load(Section$, ByVal cbBox As ComboBox)
   With cbBox
      Dim i
      For i = 0 To ConfigValue_Load(Section, cbBox.Name & "_ListCount", 0) - 1
         .AddItem ConfigValue_Load(Section, cbBox.Name & "_" & i, "")
      Next
   End With
 End Sub

Sub ComboBox_Save(Section$, ByVal cbBox As ComboBox)
   With cbBox
      Dim i
      For i = 0 To .ListCount - 1
         ConfigValue_Save(Section, cbBox.Name & "_" & i) = .List(i)
      Next
      
      If .ListCount > 0 Then
         ConfigValue_Save(Section, cbBox.Name & "_ListCount") = .ListCount
      End If
   End With
End Sub



Sub TextBox_Load(Section$, ByVal Txt As TextBox)
   With Txt
      'signal [txt]_change that were and load the settings
      'so it might react on this i.e. like not the execute the event handler code
      .Enabled = False
         .Text = ConfigValue_Load(Section, Txt.Name, Txt.Text)
      .Enabled = True
   End With
 End Sub
Sub TextBox_Save(Section$, ByVal Txt As TextBox)
  'don't save Multiline Textbox
   If Txt.MultiLine = False Then
      ConfigValue_Save(Section, Txt.Name) = Txt.Text
   End If
End Sub


Sub Checkbox_TriStateToggle(CheckBox As CheckBox, value)
   Static Block_Click As Boolean
   If Block_Click = False Then
      Block_Click = True
      
      With CheckBox

         If value = vbGrayed Then
            value = vbUnchecked
         Else
            value = value + 1
         End If
         .value = value
         
      End With
      
      Block_Click = False
   End If
End Sub


Public Function MakePrintable$(str$)
   
   MakePrintable = str
   Dim i
   For i = 1 To Len(str)
      
      Dim char$
      char = Mid(str, i, 1)
      Select Case char
         Case vbNullChar To " "
            char = "."
      End Select
      
      
      Mid(MakePrintable, i, 1) = char
   Next

End Function

Function Left2$(str$, Optional Length_SeenFromEnd& = 1)
    Left2 = Left(str$, Len(str$) - Length_SeenFromEnd)
End Function



Public Function RE_FindPattern$(Data$, Pattern$, Optional Match As Match)
       
   With New RegExp
      .IgnoreCase = True
      .Global = False
      .MultiLine = False
      .Pattern = Pattern
      
      Dim matches As MatchCollection
      
      Set matches = .Execute(Data)
      If matches.Count = 1 Then
         'Dim match As match
         Set Match = matches(0)
         If Match.SubMatches.Count = 1 Then
            RE_FindPattern = matches.item(0).SubMatches(0)
         End If
      End If
   End With
End Function




Public Function RE_FindPatterns(Data, Pattern$)
       
   With New RegExp
      .IgnoreCase = True
      .Global = True
      .MultiLine = False
      .Pattern = Pattern
      
      Set RE_FindPatterns = .Execute(Data)
   End With
End Function

