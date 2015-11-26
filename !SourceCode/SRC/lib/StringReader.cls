VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Important ref to clean up the Chr, ChrB, ChrW Hassle
'http://support.microsoft.com/kb/145745/



Option Explicit
Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, src As Any, ByVal Length&)
Private Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (Dest As Long, ByVal src As String, ByVal Length&)
Private Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Dest As String, src As Long, ByVal Length&)
Private Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (Dest As Long, ByVal src As Integer, ByVal Length&)

Public bSearchBackward As Boolean
'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarPosition As Long 'lokale Kopie
Public mvardata As String  'lokale Kopie
'Public Position As Long 'lokale Kopie
Public mStorePos As Long
Public DisableAutoMove As Boolean

Public Sub Truncate()
   Data = Left(Data, Position)
End Sub


Public Property Get Stream(Length As Long) As StringReader
    Set Stream = New StringReader
    Stream.Data = FixedString(Length)
End Property





Public Property Let zeroString(ByVal vData As String)
    Stop
End Property


Public Property Get zeroString() As String
   Dim EndOfZeroString As Long
   EndOfZeroString = InStr(mvarPosition, mvardata, vbNullChar)
   zeroString = FixedString(EndOfZeroString - mvarPosition)
 ' Skip over zerobyte at the end of the ZeroTerminatedString
   mvarPosition = mvarPosition + 1
   
End Property






Public Property Let Length(ByVal vData As Long)
    Stop 'not implemented
End Property


Public Property Get Length() As Long
    Length = Len(mvardata)
End Property




Public Sub RestorePos()
   Position = mStorePos
End Sub

Public Sub StorePos()
   mStorePos = Position
End Sub


Public Property Let ToMove(ByVal vData As Long)
   If DisableAutoMove Then Exit Property
   Move (vData)
End Property


Public Property Let Position(vData As Long)
    mvarPosition = limit(vData, Len(mvardata), 0) + 1
End Property


Public Property Get Position() As Long
    Position = mvarPosition - 1
End Property


Private Function StrtoLng(ByVal value$) As Long
   MemCopyStrToLng StrtoLng, value, 4
End Function
'
Private Function LngtoStr(ByVal value&) As String
   Dim tmp$
   tmp = Space(4)
   MemCopyLngToStr tmp, value, 4
   LngtoStr = tmp
End Function

'/////////////////////////////////////////////////////////
'// set_EOS - Returns True if Position is at the End Of String
Public Property Let EOS(ByVal vData As Boolean)
    Position = Len(mvardata) And vData
End Property
'// get_EOS - Forward to End Of String
Public Property Get EOS() As Boolean
 '  EOS = Position >= Len(mvardata)
 'optimised
 '  EOS = (mvarPosition - 1) >= Len(mvardata)
    EOS = mvarPosition > Len(mvardata)
End Property


Public Sub Move(Chars&)
   Position = Position + Chars
'   Debug.Print "Move: ", Chars
End Sub

Public Function FindByte(ByVal int32_To_Find$, Optional Range) As Long
   FindByte = FindString(Left(LngtoStr(int32_To_Find), 1))
End Function

Public Function FindInt(ByVal int32_To_Find$, Optional Range) As Long
   FindInt = FindString(Left(LngtoStr(int32_To_Find), 2))
End Function

Public Function FindLong(ByVal int32_To_Find$, Optional Range) As Long
   FindLong = FindString(LngtoStr(int32_To_Find))
End Function

Public Function FindString(String_To_Find$, Optional Range)  ', Optional Alternativ_String_To_Find) As Long
   If IsMissing(Range) Then Range = Len(mvardata)
   
'   Findstring = InStr(1, mid$(mvarData, mvarPosition, Range), String_To_Find)
    If bSearchBackward Then
      FindString = InStrRev(mvardata, String_To_Find, mvarPosition)
      
    Else
      FindString = InStr(mvarPosition, mvardata, String_To_Find)
   
    End If
      
   'Test if out of range
    If Abs(mvarPosition - FindString - 1) > Range Then
       FindString = 0
    End If
  
   
'   If IsMissing(Alternativ_String_To_Find = False) And (Findstring = 0) Then
'   Findstring = InStr(1, mid$(mvarData, Position, Range), String_To_Find)
   
 
 ' If string was found
   If FindString Then
     
     'Return start of String
'      Findstring = (Findstring - 1) + Position
      FindString = (FindString - 1)
 
     
     'seek at the end of found String
      ToMove = (FindString - Position) + Len(String_To_Find)

     
   End If
End Function

Public Property Let FixedString(Optional ByVal Length& = -1, vData As String)
    If Length <= 0 Then Length = Len(vData)
    If Length <= 0 Then Exit Property
    
   'Enlarge Buffer if necessary
    Dim enlarge&
    enlarge = (Length + mvarPosition - 1) - Len(mvardata)
    If enlarge >= 1 Then mvardata = mvardata & Space(enlarge)
    
    Mid$(mvardata, mvarPosition, Length) = vData
    ToMove = Length
End Property


Public Property Get FixedString(Optional ByVal Length& = -1) As String
    If Length <= -1 Then Length = Len(mvardata)
    FixedString = Mid$(mvardata, mvarPosition, Length)
   
    ToMove = Length
End Property

Public Property Get FixedStringW(Optional ByVal Length& = -1) As String
    Length = Length * 2
    If Length <= -1 Then Length = Len(mvardata)
    FixedStringW = Mid(mvardata, mvarPosition, Length)
   
    ToMove = Length
End Property



Public Property Let int32(vData As Long)
    FixedString = LngtoStr(vData)
End Property


Public Property Get int32() As Long
    int32 = StrtoLng(FixedString(4))
End Property



Public Property Let int16(vData As Long)
    FixedString(2) = LngtoStr(vData)
End Property


Public Property Get int16() As Long
    int16 = StrtoLng(FixedString(2))
End Property

'Public Property Let int16Sig(vData As Integer)
'    FixedString(2) = LngtoStr(vData)
'End Property


Public Property Get int16Sig() As Integer
    Dim value&
    value = StrtoLng(FixedString(2))
    int16Sig = value Or -(value And &H8000)  '-32768 '&H8000




End Property


Public Property Let int8(vData As Long)
    FixedString(1) = LngtoStr(vData)
'    FixedString(1) = ChrB(vData) & ChrB(0)
'    FixedString(1) = ChrW(vData) ' & ChrB(0)
End Property



Public Property Get int8() As Long
'    int8 = StrtoLng(FixedString(1))
'    int8 = Asc(FixedString(1))
'    Debug.Assert AscB(Mid(mvardata, mvarPosition, 1)) = Asc(Mid(mvardata, mvarPosition, 1))
    int8 = Asc(Mid$(mvardata, mvarPosition, 1))
'    int8 = AscB(MidB(StrConv(mvardata, vbFromUnicode), mvarPosition, 1))

    ToMove = 1



End Property

'Public Property Let int8Sig(vData As Long)
'    FixedString(1) = LngtoStr(vData)
'End Property


Public Property Get int8Sig() As Long

   int8Sig = StrtoLng(FixedString(1))
   int8Sig = int8Sig Or -(int8Sig And &H80)

End Property


Public Property Let Data(vData As String)
    mvardata = vData
    Position = 0
End Property


Public Property Get Data() As String
Attribute Data.VB_UserMemId = 0
    Data = mvardata
End Property

Public Sub CopyData(src&, Dest&, Optional size = -1)
   Dim tmpstr$
   Position = src
   tmpstr = FixedString(size)
   
   Position = Dest
   FixedString(size) = tmpstr
   
   
End Sub

Private Sub Class_Initialize()
'   mvarPosition = 1
End Sub
