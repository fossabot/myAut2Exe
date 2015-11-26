Attribute VB_Name = "UTF8"
Option Compare Binary
Option Explicit


'-------------------------------------------------------------------------
' Konstanten
'-------------------------------------------------------------------------
Private Const CP_ACP = 0
Private Const CP_UTF8 = 65001


'-------------------------------------------------------------------------
' API-Deklarationen
'-------------------------------------------------------------------------
Private Declare Function GetACP Lib "kernel32" () As Long

'Run after Loading
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

'Run before Saving
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As String, ByVal cchWideChar As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, lpUsedDefaultChar As Long) As Long



'-------------------------------------------------------------------------
' DecodeUTF8
'-------------------------------------------------------------------------
Public Function DecodeUTF8(ByVal sValue As String) As String

  If Len(sValue) = 0 Then Exit Function
'  DecodeUTF8 = WToA(StrConv(sValue, vbUnicode), CP_ACP)
  DecodeUTF8 = AToW(StrConv(sValue, vbFromUnicode), CP_UTF8)

End Function


'-------------------------------------------------------------------------
' EncodeUTF8
'-------------------------------------------------------------------------
Public Function EncodeUTF8(ByVal sValue As String) As String


  If Len(sValue) = 0 Then Exit Function
  EncodeUTF8 = WToA(StrConv(sValue, vbUnicode), CP_UTF8)

End Function

'Run before Saving
'-------------------------------------------------------------------------
'   WToA
'   UNICODE to ANSI conversion, via a given codepage
'-------------------------------------------------------------------------
Private Function WToA(ByVal sValue As String, Optional ByVal cpg As Long = -1, Optional ByVal lFlags As Long = 0) As String
Dim cwch              As Long
Dim pwz               As Long
Dim pwzBuffer         As Long
Dim sBuffer           As String

  If cpg = -1 Then cpg = GetACP()
 
'  pwz = StrPtr(sValue)
'  cwch = WideCharToMultiByte(cpg, lFlags, pwz, -1, 0&, 0&, ByVal 0&, ByVal 0&)
  cwch = WideCharToMultiByte(cpg, lFlags, sValue, -1, 0&, 0&, ByVal 0&, ByVal 0&)
  WToA = Space$(cwch)
  
  cwch = WideCharToMultiByte(cpg, lFlags, sValue, -1, WToA, Len(WToA), ByVal 0&, ByVal 0&)
  WToA = Left(WToA, cwch - 1)

End Function

'Run after Loading (cpg=0)
'-------------------------------------------------------------------------
'   AToW
'   ANSI to UNICODE conversion, via a given codepage.
'-------------------------------------------------------------------------
Private Function AToW(ByVal sValue As String, Optional ByVal cpg As Long = -1, Optional ByVal lFlags As Long = 0) As String
Dim cwch      As Long
Dim pwz       As Long
Dim pwzBuffer As Long
Dim sBuffer   As String

  If cpg = -1 Then cpg = GetACP()
  
  pwz = StrPtr(sValue)
  cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, 0&, 0&)
  
  sBuffer = String$(cwch + 1, vbNullChar)
  pwzBuffer = StrPtr(sBuffer)
  
  cwch = MultiByteToWideChar(cpg, lFlags, pwz, -1, pwzBuffer, Len(sBuffer))
  
  AToW = Left$(sBuffer, cwch - 1)

End Function
'Purpose:Returns True if string has a Unicode char.
Public Function IsUnicode(s As String) As Boolean
   Dim i As Long
   Dim bLen As Long
   Dim Map() As Byte

   If LenB(s) Then
      Map = s
      bLen = UBound(Map)
      For i = 1 To bLen Step 2
         If (Map(i) > 0) Then
            IsUnicode = True
            Exit Function
         End If
      Next

      
   End If
End Function
