Attribute VB_Name = "CustomStrCrypt"
Option Explicit
 

'Public Declare Sub CryptCall Lib "hwindr.dll" Alias "_deCode" _
   (ByVal Key1 As String, _
    ByVal Key2 As String) _

Private Declare Function myCryptCall Lib "hwindr.dll" Alias "_deCode" _
   (ByVal Key1 As String, _
    ByVal Key2 As String) _
   As Long
 
'Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As Long
 
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
 
 
''Using the GetCommandLine API function.
'Function GetCommand() As String
'  GetCommand = apiGetCommandLine()
'End Function 'GetCommand
 
Public Function CryptCall(Key1, Key2) As String
    Dim RetStr As Long
    Dim SLen As Long
    Dim Buffer As String
    'Get a pointer to a string, which contains the command line
    RetStr = myCryptCall(Key1, Key2)
    'Get the length of that string
    SLen = lstrlen(RetStr)
    If SLen > 0 Then
        'Create a buffer
        CryptCall = Space$(SLen)
        'Copy to the buffer
        CopyMemory ByVal CryptCall, ByVal RetStr, SLen
    End If
End Function


