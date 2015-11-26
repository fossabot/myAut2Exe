Attribute VB_Name = "ErrorHelper"
Option Explicit

Public ErrThrow_LastDllError&
Private Number&, Source$, Description$, HelpFile$, HelpContext&

Private Declare Function FormatMessage Lib "kernel32" _
  Alias "FormatMessageA" ( _
  ByVal dwFlags As Long, _
  lpSource As Any, _
  ByVal dwMessageId As Long, _
  ByVal dwLanguageId As Long, _
  ByVal lpBuffer As String, _
  ByVal nSize As Long, _
  Arguments As Long) As Long
 
' FormatMessage Konstanten
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H900
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
 
' FormatMessage Rückgabe-Konstante
Private Const ERROR_RESOURCE_LANG_NOT_FOUND = 1815&
 
' Einige FormatMessage Sprachkonstanten
Private Const LANG_NEUTRAL = &H0
Private Const LANG_GERMAN = &H7
Private Const LANG_FRENCH = &HC
Private Const LANG_ENGLISH = &H9
 
' Einige FormatMessage Sub-Sprach-Konstanten
Private Const SUBLANG_DEFAULT = &H1
Private Const SUBLANG_FRENCH = &H1
Private Const SUBLANG_ENGLISH_US = &H1
Private Const SUBLANG_GERMAN = &H1
 
' Eine der Get-/Set- LastError Konstanten
Private Const ERROR_ACCESS_DENIED = 5& ' Zugriff verweigert



Private FormatMessageBuff As String * 256

Public Sub RaiseDllError(Location$, FuncName$, ParamArray FuncParams())
  Dim Flags As Long, LastErr$
  Dim Retval As Long, LanguageID As Long
 
  ' Den Fehlertext mit FormatMessage in der Standardsprache ausgeben
  Flags = FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS
  LanguageID = 0 'LANG_NEUTRAL Or (SUBLANG_DEFAULT * 1024)
  LastErr = Err.LastDllError
  Retval = FormatMessage(Flags, 0&, Err.LastDllError, LanguageID, FormatMessageBuff, Len(FormatMessageBuff), 0&)
  
  If RangeCheck(Retval, 256, 1) Then
    
    Dim ErrMsg As String
    ErrMsg = Left$(FormatMessageBuff, Retval)

    Dim FunctionCall$
    FunctionCall = FuncName & Brackets(Join(FuncParams, ", "))
    
    Err.Raise vbObjectError, , _
      FunctionCall & " @ " & Location & " failed!  GetLastError: 0x" & H32(LastErr) & " - " & ErrMsg
    
  Else
    MsgBox "Whoops for some strange reason FormatMessage() failed.", vbCritical, "Error"
    Err.Raise vbObjectError, , ""
  End If
End Sub
 
Public Sub ErrThrowSimple()

   With Err
      .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
   End With

End Sub
 
 
Public Sub ErrThrow()
   
   With Err
      Dim Number&, Source$, Description$, HelpFile$, HelpContext&
      
      Number = .Number
      Source = .Source
      Description = .Description
      HelpFile = .HelpFile
      HelpContext = .HelpContext
      ErrThrow_LastDllError = .LastDllError
      
    ' disable local errorHandler
      On Error GoTo 0
      .Raise Number, Source, Description, HelpFile, HelpContext
      
   End With
End Sub


Public Sub ErrStore()
   With Err
      Number = .Number
      Source = .Source
      Description = .Description
      HelpFile = .HelpFile
      HelpContext = .HelpContext
'      ErrThrow_LastDllError = .LastDllError
   End With
End Sub

Public Sub ErrRestore()
   With Err
      .Number = Number
      .Source = Source
      .Description = Description
      .HelpFile = HelpFile
      .HelpContext = HelpContext
'      ErrThrow_LastDllError = .LastDllError
   End With
End Sub

