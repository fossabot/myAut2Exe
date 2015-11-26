VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Commandline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CommandLine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetCommandLineRaw Lib "kernel32" Alias "GetCommandLineA" () As String
Private argsWithQuotes$()
Private Args$()
Private EmptyArray$()

Public NumberOfCommandLineArgs
Public CommandLine As String


'///////////////////////////////////////////////////
'// getArgs - Get CommandLineArguments with Quotes
Public Function getArgs() As String()
   getArgs = Args
End Function

Public Function getArgsWithQuotes()
   getArgsWithQuotes = argsWithQuotes
End Function


'//////////////////////////////////////////////////////////////////////////
'// Constructor - Sets args, args and NumberOfCommandLineArgs
Private Sub Class_Initialize()
   
  GetCommandLine

'  If CommandLine = Empty Then Exit Sub
  Call commandLineSplitRek

End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
'// commandLineSplitRek -  Split the commandLine and store chunks in args() and argsWithQuotes()
'//                        should only used by the Constructor
Private Sub commandLineSplitRek(Optional ByVal startPos = 1, Optional levelCounter = 0)
  'On Error GoTo commandLineSplitRek_err
   
   Dim endPos%                ' Endposition for cut (...e.exe"<-)
'   Dim startPos%           ' Startposition for "new" commandline (->"C:\p...")
   
   endPos = Len(CommandLine) + 1
  
  'Filter Leading Space
   For startPos = startPos To endPos
     If Mid(CommandLine, startPos, 1) > " " Then Exit For
   Next
   
 ' Trivial case - Empty String
   If startPos >= endPos Then
    ' create array for saving commandline arguments
      NumberOfCommandLineArgs = levelCounter
      If NumberOfCommandLineArgs > 0 Then
         ReDim Args(NumberOfCommandLineArgs - 1)
      Else
       ' Clear Arrays
         Args = EmptyArray
      End If
 
 ' Recursion case
   Else
 
   ' does Commandline starts with "C:\P...
     If Mid(CommandLine, startPos, 1) = """" Then
           
      ' seek over quote
        Inc startPos
        
      ' find next "(DoubleQuote)
        endPos = InStr(startPos, CommandLine, """")
     Else
      ' find next ' '(Space)
        endPos = InStr(startPos, CommandLine, " ")
     End If
     
   ' we are already at the end
     If (endPos <= 0) Then
        endPos = Len(CommandLine) + 1
     End If

   ' Call commandLineSplitRek recursiv with "new" commandline 6 increase levelCounter
     commandLineSplitRek endPos + 1, levelCounter + 1
     
   ' Save Data in Array
     Args(levelCounter) = Mid(CommandLine, startPos, endPos - startPos)
   
   End If
   
'commandLineSplitRek_err:
End Sub
Private Function GetCommandLine() As String

  CommandLine = Command

  
End Function

