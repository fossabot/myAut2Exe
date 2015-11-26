Attribute VB_Name = "ExtractAHKScript"
Option Explicit
'Private Declare Function CreateProcess Lib "kernel32" Alias _
'  "CreateProcessA" (ByVal lpApplicationName As Long, _
'  ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, _
'  ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
'  ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, _
'  ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, _
'  lpProcessInformation As PROCESS_INFORMATION) As Long
'
'Private Type SECURITY_ATTRIBUTES
'   nLength As Long
'   lpSecurityDescriptor As Long
'   bInheritHandle As Long
'End Type
'
'Private Type STARTUPINFO
'   cb As Long
'   lpReserved As Long
'   lpDesktop As Long
'   lpTitle As Long
'   dwX As Long
'   dwY As Long
'   dwXSize As Long
'   dwYSize As Long
'   dwXCountChars As Long
'   dwYCountChars As Long
'   dwFillAttribute As Long
'   dwFlags As Long
'   wShowWindow As Integer
'   cbReserved2 As Integer
'   lpReserved2 As Byte
'   hStdInput As Long
'   hStdOutput As Long
'   hStdError As Long
'End Type
'
'Private Type PROCESS_INFORMATION
'   hProcess As Long
'   hThread As Long
'   dwProcessId As Long
'   dwThreadId As Long
'End Type
'Private Const CREATE_SUSPENDED As Long = &H4
'
'Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long


Private Const RT_ICON As Long = 3&
Private Const RT_RCDATA As Long = 10&
Private Const RT_STRING As Long = 6&
Private Const RT_VERSION As Long = 16
'Private Const RT_GROUP_ICON As Long = (RT_ICON + DIFFERENCE)
'Private Const RT_GROUP_CURSOR As Long = (RT_CURSOR + DIFFERENCE)
Private Const RT_CURSOR As Long = 1&
Private Declare Function FindResource Lib "kernel32.dll" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As Long) As Long


Private Declare Function LoadResource Lib "kernel32.dll" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32.dll" (ByVal hResData As Long) As Long

Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function SizeofResource Lib "kernel32.dll" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long

Function SaveAHK11_Script(ExeFilename As ClsFilename) As Boolean
   
'   Dim StartInfo As STARTUPINFO
'   Dim ProcInfo As PROCESS_INFORMATION
'
   Dim retval&
   Dim hModule&

   hModule = LoadLibraryEx( _
      ExeFilename.FileName, 0, _
      LOAD_LIBRARY_AS_DATAFILE)
      
   If hModule = 0 Then _
         showError ("LoadLibraryEx() ..._AS_DATAFILE fail")
   
   
   Const AHK_SCRIPT_RES_NAME$ = ">AUTOHOTKEY SCRIPT<"
   Dim hResInfo As Long
   hResInfo = FindRes_RCData(hModule, _
      AHK_SCRIPT_RES_NAME)
   
   If hResInfo Then
      'New AHK 1.1 Script found!
      SaveAHK11_Script = True
      
      Dim AHK_Script$
      AHK_Script = LoadRes(hModule, _
         hResInfo)
               
      
      If FreeLibrary(hModule) = 0 Then _
         showError ("FreeLibrary fail")
         
'      Dim FileName_AHKScript As New ClsFilename
'      FileName_AHKScript = ExeFilename
         
      ExeFilename.Ext = "ahk"
         
      FileSave _
         ExeFilename.FileName, _
         AHK_Script
   End If
   
End Function

Public Sub showError(Text)

End Sub

Public Function FindRes_RCData(hModule&, ResName$) As Long
   FindRes_RCData = FindResource(hModule, _
      ResName, _
      RT_RCDATA)
      
'   If hResInfo = 0 Then _
'      showError ("FindResource('" & ResName & "') fail")

End Function
   
Public Function LoadRes(hModule&, hResInfo) As String
   
   
   Dim AHKScriptSize&
   AHKScriptSize = SizeofResource(hModule, _
      hResInfo)
   
   Dim hResData&
   hResData = LoadResource(hModule, _
            hResInfo)
            
   Dim AHKScriptPtr&
   AHKScriptPtr = LockResource( _
            hResData)
            
            
    Dim AHKScript$
    AHKScript = Space(AHKScriptSize)
    
    MemCopy AHKScript, AHKScriptPtr, AHKScriptSize
    
    LoadRes = AHKScript
    
    If (LoadRes Like "; <COMPILER*") = False Then
      MsgBox _
         "This AHK_L exe seems to be packed." & vbCrLf & _
         "Try to manually unpack/dump this and then try again." & _
          vbCrLf & _
          vbCrLf & _
         "Dumping the file with 'Process Hacker' is done by selecting the running" & vbCrLf & _
         "file Properties/Memory. Select all Image(Commit) pages " & vbCrLf & _
         "( normally that is at 0x400000) right click on them and select 'Save' " & vbCrLf & _
         "to dump the uncompressed data to disk.", _
         vbExclamation, _
         "Whoops if ya ask me, that extracted script looks like garbage."


    End If
    
End Function
