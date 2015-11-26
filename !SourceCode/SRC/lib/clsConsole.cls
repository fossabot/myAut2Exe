VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'_____________________ Console ________________________________
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe _
        As Long, phWritePipe As Long, lpPipeAttributes As Any, _
        ByVal nSize As Long) As Long
        
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile _
        As Long, ByVal lpBuffer As String, ByVal _
        nNumberOfBytesToRead As Long, lpNumberOfBytesRead As _
        Long, ByVal lpOverlapped As Any) As Long
          
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
        lpApplicationName As Long, ByVal lpCommandLine As _
        String, lpProcessAttributes As Any, lpThreadAttributes _
        As Any, ByVal bInheritHandles As Long, ByVal _
        dwCreationFlags As Long, ByVal lpEnvironment As Long, _
        ByVal lpCurrentDirectory As Long, lpSTARTUPINFOupInfo As Any, _
        lpProcessInformation As Any) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
        hObject As Long) As Long

Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal _
        hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize _
        As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, _
        lpBytesLeftThisMessage As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" ( _
        ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
        
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
        
        

Private Const STATUS_PENDING As Long = &H103
Private Const STILL_ACTIVE As Long = STATUS_PENDING


Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type
      
Private Type STARTUPINFO
  cb As Long
  lpReserved As Long
  lpDesktop As Long
  lpTitle As Long
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type
      
Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessID As Long
  dwThreadID As Long
End Type
         
Const NORMAL_PRIORITY_CLASS& = &H20&
Const STARTUPINFOF_USESTDHANDLES& = &H100&
Const STARTF_USESHOWWINDOW& = 1

Public Event OnInit(ProgramName$)
Public Event OnOutput(TextLine$, ProgramName$)
Public Event OnDone(ExitCode&)

Public Function ShellExConsole(FileName$, Params$, Optional ByRef ExitCode&, Optional WinStyle As VbAppWinStyle = vbHide) As String
On Error GoTo ShellExConsole_err
  
   
   
   Dim tmp As New ClsFilename
   tmp = FileName
   
   Dim ProgramName$
   ProgramName = tmp.Name
  
   'GUI_SkipEnable
   RaiseEvent OnInit(ProgramName)
   
  
  
  ' Create CommandLine
    Dim ShellCommand$
    ShellCommand = Quote(FileName) & " " & Params
    
  ' Open Pipe
    Dim sa As SECURITY_ATTRIBUTES
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    
    
    
  ' Create hWritePipe for CreateProcess!STARTUPINFO
  ' -> creates a console we'll read the output from
    Dim retval As Long
    Dim hReadPipe As Long, hWritePipe As Long
    retval = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If retval = 0 Then
        Err.Raise 5, , "CreatePipe failed! RetVal: 0x" & H32(retval)
    End If
   
  ' Prepare STARTUPINFO for CreateProcess
  '    make it use our handle for ConsoleOutput instead of the standard one
    Dim STARTUPINFO As STARTUPINFO
    With STARTUPINFO
      .cb = Len(STARTUPINFO)
      .dwFlags = STARTUPINFOF_USESTDHANDLES
      If WinStyle = vbHide Then
         .dwFlags = .dwFlags Or STARTF_USESHOWWINDOW
      '   .wShowWindow = 0
      
      End If
      .hStdOutput = hWritePipe
 '     .hStdError = hWritePipe
    End With
    
  ' Call CreateProcess
    Dim proc As PROCESS_INFORMATION
    retval = CreateProcessA(0&, ShellCommand$, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, STARTUPINFO, proc)

    If retval <> 0 Then
        
        Do
            Dim lPeekData As Long
            PeekNamedPipe hReadPipe, ByVal 0&, 0&, ByVal 0&, _
                lPeekData, ByVal 0&
            
            If lPeekData > 0 Then
                
              ' Read and append console output data via ReadFile
                Dim L As Long
                Dim Buffer As String
                Buffer = Space$(lPeekData)
                retval = ReadFile(hReadPipe, Buffer, Len(Buffer), L, 0&)
                If retval = 1 Then
                    
                    Buffer = Left(Buffer, L)
                    
                  ' Send line / raise OnOutputLine Event
                    RaiseEvent OnOutput(Buffer, ProgramName)
                    
                    
                    Dim retText As String
                    retText = retText & Buffer
                    
                    
                Else
                    Err.Raise 5, , "ReadFile failed! RetVal: 0x" & H32(retval)
                End If
                
            Else
              ' Wait until procress finishes and get its ExitCode
                Dim GetExitCode_RetVal&
                GetExitCode_RetVal = GetExitCodeProcess(proc.hProcess, ExitCode)
             
            End If
            
            myDoEvents
        Loop While GetExitCode_RetVal And (ExitCode = STILL_ACTIVE)
    
    Else
        Err.Raise 5, , "CreateProcessA failed! RetVal: 0x" & H32(retval)
    End If
    
    CloseHandle hReadPipe
    CloseHandle hWritePipe
    
    ShellExConsole = retText

Err.Clear
ShellExConsole_err:

Select Case Err
   Case 0
   Case 5, 53
      Err.Raise vbObjectError Or Err.Number, "ShellExConsole()", "Shell(" & ShellCommand & ") [@ApiHelper.bas] FAILED! Error: " & Err.Description
      
   Case ERR_CANCEL_ALL
      retval = TerminateProcess(proc.hProcess, ExitCode)
      If retval Then
         Err.Raise ERR_CANCEL_ALL, , "User canceled process " & ProgramName & " terminated."
      Else
         Err.Raise ERR_CANCEL_ALL, , "User canceled process " & ProgramName & " terminated. FAILED! - ErrCode: " & H32(Err.LastDllError)
      End If
      
   Case ERR_SKIP
      retval = TerminateProcess(proc.hProcess, ExitCode)
      If retval Then
         Log "User skipped/canceled process " & ProgramName & " terminated."
      Else
         Log "User skipped/canceled process " & ProgramName & " terminated. FAILED! - ErrCode: " & H32(Err.LastDllError)
      End If
      
      
   Case Else
      Err.Raise vbObjectError Or Err.Number, "ShellExConsole()", Err.Description
End Select


RaiseEvent OnDone(GetExitCode_RetVal)
'GUI_SkipDisable


End Function


