Attribute VB_Name = "ApiHelper"
Option Explicit
 
  
Private Const STATUS_PENDING As Long = &H103
Private Const STILL_ACTIVE As Long = STATUS_PENDING
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
Private Const WAIT_FAILED As Long = &HFFFFFFFF
Private Const WAIT_TIMEOUT As Long = 258&
Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)

Private Const ERROR_INVALID_PARAMETER As Long = 87

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const SYNCHRONIZE As Long = &H100000
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const PROCESS_DUP_HANDLE As Long = (&H40)
Private Const PROCESS_TERMINATE As Long = (&H1)
Private Const PROCESS_VM_OPERATION As Long = (&H8)
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Const SW_HIDE As Long = 0
Public Const SW_MAXIMIZE As Long = 3
Public Const SW_MINIMIZE As Long = 6
Public Const SW_NORMAL As Long = 1
Public Const SW_RESTORE As Long = 9
Public Const SW_SHOW As Long = 5
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Console As New Console

'________________ FILETIME _______________________


Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByVal lpLastAccessTime As Long, ByRef lpLastWriteTime As FILETIME) As Long
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByVal lpLastAccessTime As Long, ByRef lpLastWriteTime As FILETIME) As Long

Public Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Public Declare Function FileTimeToSystemTime Lib "kernel32.dll" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long

Public Const UTF16_BOM$ = "ÿþ" 'Chr(&HFF) & Chr(&HFE)
Public Const UTF8_BOM$ = "ï»¿" 'Chr(&HEF) & Chr(&HBB)& Chr(&HBF)
'Public Const bUnicodeEnable As Boolean = True





'The LB_GETHORIZONTALEXTENT message is useful to retrieve the current value of the horizontal extent:
Const LB_GETHORIZONTALEXTENT = &H193
Const LB_SETHORIZONTALEXTENT = &H194




Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

' Set the horizontal extent of the control (in pixel).
' If this value is greater than the current control's width
' an horizontal scrollbar appears.



Sub Listbox_SetHorizontalExtent(lb As Listbox, ByVal newWidth As Long)
    SendMessage lb.hwnd, LB_SETHORIZONTALEXTENT, newWidth, ByVal 0&
End Sub


' Return the horizontal extent of the control (in pixel).
Function Listbox_GetHorizontalExtent(lb As Listbox) As Long
    Listbox_GetHorizontalExtent = SendMessage(lb.hwnd, LB_GETHORIZONTALEXTENT, 0, ByVal 0&)
End Function


Function ShellEx&(FileName$, Params$, Optional WinStyle As VbAppWinStyle = vbHide)
      
Dim retval&, PID&
'   On Error Resume Next
'  RetVal = ShellExecute(Me.hwnd, "open", """" & App.Path & "/" & "lzss.exe""", "-d """ & dbgFile.FileName & """ """ & outFileName & """", "", SW_NORMAL)
On Error GoTo ShellEx_err
   
   Dim ShellCommand$
   ShellCommand = Quote(FileName) & " " & Params
   PID = Shell(ShellCommand, WinStyle)
   If PID Then
    Dim hProcess&, ExitCode&
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, PID)
    If hProcess Then
    
       GUI_SkipEnable

       Do
          retval = GetExitCodeProcess(hProcess, ExitCode)
'               RetVal = WaitForSingleObject(hProcess, 100)
          myDoEvents
       Loop While retval And (ExitCode = STILL_ACTIVE)
       
       ShellEx = ExitCode
    Else
  'Commented out because sometimes there are false positives( PID get's invalid betweem Shell() and OpenProcess()
'      RaiseDllError "ShellEx()", "OpenProcess", "PROCESS_QUERY_INFORMATION", 0, "PID: " & PID
      If Err.LastDllError <> ERROR_INVALID_PARAMETER Then
         Log "OpenProcess() failed. GetLastError: 0x" & H32(Err.LastDllError)
      End If
      
'      Err.Raise vbObjectError, , "OpenProcess() failed. GetLastError: 0x" & H32(Err.LastDllError)
    End If

     
   End If
Err.Clear
ShellEx_err:

Select Case Err
   Case 0
   Case 5, 53
      Err.Raise vbObjectError Or Err.Number, "ShellEx()", "Shell(" & ShellCommand & ") [@ApiHelper.bas] FAILED! Error: " & Err.Description
   
   Case ERR_SKIP
      retval = TerminateProcess(hProcess, ExitCode)
      If retval Then
         Log "User skipped/canceled process " & FileName & " terminated."
      Else
         Log "User skipped/canceled process " & FileName & " terminated. FAILED! - ErrCode: " & H32(Err.LastDllError)
      End If
   
   Case Else
      Err.Raise vbObjectError Or Err.Number, "ShellEx()", Err.Description
End Select


End Function


'Private Sub FileRename(SourceFileName$, destinationFileName$)
'         On Error Resume Next
'         log_verbose "Copying: " & SourceFileName & " -> " & destinationFileName
'
'         VBA.FileCopy SourceFileName, destinationFileName
'
'         If Err Then log_verbose "=> FAILED - " & Err.Description
'
'End Sub
Public Function FileRename(SourceFileName$, destinationFileName$) As Boolean

      Dim retval&
'      log_verbose "Renaming: " & SourceFileName & " -> " & destinationFileName
      retval = MoveFile(SourceFileName$ & vbNullChar, destinationFileName$ & vbNullChar)
      
      If retval = 0 Then
         On Error Resume Next
         GetAttr SourceFileName
         If Err Then
            log_verbose "=> FAILED - Can't open source file!"
         Else
            GetAttr destinationFileName
            If Err = 0 Then
               log_verbose "=> FAILED - destination file already exists!"
            Else
               log_verbose "=> FAILED - source file is in use!"
            End If
         End If
      Else
         FileRename = True
      End If

End Function

Public Sub FileDelete(SourceFileName$)
   
   On Error Resume Next
   log_verbose "Deleting: " & SourceFileName
   
   Kill SourceFileName
   
   If Err Then log_verbose "=> FAILED - " & Err.Description
  
End Sub

Private Sub createBackup()
   With FileName
      On Error Resume Next
      log_verbose " Creating Backup..."
     
     'Prepare FileNames
      Dim FileExe$, FileBak$
      FileExe = .NameWithExt
      FileBak = .Name & ".vEx"
      
     'Set Workingdir
      ChDrive .Path
      ChDir .Path
      
 '    'Delete .bak
'      FileDelete FileBak
     
      On Error GoTo 0
      
     'Better we close the file before renaming...
     'in short what later will cause problems:
     'the openfilehandle which refered to winlogon.exe will
     'refered to winlogon.bak after renaming, but the File.-objekt
     'still thinks the openfilehandle belongs to winlogon.exe
      File.CloseFile
      
     'Rename .exe to .bak
      FileRename FileExe, FileBak
     
     'copy .bak to .exe
      FileCopy FileBak, FileExe
      
     'Remove readonly attrib & Test if .FileName exists => raise.Err 53
      SetAttr .FileName, vbNormal
   
   End With
End Sub


'Sub log_verbose(Text$)
'   FrmMain.Log Text
'End Sub

Function isUTF16(Text$) As Boolean
   isUTF16 = (Mid(Text, 1, Len(UTF16_BOM)) = UTF16_BOM)
End Function
Function isUTF8(Text$) As Boolean
   isUTF8 = (Mid(Text, 1, Len(UTF8_BOM)) = UTF8_BOM)
End Function


''Converts a string to Unicode if unicode is enabled
'Function T$(TextString$)
'   If bUnicodeEnable Then
'      T = StrConv(TextString, vbUnicode)
'   Else
'      T = TextString
'   End If
'
'End Function
'
'Function Accii$(TextString$)
'   If bUnicodeEnable Then
'      Accii = StrConv(TextString, vbFromUnicode)
'   Else
'      Accii = TextString
'   End If
'
'End Function

