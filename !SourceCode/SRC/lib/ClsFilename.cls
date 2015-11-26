VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFilename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Private
Public mvarFileName$


Public Property Get FileName() As String
Attribute FileName.VB_UserMemId = 0
   FileName = mvarFileName
End Property

Public Property Let FileName(ByVal vNewValue As String)
   mvarFileName = Trim(vNewValue)
End Property

'         FileName.WholeName = .lpstrFile
'         FileName.Path = Left(.lpstrFile, .nFileOffset)
'         FileName.FileName = Mid(.lpstrFile, .nFileOffset, .nFileExtension - .nFileOffset)
'         FileName.Ext = Mid(.lpstrFile, .nFileExtension, -1)

'LastDir
Public Property Get Dir() As String
On Error Resume Next
Dim tmp
   tmp = Split(mvarFileName, "\")
   Dir = tmp(UBound(tmp) - 1) & "\"
End Property


Public Property Get Path() As String
   Path = Left(mvarFileName, InStrRev(mvarFileName, "\"))
End Property

Public Property Let Path(ByVal vNewValue As String)
   If Right(vNewValue, 1) = "\" Then
      mvarFileName = vNewValue & Name & Ext
   Else
      mvarFileName = vNewValue & "\" & Name & Ext
   End If
End Property


Public Property Get Name() As String
On Error Resume Next
   Dim NameExt$
   NameExt = Mid$(mvarFileName, Len(Path) + 1)
   
   Dim DotPos&
   DotPos = InStrRev(NameExt, ".")
   If DotPos Then
      Name = Mid$(NameExt, 1, DotPos - 1)
   Else
      Name = NameExt
   End If
   
End Property

Public Property Let Name(ByVal vNewValue As String)
   mvarFileName = Path & vNewValue & Ext
End Property


Public Property Get Ext() As String
On Error Resume Next
   Ext = Mid$(mvarFileName, InStrRev(mvarFileName, "."))
End Property

Public Property Let Ext(ByVal vNewValue As String)
      
   mvarFileName = Path & Name & IIf(Left(vNewValue, 1) = ".", "", ".") & vNewValue
End Property

Public Property Get NameWithExt() As String
On Error Resume Next
   NameWithExt = Mid$(mvarFileName, Len(Path) + 1)
End Property

Public Property Let NameWithExt(ByVal vNewValue As String)
'   mvarFileName = Replace(Path & vNewValue, "\\", "\")
   mvarFileName = Path & vNewValue
End Property

Public Sub MakePath()
   ' create Dir and stop on all error other than 'dir already exists'
   
   Dim NewDir, NewPath$
   NewPath = ""
   
   For Each NewDir In Split(Path, "\")
     
     'Extent path Dir by Dir Create
      NewPath = NewPath & NewDir & "\"
      
      On Error Resume Next
      MkDir NewPath
      
     
     '75 - "Path/File access error"  Example: "\"
     '76 - "Path not found"  Example: "\\intel2400"
      If (Err <> 75) And (Err <> 0) And (Err <> 76) Then
      Dim tmpstr$
         tmpstr = Err.Description & " [" & Err.Number & "] when creating dir: " & Path
         On Error GoTo 0
         Err.Raise vbObjectError, , tmpstr
      End If
   
   Next

End Sub

