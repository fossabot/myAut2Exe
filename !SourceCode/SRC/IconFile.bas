Attribute VB_Name = "IconFile"
Private Au3Standard_IconFileCrcs As New Collection

Public Sub HandleIconFile(FileName As String)
                   
         If Frm_Options.chk_extractIcon.value <> vbUnchecked Then
                      
         ' ==> Create output fileName
           Dim IconFileName As New ClsFilename
           IconFileName = FileName      ' initialise with ScriptPath
           IconFileName.Ext = ".ico"
           
           Log "Extracting ExeIcon/s to: " & Quote(IconFileName.FileName)
           On Error Resume Next
           ShellEx App.Path & "\" & "data\ExtractIcon.exe", _
                   Quote(File.FileName) & " " & Quote(IconFileName.FileName), vbNormalFocus
           If Err Then
               FrmMain.Log "ERROR: " & Err.Description
               Exit Sub
            End If
           
         ' Test For AutoItStandard
           If Frm_Options.chk_extractIcon.value <> vbUnchecked Then
               
              'init
Au3Standard_IconFileCrcs.Add "AutoIt_Main_v10_48x48_RGB-A.ico", "E1E3EB6E"

Au3Standard_IconFileCrcs.Add "AHK_L___________48x48_RGB-A.ico", "B186AA0D"
Au3Standard_IconFileCrcs.Add "AHK_Classic_____32x32_RGB__.ico", "FCC71A4B"

 
               
              'Get Data
               Dim IconFileData As New StringReader
               IconFileData.Data = FileLoad(IconFileName.FileName)
               
              'Calc CRC
               Dim IconFileDataCrc As String
               IconFileDataCrc = ADLER32(IconFileData)
               
              'Check CRC List
               On Error Resume Next
               Dim FileName_Au3Standard_IconFile As String
               
               FileName_Au3Standard_IconFile = _
                  Au3Standard_IconFileCrcs(IconFileDataCrc)
              
              'Delete File if in CRC List
               If FileName_Au3Standard_IconFile <> "" Then
                  FileDelete IconFileName.FileName
                  FrmMain.Log "   ^- IconFile deleted since it's standard AU3-icon: (" & IconFileDataCrc & ")  '" & FileName_Au3Standard_IconFile & "'"
               End If
   
            End If
         End If

End Sub

Public Function IsStandard_IconFile(uniqueItemID$) As Boolean
   On Error Resume Next
   Au3Standard_IconFileCrcs.Add "", uniqueItemID
   IsUnique = Err = 0
End Function
