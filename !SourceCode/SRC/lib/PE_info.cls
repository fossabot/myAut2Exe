VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PE_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private PEStart&
'Private Const CHECKSUM_OPEN_FAILURE As Long = 1
'Private Const CHECKSUM_MAP_FAILURE As Long = 2
'Private Const CHECKSUM_MAPVIEW_FAILURE As Long = 3
'Private Const CHECKSUM_UNICODE_FAILURE As Long = 4
Private Declare Function MapFileAndCheckSum Lib "E:\Programmierung\Projekte\vir_clean\Imagehlp.dll" Alias "MapFileAndCheckSumA" (ByVal FileName As String, ByRef HeaderSum As Long, ByRef CheckSum As Long) As Long

Public Sub Create()

'     '--- find PE-signature ---
'     'Get First 0x400 Bytes
'      Dim tmpstr$
'      file.Position = 0
'      tmpstr = file.FixedString(&H400)
'
'     'Locate start of PE-header
'      PEStart = InStr(1, tmpstr, "PE" & vbNullChar & vbNullChar, vbBinaryCompare)
'      If PEStart = 0 Then err.Raise vbObjectError Or 1, , "No PE-Header Found"
    
     '--- find PE-signature ---
     'Check DOS Header
      Dim tmpstr$
      File.Position = 0
     
     'to skip the Error in VB-IDE Rightclick Toggle/Break on unhandled errors
     'MZ DOS-Header->e_magic
      If File.int16 <> &H5A4D Then Err.Raise vbObjectError Or 1, , "No ExeFile DOS-Header.e_magic<>""MZ"""

     
     'Locate & Validate PE-header
      File.Position = &H3C '   DOS-Header->e_lfanew
      PEStart = File.int32
      File.Position = PEStart
      PEStart = PEStart + 1
      
      If File.int32 <> &H4550 Then Err.Raise vbObjectError Or 2, , "No ExeFile 'PE-Header.Signature<>""PE"""
    
    '  --- get PE_Header  ---
      Dim hFile&
      hFile = FreeFile
      Open File.FileName For Binary Access Read As #hFile
        Get hFile, PEStart, PE_Header
      Close hFile
      
      
    ' Validate Machine Type
      If PE_Header.Machine <> &H14C Then
         If PE_Header.Machine = &H8664 Then
'            Err.Raise vbObjectError Or 4, , "PE-Header.Signature=HDR64_MAGIC!"
            IsPE64 = True
            
                '  --- get PE_Header64  ---
            hFile = FreeFile
            Open File.FileName For Binary Access Read As #hFile
              Get hFile, PEStart, PE_Header64
            Close hFile

            
            
         Else
           Err.Raise vbObjectError Or 3, , "Unsupported PE-Header.Signature 0x" & H16(PE_Header.Machine) & " <>I386(0x14C)."
         End If
      
      Else
         If PE_Header.OptionalHeaderSize <> &HE0 Then
            Err.Raise vbObjectError Or 5, , "PE_Header.OptionalHeaderSize = E0 expected but curvalue is " & H32(PE_Header.OptionalHeaderSize)
         
         End If
   
   '    ' Validate Optionaldata Type
   '      If PE_Header.Magic <> &H10B Then
   '         err.Raise vbObjectError Or 5, , "PE_Header.Magic <> PE32!"
   '      End If


      End If
     
     
      

     
'      If PE_Header.NumberofDataDirectories > &HE Then
'         err.Raise vbObjectError Or 5, , "PE_Header.NumberofDataDirectories must be greater than 14"
'      End If
'imagever 3..5 <-winXP / 6 vista
     
     
End Sub

Public Sub WriteHeader()
   If PEStart = 0 Then Err.Raise vbObjectError, , "You must call PE_info::Create first!"
    
    Dim tmpstr$
    tmpstr = Space(Len(PE_Header) - (UBound(PE_Header.Sections) - PE_Header.NumberofSections) * &H28)
'    Stop
    MemCopyAnyToStr tmpstr, PE_Header, Len(tmpstr)
    File.Position = PEStart - 1
    File.FixedString(-1) = tmpstr
        
'    '  --- get PE_Header  ---
'      Dim hFile&
'      hFile = FreeFile
'      Open file.FileName For Binary Access Write As #hFile
'        Put hFile, PEStart, PE_Header
'      Close hFile

End Sub

Public Function UpdateChecksum&() '(Optional ByRef Output_ChecksumCurrent&, Optional ByRef Output_ChecksumNew&)
   If PEStart = 0 Then Err.Raise vbObjectError, , "You must call PE_info::Create first!"
   UpdateChecksum = MapFileAndCheckSum(File.FileName, 0, PE_Header.FileChecksum)
   WriteHeader
End Function


Public Function GetChecksum&() '(Optional ByRef Output_ChecksumCurrent&, Optional ByRef Output_ChecksumNew&)
   If PEStart = 0 Then Err.Raise vbObjectError, , "You must call PE_info::Create first!"
   Dim tmp&, Retval&
   Retval = MapFileAndCheckSum(File.FileName, 0, tmp)
   GetChecksum = tmp
End Function

Public Function ComputeSizeOfHeader&()
   ComputeSizeOfHeader = PEStart + &H18 + PE_Header.OptionalHeaderSize + PE_Header.NumberofSections * &H28
End Function


Public Sub PE_GetSectionData()
   'ReDim PE_SectionData(PE_Header.NumberofSections - 1)
   
   Set PE_SectionData = New Collection
   
   Dim i&, filedata As StringReader
      For i = 0 To PE_Header.NumberofSections - 1
         With PE_Header.Sections(i)
            
            Set filedata = New StringReader
            
            File.Position = .PointertoRawData
            filedata = File.FixedString(.RawDataSize)
            PE_SectionData.Add filedata
         End With
      Next

End Sub

Public Sub PE_SetSectionData()
   Dim i&
      
     'Seek to first section
     'ToDo: or better Seek end of exact end of PE-Header
      File.Position = PE_Header.Sections(0).PointertoRawData
      
      For i = 0 To PE_Header.NumberofSections - 1
         With PE_Header.Sections(i)
            
          ' Fill Sectiongaps with 0
            Dim gapsize
            gapsize = .PointertoRawData - File.Position
            File.FixedString(-1) = String(gapsize, Chr(0))
            
          ' Seek to start of section
            File.Position = .PointertoRawData
          ' write Sectiondata
            File.FixedString(.RawDataSize) = PE_SectionData(i + 1)
         End With
      Next
'   ReDim PE_SectionData(0)
End Sub

Public Function RawToVA(Raw As Long)
   Dim i&
   RawToVA = -1
   
  'find section
   For i = 0 To PE_Header.NumberofSections - 1
      
      With PE_Header.Sections(i)
         If RangeCheck(Raw, .PointertoRawData + .RawDataSize, .PointertoRawData) Then
            RawToVA = PE_Header.ImageBase + .RVAOffset + (Raw - .PointertoRawData)
            Exit For
         End If
      End With
   
   Next

End Function



Public Function VAToRaw(VA As Long) As Long
   Dim i&, RVA&
   RVA = VA - PE_Header.ImageBase
   
   VAToRaw = -1
   
  'find section
   For i = 0 To PE_Header.NumberofSections - 1
      
      With PE_Header.Sections(i)
         If RangeCheck(RVA, .RVAOffset + .VirtualSize, .RVAOffset) Then
            VAToRaw = .PointertoRawData + (RVA - .RVAOffset)
            Exit For
         End If
      End With
   
   Next

End Function

Public Sub SeekToVA(VA As Long)
   File.Position = VAToRaw(VA)
End Sub


Public Function RVAToRaw(RVA As Long) As Long
   Dim i&
  
   RVAToRaw = -1
   
  'find section
   For i = PE_Header.NumberofSections - 1 To 0 Step -1
      
      With PE_Header.Sections(i)
         If RangeCheck(RVA, .RVAOffset + .VirtualSize, .RVAOffset) Then
            RVAToRaw = .PointertoRawData + (RVA - .RVAOffset)
            Exit For
         End If
      End With
   
   Next

End Function


'Public Function GetFirstIcon(ByRef Out_Offset&, ByRef Out_Size)
'Public Function GetFirstIcon() As String
'
'   Dim hLibrary As Long
'   hLibrary = LoadLibraryEx(File.FileName, 0, LOAD_LIBRARY_AS_DATAFILE)
'
'   Dim lpID
'   lpID = 1
'
'   Dim hRsrc As Long
'   hRsrc = FindResource(hLibrary, lpID, RT_GROUP_ICON)
'
'   Dim hGlobal As Long
'   hGlobal = LoadResource(hLibrary, hRsrc)
'
'   Dim lpIcon As Long
'   lpIcon = LockResource(hGlobal)
'
'   'Load Icons
'   '
'
'
'
'  'Get ResOffSet & Size
'   Dim ResSectionStart&
'   ResSectionStart = RVAToRaw(PE_Header.ResourceTableAddress)
'
'   Dim ResSectionSize&
'   ResSectionSize = PE_Header.ResourceTableAddressSize
'
'
'  'load ResData into ResStruct
'   Dim hFile&
'   hFile = FreeFile
'   Open File.FileName For Binary Access Read As #hFile
'      Dim ResRootDirectory As ResDirectory
'      Get hFile, ResSectionStart + 0 + 1, ResRootDirectory
'
'         Dim ResTypeStart&
'         ResTypeStart = (ResRootDirectory.ResourceEntry.OffsetToData Xor &H80000000)
'
'         Dim ResTypeDirectory As ResDirectory
'         Get hFile, ResSectionStart + ResTypeStart + 1, ResTypeDirectory
'
'            Dim ResDataEntryStart&
'            ResDataEntryStart = (ResTypeDirectory.ResourceEntry.OffsetToData Xor &H80000000)
'
'            Dim ResDataEntryDirectory As ResDirectory
'            Get hFile, ResSectionStart + ResDataEntryStart + 1, ResDataEntryDirectory
'
'
'               Dim ResDataStart&
'               ResDataStart = ResDataEntryDirectory.ResourceEntry.OffsetToData
'
'               Dim ResDataDirectory As ResourceDataEntry
'               Get hFile, ResSectionStart + ResDataStart + 1, ResDataDirectory
'
'
'              'LoadData
'               GetFirstIcon = Space(ResDataDirectory.Size)
'               Get hFile, RVAToRaw(ResDataDirectory.OffsetToData_RVA) + 1, GetFirstIcon
'
'   Close hFile
'
'
'End Function


'Public Function LoadRes()
'
'  'Get ResOffSet & Size
'   Dim ResSectionStart&
'   ResSectionStart = VAToRaw(PE_Header.ResourceTableAddress)
'
'   Dim ResSectionSize&
'   ResSectionSize = PE_Header.ResourceTableAddressSize
'
'
'  'load ResData into ResStruct
'   Dim hFile&
'   hFile = FreeFile
'   Open File.FileName For Binary Access Read As #hFile
'
'
'     'Load RootDir
'      Dim ResRoot As ResDirectory
'      Get hFile, ResSectionStart, ResRoot
'
'     'Get Type
'      Dim ResOffsetToType As Long
'      ResOffsetToType = ResRoot.ResourceEntr
'      'Xor &H80000000
'
'
'      ResRootType = ResRoot.ResourceEntry.ResType Xor &H80000000
'      Select Case ResRootType
'         Case Icon
'
'      End Select
'
'
'
'      Dim ResRootEntries As Long
'      ResRootEntries = ResRoot.ResourceEntry
'
'
'
'   Close hFile
'
'
'
'End Function
'
