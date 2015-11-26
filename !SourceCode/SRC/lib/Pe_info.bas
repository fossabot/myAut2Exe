Attribute VB_Name = "Pe_info_bas"
'Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
'Public Const LOAD_LIBRARY_AS_DATAFILE As Long = &H2
'Public Const RT_GROUP_ICON As Long = (RT_ICON + DIFFERENCE)
'Public Const RT_ICON As Long = 3&
'Public Declare Function FindResource Lib "kernel32.dll" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
'Public Declare Function LoadResource Lib "kernel32.dll" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
'Public Declare Function LockResource Lib "kernel32.dll" (ByVal hResData As Long) As Long
'Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
'



Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Type JOBOBJECT_BASIC_LIMIT_INFORMATION
  PerProcessUserTimeLimit As Long
  PerJobUserTimeLimit As Long
  LimitFlags As Long
  MinimumWorkingSetSize As Long
  MaximumWorkingSetSize As Long
  ActiveProcessLimit As Long
  Affinity As Long
  PriorityClass As Long
  SchedulingClass As Long
End Type


Public Type Section
    SectionName          As String * 8
    VirtualSize          As Long
    RVAOffset            As Long
    RawDataSize          As Long
    PointertoRawData     As Long
    PointertoRelocs      As Long
    PointertoLineNumbers As Long
    NumberofRelocs       As Integer
    NumberofLineNumbers  As Integer
    SectionFlags         As Long
End Type

Public Type PE_Header
  PESignature                    As Long
  Machine                        As Integer
  NumberofSections               As Integer
  TimeDateStamp                  As Long
  PointertoSymbolTable           As Long
  NumberofSymbols                As Long
  OptionalHeaderSize             As Integer
  Characteristics                As Integer
  Magic                          As Integer
  MajorVersionNumber             As Byte
  MinorVersionNumber             As Byte
  SizeofCodeSection              As Long
  InitializedDataSize            As Long
  UninitializedDataSize          As Long
  EntryPointRVA                  As Long
  BaseofCode                     As Long
  BaseofData                     As Long

' extra NT stuff
  ImageBase                      As Long
  SectionAlignment               As Long
  FileAlignment                  As Long
  OSMajorVersion                 As Integer
  OSMinorVersion                 As Integer
  UserMajorVersion               As Integer
  UserMinorVersion               As Integer
  SubSysMajorVersion             As Integer
  SubSysMinorVersion             As Integer
  Reserved                       As Long
  ImageSize                      As Long
  HeaderSize                     As Long
  FileChecksum                   As Long
  SubSystem                      As Integer
  DLLFlags                       As Integer
  StackReservedSize              As Long
  StackCommitSize                As Long
  HeapReserveSize                As Long
  HeapCommitSize                 As Long
  LoaderFlags                    As Long
  NumberofDataDirectories        As Long
'end of NTOPT Header
  ExportTableAddress             As Long
  ExportTableAddressSize         As Long
  ImportTableAddress             As Long
  ImportTableAddressSize         As Long
  ResourceTableAddress           As Long
  ResourceTableAddressSize       As Long
  ExceptionTableAddress          As Long
  ExceptionTableAddressSize      As Long
  SecurityTableAddress           As Long
  SecurityTableAddressSize       As Long
  BaseRelocationTableAddress     As Long
  BaseRelocationTableAddressSize As Long
  DebugDataAddress               As Long
  DebugDataAddressSize           As Long
  CopyrightDataAddress           As Long
  CopyrightDataAddressSize       As Long
  GlobalPtr                      As Long
  GlobalPtrSize                  As Long
  TLSTableAddress                As Long
  TLSTableAddressSize            As Long
  LoadConfigTableAddress         As Long
  LoadConfigTableAddressSize     As Long
  
  BoundImportsAddress            As Long
  BoundImportsAddressSize        As Long
  IATAddress                     As Long
  IATAddressSize                 As Long

  DelayImportAddress             As Long
  DelayImportAddressSize         As Long
  COMDescriptorAddress           As Long
  COMDescriptorAddressSize       As Long
  
  ReservedAddress                As Long
  ReservedAddressSize            As Long
  
'  Gap                            As String * &H28&
  Sections(64)                   As Section
End Type




Public Type PE_Header64
  PESignature                    As Long
  Machine                        As Integer
  NumberofSections               As Integer
  TimeDateStamp                  As Long
  PointertoSymbolTable           As Long
  NumberofSymbols                As Long
  OptionalHeaderSize             As Integer
  Characteristics                As Integer
  Magic                          As Integer
  MajorVersionNumber             As Byte
  MinorVersionNumber             As Byte
  SizeofCodeSection              As Long
  InitializedDataSize            As Long
  UninitializedDataSize          As Long
  EntryPointRVA                  As Long
  BaseofCode                     As Long
  BaseofData                     As Long

' extra NT stuff
  ImageBase                      As Long
'  ImageBase64                      As Long
  SectionAlignment               As Long
  FileAlignment                  As Long
  OSMajorVersion                 As Integer
  OSMinorVersion                 As Integer
  UserMajorVersion               As Integer
  UserMinorVersion               As Integer
  SubSysMajorVersion             As Integer
  SubSysMinorVersion             As Integer
  Reserved                       As Long
  ImageSize                      As Long
  HeaderSize                     As Long
  FileChecksum                   As Long
  SubSystem                      As Integer
  DLLFlags                       As Integer
  StackReservedSize              As Long
  StackReservedSize64              As Long
  StackCommitSize                As Long
  StackCommitSize64                As Long
  HeapReserveSize                As Long
  HeapReserveSize64                As Long
  HeapCommitSize                 As Long
  HeapCommitSize64                 As Long
  LoaderFlags                    As Long
  NumberofDataDirectories        As Long
'end of NTOPT Header
  ExportTableAddress             As Long
  ExportTableAddressSize         As Long
  ImportTableAddress             As Long
  ImportTableAddressSize         As Long
  ResourceTableAddress           As Long
  ResourceTableAddressSize       As Long
  ExceptionTableAddress          As Long
  ExceptionTableAddressSize      As Long
  SecurityTableAddress           As Long
  SecurityTableAddressSize       As Long
  BaseRelocationTableAddress     As Long
  BaseRelocationTableAddressSize As Long
  DebugDataAddress               As Long
  DebugDataAddressSize           As Long
  CopyrightDataAddress           As Long
  CopyrightDataAddressSize       As Long
  GlobalPtr                      As Long
  GlobalPtrSize                  As Long
  TLSTableAddress                As Long
  TLSTableAddressSize            As Long
  LoadConfigTableAddress         As Long
  LoadConfigTableAddressSize     As Long
  
  BoundImportsAddress            As Long
  BoundImportsAddressSize        As Long
  IATAddress                     As Long
  IATAddressSize                 As Long

  DelayImportAddress             As Long
  DelayImportAddressSize         As Long
  COMDescriptorAddress           As Long
  COMDescriptorAddressSize       As Long
  
  ReservedAddress                As Long
  ReservedAddressSize            As Long
  
'  Gap                            As String * &H28&
  Sections(64)                   As Section
End Type


' ------- Additional API declarations ---------------
Public Const IMAGE_ORDINAL_FLAG = &H80000000

Type IMAGE_IMPORT_BY_NAME
   Hint As Integer
   ImpName As String * 254
End Type

Type IMAGE_IMPORT_DESCRIPTOR
   OriginalFirstThunk As Long
   TimeDateStamp As Long
   ForwarderChain As Long
   pDllName As Long
   FirstThunk As Long
End Type

Type IMAGE_BASE_RELOCATION
   VirtualAddress As Long
   SizeOfBlock As Long
End Type


Public IMAGE_IMPORT_DESCRIPTOR As IMAGE_IMPORT_DESCRIPTOR
Public IMAGE_BASE_RELOCATION As IMAGE_BASE_RELOCATION
Public IMAGE_IMPORT_BY_NAME As IMAGE_IMPORT_BY_NAME


'Public Enum ResTypes
'   Icon ' As Integer
'End Enum
'
'Public Type ResourceEntry
'     ResType As Long            '0x00000003  (ICON)
'     OffsetToData As Long       '0x80000048  (DATA_IS_DIRECTORY)
'End Type
'
'Public Type ResDirectory
'   Characteristics As Long
'   TimeDateStamp As Long            '0x486140F5  (Tue Jun 24 18:46:13 2008)
'   MajorVersion As Integer          '0x0053
'   MinorVersion As Integer          '0x0002  -> 83.02
'   NumberOfNamedEntries As Integer  '0x0000
'   NumberOfIdEntries As Integer     '0x0003
'   ResourceEntry As ResourceEntry
'End Type
'
'
'Public Type ResourceDataEntry
'   OffsetToData_RVA As Long
'   Size As Long
'   CodePage As Long
'   Reserved As Long
'End Type
'
'
'Public Type ICONDIRENTRY
'   bWidth As Byte
'   bHeight As Byte
'   bColorCount As Byte
'   bReserved As Byte
'
'   wPlanes As Integer
'   wBitCount As Integer
'   dwBytesInRes As Long
'   dwImageOffset As Long
'End Type
'
'
'Public Type ICONDIR
'   idReserved As Integer   '// Reserved (must be 0)
'   idType As Integer       '// Resource Type (1 for icons)
'   idCount As Integer      '// How many images?
'   idEntry As ICONDIRENTRY '// An entry for each image (idCount of 'em)
'End Type
'
'
'
'
'







'assumption the .text Sections ist the first and .data Section the second in pe_header
Public Const TEXT_SECTION& = 0
Public Const DATA_SECTION& = 1

Public PE_info As New PE_info
Public PE_Header As PE_Header
Public PE_Header64 As PE_Header64

Public IsPE64 As Boolean

'Public file As New FileStream
Public file_readonly As New FileStream
'Public FileName As New ClsFilename
Public PE_SectionData As Collection



Function AlignForFile(value&)
   AlignForFile = Align(value, PE_Header.FileAlignment)
End Function

Function AlignForSection(value&)
   AlignForSection = Align(value, PE_Header.SectionAlignment)
End Function


Function Align&(value&, alignment&)
   '                  keep if equal  |  round up to next bounary
   Align = IIf(value, alignment * (((value - 1) \ alignment) + 1), 0)
End Function


