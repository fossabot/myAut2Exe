#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Change2CUI=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

;===============================================================================
;
; Function Name:    _GetExeSubsystem
; Description:      Get's the subsystem byte in specified executeable
; Parameter(s):     $sExeFile, Relative or absolute path to exe file
;
; Requirement(s):
; Return Value(s):  True on Success or sets @error to:
;                       1 - Error calling MapAndLoad
;                       2 - No PE file
;                       3 - Error calling UnMapAndLoad
; Author(s):        piccaso, KaFu
;
;===============================================================================
Const $gi_SetPeSubsystem_dwSignature = 0x4550 ; "PE"
Const $Characteristics_IMAGE_FILE_DLL = 0x2000
Const $struct_IMAGE_NT_HEADERS = _
		"DWORD Signature;" & _
		"SHORT Machine;" & _
		"SHORT NumberOfSections;" & _
		"DWORD TimeDateStamp;" & _
		"DWORD PointerToSymbolTable;" & _
		"DWORD NumberOfSymbols;" & _
		"SHORT SizeOfOptionalHeader;" & _
		"SHORT Characteristics;" & _
		"" & _
		"SHORT Magic;" & _
		"BYTE MajorLinkerVersion;" & _
		"BYTE MinorLinkerVersion;" & _
		"DWORD SizeOfCode;" & _
		"DWORD SizeOfInitializedData;" & _
		"DWORD SizeOfUninitializedData;" & _
		"DWORD AddressOfEntryPoint;" & _
		"DWORD BaseOfCode;" & _
		"DWORD BaseOfData;" & _
		"DWORD ImageBase;" & _
		"DWORD SectionAlignment;" & _
		"DWORD FileAlignment;" & _
		"SHORT MajorOperatingSystemVersion;" & _
		"SHORT MinorOperatingSystemVersion;" & _
		"SHORT MajorImageVersion;" & _
		"SHORT MinorImageVersion;" & _
		"SHORT MajorSubsystemVersion;" & _
		"SHORT MinorSubsystemVersion;" & _
		"DWORD Win32VersionValue;" & _
		"DWORD SizeOfImage;" & _
		"DWORD SizeOfHeaders;" & _
		"DWORD CheckSum;" & _
		"SHORT Subsystem;" & _
		"SHORT DllCharacteristics;" & _
		"DWORD SizeOfStackReserve;" & _
		"DWORD SizeOfStackCommit;" & _
		"DWORD SizeOfHeapReserve;" & _
		"DWORD SizeOfHeapCommit;" & _
		"DWORD LoaderFlags;" & _
		"DWORD NumberOfRvaAndSizes"

Const $struct_LOADED_IMAGE = _
		"PTR ModuleName;" & _
		"PTR hFile;" & _
		"PTR MappedAddress;" & _
		"PTR FileHeader;" & _
		"PTR LastRvaSection;" & _
		"UINT NumberOfSections;" & _
		"PTR Sections;" & _
		"UINT Characteristics;" & _
		"BYTE fSystemImage;" & _
		"BYTE fDOSImage;" & _
		"BYTE fReadOnly;" & _
		"UBYTE Version;" & _
		"BYTE Links[8];" & _
		"UINT SizeOfImage"

Const $IMAGE_SUBSYSTEM_NATIVE = 1
Const $IMAGE_SUBSYSTEM_WINDOWS_GUI = 2
Const $IMAGE_SUBSYSTEM_WINDOWS_CUI = 3
Const $IMAGE_SUBSYSTEM_OS2_CUI = 5
Const $IMAGE_SUBSYSTEM_POSIX_CUI = 7
Const $IMAGE_SUBSYSTEM_NATIVE_WINDOWS = 8
Const $IMAGE_SUBSYSTEM_WINDOWS_CE_GUI = 9

Func PEFile_Map($sExeFile)
	Global $IMAGE_NT_HEADERS
	Global $LOADED_IMAGE
	$LOADED_IMAGE = DllStructCreate($struct_LOADED_IMAGE)

	Global $himagehlp_dll
	if $himagehlp_dll<>0 then PEFile_UnMap()

	$himagehlp_dll=DllOpen("imagehlp.dll")

	Global $Ret
	$Ret = DllCall( _
			$himagehlp_dll, _
			"int", "MapAndLoad", _
			"str", $sExeFile, _
			"str", "", _
			"ptr", DllStructGetPtr($LOADED_IMAGE), _
			"int", 0, _
			"int", 0) ;ReadOnly

	If @error Then _
			Return SetError(1, 0, False)
	If $Ret[0] = 0 Then _
			Return SetError(1, 0, False)
	$PointerTo_FileHeader = DllStructGetData($LOADED_IMAGE, "FileHeader")
	$IMAGE_NT_HEADERS = DllStructCreate($struct_IMAGE_NT_HEADERS, $PointerTo_FileHeader)
	If NT_hdr_get("Signature") <> $gi_SetPeSubsystem_dwSignature Then
		PEFile_UnMap()
		Return SetError(2, 0, False)
	EndIf
EndFunc   ;==>PEFile_Map

Func PEFile_UnMap()

	$Ret = DllCall($himagehlp_dll, _
			"int", "UnMapAndLoad", _
			"ptr", DllStructGetPtr($LOADED_IMAGE))
	If @error Then
			$himagehlp_dll=0
			Return SetError(3, 0, False)
	EndIf
	$himagehlp_dll=0

EndFunc   ;==>PEFile_UnMap

Func NT_hdr_get($Fieldname)
	Return (DllStructGetData($IMAGE_NT_HEADERS, $Fieldname))
EndFunc   ;==>NT_hdr_get
Func NT_hdr_set($Fieldname, $value)
	Return (DllStructSetData($IMAGE_NT_HEADERS, $Fieldname, $value))
EndFunc   ;==>NT_hdr_set

Func _GetExeSubsystem()
	$iStubsystem = NT_hdr_get("Subsystem")
	If $Ret[0] <> 0 Then
		Switch $iStubsystem
			Case $IMAGE_SUBSYSTEM_NATIVE
				$sStubsystem = "Native"
			Case $IMAGE_SUBSYSTEM_WINDOWS_GUI
				$sStubsystem = "Windows GUI"
			Case $IMAGE_SUBSYSTEM_WINDOWS_CUI
				$sStubsystem = "Windows CUI"
			Case $IMAGE_SUBSYSTEM_OS2_CUI
				$sStubsystem = "OS2 CUI"
			Case $IMAGE_SUBSYSTEM_POSIX_CUI
				$sStubsystem = "POSIX CUI"
			Case $IMAGE_SUBSYSTEM_NATIVE_WINDOWS
				$sStubsystem = "Native Windows"
			Case $IMAGE_SUBSYSTEM_WINDOWS_CE_GUI
				$sStubsystem = "Windows CE GUI"
		EndSwitch
		Return $sStubsystem
	EndIf
EndFunc   ;==>_GetExeSubsystem

; http://www.autoitscript.com/forum/index.php?showtopic=47809&view=findpost&p=369442
; piccaso

; HOWTO: How To Determine Whether an Application is Console or GUI
; http://support.microsoft.com/kb/90493/en-us?fr=1

Main()
Func Main()

	; Get TargetFile
	Local $PE_FileName = GetPEFile()
	If $PE_FileName = False Then Exit

	; verify TargetFile
	If OpenPEFile($PE_FileName) = False Then Exit
	PEFile_UnMap()

	; Copy *.* to *.dll and open it
	Local $Dll_FileName = $PE_FileName
	ChangeFileExt ($Dll_FileName,"dll")

	FileCopy ($PE_FileName, $Dll_FileName, 1)
	OpenPEFile($Dll_FileName)

	; Set Dll Flag in PE_Header/Characteristics [via BitOR]
	NT_hdr_set("Characteristics", BitOR(NT_hdr_get("Characteristics"), _
			$Characteristics_IMAGE_FILE_DLL) )

	; Save Changes
	PEFile_UnMap()

	Logtxt($Dll_FileName & " created.")

EndFunc   ;==>Main


#include <File.au3>
Func ChangeFileExt(ByRef $Filename, $NewExt)

	Local $szPath, $szDrive, $szDir, $szFName, $szExt

	_PathSplit($Filename, $szDrive, $szDir,$szFName, $szExt)

	$szExt = $NewExt

	$Filename = _PathMake($szDrive, $szDir,$szFName, $szExt)

EndFunc



Func GetPEFile()

	Global $GUIMODE = ($CmdLine[0] == 0)

	If $GUIMODE Then
		$message = "Select Exe to convert To Dll"
		$sFile = FileOpenDialog($message, @ScriptDir & "\", "Exe (*.exe)", 1)

	Else
		$sFile = $CmdLine[1]

	EndIf

	If @error Then
;		GUI_LogErr("No File chosen")
		Return False
	Else
		$sFile = StringReplace($sFile, "|", @CRLF)
	EndIf

	Return ($sFile)

EndFunc

Func OpenPEFile($sFile)

	PEFile_Map($sFile)
	If @error Then
		LogErr("No valid PE-File.")
		Return False
	EndIf
	If (BitAND(NT_hdr_get("Characteristics"), $Characteristics_IMAGE_FILE_DLL)) Then
		LogErr("Already a Dll!")
		Return False
	EndIf

	Return True

EndFunc   ;==>GetAndOpenPEFile


;==========================================================
; LOGGING
;
Func Logtxt($Text)
	if $GUIMODE then
		GUI_Log($Text)
	Else
		ConsoleWrite($Text & @CRLF)
	EndIf
EndFunc   ;==>Logtxt

Func LogErr($Text)
	if $GUIMODE then
		GUI_LogErr($Text)
	Else
		ConsoleWriteError($Text & @CRLF)
	EndIf
EndFunc   ;==>LogErr


Func GUI_Log($Text)
	MsgBox(4096, "", $Text)
EndFunc   ;==>GUI_LogErr

Func GUI_LogErr($Text)
	MsgBox(4096, "", $Text)
EndFunc   ;==>GUI_LogErr