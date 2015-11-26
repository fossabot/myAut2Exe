;====================================================
;============= au3 Resources Manager  0.3 ===========
;====================================================
; This script shows the very basics about how to Extract
; the main au3-script from a 'compiled' au3 exe
; Yes this is what FileInstall() does
; but here you see what FileInstall hides from you +
; you can easily extent this seed.
; Like to make it to Delete or Add new Resources to a au3-exe file
; (encryption is a simple XOR that means:
; apply decryption on decrypted code and it get encrypted again and ready to add)
;
;  A3X_Extract($fileInputScript)                   Extracts all Files from an AutoIT3
;  [A3X_Add($fileInputScript,$NewFileName, byref $FileData) Adds fileData to fileInputScript with Name given by $NewFileName]
;  [A3X_Remove($fileInputScript,$FileName)               Removes File that matches with FileName]
;  [A3X_List($fileInputScript,byref $Out_ScriptInfoArr)     List all Files from an AutoIT3 file]


#include <guiconstants.au3>
#include <string.au3>
#Include <GUIConstants.au3>
#Include <GuiListView.au3>
#Include <File.au3>
#include <Date.au3>
#include <Array.au3>
#include <Date.au3>


;  ==== myStatic File StreamReader ====
;  Before use these Data must be set:
   Global $filedataPos  ; Current Position in String
   Global $filedata     ; Data that is modified
;---------------------------------------------

;  Converts Bytes that represent a int number to some it's value
   Func StringToInt($IntString)
      return(String(Binary($IntString)))
   EndFunc


   Func GetUInt8()
      Move (1)
      return StringToInt( StringMid($filedata, $filedataPos-1, 1))
   EndFunc

   Func GetUInt16()
      Move (2)
      return StringToInt( _StringReverse(StringMid($filedata, $filedataPos-2, 2)))
   EndFunc

   Func GetUInt32()
      Move (4)
      return StringToInt( _StringReverse(StringMid($filedata, $filedataPos-4, 4)))
   EndFunc



   Func Move($BytesToMove)
      $filedataPos+=$BytesToMove
   EndFunc

   Func GetString($StrLen)

    ; $StrLen -1 means extract until end of string
      if $StrLen=-1 then $StrLen=StringLen($filedata)-$filedataPos

      Move ($StrLen)
      return StringMid($filedata, $filedataPos-$StrLen, $StrLen)
   EndFunc


; ----------------------------------------------------------------------------
; Main Script Start
; ----------------------------------------------------------------------------
;//////////////////////////////////////////////////////////////////////////////////
;///
;///                         S T A R T
;

   $fileInputScript=FileOpenDialog("Choose some unmodified Au3-Compiled exe","","Compiled au3-exe(*.exe)|Compiled script(*.a3x)|All files(*.*)")
   A3X_Extract($fileInputScript)

exit
; ----------------------------------------------------------------------------
; Main Script End
; ----------------------------------------------------------------------------

;///////////////////////////////////////////////////////////
;/// A3X_Init -  Opens a A3X File
;//
Func A3X_Open($fileInputScript)

Global $fileInputScript_Drive, $fileInputScript_Dir, $fileInputScript_Name, $fileInputScript_ext
   _PathSplit($fileInputScript, $fileInputScript_Drive, $fileInputScript_Dir, $fileInputScript_Name, $fileInputScript_ext)

;* Open Compiled Autoit Exe File & Read in FileData...
   logAdd("opening: " & $fileInputScript)
   $hfileInputScript = FileOpen($fileInputScript,16)
   $filedata = FileRead($hfileInputScript,)
   if @error Then CriticalExit ("Error: Opening/Reading file failed with error: " & @error, "For some reason the File: " & $fileInputScript & " could not be be opened/read.")
   FileClose($hfileInputScript)

  ;   Convert read data to a String
   $filedata = BinaryToString($filedata)

;* Find AutoItSignature (seperate Interpreter & Script Part)
   const $AU3Sig  = _HexToString("A3484BBE986C4AA9994C530A86D6487D")
   $Scripts = StringSplit($filedata,$AU3Sig,1)
   if @error Then CriticalExit ("Error: Unsupported Version of AutoIt script.","Au3 signature scan failed")



 ; Set Encrypted ScriptData as Input DataStream & Seek to it's beginning
   $filedata=$Scripts[$Scripts[0]]
   $filedataPos=1


;* Process Header
   Const $SubTypeSig = "AU3!EA05"
   $SubType = GetString(0x8)
   if ($SubType <> $SubTypeSig) then
      CriticalExit("Unexpected subType"  ,"Expected SubType: "& $SubTypeSig & @CRLF & "Found SubType: " & $SubType & @CRLF & "Probably this is an unsupported version of an AutoIt script.")
   endif

   $MD5PassphraseHash = getString(0x10)
   LogAddWithOffset("MD5PassphraseHash: " & _StringToHex ($MD5PassphraseHash) )

   global $MD5PassphraseHash_ByteSum
   for $i=1 to StringLen($MD5PassphraseHash)
      $MD5PassphraseHash_ByteSum += StringtoInt( StringMid($MD5PassphraseHash,$i,1) )
   next
   LogAdd("$MD5PassphraseHash_ByteSum: " & hex($MD5PassphraseHash_ByteSum))

EndFunc
;///////////////////////////////////////////////////////////
;/// A3X_Extract -  Extracts all Files from an AutoIT3
;//
Func A3X_Extract($fileInputScript)
   A3X_Open($fileInputScript)

   For $FileCount=1 to 0x7ffffff
      if A3X_Process()==false then ExitLoop
      a3x_save()
   next

EndFunc

Func A3X_Process()
;* Process Header

   $ResType = ScriptEncrypt(getString(0x4),5882)
   LogAddWithOffset("ResType: " & $ResType )

   if $ResType<>"FILE" then
      move(-0x4)
      LogAddWithOffset("Extraction Compiled! ")
      Return False
   endif

   global $SrcFile_FileInst = GetCryptedStr(0x29BC, 0xA25E) ;0x29BC A25E
   LogAddWithOffset("SrcFile_FileInst: "& $SrcFile_FileInst)   ;StringLen(">AUTOIT UNICODE SCRIPT<") is 23 Bytes ! - Illumniati agghhr !!!

   $CompiledPathName = GetCryptedStr(10668, 62046) ;29AC  F25E
   LogAddWithOffset("CompiledPathName: " & $CompiledPathName)

   ; ==> Is script compressed
   Global $IsCompressed = getUInt8()
   LogAddWithOffset ("IsCompressed: " & $IsCompressed )

   ; ==> Get size of compressed script data
   $ScriptSize = BitXOR(getUInt32(), 0x45AA ) ;Xor 17834 '45AA
   LogAddWithOffset ("ScriptSize Compressed: " & Hex($ScriptSize) & "  Decimal:" & $ScriptSize)

   $SizeUncompressed = BitXOR(getUInt32() , 0x45AA) ;Xor 17834 '45AA
   LogAddWithOffset ("ScriptSize UnCompressed(not used by aut2exe so far): " & Hex($SizeUncompressed) & "  Decimal:" & $SizeUncompressed)


   ; ==> CRC32 value of uncompressed script data
   $ScriptData_CRC =  BitXOR(getUInt32() ,50130 );'0C3D2
   LogAddWithOffset ("ADLER32 check value for uncompressed script data: " & Hex($ScriptData_CRC))

   LogAdd ("FileTime (number of 100-nanosecond intervals since January 1, 1601) ")
   ;Dim pCreationTime As FILETIME, pLastWrite As FILETIME
   $pCreationTime_dwHighDateTime = getUInt32()
   $pCreationTime_dwLowDateTime = getUInt32()
   LogAddWithOffset ("    pCreationTime:  " & Hex($pCreationTime_dwHighDateTime) & Hex($pCreationTime_dwLowDateTime)) ;& "  " & FormatFileTime(pCreationTime)

   $pLastWrite_dwHighDateTime = getUInt32()
   $pLastWrite_dwLowDateTime = getUInt32()
   LogAddWithOffset ("    pLastWrite   :  " & Hex($pLastWrite_dwHighDateTime) & Hex($pLastWrite_dwLowDateTime)) ;& "  " & FormatFileTime(pLastWrite)

; '==> Read encrypted script data
   LogAddWithOffset ("Begin of script data")
   $ScriptData = getString($ScriptSize)

;  ' ~~~ Process decrypted scriptdata ~~~
   LogAddWithOffset ("Decrypting script data...")


;  Use EncryptionKey to initialise Mersenne Twister random number generator, MT19937
   SRandom ($MD5PassphraseHash_ByteSum + 8879) ;'&H22AF


; '==> Decrypted/encrypted script data
   ;Benchmark execution time with StringMid    :   ~3800
   ;Benchmark execution time with Array(as now):   ~1800
   Global $DecScriptData=""
   $ScriptData = StringSplit($ScriptData,"")
   for $i=1 to $ScriptData[0]
   ;  XOR Each Char of $EncStr with Byte from Random(); Random sequence given through startvalue in SRandom
      $DecScriptData &= chr(BitXOR(asc($ScriptData[$i]), Random(0x0,0xFF,1)))
   next
EndFunc

Func A3x_Save()
;  Generate Output Filename
   local $CompiledPathName, $CompiledName_Drive, $CompiledName_Dir, $CompiledName_Name, $CompiledName_ext
   _PathSplit($SrcFile_FileInst, $CompiledName_Drive, $CompiledName_Dir, $CompiledName_Name, $CompiledName_ext)
;  test if it's ">AUTOIT SCRIPT<" or ">AUTOIT UNICODE SCRIPT<" which is the mainscript ...
   if StringInStr($SrcFile_FileInst,">")<>0 then
   ;  It's the MainScript use Name from Input File and append .au3
      $fileOutScript =         _PathMake($fileInputScript_Drive, $fileInputScript_Dir, $fileInputScript_Name, ".AU3")
   else
   ;  Use Name from $CompiledPathName $SrcFile_FileInst as savename
      $fileOutScript =         _PathMake($fileInputScript_Drive, $fileInputScript_Dir, $CompiledName_Name, ".AU3")
   endif


;  Save Filedata to *.AU3
   $hfileOutScript = FileOpen($fileOutScript,2 + 16)
   FileWrite($hfileOutScript,$DecScriptData)
      if @error Then CriticalExit("Error: Create or WriteFile failed with error: " & @error,"For some reason the decrypted scriptdata cannot be written to file: " & $fileInputScript)
   FileClose($hfileOutScript)


;  Decompress scriptdata if needed
   if $IsCompressed then
   ;  Rename *.au3 -> "*.EA05"
      $fileOutCompressedScript=$fileOutScript & ".EA05"
      FileMove($fileOutScript,$fileOutCompressedScript,1)

   ;  Call 'LZSS.exe' to unpack script data
      $LZSS_exePath= @ScriptDir & "\LZSS.exe"
      if not(FileExists($LZSS_exePath)) Then $LZSS_exePath=FileOpenDialog("LZSS.exe decompression tool not found. Please copy it in here or specify path to where it is.",@ScriptDir,"LZSS decompression Tool(LZSS.exe)|All Files(*.*)")
    ;    Call Command: "LZSS.exe -d *.EA05 *.au3"
      $ShellCommandline = $LZSS_exePath & ' -d "' & $fileOutCompressedScript & '" "' & $fileOutScript &'"'
      $RetVal=RunWait($ShellCommandline,"")
   ;  Note: Decompression is simple and could be also done as AutoIt script
   ;  But the major problem will be performance especially for the bit reader
   ;  since performance of decryption that works on byte level is worse enough.

      LogAdd($ShellCommandline & " called with exitcode: " & $RetVal)
   endif
EndFunc


;///////////////////////////////////////////////////////////
;/// GetCryptedStr - Returns decrypted string
;//
;// Note: Uses GetUInt32() and  GetString() to get inputdata
;//
func GetCryptedStr( Const $LenEncryptionSeed, Const $StrEncryptionSeed)

 ; Get encrypted length from File and Xor it with $LenEncryptionSeed to decrypt it
   local $StrLen = BitXOR(GetUInt32() ,$LenEncryptionSeed)

 ; Now Read that many byte $StrLen tells and Decrypt it
 ; Decryption Key depends on Length of String since it is $StrEncryptionSeed + Length of String
   Return( ScriptEncrypt(GetString($StrLen), $StrEncryptionSeed + $StrLen))

EndFunc



;///////////////////////////////////////////////////////////
;/// ScriptEncrypt - Used to decrypt/encrypt scriptdata
;//
func ScriptEncrypt($EncStr, $EncKey)

  ;Use EncryptionKey to initialise Mersenne Twister random number generator, MT19937
   SRandom($EncKey)

 ; XOR Each Char of $EncStr with Byte from Random(); Random sequence given through startvalue in SRandom
   local $DecStr
   for $i=1 to StringLen($EncStr)
      $DecStr &= chr(BitXOR(Binary(StringMid($EncStr,$i,1)), Random(0x0,0xFF,1)))
   next

   return($DecStr)

EndFunc


;////////////////////////////////////////////////////////////////////////////////
;/// A3X_Add -  Adds fileData to fileInputScript with Name given by $NewFileName
;//
Func A3X_Add($fileInputScript,$NewFileName, byref $FileData)
   CriticalExit("STOP: Sorry not fully implemented yet")
EndFunc


;///////////////////////////////////////////////////////////
;/// A3X_Remove -  Removes File that matches with FileName
;//
Func A3X_Remove($fileInputScript,$FileName)
   CriticalExit("STOP: Sorry not implemented yet")
EndFunc



;///////////////////////////////////////////////////////////
;/// A3X_List -  List all Files from an AutoIT3 file
;//
Func A3X_List($fileInputScript,byref $Out_ScriptInfoArr)
   CriticalExit("STOP: Sorry not implemented yet")
EndFunc




;///////////////////////////////////////////////////////////
;/// LogAddWithOffset - Outputs a log entry with Offset
;//
func LogAddWithOffset($Text)
;TODO Add Logging Code Here
   ;MsgBox(0,"Log With Offset",hex($filedataPos) & "   " &$Text)
   ;DllCall("kernel32.dll", "none", "OutputDebugString", "str", hex($filedataPos) & "   " &$Text)
   ConsoleWrite( hex($filedataPos) & @TAB & $Text & @CRLF )
EndFunc

;//////////////////////////////////
;/// LogAdd - Outputs a log entry
;//
func LogAdd($Text)
;TODO Add Logging Code Here
   ;MsgBox(0,"Log",$Text)
    ;DllCall("kernel32.dll", "none", "OutputDebugString", "str", ($Text))
   ConsoleWrite  (@TAB & $Text & @CRLF )
EndFunc


;DebugMessage
func dm($Msg,$Title)
   MsgBox(0,$Title,$Msg)
EndFunc

;////////////////////////////////////////////////////////////////////
;/// CriticalExit - Show Message with critical error and exit
;//
Func CriticalExit($MessageTitle, $MessageText)
      MsgBox(0x10, $MessageTitle, $MessageText,)
      exit
EndFunc



Exit
;===========================================================
;==== End of Sample - sorry the rest isn't intergated yet ==
; connection to the source of this source:
;    CcWw22Kk@gmx.de <-Please remove double chars before you mail!


$WinMain = GuiCreate('Encryption tool', 300, 100)

$EditText = GuiCtrlCreateEdit('',1,1,1,1)

$InputLevel = GuiCtrlCreateInput(1, 10, 10, 50, 20, 0x2001)
$UpDownLevel = GUICtrlSetLimit(GuiCtrlCreateUpDown($inputlevel),10,1)

$EncryptButton = GuiCtrlCreateButton('Encrypt', 65, 10, 105, 35)
$DecryptButton = GuiCtrlCreateButton('Decrypt', 180, 10, 105, 35)

GuiCtrlCreateLabel('Level',10,35)

GuiSetState()



Do
   $Msg = GuiGetMsg()
   If $msg = $EncryptButton Then

      GuiSetState(@SW_DISABLE,$WinMain)

      $string = FileOpenDialog("Select a file to encrypt","C:\","All (*.*)")
      $fileopen = FileRead($string)

      GuiCtrlSetData($EditText,_StringEncrypt(1,$fileopen,"nothing",GuiCtrlRead($InputLevel)))
      FileDelete($string)
        If Not FileExists($string) Then
            _FileCreate($string)
        EndIf
      Sleep(500)
      FileWrite($string,GUICtrlRead($EditText))

      GuiSetState(@SW_ENABLE,$WinMain)

   EndIf
   If $msg = $DecryptButton Then

      GuiSetState(@SW_DISABLE,$WinMain)

      $string = FileOpenDialog("Select a file to de crypt","C:\","All (*.*)")
      $fileopen = FileRead($string)

      GuiCtrlSetData($EditText,_StringEncrypt(0,$fileopen,"nothing",GuiCtrlRead($InputLevel)))
      FileDelete($string)
      If Not FileExists($string) Then
            _FileCreate($string)
      EndIf
      Sleep(500)
      FileWrite($string,GUICtrlRead($EditText))

      GuiSetState(@SW_ENABLE,$WinMain)

   EndIf
Until $msg = $GUI_EVENT_CLOSE