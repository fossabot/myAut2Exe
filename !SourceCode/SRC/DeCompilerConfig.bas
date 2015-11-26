Attribute VB_Name = "DeCompilerConfig"
Option Explicit

'used to loacted script start via the "FILE"-Marker Normally this is 6382 or 0x18EE
Public FILE_DecryptionKey As Long

'----------------------------------------

Public Const Script_KEY& = &HAAAAAAAA ' for AHK to get decrypted checksum and scriptStart that are attacted to the end of the file
'^Note: uncritical since myauttoexe uses other methodes to get the start of the scriptfile

Public AU3Sig_HexStr$
'Public Const AU3Sig_HexStr$ = "A3 48 4B BE 98 6C 4A A9 99 4C 53 0A 86 D6 48 7D" ' "£HK¾˜lJ©™LS.†ÖH}"
'Public Const AU3Sig_HexStr$ = "8B 68 83 7B 68 62 9C AD BB 55 15 8C 75 41 94 FA"

Public AU3_TypeStr$
'Public Const AU3_TypeStr$ = "AU3!"
'Public Const AU3_TypeStr$ = "ºin"

'Public AU2_TypeStr$
'Public Const AU2_TypeStr$ = "AU2!"

Public AU3_SubTypeStr$
'Public Const AU3_SubTypeStr$ = "EA06"
'Public Const AU3_SubTypeStr$ = "*¿‰k"

Public AU3_SubTypeStr_old$
'Public Const AU3_SubTypeStr_old$ = "EA05"

Public AU3_ResTypeFile$
'Public Const AU3_ResTypeFile$ = "FILE"
'Public Const AU3_ResTypeFile$ = "Í+åÀ"

Public XORKey_MD5PassphraseHashText_Len&
'Public Const XORKey_MD5PassphraseHashText_Len& = 64193  '&HFAC1
'Public Const XORKey_MD5PassphraseHashText_Len& = &H75D6

Public XORKey_MD5PassphraseHashText_Data&
'Public Const XORKey_MD5PassphraseHashText_Data& = 50130 '&HC3D2
'Public Const XORKey_MD5PassphraseHashText_Data& = &H11DD

Public Data_DecryptionKey&
'Public Const Data_DecryptionKey& = 8879  ' &H22AF
'Public Const Data_DecryptionKey& = &H93DD


Public XORKey_MD5PassphraseHashText_DataNEW&
'Public Const XORKey_MD5PassphraseHashText_DataNEW& = 39410 '&H99F2
'Public Const XORKey_MD5PassphraseHashText_DataNEW& = &H13EE5574

Public Data_DecryptionKey_NewConst&

'Public Const Data_DecryptionKey_NewConst& = 9335 ' &H2477
'Public Const Data_DecryptionKey_NewConst& = &H524ECE35

Public Xorkey_SrcFile_FileInstNEW_Len&
'Public Const Xorkey_SrcFile_FileInstNEW_Len& = 44476 '0xADBC
'Public Const Xorkey_SrcFile_FileInstNEW_Len& = &H26187A62 '0xADBC

Public Xorkey_SrcFile_FileInstNEW_Data&
'Public Const Xorkey_SrcFile_FileInstNEW_Data& = 45887 '0xB33F
'Public Const Xorkey_SrcFile_FileInstNEW_Data& = &HE3F04F31 '0xB33F


Public Xorkey_SrcFile_FileInst_Len&
'Public Const Xorkey_SrcFile_FileInst_Len& = 10684 '0x29BC
Public Xorkey_SrcFile_FileInst_Data&
'Public Const Xorkey_SrcFile_FileInst_Data& = 41566 '0xA25E


Public Xorkey_CompiledPathName_Len&
'Public Const Xorkey_CompiledPathName_Len& = 10668 '29AC
Public Const Xorkey_CompiledPathName_Data& = 62046 'F25E
'Public Const Xorkey_CompiledPathName_Data& = 62046 'F25E

Public Xorkey_CompiledPathNameNEW_Len&
'Public Const Xorkey_CompiledPathNameNEW_Len& = 63520 '0F820
'Public Const Xorkey_CompiledPathNameNEW_Len& = &HC58D486E

Public Xorkey_CompiledPathNameNEW_Data&
'Public Const Xorkey_CompiledPathNameNEW_Data& = 62585 '0F479
'Public Const Xorkey_CompiledPathNameNEW_Data& = &HE6CDCA8F






Public Const AHK_ForceNAPassword As Boolean = False


