Attribute VB_Name = "GetCamo"
Option Explicit




Private Function z_szToGREPHex(IsUnicode As Boolean, ParamArray Args())
   
   Dim Texts
   If UBound(Args) = 0 Then
      Texts = Args(0)
   Else
      Texts = Args
   End If
   
   Dim seperator$
   seperator = "\x"
   
   Dim ret As New clsStrCat
   Dim Data() As Byte
   
   Dim Text
   For Each Text In Texts
      Data = Text & IIf(IsUnicode, vbNullChar, "")
   
      Dim i&
      For i = LBound(Data) To UBound(Data) Step IIf(IsUnicode, 1, 2)
         ' ... due to unsafe Unicode convert
         If Not (IsUnicode) Then Debug.Assert Data(i + 1) = 0
         ret.Concat seperator & H8(Data(i))
      Next
   Next
   
   z_szToGREPHex = ret.value
End Function
Public Function szToUnicodeGREPHex(ParamArray Texts() As Variant)
   szToUnicodeGREPHex = z_szToGREPHex(True, Texts)
End Function

Public Function sToGREPHex(ParamArray Texts())
   sToGREPHex = z_szToGREPHex(False, Texts)
End Function

'Public Function REszToGREPHex(Text_RE, Optional Seperator = "\x")
'   Dim Texts()
'   For Each item In Text_RE
'
'      Dim RE
'      RE = False
'      RE = item(2)
'
'      If RE Then
'         Text = RE
'
'      Else
'
'         Dim Text
'         Text = item(1)
'
'      End If
'
'
'      ArrayAdd Texts
'
'   Next
'   szToUnicodeGREPHex = szToGREPHex(Texts)
'
'End Function



Public Sub CamoGet()


 '  Dim FileName$
  ' FileName = "e:\Ablage\4\_dump.exe"
   
   log_verbose "GetCamo's: LoadingFile: " & Frm_Options.Txt_GetCamoFileName
   
   Dim filedata As New StringReader
   filedata = FileLoad(Frm_Options.Txt_GetCamoFileName)
    
    
   On Error Resume Next


' .rdata
'  000056B2                                          25 00 30 00  32 00               % 0 2
'  000056C4   64 00 00 00  72 00 62 00  00 00 00 00  77 00 2B 00  62 00   d   r b     w + b
'  000056D6   00 00 45 41  30 36 00 00  00 00 25 30  32 58 00 00  00 00     EA06    %02X
'  000056E8   41 55 33 21  00 00 00 00  61 00 75 00  74 00 00 00  2A 00   AU3!    a u t   *
'  000056FA   00 00 77 00  62 00 00 00  00 00 46 49  4C 45 00 00  00 00     w b     FILE
'  0000570C   41 00 42 00  53 00 00 00                                    A B S
   
   
Dim Pattern As New clsStrCat

  Pattern.Clear
  
  Pattern.Concat szToUnicodeGREPHex( _
                  "%02d")
                  
  Pattern.Concat RE_Group_NonCaptured( _
         szToUnicodeGREPHex( _
                  "rb", _
                  "", _
                  "w+b" _
                  ) _
               )
   Pattern.Concat "?"
                  
   '#1 AU3_SubType
   Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4))
                  'sToGREPHex( _
                  "EA06")
   Pattern.Concat szToUnicodeGREPHex( _
                  vbNullChar)
               
   Pattern.Concat sToGREPHex( _
                  "%02X")
   Pattern.Concat szToUnicodeGREPHex( _
                  vbNullChar)

   '#2 AU3_Type
   Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4))
                  'szToGREPHex( _
                  "AU3!")

   Pattern.Concat szToUnicodeGREPHex( _
                  vbNullChar, _
                  "aut" _
                  )
   Pattern.Concat RE_Group_NonCaptured(szToUnicodeGREPHex( _
                  "*" _
                  ))
   Pattern.Concat "?"
                  
   Pattern.Concat szToUnicodeGREPHex( _
                  "wb" _
                  )

   
   Pattern.Concat sToGREPHex( _
                  vbNullChar, _
                  vbNullChar)
   '#3 AU3_ResTypeFile
   Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4))
                  'szToGREPHex( _
                  "FILE")

   Pattern.Concat szToUnicodeGREPHex( _
                  vbNullChar)

  '("Wow64DisableWow64FsRedirection  Wow64RevertWow64FsRedirection")?
  ' Pattern.Concat szToUnicodeGREPHex("ABS")

  
  
    
   Dim myRegExp  As New RegExp
   myRegExp.IgnoreCase = False
   myRegExp.Global = False
   myRegExp.MultiLine = True
   
   myRegExp.Pattern = Pattern
   Dim Match  As Match
   Set Match = myRegExp.Execute(filedata.Data)(0)
   
   
   Dim mymatch As SubMatches
   Set mymatch = Match.SubMatches
   
   Debug.Assert mymatch.Count = 3
   With Frm_Options
      .Txt_AU3_SubType_hex = ToHexStr(mymatch.item(0))
      log_verbose H32(Match.FirstIndex) & " ->  Found  AU3_SubType: " & .txt_AU3_SubType
      
      .txt_AU3_Type_hex = ToHexStr(mymatch.item(1))
      log_verbose "Found  AU3_Type : " & .txt_AU3_Type
      
      .txt_AU3_ResTypeFile_hex = ToHexStr(mymatch.item(2))
      log_verbose "Found  AU3_ResTypeFile :" & .txt_AU3_ResTypeFile
   
   End With
   
   
'.data
'00002088   37 BE 0B B4 A1 8E 0C C3   7¾ ´¡Ž Ã
'00002090   1B DF 05 5A 8D EF 02 2D    ß Z ï -
'00002098   28 58 49 00 00 00 00 00   (XI
'000020A0   1C 58 49 00 01 00 00 00    XI
'000020A8   10 58 49 00 02 00 00 00    XI
'000020B0   00 58 49 00 03 00 00 00    XI
'...
'000020F8   3C 57 49 00 0C 00 00 00   <WI
'00002100   E8 59 49 00 01 00 00 00   èYI
'00002108   28 58 49 00 00 00 00 00   (XI
'00002110   1C 58 49 00 01 00 00 00    XI
'00002118   10 58 49 00 02 00 00 00    XI
'00002120   00 58 49 00 03 00 00 00    XI
'00002128   E0 57 49 00 04 00 00 00   àWI
'...
'00002168   3C 57 49 00 0C 00 00 00   <WI
'00002170   28 58 49 00 00 00 00 00   (XI
'00002178   1C 58 49 00 01 00 00 00    XI
'00002180   10 58 49 00 02 00 00 00    XI
'...
'000021D0   3C 57 49 00 0C 00 00 00   <WI
'000021D8   28 58 49 00 00 00 00 00   (XI
'000021E0   1C 58 49 00 01 00 00 00    XI
'000021E8   10 58 49 00 02 00 00 00    XI
'000021F0   00 58 49 00 03 00 00 00    XI
'000021F8   E0 57 49 00 04 00 00 00   àWI
'00002200   C8 57 49 00 05 00 00 00   ÈWI
'00002208   B8 57 49 00 06 00 00 00   ¸WI
'00002210   A4 57 49 00 07 00 00 00   ¤WI
'00002218   8C 57 49 00 08 00 00 00   ŒWI
'00002220   70 57 49 00 09 00 00 00   pWI
'00002228   58 57 49 00 0A 00 00 00   XWI
'00002230   48 57 49 00 0B 00 00 00   HWI
'00002238   3C 57 49 00 0C 00 00 00   <WI
'00002240   99 4C 53 0A 86 D6 48 7D   ™LS †ÖH}
'00002248   A3 48 4B BE 98 6C 4A A9   £HK¾˜lJ©
'00002250   80 00 00 00 00 00 00 00   €
'00002258   00 00 00 00 00 00 00 00
   Pattern.Clear
 
 
 ' myRegExp has a stupid bug - it doesn't matches \x00 !!!
 ' ^-so I used '.' instead
   Pattern.Concat ("\x01...\x02...\x03...\x03...........\x07") '\xBE '\x8E
 
'   Pattern.Concat ("\x37.\x0B\xB4\xA1.\x0C\xC3\x1B\xDF\x05\x5A\x8D\xEF\x02\x2D") '\xBE '\x8E
 '  Pattern.Concat RE_Group_NonCaptured(RE_AnyCharRepeat(429, 429))  '("\x99\x4C\x53\x0A\x86\xD6\x48\x7D")
 
   myRegExp.Pattern = Pattern
   Set Match = Nothing
   Set Match = myRegExp.Execute(filedata.Data)(0)
   If (Match Is Nothing) = False Then
      filedata.Position = Match.FirstIndex
      
      filedata.bSearchBackward = True
      filedata.FindByte &H80
      filedata.bSearchBackward = False
      
      
'    ' Subpattern
'      Pattern.Clear
'      Pattern.Concat ("\x01..." & _
'                      "\x02..." & _
'                      "\x03...")
'      myRegExp.Pattern = Pattern
'
'
'
'      filedata.DisableAutoMove = True
'
'    ' get 1KB tmp buffer
'      Dim filedataTmpBuff As New StringReader
'      filedataTmpBuff.Data = filedata.FixedString(1024)
'      Set Match = Nothing
'      Set Match = myRegExp.Execute(filedataTmpBuff.Data)(0)
'      filedata.DisableAutoMove = False
'
'    ' Seek to 80 00 00 ...
'      filedataTmpBuff.bSearchBackward = True
'      filedataTmpBuff.Position = Match.FirstIndex
'      filedataTmpBuff.FindByte &H80
'
    ' Seek to au3sig start
      filedata.Move -1 - 2 * 8
       
      
   '   Pattern.Clear
   '   Pattern.Concat RE_Group(RE_AnyCharRepeat(8, 8)) '("\x99\x4C\x53\x0A\x86\xD6\x48\x7D")
   '   Pattern.Concat RE_Group(RE_AnyCharRepeat(8, 8)) ' ("\xA3\x48\x4B\xBE\x98\x6C\x4A\xA9")
   ''   Pattern.Concat "[^\x00]\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00"
   '
   '   myRegExp.Pattern = Pattern
   '   Set mymatch = myRegExp.Execute(filedata.FixedString(-1)).item(0).SubMatches
   '
   '
   '
      Dim hex1 As New StringReader
      hex1.Data = filedata.FixedString(8)
      
      Dim hex2 As New StringReader
      hex2.Data = filedata.FixedString(8)
      
      Frm_Options.txt_AU3Sig_Hex = RTrim(ValuesToHexString(hex2) & ValuesToHexString(hex1))
      
      Frm_Options.Chk_NormalSigScan.value = vbChecked
      
   End If
'---------------------------------------------------
   Pattern.Clear
'18EE
   Pattern.Concat "\xE8...\xFF"
   Pattern.Concat (".\xC4.")        ' \x83ADD     ESP, 10
   Pattern.Concat ("\x68")                ' PUSH
   Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4)) '("\x11\x2B\x04\x7F")    '        18EE
   Pattern.Concat ("\x6A\x04")            ' PUSH    4
   Pattern.Concat ("\x8D") '\x54\x24")        '8D55 F4         LEA     EDX, [EBP-C]
   '004572CE    52              PUSH    EDX

   myRegExp.Pattern = Pattern
   Set mymatch = Nothing
   Set mymatch = myRegExp.Execute(filedata.Data).item(0).SubMatches
   
   Dim tmpstr As New StringReader
   tmpstr.Data = mymatch.item(0)

   With Frm_Options
      .txt_FILE_DecryptionKey = H32(tmpstr.int32)
   End With
 
'---------------------------------------------------
   Pattern.Clear
'99F2
   Pattern.Concat "\xE8...\xFF"
   Pattern.Concat (".\xC4.")        ' \x83ADD     ESP, 10
   Pattern.Concat ("\x68")                ' PUSH 99f2
   Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4))
   Pattern.Concat ("\x6A\x10")            ' PUSH    10
   Pattern.Concat ("\x8D") '\x54\x24")        '8D55 F4         LEA     EDX, [EBP-C]
   '004572CE    52              PUSH    EDX

   myRegExp.Pattern = Pattern
   Set mymatch = Nothing
   Set mymatch = myRegExp.Execute(filedata.Data).item(0).SubMatches
   
   tmpstr = mymatch.item(0)

   With Frm_Options
      .txtXORKey_MD5PassphraseHashText_DataNew = H32(tmpstr.int32)
   End With
 
 
 
'---------------------------------------------------
   Pattern.Clear
 
  ' Pattern.Concat ("\x8B\x06")             ' MOV     EAX, [ESI]
   Pattern.Concat ("\x50")                  ' PUSH    EAX
   Pattern.Concat ("\x81\xF7")              ' XOR     EDI,
   Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4)) '("\xBC\xAD\x00\x00") '0ADBC
   Pattern.Concat ("\x8D\x1C\x3F")          ' LEA     EBX, [EDI+EDI]
   Pattern.Concat ("\x53")                  ' PUSH    EBX
     Pattern.Concat ("\x8D") '\x4C\x24\x38")    ' LEA     ECX, [ESP+38]
           '00457313    8D  8D  E0FDFFFF   LEA     ECX, [EBP-220]
   Pattern.Concat RE_AnyCharRepeat(3, 5)
  
     Pattern.Concat ("\x6A\x01")            ' PUSH    1
     Pattern.Concat ("\x51")                ' PUSH    ECX
     Pattern.Concat ("\xE8...\xFF")         ' CALL    004151B0FC\xFF")     '  CALL    004151B0
     Pattern.Concat (".\xC4.")              ' \x83 ADD     ESP, 20
     Pattern.Concat ("\x81\xC7")            '  ADD     EDI,
   Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4)) '("\x3F\xB3\x00\x00") '0B33F
     Pattern.Concat ("\x57")                '  PUSH    EDI
     Pattern.Concat ("\x53")                '  PUSH    EBX
     Pattern.Concat ("\x8D") '\x54\x24.")       '  LEA     EDX, [ESP+28]
 '    Pattern.Concat ("\x52")                '  PUSH    EDX
   myRegExp.Pattern = Pattern
   
   Set mymatch = Nothing
   Set mymatch = myRegExp.Execute(filedata.Data).item(0).SubMatches
   
   
   With Frm_Options
       tmpstr = mymatch.item(0)
      .txtSrcFile_FileInst_LenNew = H32(tmpstr.int32)
      
       tmpstr = mymatch.item(1)
      .txtSrcFile_FileInst_DataNew = H32(tmpstr.int32)

      
   End With
 
 
'---------------------------------------------------
   Pattern.Clear
 
 
                                 '8B7C24 28       MOV     EDI, [ESP+28]
'       Pattern.Concat ("\x8B\x16")             ' MOV     EDX, [ESI]
       Pattern.Concat ("\x52")                 ' PUSH    EDX
       Pattern.Concat ("\x81\xF7")             ' XOR     EDI,
       Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4)) '("\x20\xF8\x00\x00")  0F820
       Pattern.Concat ("\x8D\x1C\x3F")         ' LEA     EBX, [EDI+EDI]
       Pattern.Concat ("\x53")                 ' PUSH    EBX
       Pattern.Concat ("\x8D") '\x44\x24.")      ' LEA     EAX, [ESP+40]
       Pattern.Concat RE_AnyCharRepeat(3, 5)
       Pattern.Concat ("\x6A\x01")              ' PUSH    1
       Pattern.Concat ("\x50")                  ' PUSH    EAX
       Pattern.Concat ("\xE8...\xFF")  ' CALL    004151B0
       Pattern.Concat (".\xC4.")          '\x83 ADD     ESP, 28
       Pattern.Concat ("\x81\xC7")              '  ADD     EDI,
       Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4)) '("\x79\xF4\x00\x00")            '0F479
       Pattern.Concat ("\x57")                  ' PUSH    EDI
   
   myRegExp.Pattern = Pattern
   Set mymatch = Nothing
   Set mymatch = myRegExp.Execute(filedata.Data).item(0).SubMatches
   

   
   With Frm_Options
      tmpstr = mymatch.item(0)
      .txtCompiledPathName_LenNew = H32(tmpstr.int32)
      
      tmpstr = mymatch.item(1)
      .txtCompiledPathName_DataNew = H32(tmpstr.int32)
   End With
 
'---------------------------------------------------
'   Pattern.Clear
'
'
'
'       Pattern.Concat ("\xE8...\xFF")            'E8 11DDFBFF     CALL    0041527B
'
'       Pattern.Concat ("\x8B.\x08")            '\x8B\x46\x08 MOV     EAX, [ESI+8]
'        '                 8B4E 08         MOV     ECX, [ESI+8]
'
'       Pattern.Concat (".\xC4\x10")            '\x83 ADD     ESP, 10
'       Pattern.Concat RE_AnyCharRepeat(1, 2) ' ("\x05") ADD     EAX,   | 81C1 77240000   ADD     ECX, 2477
'       Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4)) '("\x77\x24\x00\x00") ' 2477
'       Pattern.Concat (".") ' ("\x50")                    'PUSH    EAX
'       Pattern.Concat ("\x57")                    'PUSH    EDI
'       Pattern.Concat (".") '\x55")                    'PUSH    EBP
'       Pattern.Concat ("\xE8...\xFF")
   
   
'   Pattern.Clear


'004578A3    C2 0C00         RETN    0C
'004578A6    8B56 08         MOV     EDX, [ESI+8]
'004578A9    81C2 77240000   ADD     EDX, 2477
'004578AF    52              PUSH    EDX
'004578B0    8D8424 BC060000 LEA     EAX, [ESP+6BC]
'004578B7    50              PUSH    EAX
'004578B8    E8 90F6FEFF     CALL    00446F4D
'

'2477
   Pattern.Clear

       Pattern.Concat ("\xC2..")                  'C2 0C00         RETN    0C

       Pattern.Concat ("‹.\x08")                 ' /x8b -> ‹   8B56 08         MOV     EDX, [ESI+8]
       Pattern.Concat RE_AnyCharRepeat(1, 2)     '  05   77240000  ADD     EAX, 2477 |
                                                '   81C1 77240000  ADD     ECX, 2477
       Pattern.Concat RE_Group(RE_AnyCharRepeat(4, 4)) '("\x77\x24\x00\x00") ' 2477
       Pattern.Concat "."                        ' ("\x50")                     'PUSH    EAX | 51              PUSH    ECX
       Pattern.Concat "\x8D" & RE_AnyCharRepeat(5, 6) '  8D8D E0FDFFFF   LEA     ECX, [EBP-220]
       Pattern.Concat "."                       '  "\x50"       '50              PUSH    EAX | EDX
       Pattern.Concat "\xE8...\xFF"
       Pattern.Concat "\x33\xC0"
   
   myRegExp.Pattern = Pattern
   Set mymatch = Nothing
   Set mymatch = myRegExp.Execute(filedata.Data).item(0).SubMatches
   
   With Frm_Options
      tmpstr = mymatch.item(0)
      .txtData_DecryptionKey_New = H32(tmpstr.int32)
   End With
  
  
'---------------------------------------------------------
'00402851    803408 2F       XOR     [BYTE EAX+ECX], 2F
'00402855    41              INC     ECX
'00402856    3B4D 10         CMP     ECX, [EBP+10]
'00402859  ^ 75 F6           JNZ     SHORT 00402851
'0040285B    E9 A1020000     JMP     00402B01
'80 34 08 2F 41 3B 4D 10 75 F6 E9 A1 02 00 00

'

'2477
   Pattern.Clear

       Pattern.Concat (".\x34\x08" & RE_Group(RE_AnyChar))                  '803408 2F       XOR     [BYTE EAX+ECX], 2F
       Pattern.Concat ("\x41\x3B\x4D\x10\x75") '\xF6\xE9\xA1\x02\x00\x00")
       
       
       
   myRegExp.Pattern = Pattern
   Set mymatch = Nothing
   Set mymatch = myRegExp.Execute(filedata.Data).item(0).SubMatches
   
   Dim XORCryptkey&
   XORCryptkey = Asc(mymatch.item(0))
   If XORCryptkey Then
      Log "XORCryptkey: 0x" & H8(XORCryptkey) _
          & "    as char '" & mymatch.item(0) & "'"
      Log "Custom ReadFileHook with XORCryptkey found !!!"
      
    ' Xor & save as *.a3x
      FileName.Ext = "a3x"
      FileSave FileName.FileName, _
         SimpleXor(filedata.Data, XORCryptkey)
      Log "XOR'ed whole file and saved it to " & FileName.FileName
      
      MsgBox "Press Ok to reload " & FileName.NameWithExt & " now ! ", vbInformation, "Xor decrypt done. "
      
    ' Open File
      FrmMain.Combo_Filename = FileName.FileName
      
   End If

 
'---------------------------------------------------
  
   
   Pattern.Clear
 
   Frm_Options.CommitChanges
 
End Sub

Public Function ToHexStr(Data As String) As String
   Dim tmp As New StringReader
   tmp.Data = Data
   ToHexStr = RTrim(ValuesToHexString(tmp))
End Function


Public Function SimpleXor(ScriptData$, ByVal Xor_Key&) As String
   
      
      Dim tmpBuff() As Byte
      tmpBuff = StrConv(ScriptData, vbFromUnicode, LocaleID)
      Dim tmpByte As Byte
      
      Dim StrCharPos&
      For StrCharPos = 0 To UBound(tmpBuff)
         tmpByte = tmpBuff(StrCharPos)
         tmpByte = (tmpByte Xor Xor_Key) And &HFF
         tmpBuff(StrCharPos) = tmpByte
      
         If 0 = (StrCharPos Mod &H8000) Then myDoEvents
         
      Next
      
      SimpleXor = StrConv(tmpBuff, vbUnicode, LocaleID)
      
      
End Function

