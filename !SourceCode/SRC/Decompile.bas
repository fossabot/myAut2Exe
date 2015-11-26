Attribute VB_Name = "DeCompiler"
Option Explicit





'Mersenne Twister
Private Declare Function MT_Init_Raw Lib "data\RanRot_MT.dll" Alias "MT_Init" (ByVal initSeed As Long) As Long
Private Declare Function MT_GetI8 Lib "data\RanRot_MT.dll" () As Long

Private Declare Function RanRot_Init_Raw Lib "data\RanRot_MT.dll" Alias "RanRot_Init" (ByVal initSeed As Long) As Long
Private Declare Function RanRot_GetI8 Lib "data\RanRot_MT.dll" () As Long


'Private Declare Function Uncompress Lib "LZSS.DLL" (ByVal CompressedData$, ByVal CompressedDataSize&, ByVal OutData$, ByVal OutDataSize&) As Long
'Private Declare Function GetUncompressedSize Lib "LZSS.DLL" (ByVal CompressedData$, ByRef nUncompressedSize&) As Long

Private RandSeed As Long
Private isAutoIT2Script As Boolean




'    Sub Main()
'        Dim i As Integer
'
'        TRandomInit (Environment.TickCount) ' initialize with time as seed
'
'        Console.WriteLine ("Random integers in interval from 0 to 99:")
'        For i = 1 To 40
'            Console.Write (TIRandom(0, 99).ToString("00  "))
'            If i Mod 10 = 0 Then
'                Console.WriteLine()
'            End If
'        Next i
'        Console.WriteLine()
'
'        Console.WriteLine ("Random floating point numbers in interval from 0 to 1:")
'        For i = 1 To 32
'            Console.Write (TRandom().ToString("0.000000 "))
'            If i Mod 8 = 0 Then
'                Console.WriteLine()
'            End If
'        Next i
'        Console.WriteLine()
'
'        Console.WriteLine ("Random bits (Hexadecimal):")
'        For i = 1 To 32
'            Console.Write (TBRandom().ToString("X8") + " ")
'            If i Mod 8 = 0 Then
'                Console.WriteLine()
'            End If
'        Next i
'
'    End Sub
'
'End Module




Dim AU3Sig As New StringReader, AU3SigSize&
'Dim PE As New PE_info


Dim FilePath_for_Txt$


Public MD5PassphraseHashText$
Const MD5_HASH_EMPTY_STRING$ = "D41D8CD98F00B204E9800998ECF8427E"

'Const MD5_CRACKER_URL$ = "http://gdataonline.com/qkhash.php?mode=txt&hash="
Const MD5_CRACKER_URL$ = "http://www.md5cracker.de/crack.php?form=Cracken&md5="
'   http://www.milw0rm.com/cracker/info.php?'

Dim bIsProbablyOldScript As Boolean

Dim bIsNewScriptType As Boolean

Dim PEFile_EOF_Offset&
Dim PEFile_EndOfResourceData_Offset&

Dim ScriptData As StringReader

Dim ScriptStartPos&

Private Function MT_Init(ByVal initSeed As Long) As Long
'Fix ugly VB-Bug the makes -27813 to "FFFF935B"!!!
'instead it should be "935B" that is 37723!!!
'Happens instead 7FFF..FFFF
   Debug.Assert initSeed > -65535
   If (initSeed > -65535) And (initSeed < -32768) Then
      initSeed = initSeed And 65535
   End If
   
   MT_Init = MT_Init_Raw(initSeed)
End Function


Private Function RanRot_Init(ByVal initSeed As Long) As Long
'   If (initSeed > -65535) And (initSeed < -32768) Then
'      initSeed = initSeed And 65535
'   End If
   RanRot_Init = RanRot_Init_Raw(initSeed And 65535)
'Note 'And 65535' is the actual Implementation of Ranrot in AU3.4.6
'0044313A >  0FB74424 08     MOVZX   EAX, [WORD ESP+8]
'0044315A >  69C0 FBB4A953   IMUL    EAX, EAX, 53A9B4FB

   
End Function


Sub FL_verbose(Text)
   FrmMain.FL_verbose Text
End Sub

Sub log_verbose(TextLine$)
   FrmMain.log_verbose TextLine
End Sub


Sub FL(Text)
   FrmMain.FL Text
End Sub

Public Sub log2(TextLine$)
   Log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Public Sub Log(TextLine$, Optional LinePrefix$)
   FrmMain.Log TextLine, LinePrefix
End Sub

'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Public Sub Log_Clear()
   FrmMain.Log_Clear
End Sub

'
'Private Sub DeleteBackup()
'     FileRename FileName.Name & ".vEx", FileName.Name & ".del"
'     FileDelete FileName.Name & ".del"
'End Sub

'Working but not need anymore
'Private Sub mt_MT_Init(Key)
'
'
'   Dim Table
'   ReDim Table(624) '0x270
'   Dim v1&, v2&
'   Table(1) = Key
'
'   For i = 1 To UBound(Table) - 1
'     v1 = Table(i)
'     Debug.Assert i <> 5
' ' Cutoff + rotate last 30 bits
' ' v2 = v1 \ &H40000000 '2^30
'   If (v1 >= 0) Then
'      If (v1) < &H40000000 Then '2^30
'         v2 = 0
'      Else
'         v2 = 1
'      End If
'   Else
'      If v1 < &HC0000000 Then '2^30
'         v2 = 2
'      Else
'         v2 = 3
'      End If
'   End If
'
'   v1 = v1 Xor v2
'
'
''    v1 = v1 * 1812433253 '6C078965
'     v1 = Mul(v1, 0, 1812433253, 0) '6C078965
'
''     MsgBox v1
''     v2 = Int(v1 / &H40000000 / 4)
''     ' 9B2 252ADAA2            '2482 623565474
''     ' 9B2 252ADAA2- 9B2 00000000
''     v1 = v1 - (v2 * &H40000000 * 4)
'
'     v1 = v1 + i
'
'     Table(i + 1) = v1
'   Next
'
'End Sub

Private Function GetEncryptStrLen&(LenEncryptionSeed&, hFile As FileStream)
      GetEncryptStrLen = hFile.int32
      GetEncryptStrLen = GetEncryptStrLen Xor LenEncryptionSeed
      
      RangeCheck GetEncryptStrLen, hFile.Length - hFile.Position, 0, "Invalid InputData - StringEncryption length(" & GetEncryptStrLen & ") is bigger than the file"

End Function

Private Function GetEncryptStrNew(LenEncryptionSeed&, StrEncryptionSeed, _
         hFile As FileStream, _
         Optional ConvertOutPutToUTF8 As Boolean = True) As String
      
      Dim StrLen&
      StrLen = GetEncryptStrLen(LenEncryptionSeed, hFile)
      
     'Double size on new type because of Unicode
      Dim StrLenToRead
      StrLenToRead = StrLen + StrLen
      
'      RangeCheck StrLenToRead, hFile.Length - hFile.Position, 0, "GetEncryptStrNew() tried to read a string of is 0x" & H32(StrLenToRead) & " byte thats bigger than the file."
      
      GetEncryptStrNew = StrConv( _
            DeCryptNew(hFile.FixedString(StrLenToRead), StrEncryptionSeed + StrLen) _
                         , vbFromUnicode, LocaleID)

     'Unicode to Accii
      If ConvertOutPutToUTF8 Then
         GetEncryptStrNew = EncodeUTF8(GetEncryptStrNew)
      End If
      
End Function

Private Function DeCryptNew(ByVal Data$, Seed&)
   
   
   RanRot_Init Seed
   
   Dim inBuff As New StringReader
   Dim OutBuff As New StringReader
   inBuff.Data = Data
   OutBuff.Data = Data

 ' Decrypt/Encrypt by  Xor Data from MT with inData
   Do While inBuff.EOS = False
      Dim DataIn&, Key&
      DataIn = inBuff.int8
      Key = RanRot_GetI8()
      
      OutBuff.int8 = DataIn Xor (Key And &HFF)

'      OutBuff.int8 = inBuff.int8 Xor (RanRot_GetI8() And &HFF)
      'DeCrypt = DeCrypt & Chr(inBuff.int8 Xor (MT_GetI8 And &HFF))
   Loop
   
   DeCryptNew = OutBuff.Data
   
   
   
 '  MsgBox _
      "Sorry Decryptions for new au3 Files is not implemented yet." & vbCrLf & _
      "(...and so you can't extract files whose source you don't have.)" & vbCrLf & _
      "" & vbCrLf & _
      "But you can test the TokenDecompiler that is already finished!" & vbCrLf & _
      "" & vbCrLf & _
      "1. add this line at the beginning of the your au3-sourcecode:" & vbCrLf & _
      "  FileInstall('>>>AUTOIT SCRIPT<<<', @ScriptDir & '\ExtractedSource.au3')" & vbCrLf & _
      "2. Compile it with the AutoIt3Compiler." & vbCrLf & _
      "3. Run the exe -> 'ExtractedSource.au3' get's extracted." & vbCrLf & _
      "4. Now open 'ExtractedSource.au3' with this decompiler." & vbCrLf & _
      "" & vbCrLf, _
      vbInformation, "Decryptions for new au3 Files is not implemented yet"
      
'   Err.Raise ERR_NO_AUT_EXE + 100, , "Sorry Decryptions for new Au3 files is not implemented yet :("
End Function



Private Function GetEncryptStr(LenEncryptionSeed&, StrEncryptionSeed, hFile As FileStream) As String
      Dim StrLen&
      StrLen = GetEncryptStrLen(LenEncryptionSeed, hFile)
            
      GetEncryptStr = DeCrypt(hFile.FixedString(StrLen), StrEncryptionSeed + StrLen)
End Function

Private Function DeCrypt(ByVal Data$, Key&)
   'Mersenne Twister (MT) to generate 'random' values
   'http://eprint.iacr.org/2005/165.pdf page 4
   'http://www.ecrypt.eu.org/stream/svn/viewcvs.cgi/ecrypt/trunk/submissions/cryptmt/cryptmt.c?rev=1&view=markup
   'http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/MT2002/emt19937ar.html
   
   If isAutoIT2Script Then
      RandV2_Init Key
   Else
   ' Key->StartSeed for MT
      MT_Init Key
   End If
   
   Dim inBuff As New StringReader
   Dim OutBuff As New StringReader
   inBuff.Data = Data
   OutBuff.Data = Data

 ' Decrypt/Encrypt by  Xor Data from MT with inData
   Do While inBuff.EOS = False
      If isAutoIT2Script Then
         OutBuff.int8 = inBuff.int8 Xor (RandV2 And &HFF)
      Else
         OutBuff.int8 = inBuff.int8 Xor (MT_GetI8 And &HFF)
      End If
         
         'DeCrypt = DeCrypt & Chr(inBuff.int8 Xor (MT_GetI8 And &HFF))
   Loop
   
   DeCrypt = OutBuff.Data
End Function


Private Sub RandV2_Init(Seed&)
   RandSeed = Seed
End Sub


Private Function RandV2&()
   RandSeed = AddInt32(MulInt32(RandSeed, 214013), 2531011) '&H343FD 214013  &H269EC3
    RandV2 = HexToInt(Left(H32(RandSeed), 4)) ' & &H7FFF
   
End Function




Private Function TestForV3_26() As Boolean
   FL_verbose "Testing for AutoIT3.26 Script..."
   With File
      .Position = .Length - 4 - 4
      TestForV3_26 = .FixedString(8) = AU3_TypeStr & AU3_SubTypeStr  '"AU3!EA06"

      If TestForV3_26 = False Then
         FL_verbose "...FAILED!"
      Else
         
'         .Position = .Length - 8 - 4 - 4
         
         Dim Start&
'         Start = .int32
'
'         Dim ScriptEnd&
'         ScriptEnd = .int32
'
'         FL "ScriptStart: 0x" & H32(Start)
'         FL "ScriptEnd: 0x" & H32(ScriptEnd)
         
         .Position = Start
         
'         Err.Raise -1, , "ScriptStart Found"
      End If
   End With
End Function


Private Function TestForV3_2() As Boolean
   FL_verbose "Testing for AutoIT3.2 Script..."
   With File
      .Position = .Length - 4 - 4
      TestForV3_2 = .FixedString(8) = AU3_TypeStr & AU3_SubTypeStr_old '"AU3!EA05"
   End With
   If TestForV3_2 = False Then FL_verbose "...FAILED!"
End Function

Private Function TestForV3_1() As Boolean
   FL_verbose "Testing for AutoIT3.1 Script..."
   With File
      .Position = .Length - 4 - 4 - 4

      Dim Script_End&
      Script_End = .int32 Xor Script_KEY
'      FL "Script_End: " & H32(Script_End) & "  (XORed with 0x" & H32(Script_KEY)

      Dim Script_Start&
      Script_Start = .int32 Xor Script_KEY
'      FL "Script_Start: " & H32(Script_Start) & "  (XORed with 0x" & H32(Script_KEY)

      Dim Script_CRC&
      Script_CRC = .int32 Xor Script_KEY
'      FL "Script_CRC: " & H32(Script_CRC) & "  (XORed with 0x" & H32(Script_KEY)
      
      If (Script_Start < Script_End) Then
         If RangeCheck(Script_Start, .Length, &H1001) And RangeCheck(Script_End, .Length, &H1001) Then
            bIsProbablyOldScript = True
            Dim Script_Data As New StringReader
            .Position = 0
            Script_Data = .FixedString(Script_End)
      
            Dim Script_CRC_Calculated&
            Script_CRC_Calculated = HexToInt(ADLER32(Script_Data))
'            log "Script_CRC_Calculated: " & H32(Script_CRC_Calculated)
            
            TestForV3_1 = Script_CRC_Calculated = Script_CRC
            If TestForV3_1 Then
                  .Position = Script_Start
                  Dim Script_lengh&
                  Script_lengh = .int32 Xor 44460 '&HADAC
               FL "skipping " & H16(Script_lengh) & " byte of random fill data"
                  Dim FillData As New StringReader
                  FillData = .FixedString(Script_lengh)
               Log ValuesToHexString(FillData)
'               log FillData.mvardata
            
            End If   'CRC_test
         End If   'RangeCheck
      End If   'Script_Start < Script_End
   
   End With
   
   If TestForV3_1 = False Then FL_verbose "...FAILED!"

End Function

Private Function TestForV3_0() As Boolean
   FL_verbose "Testing for AHK/AutoIT3.0 Script..."

   With File
      .Position = .Length - 4 - 4
      
      ' ==== Handler for old Scripts ====
      Dim Script_Start&
      Script_Start = .int32
      FL_verbose "Script_Start: " & H32(Script_Start)
      
      Dim Script_CRC&
      Script_CRC = .int32 Xor Script_KEY&
      FL_verbose "Script_CRC: " & H32(Script_CRC) & "  (XORed with 0x" & H32(Script_KEY)
      
      Dim Script_End&
      Script_End = .Length - 4
      log_verbose " ===> Script_End: " & H32(Script_End)
      
      If RangeCheck(Script_Start, .Length, &H1001) Then
         

         bIsProbablyOldScript = True
       
       ' Check CRC32 to be sure that it's in the right format
         CRCInit 79764919 '&H4C11DB7)
      
         Dim Script_CRC_Calculated&
         .Position = 0
         Log "Calculating CRC"
         Script_CRC_Calculated = CRC32(StrConv(.FixedString(Script_End), vbFromUnicode, LocaleID))
         log_verbose "            Script_CRC_Calculated: " & H32(Script_CRC_Calculated)
      
         TestForV3_0 = Script_CRC_Calculated = Script_CRC
         If TestForV3_0 Then
            .Position = Script_Start
            
            Dim modified_AU3_Signature As New StringReader
            modified_AU3_Signature = .FixedString(AU3SigSize)
            Log IIf(modified_AU3_Signature <> AU3Sig, "Modified ", "") & "AU3_Signature: " & ValuesToHexString(modified_AU3_Signature) & "  " & modified_AU3_Signature
         
         ElseIf FrmMain.Chk_verbose.value = vbChecked Then
            Script_CRC_Calculated = Script_CRC_Calculated Xor Script_KEY
            log_verbose "Writing back corrected CRC: " & H32(Script_CRC_Calculated)
            If vbYes = MsgBox("Do you like to write back corrected CRC-value to " & .FileName & " ? ", vbYesNo Or vbDefaultButton2, "Testing for AHK/AutoIT3.0 Script") Then
               .Readonly = False
               .CloseFile
            
               .Position = .Length - 4
               .int32 = Script_CRC_Calculated
               TestForV3_0 = True
            End If

         End If
      End If
   End With
   If TestForV3_0 = False Then FL_verbose "...FAILED!"
   
End Function


Private Function TestForV2_0() As Boolean
   
   FL_verbose "Testing for AutoIT2 Script..."
   
   With File
      .Position = .Length - 4
      
      ' ==== Handler for old Scripts ====
      Dim Script_Start&
      Script_Start = .int32
      FL_verbose "Script_Start: " & H32(Script_Start)
      
      Dim Script_End&
      Script_End = .Length - 4
      log_verbose " ===> Script_End: " & H32(Script_End)
      
      If RangeCheck(Script_Start, .Length, &H1001) Then
         
         .Position = Script_Start
          
         Dim modified_AU3_Signature As New StringReader
         modified_AU3_Signature = .FixedString(AU3SigSize)
         Log IIf(modified_AU3_Signature <> AU3Sig, "Modified ", "") & "AU3_Signature: " & ValuesToHexString(modified_AU3_Signature) & "  " & modified_AU3_Signature
         
         TestForV2_0 = True
         
       Else
         FL_verbose "...FAILED!"
       End If
   End With
   
End Function




Private Sub FindStartOfScriptAlternative()
   With File
      
      bIsProbablyOldScript = Frm_Options.Chk_ForceOldScriptType.value = vbChecked
      If Frm_Options.Chk_ForceOldScriptType.value = CheckBoxConstants.vbGrayed Then
      
         bIsNewScriptType = False
         
         If TestForV3_26 Then
            Log "Script Type 3.2.5+ found."
            bIsNewScriptType = True
            
         
         ElseIf TestForV3_2 Then
            Log "Script Type 3.2 found."
         
         ElseIf TestForV3_1 Then
            Log "Script Type 3.1 found."
            'log "Script_Start is invalid trying something else..."
            Exit Sub
         
         ElseIf TestForV3_0 Then
            Log "Script Type 3.0 found."
            Exit Sub
            
         ElseIf TestForV2_0 Then
            Log "Script Type 2.0 found."
            isAutoIT2Script = True
            Exit Sub
            
         End If 'of New ScriptType
      
      End If 'of Force ScriptType
      
          
' === Alternativ Scan ===
          
    '  Signature not found - try alternative search...
      'Err.Raise vbObjectError Or 41022, , "Error: The executable file is not recognised as a compiled AutoIt script."
Log "AlternativeSigScan for 'FILE'-signature in au3-body..."
'The Compiled Script AutoIT File format
'--------------------------------------
'
'AutoIt_Signature        size 0x14 Byte  String "£HK...AU3!"
'MD5PassphraseHash       size 0x10 Byte                      [LenKey=FAC1, StrKey=C3D2 AHK only]
'ResType                 size 0x4 Byte   eString: "FILE"     [             StrKey=16FA]
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ Search for that !
     
'     .FindString HexvaluesToString("FF 6D B0 CE")  'FF 6D B0 CE       ÿm°Î
      If FindLocation(DeCrypt(AU3_SubTypeStr, 5882), "FILE-(old)signature", True) = -1 Then  '16FA
                 '.FindString HexvaluesToString("6B 43 CA 52")
         
         If FindLocation(DeCryptNew(AU3_SubTypeStr, FILE_DecryptionKey), "FILE-(new)signature", True) = -1 Then  '6382) '18EE
            
'            If LongValScan = False Then
            
               ' Not Found - Search for signature of new Aut3Script
                 If IsValidPEFile Then
                    Log "Alternative search fail - assuming end of exe-stub as start of script. This is very vague but may work..."
                    If (File.Length - PEFile_EOF_Offset) < &H40 Then
                       Err.Raise ERR_NO_AUT_EXE, , "At the end must be at least 0x40 bytes at the end... Please enter start of script manually."
                    Else
                       File.Position = PEFile_EOF_Offset + AU3SigSize
                    End If
                    
                    
                    
                    
                 Else
                    Err.Raise ERR_NO_AUT_EXE, , "'FILE'-signature not found. Please enter start of script manually."
                 End If
               
 '           End If
            Exit Sub
            
         
         Else
         
            '...Finally found :)
            If bIsNewScriptType = False Then
               Log "Modified Script Type 3.2.5+ found."
            End If
            
            bIsNewScriptType = True
      
         End If
         
      End If
   
   '  FILE-Signature found ! Now move back to script start...
'      SeekBackwardsScriptStart
'End Sub
'Private Sub SeekBackwardsScriptStart()

      Dim FilePos_BodyStart&
      FilePos_BodyStart = .Position - 4
      
   '  Determinate if it's an AUTOHOTKEY or AUTOIT SCRIPT ...
'      .Move 4 ' skip over string 'FILE'
   
      Dim Au3SrcFile_FileInst As Boolean
      Dim SrcFile_FileInst$
      If bIsNewScriptType Then
         SrcFile_FileInst = GetEncryptStrNew(Xorkey_SrcFile_FileInstNEW_Len, Xorkey_SrcFile_FileInstNEW_Data, File)  '0xADBC B33F
      Else
         SrcFile_FileInst = GetEncryptStr(Xorkey_SrcFile_FileInst_Len, Xorkey_SrcFile_FileInst_Data, File)  '0x29BC A25E
      End If
      

FL "SrcFile_FileInst: " & SrcFile_FileInst
      If SrcFile_FileInst = ">>>AUTOIT SCRIPT<<<" Then
          Au3SrcFile_FileInst = True
          
      ElseIf SrcFile_FileInst = ">>>AUTOIT NO CMDEXECUTE<<<" Then
          'Note: Script was compiled using '#NoAutoIt3Execute' to block
          Au3SrcFile_FileInst = True
            
            
      ElseIf SrcFile_FileInst = ">AUTOIT UNICODE SCRIPT<" Then
         Au3SrcFile_FileInst = True
         
      ElseIf SrcFile_FileInst = ">AUTOIT SCRIPT<" Then
       ' use AHK_Mode for old scripts
         Au3SrcFile_FileInst = Not (bIsProbablyOldScript)
         
      ElseIf SrcFile_FileInst = ">AUTOHOTKEY SCRIPT<" Then
         Au3SrcFile_FileInst = False
         
      ElseIf SrcFile_FileInst = ">AHK WITH ICON<" Then
         Au3SrcFile_FileInst = False
         
      ElseIf SrcFile_FileInst = ">" Then
         Au3SrcFile_FileInst = False
         
      ElseIf SrcFile_FileInst = "<" Then
         Au3SrcFile_FileInst = False
         
      Else
Log "WARNING: unknown SrcFile_FileInst!"
         Au3SrcFile_FileInst = vbYes = MsgBox("Press YES to process this as an AUTOIT SCRIPT." & vbCrLf & "Press NO to process this as an AUTOHOTKEY SCRIPT.", vbQuestion + vbYesNo, "Unknown SrcFile_FileInst : " & SrcFile_FileInst)
      End If

    ' Now seek back to script start position....
Log "Seeking back to script start position..."
      .Position = FilePos_BodyStart
      If Au3SrcFile_FileInst Then
       ' MD5PasswordHash
         .Move -&H10
       
       ' "EA05"
         .Move -4
       
       ' SubType ["AU3!"]
         .Move -4
       
       ' AU3Signature ["£HK¾..."]
         .Move -AU3SigSize
         
      Else

       '  working but to cryptic
'                  .Move -4
'                  Do Until (4 + .Position + (.int32 Xor XORKey_MD5PassphraseHashText_Len)) = FilePos_BodyStart
'                     .Move (-1 - 4)
'                  Loop
'                  .Move -4
        
         
       ' Determinating length of au3-password
       ' Expected format:
       '   <DWORD>Len  <String>Password   <Offset>FilePos_BodyStart...
         
         Const AU3_MAX_PASSWORDLEN = 263
         Do
          ' Get Length
            Dim PasswordLen As Double
            .Move -4
            PasswordLen = .int32 Xor XORKey_MD5PassphraseHashText_Len

          ' If Current File Position + Length is FilePos_BodyStart...
            If .Position + PasswordLen = FilePos_BodyStart Then
              '...it's the correct length; so seek back a Dword[4byte]
               .Move -4
              'Exit Loop
               Exit Do
            ElseIf (FilePos_BodyStart - .Position) >= AU3_MAX_PASSWORDLEN Then
               Err.Raise vbObjectError, , "Determinating length of au3-password failed (length >= " & AU3_MAX_PASSWORDLEN & ")"
            End If

           ' Seek to next position to try
            .Move -1
         Loop While True

       ' SubType
         .Move -1
       
       ' AU3Signature
         .Move -AU3SigSize
      End If

      
      Dim modified_AU3_Signature As New StringReader
      modified_AU3_Signature = .FixedString(AU3SigSize)
      
      Log IIf(modified_AU3_Signature <> AU3Sig, "Modified ", "") & "AU3_Signature: " & ValuesToHexString(modified_AU3_Signature) & "  " & modified_AU3_Signature

'            log "Not found trying heuristic search..."
'            PE.Create
'            Dim LastSection As PE_Section
'            With LastSection
'               LastSection = PE_Header.Sections(PE_Header.NumberofSections - 1)
'               log "LastSection in PE_Header is: " & szNullCut(.SectionName) & " at: " & H32(.PointertoRawData) & " Size: " & H32(.RawDataSize)
'
'               Dim ScriptStart
'
'            End With
   End With

End Sub
Private Sub FindStartOfScript()
   
   If Frm_Options.Chk_NormalSigScan.value = vbChecked Then
      
      Dim Location&
      Location = FindLocation(AU3Sig.Data & AU3_TypeStr, "AutoIt3 Signature")
      If Location = -1 Then
         Location = FindLocation(AU3Sig.Data, "AutoIt2 Signature")
         If Location = -1 Then
            FindStartOfScriptAlternative
         End If
      Else
         File.Move -4 ' rewind "AU3!"
      End If
      
   Else
      FindStartOfScriptAlternative
   End If
   


End Sub

Private Function FindLocation(SearchPattern$, Optional PatternName$ = "", Optional AlwaysUseFirstLocation As Boolean = False) As Long
   
   With File
       
      Dim tmp As New StringReader
      tmp = SearchPattern
       
     ' ===> Find Script Signature in FileData  (and place FileReadPointer behind it)
      Log "Scanning for " & PatternName & ": " & ValuesToHexString(tmp) & "   " & SearchPattern
     
     ' .Position = 0
     
     'Search for AutoIt Signature( from behind)
      Dim Locations As Collection
      Set Locations = .FindStrings(SearchPattern)

      
      
       ' and check if Findpattern was found more than one time
      If Locations.Count = 0 Then
       '  Signature not found - try alternative search...
         'Err.Raise vbObjectError Or 41022, , "Error: The executable file is not recognised as a compiled AutoIt script."
   Log "...not found."
         FindLocation = -1
        
      Else
      
         If (Locations.Count = 1) Or AlwaysUseFirstLocation Then
         'Okay one occurance - as it should be
            Dim SeektoLocation&
            SeektoLocation = 1
         
         Else
            Set frmSearchResults.Locations = Locations
            FrmMain.Hide
            frmSearchResults.Show vbModal
            FrmMain.Show
            SeektoLocation = frmSearchResults.SelectedLocation
            If SeektoLocation = -1 Then Err.Raise ERR_CANCEL_ALL, , "Cancel during selecting script location"

            
'            SeektoLocation = InputBox("There are " & Locations.Count & " possible location were the Script starts, please choose one to try:", , 1)
'            RangeCheck SeektoLocation, Locations.Count, 1, "Invalid location value!"
         End If
         
         .Position = Locations(SeektoLocation)
         .Move Len(SearchPattern)
      
      
         FindLocation = .Position

      End If
         
         
   End With

End Function

Private Function FormatFileTime(TimeStamp As FILETIME) As String
   Dim SysTime As SYSTEMTIME
   With SysTime
      FileTimeToSystemTime TimeStamp, SysTime
      FormatFileTime = Format(.wDay & "." & .wMonth & "." & .wYear & " " & .wHour & ":" & .wMinute & ":" & .wSecond, "dd.mm.yyyy hh:mm:ss") & " [" & .wMilliseconds & "]"
   End With
End Function

Private Sub UserPassWordCheck(MD5PassphraseHashText$, bIsClearTextPwd As Boolean)
   #If DoUserPassWordCheck Then
'////////////////////////////////////////////////////////////////////
'//
'//  A t t e n t i o n , W a r n i n g , A t t e n t i o n , W a r n i n g
'//                P r o t e c t e d  C o d e
'// It is strictly FORBIDDEN to REMOVE or modify the following code:
         
         Dim md5 As ClsMD5
         Set md5 = New ClsMD5
         Dim userPassword$, userPassword_Hash$, scriptPassword_Hash$
         scriptPassword_Hash = LCase(MD5PassphraseHashText)
         Do
            userPassword = InputBox("Please Password:", "Script File is Password Protected", "Sorry but for legal reason you must enter a valid password to continue.")
            If userPassword = "" Then Err.Raise vbObjectError, , "Stopped because user didn't entered a valid password!"
            
              'According to type test clearTextPWD or Hash
               If bIsClearTextPwd Then
                  userPassword_Hash = userPassword
               Else
                  userPassword_Hash = md5.md5(userPassword)
               End If
         
         Loop Until userPassword_Hash = scriptPassword_Hash

'//                   E N D  O F  'untouchable code'               //
'////////////////////////////////////////////////////////////////////
#End If


End Sub



'////////////////////////////////////////////////////
'/// Decompile - Decompiles .exe[->File] to .au3 or .ahk
'//
'//  Notes:
'//   Not indented lines are for log purpose only (and not so important)
Public Sub Decompile()
   Dim FilePointerFallBackOnError&
   
   
   isAutoIT2Script = False
   AU3Sig = HexvaluesToString(AU3Sig_HexStr) ' & "AU3!"  ' "£HK¾˜lJ©™LS.†ÖH}"
   AU3SigSize = Len(AU3Sig)

'log "---------------------------------------------------------"
   
  'Clear ExtractedFiles
   Set ExtractedFiles = New Collection
   
   With File
    
      Log "Unpacking: " & FileName.FileName
      .Create FileName.FileName, False, False, True
      .Position = 0
      
    ' Chk_NormalSigScan is disabled when Txt_Scriptstart is set
      If Frm_Options.Chk_NormalSigScan.Enabled = False Then
         .Position = HexToInt(FrmMain.Txt_Scriptstart)
         .Move AU3SigSize
     
     'Find start of script and quit this function with runtime error if search fails
      'ElseIf Frm_Options.Chk_NormalSigScan = vbChecked Then
      '   FindStartOfScript
      Else
         
         On Error Resume Next
         
         FindStartOfScript
         
         
         Dim FindStartOfScript_err_Number&
         FindStartOfScript_err_Number = Err.Number
         
         Dim FindStartOfScript_err_Description$
         FindStartOfScript_err_Description = Err.Description

         On Error GoTo 0
         
      End If
      
      
      ScriptStartPos = .Position - AU3SigSize
      Log ""
      Log " ---> ScriptStartOffset: " & H32(ScriptStartPos)
      FilePointerFallBackOnError = ScriptStartPos
      
      RangeCheck .Position, .Length, 0, "ERROR: ScriptStartPosition is outside the file! -", "Decompile"
      
      
    ' --- Save Stub  - if not PEFile ---
      If Not IsValidPEFile Then
         If ScriptStartPos > 0 Then
            Log "This is no PE-Exe File & Script don't start at Offset 0 -> Saving StubData"
        
           
            Dim FileName_FileStub$
            FileName_FileStub = FileName.NameWithExt & ".stub"
            Log "Copy FileStubData into: " & FileName_FileStub
            
            FileSave FileName.Path & FileName_FileStub, _
                     FileReadPart(.FileName, 0, ScriptStartPos)
          
         End If
       
      Else
         Log "      EndOf_PE-ExeFile : " & H32(PEFile_EOF_Offset)
         Log "      EndOf_PE-ExeFile_ResourceData : " & H32(PEFile_EndOfResourceData_Offset)
   
         
         HandleIconFile File.FileName
         
         
         Dim isAHK11_Script As Boolean
         isAHK11_Script = SaveAHK11_Script(FileName)
      
         If isAHK11_Script Then
            ExtractedFiles.Add FileName.FileName
            Exit Sub
         End If
         
         
         On Error GoTo 0

  
      End If
    
      If FindStartOfScript_err_Number Then
          On Error GoTo 0
          Err.Raise _
            FindStartOfScript_err_Number, "", _
            FindStartOfScript_err_Description
      End If
    
    
    'File
    ' ===> Check if it's an old or New AutoIt Script
      Dim SubType As New StringReader:   SubType.DisableAutoMove = True
      SubType = .FixedString(4)
      FL "SubType: 0x" & H8(SubType.int8) & "  " & SubType.mvardata
      
      Dim bIsOldScript As Boolean
      If SubType.Data = AU3_TypeStr Then
         bIsOldScript = False
    
    ' the offical AutoHotkey Script Decompiler checks this to be '3'
      ElseIf SubType.int8 = 3 Then
         bIsOldScript = True
      
      ElseIf SubType.int8 = 4 Then
         bIsOldScript = True
      
    ' AutoIT2 script
      ElseIf SubType.int8 = 1 Then
         bIsOldScript = True
         isAutoIT2Script = True
      
      
      Else
         'err.Raise vbObjectError,,"Unexpected Script subtype"
         FL "Unexpected Script subtype: " & "0x" & H32(SubType.int32) & " " & SubType.Data
        'Ask user for Script subtype
         Dim MsgBoxResult&
         MsgBoxResult = MsgBox("Is this an AutoIT3 file?", vbQuestion + vbYesNoCancel + _
                        IIf(IsAutoIT3File, vbDefaultButton1, vbDefaultButton2))
         If MsgBoxResult = vbCancel Then
            Err.Raise ERR_CANCEL_ALL, , "User hit CANCEL on the question: ""Is this an AutoIT3 file?"""
         End If

         bIsOldScript = vbNo = MsgBoxResult
         
         If bIsOldScript Then
            isAutoIT2Script = vbYes = MsgBox("Is this an AutoIT 2 file?", vbQuestion + vbYesNo + _
                        IIf(IsAutoIT2File, vbDefaultButton1, vbDefaultButton2))
         End If
         
      End If
      


      Log "~ Note:  The following offset values are were the data ends (and not were it starts) ~"


    
    ' ===> Get Script Password
      Dim MD5PassphraseHash As New StringReader
      If bIsOldScript Then
       ' Old AutoIT Script if branch...
       ' Move three bytes back since SubType is only 1 Byte but before we read 4 byte
         .Move -3
         MD5PassphraseHash = ReadPassword 'GetEncryptStr(XORKey_MD5PassphraseHashText_Len, XORKey_MD5PassphraseHashText_Data, File) '&HFAC1, &HC3D2
         'MD5PassphraseHash = GetEncryptStr(&H36A73993, XORKey_MD5PassphraseHashText_Data, File)
         MD5PassphraseHashText = MD5PassphraseHash
         
      
      Else
       ' New AutoIT script if branch...
         
         
         Dim Type2$
         Type2 = .FixedString(4)
         
         bIsNewScriptType = Type2 = AU3_SubTypeStr
         If bIsNewScriptType Then
            FL "New tokenised AutoIt script found."
         
         ElseIf Type2 <> AU3_SubTypeStr_old Then
            FL "Type2 = " & Type2 & "  Normally you would get 'Error: Unsupported Version of AutoIt script.' here"
           
          ' Ask user for Script Type2
            bIsNewScriptType = vbYes = MsgBox("Is this a new tokenise AutoIT3 file(=Ver 3.2.6 -Aug2007- and higher) ?", vbQuestion + vbYesNo)


         Else
            FL "AutoIt Script Found.  - Type2 = " & Type2
         End If
         
         
         'Err.Raise vbObjectError Or 41022, , "Error: Unsupported Version of AutoIt script."

      
         ' GetPassword Hash from with later the key to decrypt the script is calculated
         MD5PassphraseHash = .FixedString(&H10)
         If bIsNewScriptType Then MD5PassphraseHash = DeCryptNew(MD5PassphraseHash, XORKey_MD5PassphraseHashText_DataNEW) '&H99F2
         
         MD5PassphraseHashText = ValuesToHexString(MD5PassphraseHash, "")
           
         Dim IsHashForEmptyPassword As Boolean
         IsHashForEmptyPassword = MD5PassphraseHashText = MD5_HASH_EMPTY_STRING$
         If IsHashForEmptyPassword Then MD5PassphraseHashText = ""
            
      End If
      
      
     '==> Ask User For Password
      If (MD5PassphraseHashText = "") Then
         Log "Script has no password (MD5PassphraseHash for password """" )"

      Else
         Log "Script is password protected!"

         #If DoUserPassWordCheck Then
         '////////////////////////////////////////////////////////////////////
         '//
         '//  A t t e n t i o n , W a r n i n g , A t t e n t i o n , W a r n i n g
         '//                P r o t e c t e d  C o d e
         '// It is strictly FORBIDDEN to REMOVE or modify the following code:
                  
          UserPassWordCheck MD5PassphraseHashText$, bIsOldScript
          
         
         '//                   E N D  O F  'untouchable code'               //
         '////////////////////////////////////////////////////////////////////
         #End If
      
      End If

      FL "Password/MD5PassphraseHash: " & ValuesToHexString(MD5PassphraseHash, "")
      Log Space(8 + 4) & MD5PassphraseHash.Data
      
      FrmMain.mi_MD5_pwd_Lookup.Visible = (IsHashForEmptyPassword = False) And (bIsOldScript = False)


    
    ' ==> Prepare decryption of script...
    ' A 32 bit checksumvalue over all bytes from the MD5PassphraseHash is the decryptionkey
      Dim MD5PassphraseHash_ByteSum&
      MD5PassphraseHash_ByteSum = 0
      
      MD5PassphraseHash.EOS = False
      Do Until MD5PassphraseHash.EOS

         If bIsNewScriptType Then
          ' For AHK scripts use signed int8 to multiply
          ' Note: as bug or with intention startvalue is 0 so MD5PassphraseHash_ByteSum will be also always 0.
            MD5PassphraseHash_ByteSum = MD5PassphraseHash_ByteSum * MD5PassphraseHash.int8Sig
         ElseIf bIsOldScript Then
          ' For AHK scripts use signed int8 to also compute äöü correct
            MD5PassphraseHash_ByteSum = MD5PassphraseHash_ByteSum + MD5PassphraseHash.int8Sig
         Else
          ' For new MD5 scripts use unsigned int8 to compute
            MD5PassphraseHash_ByteSum = MD5PassphraseHash_ByteSum + MD5PassphraseHash.int8
         End If
         
'         Debug.Print MD5PassphraseHash.Position, H32(MD5PassphraseHash_ByteSum)
      Loop
      Log "MD5PassphraseHash_ByteSum: " & H32(MD5PassphraseHash_ByteSum) & "  '+ " & IIf(bIsNewScriptType, H32(Data_DecryptionKey_NewConst), H32(Data_DecryptionKey)) & "' => decryption key!"

   
   
   
      Log "------------ Processing Body -------------"
      On Error GoTo err_ProcessingBody
      ErrStore 'Clear
      
    ' Set default
      Dim ResTypeFILE$
      ResTypeFILE = AU3_ResTypeFile
      
      Dim FileCount&
      For FileCount = 1 To &H7FFFFFF
         
       ' Used for saving overlaydata
         FilePointerFallBackOnError = .Position
      
        'so the rare case, that we're already at the end
         If .EOS Then Exit For

      '===> read various Header data
         Dim ResType$
         If bIsNewScriptType Then
            ResType = DeCryptNew(.FixedString(4), FILE_DecryptionKey) '6382) '18EE
         Else
            ResType = DeCrypt(.FixedString(4), 5882) '000016FA
         End If
         If ResType <> ResTypeFILE Then
         
           ' Is checkbox normal signature scan is not greyed out(disabled) OR
           ' minimal Overlay(0x40Bytes)
           ' this is not the first File (because the ResType -even if it's not 'FILE' - must be the same for all following files)
            If (Frm_Options.Chk_NormalSigScan.value = vbGrayed) Or _
               ((.Length - .Position <= &H40) Or _
               (FileCount > 1)) Then
Processing_Finished:
                  Log "Processing Finished!"
               ' No valid FILE Marker so seek back
                  .Move -4
                  Exit For
                        
            Else
            
               FrmMain.Txt_Scriptstart.FontBold = True
               FrmMain.Txt_Scriptstart.ForeColor = vbRed
               Dim msgboxResult_InvalidFileMaker&
               '(Please delete script start value textbox to disable.)
               msgboxResult_InvalidFileMaker = MsgBox("Invalid File Maker found - continue anyway?", vbYesNoCancel, "Manually extract mode enabled.")
               If vbYes = msgboxResult_InvalidFileMaker Then
                  ResTypeFILE = ResType
                  
               ElseIf vbNo = msgboxResult_InvalidFileMaker Then
'                  ExtractedFiles.Add File.FileName, "MainScript"
                  GoTo Processing_Finished
                  
               ElseIf vbCancel = msgboxResult_InvalidFileMaker Then
                  Err.Raise ERR_CANCEL_ALL, , "Decompilation canceled because of InvalidFileMaker"
                  
               End If
            End If
      
      End If
   
         Log "=== > Processing FILE: #" & FileCount
         FL "ResType: " & ResType
      
      
         Dim SrcFile_FileInst$
         If bIsNewScriptType Then
            SrcFile_FileInst = GetEncryptStrNew(Xorkey_SrcFile_FileInstNEW_Len, Xorkey_SrcFile_FileInstNEW_Data, File, False) 'ADBC 0B33F
         Else
            SrcFile_FileInst = GetEncryptStr(Xorkey_SrcFile_FileInst_Len, Xorkey_SrcFile_FileInst_Data, File) '0x29BC A25E
         End If
         
         FL "SrcFile_FileInst: " & SrcFile_FileInst
      
         Dim CompiledPathName As New ClsFilename
         If bIsNewScriptType Then
            CompiledPathName = GetEncryptStrNew(Xorkey_CompiledPathNameNEW_Len, Xorkey_CompiledPathNameNEW_Data, File, False) '0F820  0F479
         Else
            CompiledPathName = GetEncryptStr(Xorkey_CompiledPathName_Len, Xorkey_CompiledPathName_Data, File) '29AC  F25E
         End If
         FL "CompiledPathName: " & CompiledPathName
         
         
         Dim bIsAHK_Script As Boolean, bIsAHK_NoDeCompileScript As Boolean
         bIsAHK_Script = False: bIsAHK_NoDeCompileScript = False
         
         If SrcFile_FileInst = ">>>AUTOIT SCRIPT<<<" Then
         ElseIf SrcFile_FileInst = ">>>AUTOIT NO CMDEXECUTE<<<" Then
          'Note: Script was compiled using '#NoAutoIt3Execute' to block
         
         ElseIf SrcFile_FileInst = ">AUTOIT UNICODE SCRIPT<" Then
         ElseIf SrcFile_FileInst = ">AUTOIT SCRIPT<" Then
         
         ElseIf SrcFile_FileInst = ">AUTOHOTKEY SCRIPT<" Then
            bIsAHK_Script = True
            
         ElseIf SrcFile_FileInst = ">AHK WITH ICON<" Then
            bIsAHK_Script = True

      '; <COMPILER: v1.0.46.15> (May'07)    [previous version 1.0.46.09 March'07]
      '  you will get here when AHK was Compiled with N/A as Passphrase to prevent decompiling
      '  Ahk2Exe.exe will show: "Read: The following error occurred: FileNotFound"
      
      '  Note: AHK_ExtraDecryption is Applied after script is Decrypted and Decompressed
         ElseIf SrcFile_FileInst = ">" Then
            Log "Note: This AHK SCRIPT was compiled with 'N/A' as passphrase"
            bIsAHK_NoDeCompileScript = True
            bIsAHK_Script = True
         
         ElseIf SrcFile_FileInst = "<" Then 'like AHK WITH ICON
            Log "Note: This AHK SCRIPT(with icon) was compiled with 'N/A' as passphrase"
            bIsAHK_NoDeCompileScript = True
            bIsAHK_Script = True
         
         Else
            'If it's like this everything is as unusal
            ' CompiledPathName = "d:\ahk\compile_ahk\compile_ahk.exe" &
            ' SrcFile_FileInst = "Compile_AHK.exe"
              If 0 = InStr(1, CompiledPathName, SrcFile_FileInst, vbTextCompare) Then
                 Log Space(8 + 4) & "WARNING: unknown SrcFile_FileInst(should something like >AUTOIT SCRIPT< or >AUTOHOTKEY SCRIPT<)!"
                     If AHK_ForceNAPassword Then
                     'If vbYes = MsgBox("Do you like to force it to be an AHK-Script with 'N/A' as passphrase?", vbYesNo, "Force AHK-Script") Then
                        Log "User Forced: AHK SCRIPT compiled with 'N/A' as passphrase"
                        bIsAHK_NoDeCompileScript = True
                        bIsAHK_Script = True
                     End If

              End If
         End If
            
            
      ' ==> Is script compressed
         Dim IsCompressed&
         IsCompressed = .int8
         FL "IsCompressed: " & CBool(IsCompressed) & "  (" & H8(IsCompressed) & ")"
        
      ' ==> Get size of compressed script data
        Dim ScriptSize&
        ScriptSize = .int32
        ScriptSize = ScriptSize Xor IIf(bIsNewScriptType, 34748, 17834)      'New: 87BC | Old: 45AA
        FL "ScriptSize Compressed: " & H32(ScriptSize) & "  Decimal:" & ScriptSize & "  " & FormatSize(ScriptSize)
   
        Dim SizeUncompressed&
        SizeUncompressed = .int32 Xor IIf(bIsNewScriptType, 34748, 17834)      'New: 87BC | Old: 45AA
        FL "ScriptSize UnCompressed(used to seek to next file): " & H32(SizeUncompressed) & "  Decimal:" & SizeUncompressed & "  " & FormatSize(SizeUncompressed)
        
         If bIsOldScript = False Then
         ' ==> CRC32 value of uncompressed script data
            Dim ScriptData_CRC&
            ScriptData_CRC = .int32 Xor IIf(bIsNewScriptType, 42629, XORKey_MD5PassphraseHashText_Data)      'New: 0A685 | Old: 0C3D2

'            If &H1C00000 = (ScriptData_CRC And &HFFF00000) Then
'               log "Rewinded due to suspiciously CRC that is probably a date"
'               .Move -4
''                 bIsOldScript = True
'            End If

            FL "ADLER32 CRC of unencrypted script data: " & H32(ScriptData_CRC)
         End If
         
         If isAutoIT2Script = False Then
           Dim pCreationTime As FILETIME, pLastWrite As FILETIME
           pCreationTime.dwHighDateTime = .int32
           pCreationTime.dwLowDateTime = .int32
           pLastWrite.dwHighDateTime = .int32
           pLastWrite.dwLowDateTime = .int32
           FL "FileTime (number of 100-nanosecond intervals since January 1, 1601) "
           Log Space(4) & "pCreationTime:  " & H32(pCreationTime.dwHighDateTime) & H32(pCreationTime.dwLowDateTime) & "  " & FormatFileTime(pCreationTime)
           Log Space(4) & "pLastWrite   :  " & H32(pLastWrite.dwHighDateTime) & H32(pLastWrite.dwLowDateTime) & "  " & FormatFileTime(pLastWrite)
         End If
         
         If ScriptSize > 0 Then
         
           '==> Read encrypted script data
            FL "Begin of script data"
            
            Set ScriptData = New StringReader
            ScriptData = .FixedString(ScriptSize)
      
            ' ==> Create output fileName
            Dim OutFileName As ClsFilename
            Set OutFileName = New ClsFilename
            
            ' initialise with ScriptPath
            OutFileName = File.FileName
            
            
            'Note: AHK saves the mainscript as *.tmp
            If (CompiledPathName.Name Like "*>*") Or _
               (CompiledPathName.Ext Like "*tmp*") Or _
               (FileCount = 1) Then
               
               OutFileName.Ext = Switch(bIsAHK_Script, ".ahk", _
                                        bIsNewScriptType, ".tok", _
                                        isAutoIT2Script, ".aut", _
                                        True, ".au3")
               If IsAlreadyInCollection(ExtractedFiles, "MainScript") Then
                  OutFileName.Name = OutFileName.Name & "_" & ExtractedFiles.Count
                  ' Add extracted FileName to global ExtractedFiles List
                  ExtractedFiles.Add OutFileName
   
               Else
                ' Add extracted FileName to global ExtractedFiles List as 'MainScript'
                  ExtractedFiles.Add OutFileName, "MainScript"
               
               End If
               
   
            Else
               
               'if its an absolute path like "C:\Documents and Settings\EnCodeItInfo\Restart_EnCoded1.au3"
               'Just use the filename and don't create subdirs
               If InStr(SrcFile_FileInst, ":") Then
                  OutFileName.NameWithExt = CompiledPathName.Dir & CompiledPathName.NameWithExt
               Else
               ' Set Dir
                 If SrcFile_FileInst <> "" Then
                    OutFileName.NameWithExt = SrcFile_FileInst
                 Else
                    OutFileName.NameWithExt = OutFileName.Name & "_" & H16(FileCount) & ".tmp"
                 End If
               End If
               
               ' create Dir if it doesn't exists
               OutFileName.MakePath
               
               If IsValidFileName(OutFileName.FileName) = False Then
                  OutFileName.Name = "FileWithInvalidName_" & H16(FileCount)
                  
                  If IsValidFileName(OutFileName.FileName) = False Then
                     OutFileName = File.FileName
                     OutFileName.NameWithExt = "FileWithInvalidNameAndPath_" & H16(FileCount)
                  End If
                  
               End If
               
               
             ' Add extracted FileName to global ExtractedFiles List
               ExtractedFiles.Add OutFileName
               
            End If
                 
      ' ~~~ Saving Raw encrypted scriptdata ~~~
            Dim RawScriptFileName As New ClsFilename
            RawScriptFileName = OutFileName
            RawScriptFileName.Ext = ".raw"
            
            Dim RawScriptFile As New FileStream
            With RawScriptFile
               .Create RawScriptFileName.FileName, True, FrmMain.Chk_TmpFile.value = vbUnchecked, False
               .Data = ScriptData.Data
               .CloseFile
            End With
            
            
            
      ' ~~~ Process decrypted scriptdata ~~~
            Log "Decrypting script data..."
          
            If bIsNewScriptType Then
               RanRot_Init MD5PassphraseHash_ByteSum + Data_DecryptionKey_NewConst ' &H2477
            ElseIf isAutoIT2Script Then
               RandV2_Init MD5PassphraseHash_ByteSum + Data_DecryptionKey  ' &H22AF
            Else
               MT_Init (MD5PassphraseHash_ByteSum + Data_DecryptionKey)  ' &H22AF
            End If
               
         
            With ScriptData
              
                    
              ' ==> Decrypt scriptdata
   
   'BenchStart
               Dim StrCharPos&, tmpBuff() As Byte
               tmpBuff = StrConv(.mvardata, vbFromUnicode, LocaleID)

               'tmpBuff = ReadRawFile(RawScriptFileName.FileName)
               For StrCharPos = 0 To UBound(tmpBuff)
                  
                  
                  Dim KeyByte&
                  If bIsNewScriptType Then
                     KeyByte = RanRot_GetI8
                  ElseIf isAutoIT2Script Then
                     KeyByte = RandV2
                  Else
                     KeyByte = MT_GetI8
                  End If
                  
                  tmpBuff(StrCharPos) = tmpBuff(StrCharPos) _
                        Xor (KeyByte And &HFF)
                        
   
                  If 0 = (StrCharPos Mod &H8000) Then myDoEvents
   
                  
               Next
             ' Prevent some stupid Memory error in StrConv() if tmpBuff is empty
             ' Note empty means an array from 0 to -1; StrConv maybe handles that -1 as 0xFFFFFFFF what explains the 'Memory Error'
               If UBound(tmpBuff) > 0 Then
                  .mvardata = StrConv(tmpBuff, vbUnicode, LocaleID)
               End If

   'BenchEnd
   '            Debug.Print GetTickCount - a 'Benchmark:4453 (6171 mid version)
   'Note: This Version is 4x slower
   '            Dim Benchmark&
   '            Benchmark = GetTickCount
   
   
   '            .EOS = False
   '            .DisableAutoMove = True
   '            Do Until .EOS
   '               .int8 = .int8 Xor (MT_GetI8 And &HFF)
   '               .Move 1
   '            Loop
   
   '            Debug.Print GetTickCount - Benchmark 'Benchmark:24063
               
         
             ' Do ADLER32 CRCTest for AutoIT Scripts
               If bIsOldScript = False Then
         
                  Log "Calculating ADLER32 checksum from decrypted scriptdata"
                  
                  Dim ScriptData_CRC_Calculated&
                  ScriptData_CRC_Calculated = HexToInt(ADLER32(ScriptData))
                  If ScriptData_CRC = ScriptData_CRC_Calculated Then
                     Log "   OK."
                  Else
                     Log "   FAILED!"
                     Log "   Calculate ADLER32: " & H32(ScriptData_CRC_Calculated)
                     Log "   CRC from script  : " & H32(ScriptData_CRC)
                     
                     MsgBox "The checksum from the ExeArc_Header and" & vbCrLf & _
                              "the calculated checksum on the decrypted scriptdata differs." & vbCrLf & _
                              "Well either decryption failed or the scriptdata is corrupted." & vbCrLf & _
                               vbCrLf & _
                              "Note: Often this error is caused by a AutoIT-Exe that was compressed with Armadillon." & vbCrLf & _
                              "Armadillon just lightly 'compresses' the script so myAutToExe finds the header - but" & vbCrLf & _
                              "later the scriptdata gets 'corrupted' through this compression." & vbCrLf & _
                               vbCrLf & _
                              "To fix this error, dump the decompressed data from memory to a file." & vbCrLf & _
                              "For more details see 'readme.txt'.", vbCritical, "Warning checksum failure"
                  End If
               End If
               
         
               If IsCompressed Then
                  Uncompress OutFileName, bIsOldScript
               
                  
         
                ' Read data from new script file
                  .Data = FileLoad(OutFileName.FileName)
   
                ' Handle AHK-Scripts
                  If bIsAHK_Script Then
                     If bIsAHK_NoDeCompileScript And Not (.mvardata Like "; <COMPILER*") Then
                        Decompile_HandleAHK_ExtraDecryption SizeUncompressed
                     End If
                
                   ' Delete empty lines after "; <COMPILER: v1.0.48.2>"
                     If FrmMain.Chk_TmpFile.value = vbUnchecked Then
                        Log "Removing line breaks at the beginning..."
                        AHK_RemoveLineBreaks ScriptData
                     End If
   
                     
                     If Frm_Options.Chk_RestoreIncludes.value = vbChecked Then
                        Log "Seperating includes..."
                        AHK_SeperateIncludes ScriptData, OutFileName.Path
                        
                     End If
                                       
                     Log "Saving decrypted data to """ & OutFileName.NameWithExt & """ at " & OutFileName.Path
                     FileSave OutFileName.FileName, .Data
   
                  End If
               
               Else
               '... data was not compress, so just save the script data
                  Log "Saving script to """ & OutFileName.NameWithExt & """ at " & OutFileName.Path
               
                  FileSave OutFileName.FileName, .Data
               
               End If
               
   
               Log "Setting Creation and LastWrite time for: " & OutFileName.NameWithExt
               SetCreationNLastWriteTime OutFileName, pCreationTime, pLastWrite
               
               
             ' Show scriptdata
               If SrcFile_FileInst = ">AUTOIT UNICODE SCRIPT<" Then
                  Log "Convert from FromUnicode to Accii and write data in textbox"
                  FrmMain.Txt_Script = StrConv(.Data, vbFromUnicode, LocaleID)
               Else
                  Log "Write data in textbox"
                  FrmMain.Txt_Script = .Data
               End If
      
            End With 'ScriptData
         
         Else
            Log "Skipped processing because script size is 0 ..."
         End If
         
'        'Run Tidy on script
'         Dim tmpob As ClsFilename
'         Set tmpob = FileName
'         Set FileName = OutFileName
'            SaveScriptData Txt_Script
'         Set FileName = tmpob
         
         
         Log String(79, "-")
      
   Next
   
GoTo Finally

err_ProcessingBody:
   ErrStore
   Log "ERROR - processing body stopped: " & Err.Description
Resume Finally

Finally:
   On Error Resume Next

   If FileCount > 1 Then
   
      If .Position <> FilePointerFallBackOnError Then
         Log "User cancel processing or some unexpected error happend:"
         Log "Current FilePosition: 0x" & H32(.Position) & " FilePointerFallBackOnError: 0x" & H32(FilePointerFallBackOnError)
         Log "Notice using 'FilePointerFallBackOnError' for saving overlaydata."
         .Position = FilePointerFallBackOnError
      End If
      FL "End of script data"
      ' if there are more than 8 bytes overlay save them to *.overlay file
      ' For clearity reason I pasted overlay logging to a seperated function
      Decompile_Log_ProcessOverlay .Length - .Position, .FixedString(-1), bIsOldScript
      ' ==> Exe Processing finished
   Else
      Log "Skip saving overlay at " & H32(.Position) & " since there were no files extracted so far."
   End If
   .CloseFile
   
   Log String(79, "=")

   On Error GoTo 0
   ErrRestore
   If (Err = ERR_CANCEL_ALL) Or _
      (Err = ERR_NO_AUT_EXE) Then
      ErrThrowSimple
   End If

End With

End Sub
Private Sub Uncompress(OutFileName As ClsFilename, bIsOldScript As Boolean)
With ScriptData
     ' ==> Decompress Script
      .EOS = False
      .DisableAutoMove = False
      
      Dim LZSS_Signature$
      LZSS_Signature = .FixedString(4)
      Log "JB LZSS Signature:" & LZSS_Signature

      If LZSS_Signature = "EA04" Then
         OverWriteSignature AU3_SubTypeStr_old '"EA05"
      Else

         ' Check signature of compressed data
         Dim ExpectedSignature$
         ExpectedSignature = Switch(bIsOldScript, "JB01", _
                                    bIsNewScriptType, AU3_SubTypeStr, _
                                    isAutoIT2Script, "JB01", _
                                    True, AU3_SubTypeStr_old)
         If LZSS_Signature <> ExpectedSignature Then
         Log "WARNING: Normally signature is '" & ExpectedSignature & "' - possible reasons: 'modified' AutToExe, decryption failure, new version..."
            'If signature looks weird probably decryption fail and this is of no use

            Do
               Dim LZSS_Signature_new$
               LZSS_Signature_new = InputBox("Current value is '" & LZSS_Signature & "'" & vbCrLf & "Valid values are '" & _
                     "JB01', '" & _
                     AU3_SubTypeStr_old & "' and '" & _
                     AU3_SubTypeStr & "." & vbCrLf & "Note: If current value looks weird probably decryption fail and so data might be garbage." & vbCrLf & vbCrLf & "Since this is an Auto" & IIf(bIsOldScript, "HotKey", "IT") & " Script the recommanded value is '" & ExpectedSignature & "'" & vbCrLf & vbCrLf & "Press >OK< to change this value or" & vbCrLf & ">Cancel< to keep this it unchanged.", "Compression signature is invalid !", ExpectedSignature)
            Loop Until (Len(LZSS_Signature_new) = 4) Or (Len(LZSS_Signature_new) = 0)
            
            If (Len(LZSS_Signature_new) = 4) Then
'                  If vbYes = MsgBox("Do you want to force it to : " & ExpectedSignature & " so this stream can be decompressed?" & vbCrLf & vbCrLf & "Note: If signature looks weird probably decryption fail and this is of no use", vbYesNo + vbDefaultButton1 + vbExclamation, "LZSS_Signature of decrypted data is '" & LZSS_Signature & "'") Then
               OverWriteSignature LZSS_Signature_new
            End If
         End If
         
      End If

    ' Change AutoIT2 To "JB00" so LZSS.exe can differ between AutoIT2 and AutoHotKey
      If LZSS_Signature = "JB01" And isAutoIT2Script Then
         OverWriteSignature "JB00"
      End If

   'Translate Signature for LZSS
      Select Case LZSS_Signature
         
         Case AU3_SubTypeStr
            OverWriteSignature "EA06"
         
         Case AU3_SubTypeStr_old
            OverWriteSignature "EA05"
            
      End Select
      
'         Dim SizeUncompressed& ', w1&, w2&
'         SizeUncompressed = .int8
'         SizeUncompressed = .int8 Or (SizeUncompressed * &H100)
'         SizeUncompressed = .int8 Or (SizeUncompressed * &H100)
'         SizeUncompressed = .int8 Or (SizeUncompressed * &H100)

'         RetVal = GetUncompressedSize(.data, SizeUncompressed)
'         If RetVal <> 0 Then Err.Raise 0, , "GetUncompressedSize() failed"
'log "Uncompressed script size:" & H32(SizeUncompressed)

'
    ' save compressed script data to *.pak in current Dir
    '    if 'Create DebugFile' was not checked it will be delete on close
      Dim tmpFile As New FileStream
      With tmpFile
         .Create OutFileName.Path & OutFileName.Name & ".pak", True, FrmMain.Chk_TmpFile.value = vbUnchecked, False
         .Data = ScriptData.Data
          Log "Compressed scriptdata written to " & .FileName

         
         Dim retval&
       ' About LZSS see: http://de.wikipedia.org/wiki/Lempel-Ziv-Storer-Szymanski-Algorithmus

'         Dim tmpstr$
'         tmpstr = Space(SizeUncompressed)
'         RetVal = Uncompress(.data, .Length, tmpstr, SizeUncompressed)
        ' write decompressed Data back to stream
'         .data = tmpstr
        
      Log "Expanding script data to """ & OutFileName.NameWithExt & """ at " & OutFileName.Path
         
       ' Run "LZSS.exe -d *.debug *.au3" to extract the script (...and wait for its execution to finish)
         Dim LZSS_Output$, ExitCode&
         LZSS_Output = Console.ShellExConsole( _
                  App.Path & "\" & "data\LZSS.exe", _
                  "-d " & Quote(.FileName) & " " & Quote(OutFileName.FileName), _
                  ExitCode)
      
         If ExitCode <> 0 Then Log LZSS_Output, "LZSS_Output: "
         
'                  ShellEx App.Path & "\" & "lzss.exe", _
               "-d " & Quote(.FileName) & " " & Quote(OutFileName.FileName)
         
       ' Closes and deletes TmpFile
         .CloseFile
      End With
   End With
End Sub
Private Sub SetCreationNLastWriteTime(OutFileName As ClsFilename, pCreationTime As FILETIME, pLastWrite As FILETIME)
'   Err.Clear
   Dim outFile As New FileStream
   With outFile
      
      .Create OutFileName.FileName, False, False, False
      Dim retval&
      retval = SetFileTime(outFile.hFile, pCreationTime, 0, pLastWrite)
      If retval = 0 Then
         retval = Err.LastDllError
         Log "LastDllError: " & retval
      End If
      
      .CloseFile
   End With

End Sub


Private Sub Decompile_Log_ProcessOverlay(overlaySize&, overlaybytes$, bIsOldScript As Boolean)
   
   With File
      
Log "  FileLen: " & H32(.Length) & "  => Overlay: " & H32(overlaySize)

Dim tmp As New StringReader
tmp = Left(overlaybytes, &H20)
Log "  overlaybytes: " & ValuesToHexString(tmp) & "  " & tmp
      Dim overlaySkipBytes As Long
      overlaySkipBytes = (IIf(bIsOldScript, 3, 2) * 4)
      If overlaySize > overlaySkipBytes Then
         
         Log ">>>ATTENTION: There are more overlay data than usual <<<"
         Dim FileName_Overlay$
         FileName_Overlay = .FileName & ".overlay"
         Log "saving overlaydata to: " & FileName_Overlay
         
         FileSave FileName_Overlay, _
                  Mid(overlaybytes, overlaySkipBytes + 1) ' +1 since mid starts counting at 1
      
      End If
   
   End With

End Sub
Private Function ReadPassword() As String
   Dim PassLenXorKey&
   PassLenXorKey = XORKey_MD5PassphraseHashText_Len '&HFAC1
   
   On Error Resume Next
   ReadPassword = GetEncryptStr(PassLenXorKey, XORKey_MD5PassphraseHashText_Data, File)
    
   If Err = 0 Then Exit Function
   
'  -------- Error Handlers -------
   Log "Error occured on reading password: " & Err.Description
   File.Move -4
   Dim Password_PosStart&
   Password_PosStart = File.Position
   
 
   Log "Trying Heuristic #1 (via scriptType LenKey)"
   Dim PassLen&
   
     
   'The File format is like this:
   '...
   '  PassphraseLen           size 0x4 Bytes   [XorKey=0x000FAC1]
   '  Passphrase              size (depends on PassphraseLen)[StrKey=C3D2]
   '  ResType                 size 0x4 Byte   eString: "FILE"     [StrKey=16FA]
   '   ScriptType              eString ">AUTOIT SCRIPT<"           [LenKey=00 00 29BC,
   ' ...
   'this looks for LenKey=000029BC of ScriptType and assumes that it is unchange and that scripttype is not longer that 0xFF chars
     Const SizeOf_SearchString& = 3
     File.FindBytes &HF9, 0, 0
     
     If File.EOS = False Then
     
        Dim Password_PosEnd&
        Const SizeOf_ResType& = 4
        Const SizeOf_PassphraseLen& = 4
     
   
        Password_PosEnd = File.Position - SizeOf_SearchString - SizeOf_ResType - SizeOf_PassphraseLen
        File.Position = Password_PosStart
        
        PassLen = Password_PosEnd - Password_PosStart - 1
        GoTo PassLenFound
     End If
   
   Log "Trying Heuristic #2 (requires uncompressed script)"
 ' get actual value from script
   
   File.Position = Password_PosStart
   
   Dim PassLenXoredWithKey
   PassLenXoredWithKey = File.int32
   
   
   Log "PasswordLength xored with key is: " & H32(PassLenXoredWithKey)
   
 ' Since the PasswordLength is usually not bigger than 0x20 we can use the other
 ' 3 bytes( 0x000000XX) to maybe find and extract the correct Xorkey.
   Dim PassLenXorKey_First3DigetsAsString$
   File.Position = Password_PosStart + 1
   PassLenXorKey_First3DigetsAsString = File.FixedString(3)
   
   File.Position = Password_PosStart
   
 ' in case the script interpreter is packed, the value is not there in cleartext...

   PassLenXorKey = GetPassLenXorKey(PassLenXorKey_First3DigetsAsString)
   If Err Then
      '... we'll get here

      PassLen = 32 ' <= MD5 Hashlength
      PassLen = InputBox("Please guess how long the password is:", "Error occured on reading password", PassLen)

    Else
    '... or there's a false positive(a only 3 byte pattern is not very unique)
    ' the user may check/correct the value
      Log "Heuristically found PasswordLengthkey is: " & H32(PassLenXorKey)
      PassLen = PassLenXorKey Xor PassLenXoredWithKey
PassLenFound:
      PassLen = InputBox("Correct the password length if necessary:", "Heuristically found password length is", PassLen)
   
   End If
      
   PassLenXorKey = PassLen Xor PassLenXoredWithKey
   
   Log "User guessed the password is " & PassLen & " chars."
   
   On Error GoTo 0
   ReadPassword = GetEncryptStr(PassLenXorKey, XORKey_MD5PassphraseHashText_Data, File)

End Function

'Private Sub TestCRC()
'
'End Sub
'

'Private Sub UncompressLZSS(InData, DeComp As StringReader)
'
'
'     'BitStreamRead.data=InData
'
''     Dim DeComp As New StringReader
'     Dim BitsLeft
'     Do While BitsLeft 'BitStreamRead.BitsLeft
'        If GetBits(1) = 0 Then
'        ' literal
'           DeComp.int8 = GetBits(8)
'        Else
'        '  Tupel
'           Dim RewindBytes&, size&
'           RewindBytes = GetBits(15)
'
'         ' Handle Size
'           Dim SizePlus
'           size = GetBits(2): SizePlus = &H0
'           If size = 3 Then
'
'              size = GetBits(3): SizePlus = &H3
'              If size = 7 Then
'
'                 size = GetBits(5): SizePlus = &HA
'                 If size = &H1F Then
'
'                    size = GetBits(8): SizePlus = &H29
'                    If size = &HFF Then
'
'                       size = GetBits(8): SizePlus = &H128
'                       Do While size = &HFF
'                          size = GetBits(8): SizePlus = SizePlus + &HFF
'                       Loop
'
'                    End If
'                 End If
'              End If
'           End If
'
'         ' Duplicate/Copy String
'           DeComp.FixedString = Mid(DeComp.data, DeComp.Length - RewindBytes, size + SizePlus + 3)
'
'        End If
'
'      Loop
'
'End Sub
     
     

      
      
'Private Function GetBits(NumOfBit) As Long
'TODO       : GetBits implementation
'TODO Status: incomplete
''         Dim bits%
''         For i = 0 To bits
''            Dim CompData&
''            CompData = .int16
''            CompData = CompData * 2 'shl 1
''            Bitcount = 16
''            Bitcount = Bitcount - 1
''         Next
''         CompData = CompData \ &H10000 'shr 0x10
'
'End Function
Private Function IsAutoIT3File() As Boolean
   Dim WholeFile As New FileStream
   With WholeFile
      .Create File.FileName, False, False, True
      IsAutoIT3File = .FindString("AutoIt3") >= 0
      .CloseFile
   End With
End Function
Private Function IsAutoIT2File() As Boolean
   Dim WholeFile As New FileStream
   With WholeFile
      .Create File.FileName, False, False, True
      IsAutoIT2File = .FindString("AutoIt Main Icon") >= 0
      .CloseFile
   End With
End Function
Private Function GetPassLenXorKey(FirstDigits As String) As Long
   
   Const INT32LEN& = 4
   Debug.Assert Len(FirstDigits) < INT32LEN
   
   Dim WholeFile As New FileStream
   With WholeFile
      .Create File.FileName, False, False, True
    ' Search whole file for first three digits
      Dim items As Collection
      Set items = .FindStrings(FirstDigits)
      
    ' There must be more then 2 findings
      If items.Count < 2 Then
      '... if there are less than 2 finding
      ' the last location is the value in the script
         GetPassLenXorKey = 0
         Err.Raise vbObjectError
      Else
      '... first location (hopefully) the good one
         .Position = items(1)
      ' seek back to read the whole 4 byte(=32bit) value
         .Move Len(FirstDigits) - INT32LEN
         GetPassLenXorKey = .int32
      End If
      .CloseFile
   End With

End Function
   
Private Sub Decompile_HandleAHK_ExtraDecryption(SizeUncompressed&)
             
   On Error GoTo Decompile_HandleAHK_ExtraDecryption_err
                
 ' Just look if this is Version 1_0_48_3 or above
   Dim bIsPossiblyAboveAHK_Ver1_0_48_3
   Dim AHKStub As New StringReader
   With AHKStub
      .Data = FileReadPart(File.FileName, 0, ScriptStartPos)
      .Position = 0
'      .DisableAutoMove = False
      
      
      Dim verPos$
      verPos = .FindString("1.0.48.")
      If (verPos <> 0) Then
         Dim AHK_1_0_48_SubVer%
         AHK_1_0_48_SubVer = .FixedString(2)
         bIsPossiblyAboveAHK_Ver1_0_48_3 = (AHK_1_0_48_SubVer >= 3)
      Else
         
      End If
   End With

   
   
   Dim bIsAboveAHK_Ver1_0_48_3 As Boolean
   If FrmMain.Chk_verbose.value = vbChecked Then
      
      bIsAboveAHK_Ver1_0_48_3 = (vbYes = MsgBox( _
      "This AHK-File was compiled with Decompile Passphrase 'N/A' option. myAutToExe needs to know if that was compiled with the new AHK (= Version 1.0.48.03 and above). So is this a new AHK-File ?", _
      vbYesNo Or (vbDefaultButton2 And Not (bIsPossiblyAboveAHK_Ver1_0_48_3)), _
      "AHK-Extra Decryption"))
   Else
      bIsAboveAHK_Ver1_0_48_3 = bIsPossiblyAboveAHK_Ver1_0_48_3
      Log "bIsPossiblyAboveAHK_Ver1_0_48_3 = " & bIsPossiblyAboveAHK_Ver1_0_48_3 & ""
      Log "^- This is just a GUESS!!! Please enable verbose option be able to choose that here manually."
   End If
  ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Applied Post AHK_Sub_Key
' necessary since v1.0.47.04  aug'07   ( version before v1.0.47.00 jun'07)
' if it's "; <COMPILER: v1.0.47.00>" text is already uncrypted and so this step
' need to be skipped

    If bIsAboveAHK_Ver1_0_48_3 Then
      

       Dim AHK16_Ver_Add As Long
       Select Case AHK_1_0_48_SubVer
         Case 3
            AHK16_Ver_Add = 0

         Case 4
            AHK16_Ver_Add = InputBox("AHK v1.0.48.04 - is not known yet. But you may try to enter somekey - note: subVersion.03 has 0 and subVersion.05 has 700 as key.", , 0)
            
         Case 5
            AHK16_Ver_Add = 700
            
       End Select
       
      'init AHK_Sub_Key
       Dim AHK16_Sub_Key As Long
       AHK16_Sub_Key = (SizeUncompressed And 65535) + (AHK16_Ver_Add And 65535) '&hffff
       
       Dim AHK16_Sub_Key_Heuristic As Long
       ScriptData.Position = 0
      '"; <COMPILER: v1.0.48.5> " ->
      '"; " -> 3B 20
      '     -> 203B
       AHK16_Sub_Key_Heuristic = (ScriptData.int16 - &H203B) And &HFFFF
       
       If AHK16_Sub_Key <> AHK16_Sub_Key_Heuristic Then
         AHK16_Sub_Key = InputBox("The HeuristicAHKSub-Key is '" & AHK16_Sub_Key_Heuristic & "' and the version depending is '" & AHK16_Sub_Key & "'." & vbCrLf & _
                  "Please enter which I should use.", , AHK16_Sub_Key)
       End If

       
       
'      if SizeUncompressed =0 then AHK_Sub_Key    = &h0400
       If AHK16_Sub_Key = 0 Then AHK16_Sub_Key = &H400

                         
       Log "AHK 16bit substraction key: " & H16(AHK16_Sub_Key)
       Log "Appling AHK extra decryption(v1.0.48." & AHK_1_0_48_SubVer & ")..."
       ScriptData = AHK_ExtraDecryptionNew(ScriptData, AHK16_Sub_Key)
    
    Else


      'init AHK_Sub_Key(normal way)
       Dim AHK_Sub_Key As Byte
       AHK_Sub_Key = SizeUncompressed And 255
      
       Dim AHK_Ver_Add As Byte
'                  AHK_Ver_Add = 0    'v1.0.47.4
'                  AHK_Ver_Add = &H40 'v1.0.47.6
       AHK_Ver_Add = &H20 'v1.0.48.0..2

      
    ' Note without CInt() you get a buffer overflow (Try for ex. debug.print Cbyte(255) + Cbyte(20) )
       AHK_Sub_Key = (CInt(AHK_Sub_Key) + AHK_Ver_Add) And &HFF   '<-BugFix (That line was missing)
       If AHK_Sub_Key = 0 Then AHK_Sub_Key = &H40
      
       Log "AHK substraction key: " & H8(AHK_Sub_Key)

      
      'init AHK_Sub_Key(alternative way)
      'Alternative way to calc the XOR key
      'well this assumes that the script start like this "; <COMPILER..."
       Dim AHK_Sub_Key_Heuristic As Byte
       ScriptData.Position = 0
       AHK_Sub_Key_Heuristic = ScriptData.int8 - Asc(";") And &HFF
      
      
       If AHK_Sub_Key <> AHK_Sub_Key_Heuristic Then
          'Ask user
          FrmAHK_KeyFinder.Create ScriptData, AHK_Sub_Key_Heuristic
          FrmAHK_KeyFinder.Show vbModal
          AHK_Sub_Key = FrmAHK_KeyFinder.AHK_Key
         
         'AHK_Sub_Key = "&h" & InputBox("Hmm somehow the script is be modified." & vbCrLf & _
         "The script normal key is :" & H8(AHK_Sub_Key) & ". However the " & vbCrLf & _
         "alternative key seem to be better here. Just press enter to use it. ...or change it.", "Please enter AHK-Key", H8(AHK_Sub_Key_Heuristic))
         
          Log "AHK script stub was modified; using alterative/userdefined substraction key: " & H8(AHK_Sub_Key)
      
       End If
      
      
       Log "Appling AHK extra decryption..."
       ScriptData = AHK_ExtraDecryption(ScriptData, AHK_Sub_Key)
            
   End If '8/16bit Extra AHK_Sub_Key

Decompile_HandleAHK_ExtraDecryption_err:
End Sub
   
   
   

Private Function IsValidPEFile() As Boolean
   Dim myPEFile As New PE_info

   On Error GoTo IsValidPEFile_Err
   

       ' Store current FilePos
         Dim FilePos_old
         FilePos_old = File.Position
         myPEFile.Create

      If IsPE64 Then
         With PE_Header64
            
            Dim LastSection&
            LastSection = .NumberofSections - 1
            With .Sections(LastSection)
               PEFile_EOF_Offset = .PointertoRawData + .RawDataSize
            End With
            
         End With
      
      Else
         With PE_Header
            
            LastSection = .NumberofSections - 1
            With .Sections(LastSection)
               PEFile_EOF_Offset = .PointertoRawData + .RawDataSize
            End With
            
            PEFile_EndOfResourceData_Offset = .ResourceTableAddress + _
                                .ResourceTableAddressSize
            
            PEFile_EndOfResourceData_Offset = PE_info.RVAToRaw(PEFile_EndOfResourceData_Offset)

         End With
      End If
   
   Err.Clear
IsValidPEFile_Err:
   Select Case Err
      Case 0
         IsValidPEFile = True
         
      Case Else
'         FrmMain.Log Err.Description & " Error " & Hex(Err.Number) & "  in Modul DeCompiler.IsValidPEFile()"
         IsValidPEFile = False
   End Select
   
   File.Position = FilePos_old
   
End Function

Function IsUTF16File() As Boolean
   File.Position = 0
   
   Dim First2Byte$
   First2Byte = File.FixedString(2)
   IsUTF16File = (First2Byte = UTF16_BOM)

End Function

Function IsUTF8File() As Boolean
   File.Position = 0
   
   Dim First3Byte$
   First3Byte = File.FixedString(3)
   IsUTF8File = (First3Byte = UTF8_BOM)

End Function

Function IsTextFile() As Boolean
   Log "Testing for TextFile..."
   DoEvents
   With File
      .Create FileName.FileName, False, False, True
      
      IsTextFile = IsUTF16File
      If IsTextFile = False Then
         IsTextFile = IsUTF8File
         If IsTextFile = False Then
            
            
            .Position = 0
            
            Dim dummyLocations As Collection
            Set dummyLocations = .FindStrings(Chr(0))
            
            IsTextFile = (dummyLocations.Count = 0)
         End If
      End If
      .CloseFile
   End With
   Log "Done. (Textfile=" & IsTextFile & ")"

End Function



Sub CheckScriptFor_COMPILED_Macro()
   With File
      .Create FileName.FileName, False, False, True
      .Position = 0
      Dim FoundPos
      FoundPos = .FindString("@COMPILED", , vbTextCompare)
      If FoundPos >= 0 Then
         Log "WARNING: The '@COMPILED' was found in the script - at position: " & FoundPos & _
             " to avoid 'bad suprises' you should manually check the code at this location(and if there are more locations) before you run it."
             
       ' Show first occurence of "@COMPILED" and mark it
         If .Position > 200 Then .Move -200
         With FrmMain.Txt_Script
            .Text = File.FixedString(-1)
            .SelStart = 200
            .SelLength = 10 'Note: "@COMPILED" is 10 byte long
            .SetFocus
         End With
      End If
      .CloseFile
   End With
      
End Sub

Private Sub OverWriteSignature(LZSS_Signature_new$)
   
   With ScriptData
      
      .Move -4
      Dim SignaturePos&
      SignaturePos = ScriptData.Position
      If .FixedString(4) <> LZSS_Signature_new Then
         
         Log "Forcing/overwrite signature to '" & LZSS_Signature_new
         .Position = SignaturePos
         .FixedString(4) = LZSS_Signature_new
      
      End If
   
   End With

End Sub

Public Function AHK_ExtraDecryption(ScriptData As StringReader, ByVal AHK_Sub_Key&) As StringReader
   
   With ScriptData
   
      Dim tmpBuff() As Byte
      tmpBuff = StrConv(.mvardata, vbFromUnicode, LocaleID)
      Dim tmpByte As Byte
      
      Dim StrCharPos&
      For StrCharPos = 0 To UBound(tmpBuff)
         tmpByte = tmpBuff(StrCharPos)
         tmpByte = (tmpByte - AHK_Sub_Key) And &HFF
         tmpBuff(StrCharPos) = tmpByte
      
         If 0 = (StrCharPos Mod &H8000) Then myDoEvents
         
      Next
      
      Set AHK_ExtraDecryption = New StringReader
      AHK_ExtraDecryption.Data = StrConv(tmpBuff, vbUnicode, LocaleID)
      
      FrmMain.Txt_Script = AHK_ExtraDecryption.Data
      
   End With
End Function

Public Function AHK_ExtraDecryptionNew(ScriptData As StringReader, ByVal AHK_Sub_Key&) As StringReader
' That's how it's done in C
'      INT16 *tmpBuff;
'      Key = Size;
'      if ( !Size )
'        Key = 0x400;
'      tmpBuffSize = Size >> 1;
'      i = 0;
'      if ( tmpBuffSize )
'      {
'        Do
'          tmpBuff[i++] -= Key;
'        while ( i < tmpBuffSize );
'      }

 
   
   With ScriptData
   
      Dim tmpBuff() As Byte
      tmpBuff = StrConv(.mvardata, vbFromUnicode, LocaleID)
      
    ' Split 16bit key into low and high byte(8bit)
      Dim AHK_Sub_Key_L As Byte
      AHK_Sub_Key_L = AHK_Sub_Key And &HFF
      
      Dim AHK_Sub_Key_H As Byte
      AHK_Sub_Key_H = (AHK_Sub_Key \ &H100) And &HFF
      
      
      Dim StrCharPos&
      For StrCharPos = 0 To UBound(tmpBuff) - 1 Step 2
         
       ' Doing a subtracting of two 16-Words on byte level
       
       ' Procress lower 8 bit byte and calc carry
         Dim Byte_L As Byte
         Byte_L = tmpBuff(StrCharPos)
         
         Dim Byte_L_withCarry As Long
         Byte_L_withCarry = (CInt(Byte_L) - AHK_Sub_Key_L)
         
         Byte_L = Byte_L_withCarry And &HFF
         tmpBuff(StrCharPos) = Byte_L
         
         Dim Carry As Boolean
         Carry = (Byte_L_withCarry < 0) ' Note: false => -1;   True => 0
         
       ' Procress higher 8 bit byte and add carry
         Dim Byte_H As Byte
         Byte_H = tmpBuff(StrCharPos + 1)
         
         Byte_H = (CInt(Byte_H) - AHK_Sub_Key_H + Carry) And &HFF
         tmpBuff(StrCharPos + 1) = Byte_H
      
         If 0 = (StrCharPos Mod &H8000) Then myDoEvents
         
      Next
      
    ' convert decrypted bytearray(tmpBuff[]) back to string and display it
      Set AHK_ExtraDecryptionNew = New StringReader
      With AHK_ExtraDecryptionNew
        .Data = StrConv(tmpBuff, vbUnicode, LocaleID)
        FrmMain.Txt_Script = .Data
      End With
      
   End With
End Function



'0007F656 -> SrcFile_FileInst: >>>AUTOIT SCRIPT<<<
'0007F6B2 -> CompiledPathName: C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\aut39.tmp
'0007F6B3 -> IsCompressed: True  (01)
'44476&HADBC
'63520 '&HF820
Public Sub LongValScan_Init()

' 1. Test LongValSize; skip ">>>AUTOIT SCRIPT<<<"
' 2. Test LongValSize; skip  "C:\DOCUME~1\ADMINI~1\LOCALS~1\Temp\aut39.tmp"
' 3. Test Compressed 00 or 01
   
   
On Error GoTo LongValScanInit_err
  Set FrmMain.StartLocations = New Collection
  
  Log "Testing all possible script start locations..."
   
  Set ScriptData = New StringReader
  
' Copy filedata into String
  File.Create FrmMain.Combo_Filename, False, False, True
  File.Position = 0
  ScriptData.Data = File.FixedString(-1)
  File.CloseFile

Exit Sub
LongValScanInit_err:
   Log "ERR_LongValScanInit_err: " & Err.Description
End Sub

Public Function LongValScan(XORKEY_SrcFile_FileInstSize&, _
                            XORKEY_CompiledPathNameSize&, _
                            Optional CHARSIZE = 2) As Boolean
On Error GoTo LongValScan_err
   With ScriptData

      GUIEvent_ProcessBegin .Length

'      .DisableAutoMove = True
      .Position = 0
         
      Do
'Debug.Assert .Position <> &H7F62C
         
         Dim ScriptStartPos&
         ScriptStartPos = .Position
         
         GUIEvent_ProcessUpdate ScriptStartPos
         GUI_SkipEnable
      
            
         ' >>>AUTOIT SCRIPT<<<
         Dim SrcFile_FileInstSize&
'         SrcFile_FileInstSize = .int32 Xor 44476 ' &HADBC ('StringKey: 0x29BC_10684)
         SrcFile_FileInstSize = .int32 Xor XORKEY_SrcFile_FileInstSize ' &HADBC


     '    log_verbose "Pos: " & H32(.Position) & " - SrcFile_FileInstSize: " & H32(SrcFile_FileInstSize)

         If RangeCheck(SrcFile_FileInstSize, 19, 0) Then
            .Move SrcFile_FileInstSize * CHARSIZE
         
            Dim CompiledPathNameSize&

            CompiledPathNameSize = .int32 Xor XORKEY_CompiledPathNameSize '&HF820 ('StringKey: 0x29AC_10668)
          ' Min "C:\aut39.tmp" : Max MaxPathLen
'            log_verbose "Pos: " & H32(.Position) & " - CompiledPathNameSize: " & H32(CompiledPathNameSize)
            
            If RangeCheck(CompiledPathNameSize, 256) Then
               .Move CompiledPathNameSize * CHARSIZE
               
               Dim IsCompressed&
               IsCompressed = .int8
               If RangeCheck(IsCompressed, 1, 0) Then
                  'Found
                  '.Position = ScriptStartPos - 4 ' -4 because of 'FILE'
                  LongValScan = True
                  
                  'Exit Do
                  
                  
                  Dim Location&
                  Location = ScriptStartPos ' - _
                     Len(AU3_ResTypeFile) - _
                     Len(MD5_HASH_EMPTY_STRING) - _
                     Len(AU3_SubTypeStr) - _
                     Len(AU3_TypeStr) - _
                     AU3SigSize
                     
                  FrmMain.StartLocations.Add ScriptStartPos
                  Log "  Found #" & FrmMain.StartLocations.Count & " 0x" & H32(Location)

                  FrmMain.updateStartLocations_List
                  
               End If
               
            End If
         End If
         
         .Position = ScriptStartPos
         
         .Move 1
         
      Loop Until .EOS
      
      GUIEvent_ProcessEnd
      GUI_SkipDisable
      
'      .DisableAutoMove = False
   End With

Exit Function
LongValScan_err:

      GUIEvent_ProcessEnd
      GUI_SkipDisable
End Function



'Private Function ReadRawFile(ByVal file_name) As Variant
'
'    Dim localbyte() As Byte
'    ReDim localbyte(0 To FileLen(file_name) - 1)
'
'    Dim hFile As Integer
'    hFile = FreeFile
'
'    Open file_name For Binary As #hFile
'    Log "raw data read"
'    Get #hFile, , localbyte
'    Close hFile
'
'    ReadRawFile = localbyte
'
'End Function


Public Function FileReadPart$(FileName$, Optional Position& = 0, Optional Dst_Length& = -1)

    Dim File As New FileStream
    With File
        .Create FileName, False, False, True
        .Position = Position
        FileReadPart = .FixedString(Dst_Length)
        .CloseFile
    End With
    
End Function



'Private Sub FileCopyEx( _
'    Src_FileName$, Dst_FileName$, _
'    Optional Src_Offset& = 0, Optional Src_Length& = -1, _
'    Optional Dst_Offset& = 0, Optional Dst_Length& = -1)
'
'    Dim Src_File As New FileStream
'    With Src_File
'        .Create Src_FileName
'        .FixedString
'        .CloseFile
'
'
'    Dim Dst_File As New FileStream
'    Dst_File.Create Dst_FileName
'    Dst_File.CloseFile
'
'
'End Sub
'


Public Function IsValidFileName(FileName$) As Boolean
Attribute IsValidFileName.VB_Description = "Checks for correct FileName"
   On Error Resume Next
   
   Dim FileAlreadyThere As Boolean
   FileAlreadyThere = FileExists(FileName)
   
   Open FileName For Append As #1
   IsValidFileName = (Err = 0)
   Close #1
   
   If FileAlreadyThere = False Then
      Kill FileName
   End If
   
End Function
