Attribute VB_Name = "DeTokeniser"
Option Explicit
Public Const DETOKENISE_MAKER$ = "; DeTokenise by "
#Const LineBreak_BeforeAndAfterFunctions = False

Const AUTOIT_SourceCodeLine_MAXLEN& = 4096

Const whiteSpaceTerminal$ = " "
Const ExcludePreWhiteSpaceTerminal$ = "(["
Const ExcludePostWhiteSpaceTerminal$ = ")]."

Const TokenFile_RequiredInputExtensions = ".tok .mem"

Dim bAddWhiteSpace As Boolean


Sub DeToken()
   
   Dim bVerbose As Boolean
   
   bVerbose = FrmMain.Chk_verbose.value = vbGrayed
   With File
    
      Log "DeTokenising: " & FileName.FileName
      
      If InStr(TokenFile_RequiredInputExtensions, FileName.Ext) = 0 Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! Required FileExtension for Tokenfiles: '" & TokenFile_RequiredInputExtensions & "'" & vbCrLf & _
         "Rename this file manually to show that this should be detokenied."
      End If
      
            
'      If Frm_Options.Chk_NoDeTokenise.value = vbChecked Then
'         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! Enable DeTokenise in Options to use it." & FileName.FileName
'
'      End If
      
    ' Since that may be depend on the countrysettings...
      Dim DecimalKomma$
      Const Int64_TestValue As Currency = 1234.1234
    ' ... get it!
      DecimalKomma = Split(Int64_TestValue, "1234")(1)
      
      
      .Create FileName.FileName, False, False, True
      If .Length < 4 Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "STOPPED!!! File must be at least 4 bytes"
      End If
'   .CloseFile
'   End With
   
'   With New StringReader
'      .Data = FileLoad(FileName.FileName)
   
      
   On Error GoTo DeToken_Err
      .Position = 0
      
      Dim Lines&
      Lines = .int32
      FL "Code Lines: " & Lines & "   0x" & H32(Lines)
      
    ' File shouldn't start with MZ 00 00 -> ExeFile
    ' &HDFEFFF -> Unicodemarker
      If ((Lines And 65535) = &H5A4D) Or (Lines = &HDFEFF) Then
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "That's no Au3-TokenFile. (MZ-Exe or Dll file)"
      
      ElseIf ((Lines And &H7FFFFFF) > &H3BFEFF) Then
         'It's highly unlikly that there are more that 16 Mio lines in a Sourcefile
         Err.Raise NO_AUT_DE_TOKEN_FILE, , "This seem to be no Au3-TokenFile."
      End If
      
      
            
      
      FrmMain.List_Source.Clear
      FrmMain.List_Source.Visible = True
      
    ' ProgressBarInit
      GUIEvent_ProcessBegin Lines
   
      Dim SourceCodeLine()
      ArrayDelete SourceCodeLine
      
    ' Reset AddWhiteSpace on first item
      Dim bWasLastAnOperator As Boolean
      bWasLastAnOperator = True
      
      
      
      Dim cmd&
      Dim Size&

      Dim SourceCode ' As New Collection
      Dim SourceCodeLineCount&
      ReDim SourceCode(1 To Lines):     SourceCodeLineCount = 1:
      Dim TokenCount&: TokenCount = 0
      
      Dim RawString As StringReader: Set RawString = New StringReader
      Dim DecodeString As StringReader: Set DecodeString = New StringReader

      If bVerbose Then Frm_SrcEdit.Show
      
      Do
   
         Dim TypeName$
         TypeName = ""
   
   
         Dim Atom$
         Atom = ""
         
         If (SourceCodeLineCount > Lines) Then
            Exit Do
         End If
         
         
       ' Default
         bAddWhiteSpace = False
         
         Dim TokenOffset&
         TokenOffset = .Position
         
       ' Read Token
         cmd = .int8
         Inc TokenCount
         
         
         Dim TokenInfo$
         TokenInfo = "Token: " & H8(cmd) & "      (Line: " & SourceCodeLineCount & "  TokenCount: " & TokenCount & ")"
       ' Log it ''" & Chr(Cmd) & "'
         FL_verbose TokenInfo
'         If RangeCheck(SourceCodeLineCount, 3188, 3184) Then
'            Stop
'            If FrmMain.Chk_verbose <> vbChecked Then FrmMain.Chk_verbose = vbChecked
'         Else
'           If FrmMain.Chk_verbose <> vbUnchecked Then FrmMain.Chk_verbose = vbUnchecked
'         End If
'Debug.Assert Not (SourceCodeLine Like "*$NY*")
         
         
         Select Case cmd
         
'------- Numbers -----------
         Case &H0
            'keywords
            Dim int32_0$
            int32_0 = .int32
            TypeName = "Keyword"
            FL_verbose TypeName & ": 0x" & H32(int32_0) & "   " & int32_0

            Select Case int32_0
               Case 1
                  Atom = " AND "
               Case 2
                  Atom = " OR "
               Case 3
                  Atom = " NOT "
               Case 4
                  Atom = " IF "
               Case 5
                  Atom = " THEN "
               Case 6
                  Atom = " ELSE "
               Case 7
                  Atom = " ELSEIF "
               Case 8
                  Atom = " ENDIF "
               Case 9
                  Atom = " WHILE "
               Case 10
                  Atom = " WEND "
               Case 11
                  Atom = " DO "
               Case 12
                  Atom = " UNTIL "
               Case 13
                  Atom = " FOR "
               Case 14
                  Atom = " NEXT "
               Case 15
                  Atom = " TO "
               Case 16
                  Atom = " STEP "
               Case 17
                  Atom = " IN "
               Case 18
                  Atom = " EXITLOOP "
               Case 19
                  Atom = " CONTINUELOOP "
               Case 20
                  Atom = " SELECT "
               Case 21
                  Atom = " CASE "
               Case 22
                  Atom = " ENDSELECT "
               Case 23
                  Atom = " SWITCH "
               Case 24
                  Atom = " ENDSWITCH "
               Case 25
                  Atom = " CONTINUECASE "
               Case 26
                  Atom = " DIM "
               Case 27
                  Atom = " REDIM "
               Case 28
                  Atom = " LOCAL "
               Case 29
                  Atom = " GLOBAL "
               Case 30
                  Atom = " CONST "
               Case 31
                  Atom = " STATIC "
               Case 32
                  Atom = " FUNC "
               Case 33
                  Atom = " ENDFUNC "
               Case 34
                  Atom = " RETURN "
               Case 35
                  Atom = " EXIT "
               Case 36
                  Atom = " BYREF "
               Case 37
                  Atom = " WITH "
               Case 38
                  Atom = " ENDWITH "
               Case 39
                  Atom = " TRUE "
               Case 40
                  Atom = " FALSE "
               Case 41
                  Atom = " DEFAULT "
               Case 42
                  Atom = " NULL "
               Case Else
                  Atom = "{unknown keyword}" & H32(int32_0)
            End Select
         Case &H1
            'built-in function calls
            Dim int32_1$
            int32_1 = .int32
            TypeName = "Built-in function"
            FL_verbose TypeName & ": 0x" & H32(int32_1) & "   " & int32_1

            Select Case int32_1
               Case 1
                  Atom = "ACOS"
               Case 2
                  Atom = "ADLIBREGISTER"
               Case 3
                  Atom = "ADLIBUNREGISTER"
               Case 4
                  Atom = "ASC"
               Case 5
                  Atom = "ASCW"
               Case 6
                  Atom = "ASIN"
               Case 7
                  Atom = "ASSIGN"
               Case 8
                  Atom = "ATAN"
               Case 9
                  Atom = "AUTOITSETOPTION"
               Case 10
                  Atom = "AUTOITWINGETTITLE"
               Case 11
                  Atom = "AUTOITWINSETTITLE"
               Case 12
                  Atom = "BEEP"
               Case 13
                  Atom = "BINARY"
               Case 14
                  Atom = "BINARYLEN"
               Case 15
                  Atom = "BINARYMID"
               Case 16
                  Atom = "BINARYTOSTRING"
               Case 17
                  Atom = "BITAND"
               Case 18
                  Atom = "BITNOT"
               Case 19
                  Atom = "BITOR"
               Case 20
                  Atom = "BITROTATE"
               Case 21
                  Atom = "BITSHIFT"
               Case 22
                  Atom = "BITXOR"
               Case 23
                  Atom = "BLOCKINPUT"
               Case 24
                  Atom = "BREAK"
               Case 25
                  Atom = "CALL"
               Case 26
                  Atom = "CDTRAY"
               Case 27
                  Atom = "CEILING"
               Case 28
                  Atom = "CHR"
               Case 29
                  Atom = "CHRW"
               Case 30
                  Atom = "CLIPGET"
               Case 31
                  Atom = "CLIPPUT"
               Case 32
                  Atom = "CONSOLEREAD"
               Case 33
                  Atom = "CONSOLEWRITE"
               Case 34
                  Atom = "CONSOLEWRITEERROR"
               Case 35
                  Atom = "CONTROLCLICK"
               Case 36
                  Atom = "CONTROLCOMMAND"
               Case 37
                  Atom = "CONTROLDISABLE"
               Case 38
                  Atom = "CONTROLENABLE"
               Case 39
                  Atom = "CONTROLFOCUS"
               Case 40
                  Atom = "CONTROLGETFOCUS"
               Case 41
                  Atom = "CONTROLGETHANDLE"
               Case 42
                  Atom = "CONTROLGETPOS"
               Case 43
                  Atom = "CONTROLGETTEXT"
               Case 44
                  Atom = "CONTROLHIDE"
               Case 45
                  Atom = "CONTROLLISTVIEW"
               Case 46
                  Atom = "CONTROLMOVE"
               Case 47
                  Atom = "CONTROLSEND"
               Case 48
                  Atom = "CONTROLSETTEXT"
               Case 49
                  Atom = "CONTROLSHOW"
               Case 50
                  Atom = "CONTROLTREEVIEW"
               Case 51
                  Atom = "COS"
               Case 52
                  Atom = "DEC"
               Case 53
                  Atom = "DIRCOPY"
               Case 54
                  Atom = "DIRCREATE"
               Case 55
                  Atom = "DIRGETSIZE"
               Case 56
                  Atom = "DIRMOVE"
               Case 57
                  Atom = "DIRREMOVE"
               Case 58
                  Atom = "DLLCALL"
               Case 59
                  Atom = "DLLCALLADDRESS"
               Case 60
                  Atom = "DLLCALLBACKFREE"
               Case 61
                  Atom = "DLLCALLBACKGETPTR"
               Case 62
                  Atom = "DLLCALLBACKREGISTER"
               Case 63
                  Atom = "DLLCLOSE"
               Case 64
                  Atom = "DLLOPEN"
               Case 65
                  Atom = "DLLSTRUCTCREATE"
               Case 66
                  Atom = "DLLSTRUCTGETDATA"
               Case 67
                  Atom = "DLLSTRUCTGETPTR"
               Case 68
                  Atom = "DLLSTRUCTGETSIZE"
               Case 69
                  Atom = "DLLSTRUCTSETDATA"
               Case 70
                  Atom = "DRIVEGETDRIVE"
               Case 71
                  Atom = "DRIVEGETFILESYSTEM"
               Case 72
                  Atom = "DRIVEGETLABEL"
               Case 73
                  Atom = "DRIVEGETSERIAL"
               Case 74
                  Atom = "DRIVEGETTYPE"
               Case 75
                  Atom = "DRIVEMAPADD"
               Case 76
                  Atom = "DRIVEMAPDEL"
               Case 77
                  Atom = "DRIVEMAPGET"
               Case 78
                  Atom = "DRIVESETLABEL"
               Case 79
                  Atom = "DRIVESPACEFREE"
               Case 80
                  Atom = "DRIVESPACETOTAL"
               Case 81
                  Atom = "DRIVESTATUS"
               Case 82
                  Atom = "DUMMYSPEEDTEST"
               Case 83
                  Atom = "ENVGET"
               Case 84
                  Atom = "ENVSET"
               Case 85
                  Atom = "ENVUPDATE"
               Case 86
                  Atom = "EVAL"
               Case 87
                  Atom = "EXECUTE"
               Case 88
                  Atom = "EXP"
               Case 89
                  Atom = "FILECHANGEDIR"
               Case 90
                  Atom = "FILECLOSE"
               Case 91
                  Atom = "FILECOPY"
               Case 92
                  Atom = "FILECREATENTFSLINK"
               Case 93
                  Atom = "FILECREATESHORTCUT"
               Case 94
                  Atom = "FILEDELETE"
               Case 95
                  Atom = "FILEEXISTS"
               Case 96
                  Atom = "FILEFINDFIRSTFILE"
               Case 97
                  Atom = "FILEFINDNEXTFILE"
               Case 98
                  Atom = "FILEFLUSH"
               Case 99
                  Atom = "FILEGETATTRIB"
               Case 100
                  Atom = "FILEGETENCODING"
               Case 101
                  Atom = "FILEGETLONGNAME"
               Case 102
                  Atom = "FILEGETPOS"
               Case 103
                  Atom = "FILEGETSHORTCUT"
               Case 104
                  Atom = "FILEGETSHORTNAME"
               Case 105
                  Atom = "FILEGETSIZE"
               Case 106
                  Atom = "FILEGETTIME"
               Case 107
                  Atom = "FILEGETVERSION"
               Case 108
                  Atom = "FILEINSTALL"
               Case 109
                  Atom = "FILEMOVE"
               Case 110
                  Atom = "FILEOPEN"
               Case 111
                  Atom = "FILEOPENDIALOG"
               Case 112
                  Atom = "FILEREAD"
               Case 113
                  Atom = "FILEREADLINE"
               Case 114
                  Atom = "FILEREADTOARRAY"
               Case 115
                  Atom = "FILERECYCLE"
               Case 116
                  Atom = "FILERECYCLEEMPTY"
               Case 117
                  Atom = "FILESAVEDIALOG"
               Case 118
                  Atom = "FILESELECTFOLDER"
               Case 119
                  Atom = "FILESETATTRIB"
               Case 120
                  Atom = "FILESETEND"
               Case 121
                  Atom = "FILESETPOS"
               Case 122
                  Atom = "FILESETTIME"
               Case 123
                  Atom = "FILEWRITE"
               Case 124
                  Atom = "FILEWRITELINE"
               Case 125
                  Atom = "FLOOR"
               Case 126
                  Atom = "FTPSETPROXY"
               Case 127
                  Atom = "FUNCNAME"
               Case 128
                  Atom = "GUICREATE"
               Case 129
                  Atom = "GUICTRLCREATEAVI"
               Case 130
                  Atom = "GUICTRLCREATEBUTTON"
               Case 131
                  Atom = "GUICTRLCREATECHECKBOX"
               Case 132
                  Atom = "GUICTRLCREATECOMBO"
               Case 133
                  Atom = "GUICTRLCREATECONTEXTMENU"
               Case 134
                  Atom = "GUICTRLCREATEDATE"
               Case 135
                  Atom = "GUICTRLCREATEDUMMY"
               Case 136
                  Atom = "GUICTRLCREATEEDIT"
               Case 137
                  Atom = "GUICTRLCREATEGRAPHIC"
               Case 138
                  Atom = "GUICTRLCREATEGROUP"
               Case 139
                  Atom = "GUICTRLCREATEICON"
               Case 140
                  Atom = "GUICTRLCREATEINPUT"
               Case 141
                  Atom = "GUICTRLCREATELABEL"
               Case 142
                  Atom = "GUICTRLCREATELIST"
               Case 143
                  Atom = "GUICTRLCREATELISTVIEW"
               Case 144
                  Atom = "GUICTRLCREATELISTVIEWITEM"
               Case 145
                  Atom = "GUICTRLCREATEMENU"
               Case 146
                  Atom = "GUICTRLCREATEMENUITEM"
               Case 147
                  Atom = "GUICTRLCREATEMONTHCAL"
               Case 148
                  Atom = "GUICTRLCREATEOBJ"
               Case 149
                  Atom = "GUICTRLCREATEPIC"
               Case 150
                  Atom = "GUICTRLCREATEPROGRESS"
               Case 151
                  Atom = "GUICTRLCREATERADIO"
               Case 152
                  Atom = "GUICTRLCREATESLIDER"
               Case 153
                  Atom = "GUICTRLCREATETAB"
               Case 154
                  Atom = "GUICTRLCREATETABITEM"
               Case 155
                  Atom = "GUICTRLCREATETREEVIEW"
               Case 156
                  Atom = "GUICTRLCREATETREEVIEWITEM"
               Case 157
                  Atom = "GUICTRLCREATEUPDOWN"
               Case 158
                  Atom = "GUICTRLDELETE"
               Case 159
                  Atom = "GUICTRLGETHANDLE"
               Case 160
                  Atom = "GUICTRLGETSTATE"
               Case 161
                  Atom = "GUICTRLREAD"
               Case 162
                  Atom = "GUICTRLRECVMSG"
               Case 163
                  Atom = "GUICTRLREGISTERLISTVIEWSORT"
               Case 164
                  Atom = "GUICTRLSENDMSG"
               Case 165
                  Atom = "GUICTRLSENDTODUMMY"
               Case 166
                  Atom = "GUICTRLSETBKCOLOR"
               Case 167
                  Atom = "GUICTRLSETCOLOR"
               Case 168
                  Atom = "GUICTRLSETCURSOR"
               Case 169
                  Atom = "GUICTRLSETDATA"
               Case 170
                  Atom = "GUICTRLSETDEFBKCOLOR"
               Case 171
                  Atom = "GUICTRLSETDEFCOLOR"
               Case 172
                  Atom = "GUICTRLSETFONT"
               Case 173
                  Atom = "GUICTRLSETGRAPHIC"
               Case 174
                  Atom = "GUICTRLSETIMAGE"
               Case 175
                  Atom = "GUICTRLSETLIMIT"
               Case 176
                  Atom = "GUICTRLSETONEVENT"
               Case 177
                  Atom = "GUICTRLSETPOS"
               Case 178
                  Atom = "GUICTRLSETRESIZING"
               Case 179
                  Atom = "GUICTRLSETSTATE"
               Case 180
                  Atom = "GUICTRLSETSTYLE"
               Case 181
                  Atom = "GUICTRLSETTIP"
               Case 182
                  Atom = "GUIDELETE"
               Case 183
                  Atom = "GUIGETCURSORINFO"
               Case 184
                  Atom = "GUIGETMSG"
               Case 185
                  Atom = "GUIGETSTYLE"
               Case 186
                  Atom = "GUIREGISTERMSG"
               Case 187
                  Atom = "GUISETACCELERATORS"
               Case 188
                  Atom = "GUISETBKCOLOR"
               Case 189
                  Atom = "GUISETCOORD"
               Case 190
                  Atom = "GUISETCURSOR"
               Case 191
                  Atom = "GUISETFONT"
               Case 192
                  Atom = "GUISETHELP"
               Case 193
                  Atom = "GUISETICON"
               Case 194
                  Atom = "GUISETONEVENT"
               Case 195
                  Atom = "GUISETSTATE"
               Case 196
                  Atom = "GUISETSTYLE"
               Case 197
                  Atom = "GUISTARTGROUP"
               Case 198
                  Atom = "GUISWITCH"
               Case 199
                  Atom = "HEX"
               Case 200
                  Atom = "HOTKEYSET"
               Case 201
                  Atom = "HTTPSETPROXY"
               Case 202
                  Atom = "HTTPSETUSERAGENT"
               Case 203
                  Atom = "HWND"
               Case 204
                  Atom = "INETCLOSE"
               Case 205
                  Atom = "INETGET"
               Case 206
                  Atom = "INETGETINFO"
               Case 207
                  Atom = "INETGETSIZE"
               Case 208
                  Atom = "INETREAD"
               Case 209
                  Atom = "INIDELETE"
               Case 210
                  Atom = "INIREAD"
               Case 211
                  Atom = "INIREADSECTION"
               Case 212
                  Atom = "INIREADSECTIONNAMES"
               Case 213
                  Atom = "INIRENAMESECTION"
               Case 214
                  Atom = "INIWRITE"
               Case 215
                  Atom = "INIWRITESECTION"
               Case 216
                  Atom = "INPUTBOX"
               Case 217
                  Atom = "INT"
               Case 218
                  Atom = "ISADMIN"
               Case 219
                  Atom = "ISARRAY"
               Case 220
                  Atom = "ISBINARY"
               Case 221
                  Atom = "ISBOOL"
               Case 222
                  Atom = "ISDECLARED"
               Case 223
                  Atom = "ISDLLSTRUCT"
               Case 224
                  Atom = "ISFLOAT"
               Case 225
                  Atom = "ISFUNC"
               Case 226
                  Atom = "ISHWND"
               Case 227
                  Atom = "ISINT"
               Case 228
                  Atom = "ISKEYWORD"
               Case 229
                  Atom = "ISMAP"
               Case 230
                  Atom = "ISNUMBER"
               Case 231
                  Atom = "ISOBJ"
               Case 232
                  Atom = "ISPTR"
               Case 233
                  Atom = "ISSTRING"
               Case 234
                  Atom = "LOG"
               Case 235
                  Atom = "MAPAPPEND"
               Case 236
                  Atom = "MAPEXISTS"
               Case 237
                  Atom = "MAPKEYS"
               Case 238
                  Atom = "MAPREMOVE"
               Case 239
                  Atom = "MEMGETSTATS"
               Case 240
                  Atom = "MOD"
               Case 241
                  Atom = "MOUSECLICK"
               Case 242
                  Atom = "MOUSECLICKDRAG"
               Case 243
                  Atom = "MOUSEDOWN"
               Case 244
                  Atom = "MOUSEGETCURSOR"
               Case 245
                  Atom = "MOUSEGETPOS"
               Case 246
                  Atom = "MOUSEMOVE"
               Case 247
                  Atom = "MOUSEUP"
               Case 248
                  Atom = "MOUSEWHEEL"
               Case 249
                  Atom = "MSGBOX"
               Case 250
                  Atom = "NUMBER"
               Case 251
                  Atom = "OBJCREATE"
               Case 252
                  Atom = "OBJCREATEINTERFACE"
               Case 253
                  Atom = "OBJEVENT"
               Case 254
                  Atom = "OBJGET"
               Case 255
                  Atom = "OBJNAME"
               Case 256
                  Atom = "ONAUTOITEXITREGISTER"
               Case 257
                  Atom = "ONAUTOITEXITUNREGISTER"
               Case 258
                  Atom = "OPT"
               Case 259
                  Atom = "PING"
               Case 260
                  Atom = "PIXELCHECKSUM"
               Case 261
                  Atom = "PIXELGETCOLOR"
               Case 262
                  Atom = "PIXELSEARCH"
               Case 263
                  Atom = "PROCESSCLOSE"
               Case 264
                  Atom = "PROCESSEXISTS"
               Case 265
                  Atom = "PROCESSGETSTATS"
               Case 266
                  Atom = "PROCESSLIST"
               Case 267
                  Atom = "PROCESSSETPRIORITY"
               Case 268
                  Atom = "PROCESSWAIT"
               Case 269
                  Atom = "PROCESSWAITCLOSE"
               Case 270
                  Atom = "PROGRESSOFF"
               Case 271
                  Atom = "PROGRESSON"
               Case 272
                  Atom = "PROGRESSSET"
               Case 273
                  Atom = "PTR"
               Case 274
                  Atom = "RANDOM"
               Case 275
                  Atom = "REGDELETE"
               Case 276
                  Atom = "REGENUMKEY"
               Case 277
                  Atom = "REGENUMVAL"
               Case 278
                  Atom = "REGREAD"
               Case 279
                  Atom = "REGWRITE"
               Case 280
                  Atom = "ROUND"
               Case 281
                  Atom = "RUN"
               Case 282
                  Atom = "RUNAS"
               Case 283
                  Atom = "RUNASWAIT"
               Case 284
                  Atom = "RUNWAIT"
               Case 285
                  Atom = "SEND"
               Case 286
                  Atom = "SENDKEEPACTIVE"
               Case 287
                  Atom = "SETERROR"
               Case 288
                  Atom = "SETEXTENDED"
               Case 289
                  Atom = "SHELLEXECUTE"
               Case 290
                  Atom = "SHELLEXECUTEWAIT"
               Case 291
                  Atom = "SHUTDOWN"
               Case 292
                  Atom = "SIN"
               Case 293
                  Atom = "SLEEP"
               Case 294
                  Atom = "SOUNDPLAY"
               Case 295
                  Atom = "SOUNDSETWAVEVOLUME"
               Case 296
                  Atom = "SPLASHIMAGEON"
               Case 297
                  Atom = "SPLASHOFF"
               Case 298
                  Atom = "SPLASHTEXTON"
               Case 299
                  Atom = "SQRT"
               Case 300
                  Atom = "SRANDOM"
               Case 301
                  Atom = "STATUSBARGETTEXT"
               Case 302
                  Atom = "STDERRREAD"
               Case 303
                  Atom = "STDINWRITE"
               Case 304
                  Atom = "STDIOCLOSE"
               Case 305
                  Atom = "STDOUTREAD"
               Case 306
                  Atom = "STRING"
               Case 307
                  Atom = "STRINGADDCR"
               Case 308
                  Atom = "STRINGCOMPARE"
               Case 309
                  Atom = "STRINGFORMAT"
               Case 310
                  Atom = "STRINGFROMASCIIARRAY"
               Case 311
                  Atom = "STRINGINSTR"
               Case 312
                  Atom = "STRINGISALNUM"
               Case 313
                  Atom = "STRINGISALPHA"
               Case 314
                  Atom = "STRINGISASCII"
               Case 315
                  Atom = "STRINGISDIGIT"
               Case 316
                  Atom = "STRINGISFLOAT"
               Case 317
                  Atom = "STRINGISINT"
               Case 318
                  Atom = "STRINGISLOWER"
               Case 319
                  Atom = "STRINGISSPACE"
               Case 320
                  Atom = "STRINGISUPPER"
               Case 321
                  Atom = "STRINGISXDIGIT"
               Case 322
                  Atom = "STRINGLEFT"
               Case 323
                  Atom = "STRINGLEN"
               Case 324
                  Atom = "STRINGLOWER"
               Case 325
                  Atom = "STRINGMID"
               Case 326
                  Atom = "STRINGREGEXP"
               Case 327
                  Atom = "STRINGREGEXPREPLACE"
               Case 328
                  Atom = "STRINGREPLACE"
               Case 329
                  Atom = "STRINGREVERSE"
               Case 330
                  Atom = "STRINGRIGHT"
               Case 331
                  Atom = "STRINGSPLIT"
               Case 332
                  Atom = "STRINGSTRIPCR"
               Case 333
                  Atom = "STRINGSTRIPWS"
               Case 334
                  Atom = "STRINGTOASCIIARRAY"
               Case 335
                  Atom = "STRINGTOBINARY"
               Case 336
                  Atom = "STRINGTRIMLEFT"
               Case 337
                  Atom = "STRINGTRIMRIGHT"
               Case 338
                  Atom = "STRINGUPPER"
               Case 339
                  Atom = "TAN"
               Case 340
                  Atom = "TCPACCEPT"
               Case 341
                  Atom = "TCPCLOSESOCKET"
               Case 342
                  Atom = "TCPCONNECT"
               Case 343
                  Atom = "TCPLISTEN"
               Case 344
                  Atom = "TCPNAMETOIP"
               Case 345
                  Atom = "TCPRECV"
               Case 346
                  Atom = "TCPSEND"
               Case 347
                  Atom = "TCPSHUTDOWN"
               Case 348
                  Atom = "TCPSTARTUP"
               Case 349
                  Atom = "TIMERDIFF"
               Case 350
                  Atom = "TIMERINIT"
               Case 351
                  Atom = "TOOLTIP"
               Case 352
                  Atom = "TRAYCREATEITEM"
               Case 353
                  Atom = "TRAYCREATEMENU"
               Case 354
                  Atom = "TRAYGETMSG"
               Case 355
                  Atom = "TRAYITEMDELETE"
               Case 356
                  Atom = "TRAYITEMGETHANDLE"
               Case 357
                  Atom = "TRAYITEMGETSTATE"
               Case 358
                  Atom = "TRAYITEMGETTEXT"
               Case 359
                  Atom = "TRAYITEMSETONEVENT"
               Case 360
                  Atom = "TRAYITEMSETSTATE"
               Case 361
                  Atom = "TRAYITEMSETTEXT"
               Case 362
                  Atom = "TRAYSETCLICK"
               Case 363
                  Atom = "TRAYSETICON"
               Case 364
                  Atom = "TRAYSETONEVENT"
               Case 365
                  Atom = "TRAYSETPAUSEICON"
               Case 366
                  Atom = "TRAYSETSTATE"
               Case 367
                  Atom = "TRAYSETTOOLTIP"
               Case 368
                  Atom = "TRAYTIP"
               Case 369
                  Atom = "UBOUND"
               Case 370
                  Atom = "UDPBIND"
               Case 371
                  Atom = "UDPCLOSESOCKET"
               Case 372
                  Atom = "UDPOPEN"
               Case 373
                  Atom = "UDPRECV"
               Case 374
                  Atom = "UDPSEND"
               Case 375
                  Atom = "UDPSHUTDOWN"
               Case 376
                  Atom = "UDPSTARTUP"
               Case 377
                  Atom = "VARGETTYPE"
               Case 378
                  Atom = "WINACTIVATE"
               Case 379
                  Atom = "WINACTIVE"
               Case 380
                  Atom = "WINCLOSE"
               Case 381
                  Atom = "WINEXISTS"
               Case 382
                  Atom = "WINFLASH"
               Case 383
                  Atom = "WINGETCARETPOS"
               Case 384
                  Atom = "WINGETCLASSLIST"
               Case 385
                  Atom = "WINGETCLIENTSIZE"
               Case 386
                  Atom = "WINGETHANDLE"
               Case 387
                  Atom = "WINGETPOS"
               Case 388
                  Atom = "WINGETPROCESS"
               Case 389
                  Atom = "WINGETSTATE"
               Case 390
                  Atom = "WINGETTEXT"
               Case 391
                  Atom = "WINGETTITLE"
               Case 392
                  Atom = "WINKILL"
               Case 393
                  Atom = "WINLIST"
               Case 394
                  Atom = "WINMENUSELECTITEM"
               Case 395
                  Atom = "WINMINIMIZEALL"
               Case 396
                  Atom = "WINMINIMIZEALLUNDO"
               Case 397
                  Atom = "WINMOVE"
               Case 398
                  Atom = "WINSETONTOP"
               Case 399
                  Atom = "WINSETSTATE"
               Case 400
                  Atom = "WINSETTITLE"
               Case 401
                  Atom = "WINSETTRANS"
               Case 402
                  Atom = "WINWAIT"
               Case 403
                  Atom = "WINWAITACTIVE"
               Case 404
                  Atom = "WINWAITCLOSE"
               Case 405
                  Atom = "WINWAITNOTACTIVE"
               Case Else
                  Atom = "{unknown built-in function}" & H32(int32_1)
            End Select
         Case &H2 To &HF
            '&H5
            Dim int32$
            int32 = .int32
            Atom = int32
            
          ' Bugfix for 3.3.8.1 (29th January, 2012)
          ' Tokenoptimisation occure'+-123' -> '-123'
            Dim LastAtom
            If LastAtom = "+" Then
               If Atom <= -1 Then
                  Log " Tokenoptimisation occured '+-' -> '-'  @line: " & SourceCodeLineCount
                  Dim tmp$
                  tmp = ArrayGetLast(SourceCodeLine)
                  tmp = Left2(tmp) ' Cut last char
                  ArraySetLast SourceCodeLine, tmp
               End If
            End If
                         

            
            TypeName = "Int32"
            FL_verbose TypeName & ": 0x" & H32(int32) & "   " & int32
            
          ' So far this value has always been 5
            Debug.Assert cmd = 5
         Case &H10 To &H1F
            Dim Int64 As Currency
            Int64 = .int64Value
            'int64 = H32(.int32)
            'int64 = H32(.int32) & int64
            'Replace 123,45 -> 12345
            Atom = Replace(CStr(Int64), DecimalKomma, "")
            TypeName = "Int64"
            FL_verbose TypeName & ": " & Int64
            
            Debug.Assert cmd = &H10
         
         Case &H20 To &H2F
           'Get DoubleValue
            Dim Double_$
            Double_ = .DoubleValue
            'Replace 123,11 -> 123.11
            Atom = Replace(CStr(Double_), DecimalKomma, ".")
            
            TypeName = "64Bit-float"
            FL_verbose TypeName & ": " & Double_
         
            Debug.Assert cmd = &H20
         

'------- Strings -----------
         Case &H30 To &H3F 'Keywords
            
           'Get StrLength and load it
            Size = .int32
            FL_verbose "StringSize: " & H32(Size)
            
            If Size > (.Length - .Position) Then
               Err.Raise vbObjectError, , "Invalid string size(bigger than the file)!"
            End If

            RawString = .FixedStringW(Size)
           
           'XorDecode String
            Dim pos&, XorKey_l As Byte, XorKey_h As Byte
            
            XorKey_l = (Size And &HFF)
            XorKey_h = ((Size \ &H100) And &HFF) ' 2^8 = 256
            
            Dim tmpBuff() As Byte
            tmpBuff = RawString
            
            For pos = LBound(tmpBuff) To UBound(tmpBuff) Step 2
               tmpBuff(pos) = tmpBuff(pos) Xor XorKey_l
               tmpBuff(pos + 1) = tmpBuff(pos + 1) Xor XorKey_h
'               DecodeString = tmpBuff
               
               'If 0 = (pos Mod &H8000) Then myDoEvents
            Next
            
            DecodeString = tmpBuff
            
            
'Comment out due to bad performance
'            RawString.Position = 0
'            DecodeString = Space(RawString.Length \ 2)
'            Do Until RawString.EOS
'               DecodeString.int8 = RawString.int8 Xor Size
'               If Not (RawString.EOS) Then Debug.Assert RawString.int8 = 0
'            Loop
            
            
'------- Commands -----------
            Select Case cmd
            
            Case &H30 'BlockElement (FUNC, IF...) and the Rest of 42 Elements: "AND OR NOT IF THEN ELSE ELSEIF ENDIF WHILE WEND DO UNTIL FOR NEXT TO STEP IN EXITLOOP CONTINUELOOP SELECT CASE ENDSELECT SWITCH ENDSWITCH CONTINUECASE DIM REDIM LOCAL GLOBAL CONST FUNC ENDFUNC RETURN EXIT BYREF WITH ENDWITH TRUE FALSE DEFAULT ENUM NULL"
               TypeName = "BlockElement"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
               
               Atom = DecodeString
               bAddWhiteSpace = True
              
               #If LineBreak_BeforeAndAfterFunctions Then
                  If Atom = "ENDFUNC" Then
                     Atom = Atom & vbCrLf
                  ElseIf Atom = "FUNC" Then
                     Atom = vbCrLf & Atom
                  End If
               #End If

            
            Case &H31 'FunctionCall with params
               Atom = DecodeString
               
               TypeName = "AutoItFunction"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
               
            Case &H32 'Macro
               Atom = "@" & DecodeString
               
               TypeName = "Macro"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H33 'Variable
               Atom = "$" & DecodeString
               
               TypeName = "Variable"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H34 'FunctionCall
               Atom = DecodeString
               
               TypeName = "UserFunction"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H35 'Property
               Atom = "." & DecodeString
               
               TypeName = "Property"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H36 'UserString
               
               Atom = MakeAutoItString(DecodeString.Data)
               
               TypeName = "UserString"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            Case &H37 '# PreProcessor
               Atom = DecodeString
               bAddWhiteSpace = True
               
               TypeName = "PreProcessor"
               FL_verbose """" & DecodeString.Data & """   Type: " & TypeName
            
            
            Case Else
               'Unknown StringToken
               If HandleTokenErr("ERROR: Unknown StringToken") Then
               Else
                  Err.Raise vbObjectError Or 1, , "Unknown StringToken"
                  Stop

               End If
               
               
            End Select
            
 '           log String(40, "_")
         
'------- Operators -----------
         Case &H40 To &H58
'            Atom = Choose((Cmd - &H40 + 1), ",", "=", ">", "<", "<>", ">=", "<=", "(", ")", "+", "-", "/", "", "&", "[", "]", "==", "^", "+=", "-=", "/=", "*=", "&=")
         '                     Au3Manual AcciChar
            
            Select Case cmd
               Case &H40: Atom = ","  '        2C
               Case &H41: Atom = "="  ' 1  13  3D
               Case &H42: Atom = ">"  ' 16     3E
               Case &H43: Atom = "<"  ' 18     3C
               Case &H44: Atom = "<>" ' 15     3C
               Case &H45: Atom = ">=" ' 17     3E
               Case &H46: Atom = "<=" ' 19     3C
               Case &H47: Atom = "("  '        28
               Case &H48: Atom = ")"  '        29
               Case &H49: Atom = "+": ' 7      2B
               Case &H4A: Atom = "-": ' 8      2D
               Case &H4B: Atom = "/"  ' 10     2F
               Case &H4C: Atom = "*": ' 9      2A
               Case &H4D: Atom = "&"  ' 11     26
               Case &H4E: Atom = "["  '        5B
               Case &H4F: Atom = "]"  '        5D
               Case &H50: Atom = "==" ' 14     3D
               Case &H51: Atom = "^"  ' 12     5E
               Case &H52: Atom = "+=" '2       2B
               Case &H53: Atom = "-=" '3       2D
               Case &H54: Atom = "/=" '5       2F
               Case &H55: Atom = "*=" '4       2A
               Case &H56: Atom = "&=" '6       26
               Case &H57: Atom = "?"
               Case &H58: Atom = ":"
            End Select
            TypeName = "operator"
            FL_verbose """" & Atom & """   Type: " & TypeName
'------- EOL -----------
         Case &H7F
          ' Execute
            
            
            Dim SourceCodeLineFinal$
            SourceCodeLineFinal = Join(SourceCodeLine, whiteSpaceTerminal)
            
            LogSourceCodeLine SourceCodeLineFinal
            
            
            log_verbose ">>>  " & SourceCodeLineFinal
            log_verbose String(80, "_")
            log_verbose ""
 
          ' Test Length
            Dim SourceCodeLine_Len&
            SourceCodeLine_Len = Len(SourceCodeLineFinal)
            If SourceCodeLine_Len >= AUTOIT_SourceCodeLine_MAXLEN Then
               Log "WARNING: SourceCodeLine: " & SourceCodeLineCount & " is " & _
               SourceCodeLine_Len - AUTOIT_SourceCodeLine_MAXLEN & " chars longer than " & _
               AUTOIT_SourceCodeLine_MAXLEN & " - Please remove some spaces manually to make it shorter."
            End If

          ' Processbar update
            GUIEvent_ProcessUpdate SourceCodeLineCount
          
          ' Add SourceCodeLine to SourceCode
            SourceCode(SourceCodeLineCount) = SourceCodeLineFinal
            Inc SourceCodeLineCount
            
          ' del SourceCodeLine
            ArrayDelete SourceCodeLine
            If bVerbose Then Frm_SrcEdit.LineBreak
           
          ' Reset AddWhiteSpace on next item
            bWasLastAnOperator = True
            DelayedReturn False
           

         Case Else
            
           'Unknown Token
            Log "Unknown Token_Command: 0x" & H8(cmd) & " @ " & H32(TokenOffset)
            If HandleTokenErr("ERROR: Unknown Token") Then
            Else
               Err.Raise NO_AUT_DE_TOKEN_FILE, , "Unknown Token"
               'Exit Do
            End If
           'qw
'           Stop
           

         End Select
         
'         Debug.Assert SourceCodeLineCount <> 851

         
         If cmd <> &H7F Then
            
           
          ' Add to SourceLine
            ' Always add a whiteSpace after a command (and preprocessor)
            '    and add a whiteSpace before; except the token before is an operator (Like: [] () = ...)
            If DelayedReturn(bAddWhiteSpace) Or _
               (bAddWhiteSpace And Not (bWasLastAnOperator)) Then
             
             ' Add with whitespace
               ArrayAdd SourceCodeLine, Atom
               If bVerbose Then Frm_SrcEdit.AddItem whiteSpaceTerminal & Atom, cmd, TypeName, TokenInfo & " @ 0x" & H32(TokenOffset)
            Else
              'Append to Last
               
               ArrayAppendLast SourceCodeLine, Atom
               If bVerbose Then Frm_SrcEdit.AddItem Atom, cmd, TypeName, TokenInfo & " @ 0x" & H32(TokenOffset)

            End If
            DoEventsVerySeldom
            
            bWasLastAnOperator = RangeCheck(cmd, &H56, &H40)
'         Else
            
            
         End If
         LastAtom = Atom

      Loop Until .EOS
    
    
    
Err.Clear
DeToken_Err:
Select Case Err
   Case 0
   Case ERR_CANCEL_ALL
      ErrThrowSimple
   
   Case Else
     
     Dim ErrSourceCodeLine$
     ErrSourceCodeLine = Join(SourceCodeLine, whiteSpaceTerminal)
     
     Dim ErrText$
     ErrText = "ERROR: " & Err.Description & vbCrLf & _
      "FileOffset: " & H32(.Position) & vbCrLf & _
      " when de-tokenising script line: " & SourceCodeLineCount & vbCrLf & ErrSourceCodeLine
     Log ErrText
     MsgBox ErrText, vbCritical, "Unexpected Error during detokenising"
     
    'Set incomplete SourceCodeLine
     SourceCode(SourceCodeLineCount) = ErrSourceCodeLine & " <- " & ErrText
     Inc SourceCodeLineCount

    'Cut down SourceCodeArray to Error
     ReDim Preserve SourceCode(SourceCodeLineCount)
     
     Resume DeToken_Finally
End Select

  
  If FrmMain.Chk_TmpFile = vbUnchecked Then
     Log "Keep TmpFile is unchecked => Deleting '" & FileName.NameWithExt & "'"
     FileDelete (FileName)
  End If


DeToken_Finally:
   File.CloseFile
  End With
    
' ProgressBar Finish
  GUIEvent_ProcessEnd
  
  FileName.Ext = ".au3"
  
  
'   If bUnicodeEnable Then
      Dim ScriptData$
      ScriptData = Join(SourceCode, vbCrLf) & vbCrLf & _
                  DETOKENISE_MAKER & FrmMain.Caption & vbCrLf

'      Dim FileName_UTF16 As New ClsFilename
'      FileName_UTF16.FileName = FileName.FileName
'
'      FileName_UTF16.Name = FileName.Name & "_UTF16"
'      FrmMain.Log "Saving UTF16-Script to: " & FileName_UTF16.FileName
'
'      File.Create FileName_UTF16.FileName, True, False, False
'      File.Position = 0
'      File.FixedString(-1) = UTF16_BOM & ScriptData
'      File.setEOF
'      File.CloseFile
'
'   End If
  
  FrmMain.Log "Converting Unicode to UTF8, since Tidy don't support unicode."
  SaveScriptData UTF8_BOM & EncodeUTF8(ScriptData), True
   
  Log "Token expansion succeed."
   
  FrmMain.List_Source.Visible = False

End Sub

Private Function HandleTokenErr(ErrText$) As Boolean

   With File
   
      If vbYes = MsgBox("An Token error occured - possible due to corrupted scriptdata. Contiune?", vbCritical + vbYesNo, ErrText) Then
         HandleTokenErr = True
         
'         Dim Hexdata As New clsStrCat, HexdataLine&
'         Hexdata.Clear
'         For HexdataLine = 0 To &H100 Step &H8
'            Dim Data As New StringReader
'            Data = .FixedString(&H8)
'            Hexdata.Concat H16(HexdataLine) & ":  " & ValuesToHexString(Data) & vbCrLf
'
'         Next
'         .Move -&H100
'         Stop
'         .Move InputBox("The this is the following raw Token data: " & Hexdata.value & "How many bytes should I skip?", "Skip Tokenbytes", "0")
         
      Else
         HandleTokenErr = False
      
      End If
      
   End With
End Function

Private Sub LogSourceCodeLine(TextLine$)
   FrmMain.LogSourceCodeLine TextLine$
End Sub
'Handle UserString with Quotes...
Function MakeAutoItString$(RawString$)
             
   ' HasDoubleQuote ?
     If InStr(RawString, """") <> 0 Then
        
      ' HasSingleQuote ?
        If InStr(RawString, "'") <> 0 Then
         ' Scenario3: " This is a 'Example' on correct "Quoting" String "
           MakeAutoItString = """" & Replace(RawString, """", """""") & """"
        Else
         ' Scenario2: " This is a "Example". "
           MakeAutoItString = "'" & RawString & "'"
        End If
     Else
      ' ' Scenario1: " ExampleString "
        MakeAutoItString = """" & RawString & """"
     End If
     

End Function

' Converts an AutoIt string to a Raw String
' "Test""123""_" -> Test"123"_
Public Function UndoAutoItString$(Au3Str$)
   Dim StringTerminal$
   
  'Get stringchar ( should be " or ')
   StringTerminal$ = Left(Au3Str, 1)
   
  'Cut away Lead&Tailing " or '
  'Is length of Au3Str is smaller than 2 this will give an error
  'since it's no valid Au3String
   Au3Str = Mid(Au3Str, 2, Len(Au3Str) - 2)
   
   
  'Replaces '' -> '  or "" -> "
   UndoAutoItString = Replace(Au3Str, StringTerminal & StringTerminal, StringTerminal)
   
End Function

'   With New RegExp
'      .Global = True
'
'      Const StringTerminal$ = "(['""])"
'      Const StringTerminalBackRef$ = "\1"
'      Const StringBody$ = "(.*?)"
'
'
'      .Pattern = StringTerminal & _
'                   "(?:" & _
'                   StringBody & _
'                     StringTerminalBackRef & StringTerminalBackRef & _
'                        StringBody & _
'                   ")*" & _
'                 StringTerminalBackRef
'      '$2 is the StringBody
'      '$3 is
'      Au3StrToString = .Replace(Au3Str, "$2$3$4")
'   End With
   
'End Function



'
'' Add WhiteSpace Seperator to SourceCodeLine
'Function AddWhiteSpace$()
'
'   'No WhiteSpace at the Beginning
'   If SourceCodeLine = "" Then Exit Function
'
'   Dim LastChar$
'   LastChar = Right(SourceCodeLine, 1)
'
'   Dim NextChar$
'   NextChar = Left(Atom, 1)
'
'   'Don'Append WhiteSpace in cases like this :
'   '"@CMDLIND ["   or   "@CMDLIND [0" <-"].."
'   '         (^-PreCase)                (^-PostCase)
'   If InStr(1, ExcludePreWhiteSpaceTerminal, LastChar) Or _
'      InStr(1, ExcludePostWhiteSpaceTerminal, NextChar) Then
''      Stop
'   ElseIf whiteSpaceTerminal <> LastChar Then
'         AddWhiteSpace = whiteSpaceTerminal
'   End If
'
'End Function





Private Sub FL_verbose(Text)
   FrmMain.FL_verbose Text
End Sub
Private Sub log_verbose(TextLine$)
   FrmMain.log_verbose TextLine$
End Sub

Private Sub FL(Text)
   FrmMain.FL Text
End Sub

'/////////////////////////////////////////////////////////
'// log -Add an entry to the Log
Private Sub Log(TextLine$)
   FrmMain.Log TextLine$
End Sub

'/////////////////////////////////////////////////////////
'// log_clear - Clears all log entries
Private Sub Log_Clear()
   FrmMain.Log_Clear
End Sub

