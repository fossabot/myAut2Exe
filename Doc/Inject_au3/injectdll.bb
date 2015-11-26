typedef struct _LOADED_IMAGE {
  PSTR                  ModuleName;
  HANDLE                hFile;
  
  
  Attention that sample code has a great tendency to [b]crash [/b] >_< when it comes to call[url="http://www.autoitscript.com/autoit3/docs/functions/DllCall.htm"][color=#000090]DllCall[/color][/url][color=#FF8000][font=Consolas, 'Courier New', Courier, 'Nimbus Mono L', monospace][size=2]([/size][/font][/color][color=#FF0000][font=Consolas, 'Courier New', Courier, 'Nimbus Mono L', monospace][size=2]"imagehlp.dll"[/size][/font][/color][color=#FF8000][font=Consolas, 'Courier New', Courier, 'Nimbus Mono L', monospace][size=2],[/size][/font][/color][color=#FF0000][font=Consolas, 'Courier New', Courier, 'Nimbus Mono L', monospace][size=2]"UnMapAndLoad"... )[/size][/font][/color][list]
[*]Well one quick and dirty workaround would be to just drop/[b]delete [/b]all [color=#FF0000][font=Consolas, 'Courier New', Courier, 'Nimbus Mono L', monospace][size=2]UnMapAndLoad [/size][/font][/color]calls since we just [b]read[/b] data. Nothing is needed to write back. :thumbsup:
[/list]
But well that's somehow unsatisfying and so we've learn nothing from that.Well it took me some time to find out the reason for the crash that somehow occurs inside[i]imagehlp!UnMapAndLoad();FreeModuleName [/i]during the KERNEL32.HeapFree call.First i though there is something wrong with some unproper use of "ptr" or DllStructGetPtr() well but this seem to be fine.However to crashes here:typedef [color=blue]struct[/color] _LOADED_IMAGE {  PSTR                  ModuleName;PUSH    [DWORD EBP+8]                    ; /pMemory <-[color=#000000][font=Consolas, Courier, monospace]PSTR    ModuleName;[/font][/color]PUSH    0                                              ; |Flags = 0PUSH    [color=#ff0000][DWORD 6CC851F4] [/color]              ; [b]|hHeap = 01420000[/b]CALL    [<&KERNEL32.HeapFree>]      when it tries to free the memory [i]MapAndLoad [/i]allocated for the path of the ModuleName.And finally I found it - the is problem is that Un[i]MapAndLoad uses some other [/i][b]hHeap[/b] address than [i]MapAndLoad[/i] used.So why does that value changes ?[u]->Well because Autoit loads and unload imagehlp.dll on each call when you use the dll parameter as a String.[/u][list]
[*]So instead doing it like this
[s][autoit]
DllCall("imagehlp.dll","MapAndLoad"... )
DllCall("imagehlp.dll","UnMapAndLoad"... )
[/autoit][/s]Do it like this:[autoit]
$himagehlp_dll = DllOpen("imagehlp.dll")
DllCall(himagehlp_dll, "int", "MapAndLoad"...)
DllCall(himagehlp_dll, "int", "UnMapAndLoad"...)
[/autoit]

and everything will be alright.
[/list]
Well base on the example i made my own code that is targeting the Dll-Flag in the PE-Characteristics
[autoit]
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
[/autoit]
^-that's just the roadmap - DL ExeToDll.au3 for real use....and why convert a Exe to an DLL?Na well - just for fun. :dance:... or to inject it into another Exe - to get some code (or au3-Script :ILA3:) in, to extract  some 'virtual'/'bundled' files or data from that Process.-> Inject_au3.7z :D