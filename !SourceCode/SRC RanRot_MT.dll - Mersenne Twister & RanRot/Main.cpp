// RanRot.cpp : Definiert den Einstiegspunkt für die Konsolenanwendung.
//
#include "randomc.h"  
#include <stdio.h>

//uses RanRot_MT.def
#define DllExport  __stdcall //extern "C" __declspec(dllexport) 

#define WIN32_LEAN_AND_MEAN		// Selten verwendete Teile der Windows-Header nicht einbinden
// Windows-Headerdateien:
#include <windows.h>

TRanrotBGenerator* myRanRot;



	void DllExport RanRot_Init(int seed) {
		myRanRot->RandomInit(seed);
	}

	int DllExport RanRot_GetI8() {
		return myRanRot->I8Random();
	}


BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		myRanRot = new TRanrotBGenerator(0x0);
		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
//		delete myRanRot;
		break;
	}
    return TRUE;
}


//int  main()//int argc, _TCHAR* argv[])
//{
//	int seed = 4;
//	myRanRot = new TRanrotBGenerator(0x0);
//	RanRot_Init(0x99F2);
//
//
//
//	//WRandomInit (0x99F2);
//	for (int i=0; i<7; i++) {
////		myRanRot->Random();
////		seed=myRanRot->Random()*0x100;
//		printf ("\n %X", RanRot_GetI8() );}
////99 B5
//	return 0;
//}

