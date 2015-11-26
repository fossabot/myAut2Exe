// injectdll.cpp : Defines the entry point for the console application.
//

#include <windows.h>
#include <stdio.h>


int main(int argc, char* argv[])
{
	if (argc < 3)
	{
		printf("usage: injectdll <pid> <path to dll>\n");
		return 1;
	}

	// PROCESS_CREATE_THREAD|PROCESS_VM_WRITE|PROCESS_VM_READ|PROCESS_QUERY_INFORMATION|PROCESS_VM_OPERATION|THREAD_QUERY_INFORMATION
	HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, atoi(argv[1]));

	if (hProcess == INVALID_HANDLE_VALUE)
	{
		fprintf(stderr, "cannot open that pid\n");
		return 1;
	}

	PVOID mem = VirtualAllocEx(hProcess, NULL, strlen(argv[2]) + 1, MEM_COMMIT, PAGE_READWRITE);

	if (mem == NULL)
	{
		fprintf(stderr, "can't allocate memory in that pid\n");
		CloseHandle(hProcess);
		return 1;
	}

	if (WriteProcessMemory(hProcess, mem, (void*)argv[2], strlen(argv[2]) + 1, NULL) == 0)
	{
		fprintf(stderr, "can't write to memory in that pid\n");
		VirtualFreeEx(hProcess, mem, strlen(argv[2]) + 1, MEM_RELEASE);
		CloseHandle(hProcess);
		return 1;
	}

	HANDLE hThread = CreateRemoteThread(hProcess, NULL, 0, (LPTHREAD_START_ROUTINE) GetProcAddress(GetModuleHandle("KERNEL32.DLL"),"LoadLibraryA"), mem, 0, NULL);
	if (hThread == INVALID_HANDLE_VALUE)
	{
		fprintf(stderr, "can't create a thread in that pid\n");
		VirtualFreeEx(hProcess, mem, strlen(argv[2]) + 1, MEM_RELEASE);
		CloseHandle(hProcess);
		return 1;
	}

	WaitForSingleObject(hThread, INFINITE);

	HANDLE hLibrary = NULL;
	if (!GetExitCodeThread(hThread, (LPDWORD)&hLibrary))
	{
		printf("can't get exit code for thread GetLastError() = %i.\n", GetLastError());
		CloseHandle(hThread);
		VirtualFreeEx(hProcess, mem, strlen(argv[2]) + 1, MEM_RELEASE);
		CloseHandle(hProcess);
		return 1;
	}

	CloseHandle(hThread);
	VirtualFreeEx(hProcess, mem, strlen(argv[2]) + 1, MEM_RELEASE);

	if (hLibrary == NULL)
	{
		hThread = CreateRemoteThread(hProcess, NULL, 0, (LPTHREAD_START_ROUTINE) GetProcAddress(GetModuleHandle("KERNEL32.DLL"),"GetLastError"), 0, 0, NULL);
		if (hThread == INVALID_HANDLE_VALUE)
		{
			fprintf(stderr, "LoadLibraryA returned NULL and can't get last error.\n");
			CloseHandle(hProcess);
			return 1;
		}

		WaitForSingleObject(hThread, INFINITE);
		DWORD error;
		GetExitCodeThread(hThread, &error);

		CloseHandle(hThread);

		printf("LoadLibrary return NULL, GetLastError() is %i\n", error);
		CloseHandle(hProcess);
		return 1;
	}

	CloseHandle(hProcess);

	printf("injected %08x\n", (DWORD)hLibrary);

	return 0;
}

