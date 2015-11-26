// Test exe file for compression classes - very untidy code only used
// for testing compression functions quickly
//
// CONSOLE APP
//
// (c)2002-2003 Jonathan Bennett, jon@hiddensoft.com
//

#include <stdio.h>
#include <conio.h>
#include <windows.h>
#include <mmsystem.h>
#include "jb01_compress.h"
#include "jb01_decompress.h"


///////////////////////////////////////////////////////////////////////////////
// CompressMonitorProc() - The callback function
///////////////////////////////////////////////////////////////////////////////

int CompressMonitorProc(ULONG nBytesIn, ULONG nBytesOut, UINT nPercentComplete)
{
	static	UINT	nDelay = 0;
	static	UINT	nRot = 0;
	char	szGfx[]= "-\\|/";
	UCHAR	ch;

//	if (nDelay > 16)
//	{
		nDelay = 0;
		nRot = (nRot+1) & 0x3;
		printf("\rCompressing %c        : %d%% (%d%%)  ", szGfx[nRot], nPercentComplete, 100-((100*nBytesOut) / nBytesIn));

//	}
//	else
//		nDelay++;

	// Check if ESC was pressed and if so request stopping
	if (kbhit())
	{
		ch = getch();
		if (ch == 0)
			ch = getch();
		if (ch == 27)
			return 0;
	}

	return 1;

} // CompressProc()


///////////////////////////////////////////////////////////////////////////////
// DeompressMonitorProc() - The callback function
///////////////////////////////////////////////////////////////////////////////

int DecompressMonitorProc(ULONG nBytesIn, ULONG nBytesOut, UINT nPercentComplete)
{
	static	UINT	nDelay = 0;
	static	UINT	nRot = 0;
	char	szGfx[]= "-\\|/";
	UCHAR	ch;

//	if (nDelay > 16)
//	{
		nDelay = 0;
		nRot = (nRot+1) & 0x3;
		printf("\rDecompressing %c      : %d%%  ", szGfx[nRot], nPercentComplete);
//	}
//	else
//		nDelay++;

	// Check if ESC was pressed and if so request stopping
	if (kbhit())
	{
		ch = getch();
		if (ch == 0)
			ch = getch();
		if (ch == 27)
			return 0;
	}

	return 1;

} // DecompressMonitorProc()


///////////////////////////////////////////////////////////////////////////////
// mainGetFileSize()
//
// Uses Win32 functions to quickly get the size of a file (rather than using
// fseek/ftell which is slow on large files
//
// Use BEFORE opening a file! :)
//
///////////////////////////////////////////////////////////////////////////////

ULONG mainGetFileSize(const char *szFile)
{
	HANDLE	hFile;
	ULONG	nSize;

	hFile = CreateFile(szFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL, NULL);

	if ( hFile == INVALID_HANDLE_VALUE )
		return 0;

	nSize = GetFileSize(hFile, NULL);

	CloseHandle(hFile);

	return nSize;

} // mainGetFileSize()


///////////////////////////////////////////////////////////////////////////////
// main()
///////////////////////////////////////////////////////////////////////////////

int main(int argc, char* argv[])
{
	unsigned long		nCompressedSize;
	unsigned long		nUncompressedSize;
	int					nRes;
	JB01_Compress	oCompress;					// Our compression class
	JB01_Decompress	oDecompress;				// Our decompression class


	printf("\nHiddenSoft Compression Routine - (c)2002-2003 Jonathan Bennett\n");
	printf("--> Extended Version 0.2 by CW2K - [12.10.2007]\n");
	printf("--------------------------------------------------------------\n\n");

	// Compress file to file function
	if ((argc ==4) && (!stricmp("-c", argv[1])))
	{
		// How big is the source file?
		nUncompressedSize = mainGetFileSize(argv[2]);
		printf("Input file size      : %d\n", nUncompressedSize);

		// Do the compression
		oCompress.SetDefaults();
		oCompress.SetInputType(HS_COMP_FILE);
		oCompress.SetOutputType(HS_COMP_FILE);
		oCompress.SetInputFile(argv[2]);
		oCompress.SetOutputFile(argv[3]);
		oCompress.SetMonitorCallback(&CompressMonitorProc);
		oCompress.SetCompressionLevel(4);
		DWORD dwTime1 = timeGetTime();
		nRes = oCompress.Compress();
		DWORD dwTime2 = timeGetTime();

		printf("\rCompressed           : %d%% (%d%%)  ", oCompress.GetPercentComplete(), 100 - ( oCompress.GetCompressedSize()  / nUncompressedSize) * 100  );
		printf("\nCompression time     : %.2fs (including fileIO)\n", ((dwTime2-dwTime1)) / 1000.0);

		if (nRes != JB01_E_OK)
		{
			printf("Error compressing.\n");
			return 0;
		}

		// Print the output size
		printf("Output file size     : %d\n", oCompress.GetCompressedSize());
		printf("Compression ratio    : %.2f%%\n", 100 - ((oCompress.GetCompressedSize() / nUncompressedSize) * 100) );
		printf("Compression ratio    : %.3f bpb\n", (oCompress.GetCompressedSize() * 8) / nUncompressedSize);

		return 0;
	}




	// Uncompress file to file function
	if ((argc==4) && (!stricmp("-d", argv[1])))
	{
		// How big is the source file?
		nCompressedSize = mainGetFileSize(argv[2]);
		printf("Input file size      : %d\n", nCompressedSize);

		// Do the uncompression
		oDecompress.SetDefaults();
		oDecompress.SetInputType(HS_COMP_FILE);
		oDecompress.SetOutputType(HS_COMP_FILE);
		oDecompress.SetInputFile(argv[2]);
		oDecompress.SetOutputFile(argv[3]);
		oDecompress.SetMonitorCallback(&DecompressMonitorProc);
		DWORD dwTime1 = timeGetTime();
		nRes = oDecompress.Decompress();
		DWORD dwTime2 = timeGetTime();

		printf("\rDecompressed         : %d%%  ", oDecompress.GetPercentComplete());
		printf("\nCompression time     : %.2fs (including fileIO)\n", ((dwTime2-dwTime1)) / 1000.0);

		if (nRes != JB01_E_OK)
		{
			printf("Error uncompressing.\n");
			return 0;
		}

		// Print filesize
		printf("Output file size     : %d\n", mainGetFileSize(argv[3]));

		return 0;

	}




	// If we got here, invalid parameters
	printf("Usage: %s <-c | -d> <infile> <outfile>\n", argv[0]);
	printf("  -c performs file to file compression\n");
	printf("  -d performs file to file decompression\n\n");
	printf("Supported files type(s) 'JB01' (and 'JB00', 'EA05', 'EA06' decompression only).\n");

	return 0;
}
