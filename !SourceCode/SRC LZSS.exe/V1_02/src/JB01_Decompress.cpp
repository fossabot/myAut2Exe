//
// jb01_decompress.cpp
// (c)2002 Jonathan Bennett (jon@hiddensoft.com)
//
//

#include <stdio.h>
#include <math.h>
#include <windows.h>
#include "jb01_decompress.h"


///////////////////////////////////////////////////////////////////////////////
// GetDecompressedSize()
///////////////////////////////////////////////////////////////////////////////

ULONG JB01_Decompress::GetDecompressedSize(void)
{
	int		nRes;

	// If this is file input open our file for reading
	if (m_nInputType == HS_COMP_FILE)
	{
		if ( (m_fSrc = fopen(m_szSrcFile, "rb")) == NULL)
			return JB01_E_READINGSRC;				// Error
	}
	else
		m_fSrc = NULL;

	// Init vars
	m_nUserDataPos			= 0;				// Bytes written to output
	m_nUserCompPos			= 0;				// Bytes read from input

	// Read the compressed data header
	nRes = ReadUserCompHeader(m_nDataSize);

	// Close our file
	fclose(m_fSrc);

	// Was it a valid compressed stream?
	if ( nRes != JB01_E_OK )
		return 0;								// Wasn't a valid LZSS stream (size = 0)
	else
		return m_nDataSize;						// Return with size of decompressed data

} // GetDecompressedSize()


///////////////////////////////////////////////////////////////////////////////
// SetDefaults()
// Should be called once before first using the Compress() function
///////////////////////////////////////////////////////////////////////////////

void JB01_Decompress::SetDefaults(void)
{
	m_bUserData			= NULL;
	m_bUserCompData		= NULL;

	m_nUserDataPos		= 0;					// TOTAL data bytes read
	m_nUserCompPos		= 0;					// TOTAL compressed bytes written

	m_nDataSize			= 0;					// TOTAL file uncompressed size

	m_nInputType		= HS_COMP_FILE;
	m_nOutputType		= HS_COMP_FILE;

	m_fSrc				= NULL;
	m_fDst				= NULL;
	m_szSrcFile[0]		= '\0';
	m_szDstFile[0]		= '\0';

	m_lpfnMonitor		= NULL;					// The monitor callback function (or NULL)

} // SetDefaults()


///////////////////////////////////////////////////////////////////////////////
// DecompressFile()
//
// Only accesses m_bBlockData and m_bCompBlockData member variables!
//
///////////////////////////////////////////////////////////////////////////////

int JB01_Decompress::Decompress(void)
{
	int		nRes;

	// If this is file input open our file for reading
	if (m_nInputType == HS_COMP_FILE)
	{
		if ( (m_fSrc = fopen(m_szSrcFile, "rb")) == NULL)
			return JB01_E_READINGSRC;				// Error
	}
	else
		m_fSrc = NULL;

	// If this is file output open our file for writing
	if (m_nOutputType == HS_COMP_FILE)
	{
		if ( (m_fDst = fopen(m_szDstFile, "w+b")) == NULL)
		{
			if (m_fSrc)
				fclose(m_fSrc);
			return JB01_E_WRITINGDST;				// Error
		}
	}
	else
		m_fDst = NULL;

	// Initialise vars
	m_nUserCompPos			= 0;				// Bytes read from input
	m_nUserDataPos			= 0;				// Bytes written to output
	m_nDataPos				= 0;				// Pos in our internal data buffer
	m_nDataUsed				= 0;
	m_nDataWritePos			= 0;
	m_nCompressedLong		= 0;				// Compressed stream temporary 32bit value
	m_nCompressedBitsUsed	= 0;				// Number of bits unused in temporary value
	m_bAbortRequested		= false;


	// Check that data is a valid LZSS stream and get uncompressed size too
	nRes = ReadUserCompHeader(m_nDataSize);

	if ( nRes != JB01_E_OK )
	{
		if (m_fSrc)
			fclose(m_fSrc);
		if (m_fDst)
			fclose(m_fDst);
		return JB01_E_NOTJB01;				// Wasn't a valid LZSS stream
	}


	// Allocate the memory needed for decompression
	nRes = AllocMem();
	if (nRes != JB01_E_OK)
	{
		if (m_fSrc)
			fclose(m_fSrc);
		if (m_fDst)
			fclose(m_fDst);

		return nRes;							// Return error code
	}

	// Do the decompression depending on type
	if (m_isStreamTypeJB01)
		DecompressLoop();
	else if (m_isStreamTypeJB00)
        DecompressLoop_JB00();
    else    
        DecompressLoop_EA05();



	// Free memory used by decompression
	FreeMem();

	// Close our files if required
	if (m_fSrc)
		fclose(m_fSrc);
	if (m_fDst)
		fclose(m_fDst);

	return JB01_E_OK;							// Return with success message

} // Decompress


///////////////////////////////////////////////////////////////////////////////
// ReadUserCompHeader()
// Reads in the header (Alg ID, Rev ID, uncompressed size) from the compressed
// input (9 bytes)
///////////////////////////////////////////////////////////////////////////////

int JB01_Decompress::ReadUserCompHeader(ULONG &nSize)
{
	UCHAR	bBuffer[8];

	// Memory or file as the input?
	if (m_nInputType == HS_COMP_MEM)
		memcpy(bBuffer, &m_bUserCompData[m_nUserCompPos], 8);
	else
		fread(bBuffer, sizeof(UCHAR), 8, m_fSrc);

	// Update position (skip the header, basically)
	m_nUserCompPos += 8;


	// Get uncompressed size
	nSize = (ULONG)bBuffer[4] << 24;
	nSize = nSize | (ULONG)bBuffer[5] << 16;
	nSize = nSize | (ULONG)bBuffer[6] << 8;
	nSize = nSize | (ULONG)bBuffer[7];

	// Terminate ALG string
	bBuffer[4] = '\0';

	// Set Compress StreamType
    m_isStreamTypeJB00 = !strcmp((char*)bBuffer, JB00_ALGID);
    m_isStreamTypeJB01 = !strcmp((char*)bBuffer, JB01_ALGID);
	m_isStreamTypeEA06 = !strcmp((char*)bBuffer, EA06_ALGID);
	if (m_isStreamTypeJB00 || m_isStreamTypeEA06 || m_isStreamTypeJB01 || (!strcmp((char*)bBuffer, EA05_ALGID)) )
		return JB01_E_OK;							// Return with success message

	else
		return JB01_E_NOTJB01;

} // ReadUserCompHeader()


///////////////////////////////////////////////////////////////////////////////
// AllocMem()
//
// Allocates our block buffer
//
///////////////////////////////////////////////////////////////////////////////

int JB01_Decompress::AllocMem(void)
{
	m_bData	= (UCHAR *)malloc(JB01_DATA_SIZE * sizeof(UCHAR));

	// Huffman stuff
	// Tree can be 2n-1 elements in size
	// Number of output codes = size of alphabet
	m_HuffmanLiteralTree
		= (JB01_HuffmanDecompNode *)malloc(((2*JB01_HUFF_LITERAL_ALPHABETSIZE)-1) * sizeof (JB01_HuffmanDecompNode));

	m_HuffmanOffsetTree
		= (JB01_HuffmanDecompNode *)malloc(((2*JB01_HUFF_OFFSET_ALPHABETSIZE)-1) * sizeof (JB01_HuffmanDecompNode));

	if ( (m_bData == NULL) //|| (m_bComp == NULL)
			|| (m_HuffmanLiteralTree == NULL) || (m_HuffmanOffsetTree == NULL) )
	{
		FreeMem();
		return JB01_E_MEMALLOC;
	}
	else
		return JB01_E_OK;


} // AllocMem()


///////////////////////////////////////////////////////////////////////////////
// FreeMem()
//
// Frees our block buffer
//
///////////////////////////////////////////////////////////////////////////////

void JB01_Decompress::FreeMem(void)
{
	HS_COMP_SafeFree(m_bData);

	HS_COMP_SafeFree(m_HuffmanLiteralTree);
	HS_COMP_SafeFree(m_HuffmanOffsetTree);

} // FreeMem()


///////////////////////////////////////////////////////////////////////////////
// WriteUserData()
// Outputs data the the uncompressed data stream (file or memory)
///////////////////////////////////////////////////////////////////////////////

void JB01_Decompress::WriteUserData(void)
{
	// Write out all the data from our last write position to our current position

	// Memory or file as the input?
	if (m_nOutputType == HS_COMP_MEM)
	{
		while (m_nDataWritePos < m_nDataPos)
		{
			m_bUserData[m_nUserDataPos++] = m_bData[m_nDataWritePos & JB01_DATA_MASK];
			m_nDataWritePos++;
		}
	}
	else
	{
		while (m_nDataWritePos < m_nDataPos)
		{
			fputc(m_bData[m_nDataWritePos & JB01_DATA_MASK], m_fDst);
			m_nDataWritePos++;
			m_nUserDataPos++;					// Keep track of this even though using a file
		}
	}

	// Update totals
	m_nDataUsed		= 0;						// Update how full the buffer is

} // WriteUserData()


///////////////////////////////////////////////////////////////////////////////
// MonitorUpdate()
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Decompress::MonitorCallback(void)
{
	static	UINT nDelay = 0;
	int		nRes;

	if (nDelay > 4096)							// Call function every 4096 loops (~8192 bytes)
	{
		nDelay = 0;

		// If present call the user defined function
		if (m_lpfnMonitor)
		{
			nRes = m_lpfnMonitor(m_nUserCompPos, m_nUserDataPos,
							(m_nUserDataPos * 100) / m_nDataSize);
			if (!nRes)
				m_bAbortRequested = true;
		}
	}
	else
		++nDelay;

} // MonitorUpdate()


///////////////////////////////////////////////////////////////////////////////
// DecompressLoop()
///////////////////////////////////////////////////////////////////////////////

int JB01_Decompress::DecompressLoop(void)
{
	ULONG	nMaxPos;
	UINT	nTemp;
	UINT	nLen;
	UINT	nOffset;
	ULONG	nTempPos;


	// At the start huffman coding is off
	HuffmanInit();

	// Perform decompression until we fill our predicted size (uncompressed size)
	nMaxPos		= m_nDataSize;

	while(m_nDataPos < nMaxPos)
	{
		// Read in a literal
		nTemp = CompressedStreamReadLiteral();

		// Was it a literal byte, or a  match len?
		if (nTemp < JB01_HUFF_LITERAL_LENSTART)	// 0-255 are literals, 256-292 are lengths
		{
			// Store the literal byte
			m_bData[m_nDataPos & JB01_DATA_MASK] = (UCHAR)nTemp;
			m_nDataPos++;
			m_nDataUsed++;
		}
		else
		{
			// Decode (and read more if required) to get the length of the match
			nLen = CompressedStreamReadLen(nTemp) + JB01_MINMATCHLEN;

			// Read the offset
			nOffset = CompressedStreamReadOffset();

			// Write out our match
			nTempPos = m_nDataPos - nOffset;
			while (nLen)
			{
				--nLen;
				m_bData[m_nDataPos & JB01_DATA_MASK] = m_bData[nTempPos & JB01_DATA_MASK];
				nTempPos++;
				m_nDataPos++;
				m_nDataUsed++;
			}
		}


		// Write it out
		WriteUserData();

		MonitorCallback();
		if (m_bAbortRequested)
			return JB01_E_ABORT;
	}

	return JB01_E_OK;

} // DecompressLoop()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamReadBits()
//
// Will read up to 16 bits from the compressed data stream
//
///////////////////////////////////////////////////////////////////////////////

inline UINT JB01_Decompress::CompressedStreamReadBits(UINT nNumBits)
{

	UINT	nTemp;

	// Ensure that the high order word of our bit buffer is blank
	m_nCompressedLong = m_nCompressedLong & 0x0000ffff;

	while (nNumBits)
	{
		--nNumBits;

		// Check if we need to refill our decoding bit buffer
		if (!m_nCompressedBitsUsed)
		{
			// Yes, we need to read in another 16 bits (two bytes)
			// Fill the low order 16 bits of our long buffer

			if (m_nInputType == HS_COMP_MEM)
			{
				m_nCompressedLong = m_nCompressedLong | (m_bUserCompData[m_nUserCompPos++] << 8);
				m_nCompressedLong = m_nCompressedLong | m_bUserCompData[m_nUserCompPos++];
			}
			else
			{
				nTemp = fgetc(m_fSrc);
				m_nCompressedLong = m_nCompressedLong | (nTemp << 8);
				nTemp = fgetc(m_fSrc);
				m_nCompressedLong = m_nCompressedLong | nTemp;
				m_nUserCompPos += 2;			// Still update even though it is a file
			}

			m_nCompressedBitsUsed = 16;			// We've used 16 bits
		}

		// Shift the data into the high part of the long
		m_nCompressedLong = m_nCompressedLong << 1;
		--m_nCompressedBitsUsed;
	}

	return (UINT)(m_nCompressedLong >> 16);

} // CompressedStreamReadBits()


///////////////////////////////////////////////////////////////////////////////
// HuffmanInit()
///////////////////////////////////////////////////////////////////////////////

void JB01_Decompress::HuffmanInit(void)
{
	// Literal and match length tree
	HuffmanZero(m_HuffmanLiteralTree, JB01_HUFF_LITERAL_ALPHABETSIZE);
	HuffmanGenerate(m_HuffmanLiteralTree, JB01_HUFF_LITERAL_ALPHABETSIZE, 0);
	m_bHuffmanLiteralFullyActive = false;
	m_nHuffmanLiteralIncrement = JB01_HUFF_LITERAL_INITIALDELAY;
	m_nHuffmanLiteralsLeft	= m_nHuffmanLiteralIncrement;

	// Offset tree
	HuffmanZero(m_HuffmanOffsetTree, JB01_HUFF_OFFSET_ALPHABETSIZE);
	HuffmanGenerate(m_HuffmanOffsetTree, JB01_HUFF_OFFSET_ALPHABETSIZE, 0);
	m_bHuffmanOffsetFullyActive = false;
	m_nHuffmanOffsetIncrement = JB01_HUFF_OFFSET_INITIALDELAY;
	m_nHuffmanOffsetsLeft	= m_nHuffmanOffsetIncrement;

} // HuffmanInit()


///////////////////////////////////////////////////////////////////////////////
// HuffmanZero()
///////////////////////////////////////////////////////////////////////////////

void JB01_Decompress::HuffmanZero(JB01_HuffmanDecompNode *HuffTree, UINT nAlphabetSize)
{
	// Reset the freqencies for all entries
	// At the start all entries are equally probably for an unknown file
	// A frequency of zero at the start creates a worst case tree with 255 char codes :(
	for (UINT i=0; i<nAlphabetSize; ++i)
	{
		HuffTree[i].nFrequency	= 1;
		HuffTree[i].nChildLeft	= i;			// The children on the leaf node will ALWAYS
		HuffTree[i].nChildRight	= i;			// equal themselves to indicate a leaf!
	}
}


///////////////////////////////////////////////////////////////////////////////
// HuffmanGenerate()
///////////////////////////////////////////////////////////////////////////////

void JB01_Decompress::HuffmanGenerate(JB01_HuffmanDecompNode *HuffTree, UINT nAlphabetSize, UINT nFreqMod)
{
	UINT	i, j;
	UINT	nNextBlankEntry;
	UINT	nByte1 = 0, nByte2 = 0;
	ULONG	nByte1Freq, nByte2Freq;
	UINT	nParent;
	UINT	nRoot;
	UINT	nEndNode;

	// Reset the table so we can search the first set of elements
	// entries (actual bytes)
	for (i=0; i<nAlphabetSize; ++i)
		HuffTree[i].bSearchMe = true;

	nRoot = (nAlphabetSize << 1) - 2;

	// Next free entry in the array is now nAlphabetSize
	nNextBlankEntry = nAlphabetSize;
	nEndNode = nRoot + 1;
	while (nNextBlankEntry != nEndNode )
	{
		// Get least 2 frequent entries (byte1=least frequent, byte2= next least recent)
		nByte1Freq	= nByte2Freq	= 0xffffffff;
		for (i=0; i<nNextBlankEntry; ++i)
		{
			if ( HuffTree[i].bSearchMe != false)
			{
				if (HuffTree[i].nFrequency < nByte2Freq)
				{
					if (HuffTree[i].nFrequency < nByte1Freq)
					{
						nByte2		= nByte1;
						nByte2Freq	= nByte1Freq;
						nByte1		= i;
						nByte1Freq	= HuffTree[i].nFrequency;
					}
					else
					{
						nByte2		= i;
						nByte2Freq	= HuffTree[i].nFrequency;
					}
				}
			}
		}

		// Remove the two entries from the search list
		HuffTree[nByte1].bSearchMe = false;
		HuffTree[nByte2].bSearchMe = false;

		// Create a new parent entry with the combined frequency
		HuffTree[nNextBlankEntry].nFrequency	= HuffTree[nByte1].nFrequency + HuffTree[nByte2].nFrequency;
		HuffTree[nNextBlankEntry].bSearchMe		= true;	// Add to search list
		HuffTree[nNextBlankEntry].nChildLeft	= nByte1;
		HuffTree[nNextBlankEntry].nChildRight	= nByte2;
		HuffTree[nByte1].nParent				= nNextBlankEntry;
		HuffTree[nByte2].nParent				= nNextBlankEntry;
		HuffTree[nByte1].cValue					= 0;
		HuffTree[nByte2].cValue					= 1;

		++nNextBlankEntry;
	} // End while

	// The last array entry (JB01_HUFF_ROOTNODE) is now the parent node!

	// Check our tree to see that no codes are too long
	for (i=0; i<nAlphabetSize; ++i)				// nAlphabetSize symbols to code
	{
		nParent = i;
		j = 0;									// Number of bits long the code is
		while (nParent != nRoot)
		{
			++j;
			nParent = HuffTree[nParent].nParent;
		}

		// Ensure that codes are not too long, if they are divide the freqencies by 4
		// then start again
		if (j > JB01_HUFF_MAXCODEBITS)
		{
			//printf("\n\nDamnit - huffman code too long\n\n");
			for (i=0; i<nAlphabetSize; ++i)
				HuffTree[i].nFrequency = (HuffTree[i].nFrequency >> 2) + 1;

			HuffmanGenerate(HuffTree, nAlphabetSize, nFreqMod);
			return;
		}
	} // End For


	// Finally, reduce the probability of all the freqencies of the individual
	// bytes so that "old" frequencies are worth less than any new data we get
	if (nFreqMod)
	{
		// Divide by freqency modifier, make sure is 1 or more (zeros do bad things...)
		for (i=0; i<nAlphabetSize; ++i)
			HuffTree[i].nFrequency = (HuffTree[i].nFrequency >> nFreqMod) + 1;
	}


} // HuffmanGenerate()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamReadHuffman()
///////////////////////////////////////////////////////////////////////////////

inline UINT JB01_Decompress::CompressedStreamReadHuffman(JB01_HuffmanDecompNode *HuffTree, UINT nAlphabetSize)
{
	UINT	nCode, nTemp;

	// Start with Root node of the tree, if a child is a pointer to itself
	// then it is the leaf and we stop
	nCode = ( nAlphabetSize << 1 ) - 2;
	while (HuffTree[nCode].nChildLeft != nCode)
	{
		nTemp = CompressedStreamReadBits(1);
		if (!nTemp)
			nCode = HuffTree[nCode].nChildLeft;
		else
			nCode = HuffTree[nCode].nChildRight;
	}

	// nLiteral will now be the leaf, which index in the array=literal :)

	return nCode;

} // CompressedStreamReadHuffman()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamReadLiteral()
///////////////////////////////////////////////////////////////////////////////

inline UINT JB01_Decompress::CompressedStreamReadLiteral(void)
{
	UINT	nLiteral;

	// Read in a huffman code from the compressed stream
	nLiteral = CompressedStreamReadHuffman(m_HuffmanLiteralTree, JB01_HUFF_LITERAL_ALPHABETSIZE);

	// Update the frequency of this character
	m_HuffmanLiteralTree[nLiteral].nFrequency++;
	--m_nHuffmanLiteralsLeft;

	// If we have coded enough literals, then generate/regenerate the huffman tree
	if (!m_nHuffmanLiteralsLeft)
	{
		if (m_bHuffmanLiteralFullyActive)
		{
			m_nHuffmanLiteralsLeft	= JB01_HUFF_LITERAL_DELAY;
			HuffmanGenerate(m_HuffmanLiteralTree, JB01_HUFF_LITERAL_ALPHABETSIZE, JB01_HUFF_LITERAL_FREQMOD);
		}
		else
		{
			m_nHuffmanLiteralIncrement += JB01_HUFF_LITERAL_INITIALDELAY;
			if (m_nHuffmanLiteralIncrement >= JB01_HUFF_LITERAL_DELAY)
				m_bHuffmanLiteralFullyActive = true;

			m_nHuffmanLiteralsLeft	= JB01_HUFF_LITERAL_INITIALDELAY;
			HuffmanGenerate(m_HuffmanLiteralTree, JB01_HUFF_LITERAL_ALPHABETSIZE, 0);
		}
	}

	return nLiteral;

} // CompressedStreamReadLiteral()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamReadLen()
///////////////////////////////////////////////////////////////////////////////

inline UINT JB01_Decompress::CompressedStreamReadLen(UINT nCode)
{
	UINT	nValue;
	UINT	nExtraBits;
	UINT	nMSBValue;

	if (nCode <= 263)
		return nCode - 256;
	else
	{
		// nCode increases by 4 for every extra bit added, 264 = 1 bit
		nCode = nCode - 264;
		nExtraBits	= (nCode >> 2) + 1;
		nMSBValue	= 1 << (nExtraBits+2);
		nCode = nCode & 0x0003;

		// Read in the extra bits
		nValue = CompressedStreamReadBits(nExtraBits);

		return nValue + nMSBValue + (nCode << nExtraBits);
	}

} // CompressedStreamReadLen()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamReadOffset()
///////////////////////////////////////////////////////////////////////////////

inline UINT JB01_Decompress::CompressedStreamReadOffset(void)
{
	UINT	nValue;
	UINT	nExtraBits;
	UINT	nMSBValue;
	UINT	nCode;

	// Read in a huffman code from the compressed stream
	nCode = CompressedStreamReadHuffman(m_HuffmanOffsetTree, JB01_HUFF_OFFSET_ALPHABETSIZE);

	// Update the frequency of this character
	m_HuffmanOffsetTree[nCode].nFrequency++;

	if (nCode <= 3)
		nValue = nCode;
	else
	{
		// nCode increases by 2 for every extra bit added, 4 = 1 bit
		nCode = nCode - 4;
		nExtraBits	= (nCode >> 1) + 1;
		nMSBValue	= 1 << (nExtraBits+1);
		nCode = nCode & 0x0001;

		// Read in the extra bits
		nValue = CompressedStreamReadBits(nExtraBits);
		nValue = nValue + nMSBValue + (nCode << nExtraBits);
	}

	// Update the frequency (above)
	--m_nHuffmanOffsetsLeft;

	// If we have coded enough literals, then generate/regenerate the huffman tree
	if (!m_nHuffmanOffsetsLeft)
	{
		if (m_bHuffmanOffsetFullyActive)
		{
			m_nHuffmanOffsetsLeft	= JB01_HUFF_OFFSET_DELAY;
			HuffmanGenerate(m_HuffmanOffsetTree, JB01_HUFF_OFFSET_ALPHABETSIZE, JB01_HUFF_OFFSET_FREQMOD);
		}
		else
		{
			m_nHuffmanOffsetIncrement += JB01_HUFF_OFFSET_INITIALDELAY;
			if (m_nHuffmanOffsetIncrement >= JB01_HUFF_OFFSET_DELAY)
				m_bHuffmanOffsetFullyActive = true;

			m_nHuffmanOffsetsLeft	= JB01_HUFF_OFFSET_INITIALDELAY;
			HuffmanGenerate(m_HuffmanOffsetTree, JB01_HUFF_OFFSET_ALPHABETSIZE, 0);
		}
	}

	return nValue;

} // CompressedStreamReadOffset()



///////////////////////////////////////////////////////////////////////////////
// CompressedStreamReadMatchLen()
///////////////////////////////////////////////////////////////////////////////

UINT JB01_Decompress::CompressedStreamReadMatchLen(void)
{
	UINT	nTemp;
	UINT	nLen;

	// Read in the match length using the convention shown in the LZP
	// article by Charles Bloom

	// Value	Bitstream
	// 0		00 (bin 0)
	// 1		01 (bin 1)
	// 2		10 (bin 2)
	// 3		11 000 (bin 0)					SUBTRACT 3
	// ...
	// 9		11 110 (bin 6)
	// 10		11 111 00000 (bin 0)			SUBTRACT 10
	// ...
	// 40       11 111 11110
	// 41       11 111 11111 00000000 (bin 0)	SUBTRACT 41
	// 296		11 111 11111 11111111 00000000 (bin 0)	SUBTRACT 296
	// 551		11 111 11111 11111111 00000000 (bin 0)   SUBTRACT 551

	nLen = 0;									// Starting value
	nTemp = CompressedStreamReadBits(2);		// Read in first two bits
	if (nTemp == 3)	// Bin 11 = Dec 3
	{
		nLen = 3;
		nTemp = CompressedStreamReadBits(3);	// Read next three bits
		if (nTemp == 7) // Bin 111 = Dec 7
		{
			nLen = 10;
			nTemp = CompressedStreamReadBits(5);	// Read next five bits
			if (nTemp == 31) // Bin 11111 = Dec 31
			{
				nLen = 41;
				nTemp = CompressedStreamReadBits(8);	// Read next eight bits
				if (nTemp == 255) // Bin 11111111 = Dec 255
				{
					nLen = 0x128;
					nTemp = CompressedStreamReadBits(8);	// Read next eight bits

		            while (nTemp == 255)
					{
						nLen = nLen + 255;
						nTemp = CompressedStreamReadBits(8);	// Read next eight bits
					}
				}
			}
		}
	}

	nLen = nLen + nTemp;	// Final calculation

	// Finally adjust the range from 0-295, to 1-296, we will never
	// have a match of 0 so it would be a waste
	nLen = nLen ;

	return nLen;

} // CompressedStreamReadMatchLen()


///////////////////////////////////////////////////////////////////////////////
// DecompressLoop_JB01()
///////////////////////////////////////////////////////////////////////////////

int JB01_Decompress::DecompressLoop_EA05(void)
{
	ULONG	nMaxPos;
	UINT	nTemp;
	UINT	nLen;
	UINT	nOffset;
	ULONG	nTempPos;

	// Perform decompression until we fill our predicted size (uncompressed size)
	nMaxPos		= m_nDataSize;

	while(m_nDataPos < nMaxPos)
	{

		// Read in the 1 bit flag
		nTemp=CompressedStreamReadBits(1);

		// Was it a literal byte, or a (offset,len) match pair?
		if (nTemp == ( m_isStreamTypeEA06 ? EA06_LITERAL : EA05_LITERAL ) )			// use nTemp == 0
		{
			// Store the literal byte
			nTemp=CompressedStreamReadBits(8);
			m_bData[m_nDataPos & JB01_DATA_MASK] = (UCHAR)nTemp;
			m_nDataPos++;
			m_nDataUsed++;
		}
		else
		{

			// Read the offset
			//nOffset = CompressedStreamReadOffset();
			nOffset = CompressedStreamReadBits(0xf); //0xf= HS_LZSS_WINBITS

			// Decode (and read more if required) to get the length of the match
			nLen = CompressedStreamReadMatchLen() + JB01_MINMATCHLEN;

			// Write out our match
			nTempPos = m_nDataPos - nOffset;
			while (nLen)
			{
				--nLen;
				m_bData[m_nDataPos & JB01_DATA_MASK] = m_bData[nTempPos & JB01_DATA_MASK];
				nTempPos++;
				m_nDataPos++;
				m_nDataUsed++;
			}
		}


		// Write it out
		WriteUserData();

		MonitorCallback();
		if (m_bAbortRequested)
			return JB01_E_ABORT;
	}

	return JB01_E_OK;

} // DecompressLoop()


///////////////////////////////////////////////////////////////////////////////
// DecompressLoop_JB00()
///////////////////////////////////////////////////////////////////////////////

int JB01_Decompress::DecompressLoop_JB00(void)
{
	ULONG	nMaxPos;
	UINT	nTemp;
	UINT	nLen;
	UINT	nOffset;
	ULONG	nTempPos;

	// Perform decompression until we fill our predicted size (uncompressed size)
	nMaxPos		= m_nDataSize;

	while(m_nDataPos < nMaxPos)
	{

		// Read in the 1 bit flag
		nTemp=CompressedStreamReadBits(1);

		// Was it a literal byte, or a (offset,len) match pair?
		if (nTemp == EA05_LITERAL  )			// use nTemp == 0
		{
			// Store the literal byte
			nTemp=CompressedStreamReadBits(8);
			m_bData[m_nDataPos & JB01_DATA_MASK] = (UCHAR)nTemp;
			m_nDataPos++;
			m_nDataUsed++;
		}
		else
		{

			// Read the offset
			//nOffset = CompressedStreamReadOffset();
			nOffset = CompressedStreamReadBits(0xd) + JB01_MINMATCHLEN; //0xf= HS_LZSS_WINBITS

			// Decode (and read more if required) to get the length of the match
			nLen = CompressedStreamReadBits(0x4) + JB01_MINMATCHLEN;

			// Write out our match
			nTempPos = m_nDataPos - nOffset;
			while (nLen)
			{
				--nLen;
				m_bData[m_nDataPos & JB01_DATA_MASK] = m_bData[nTempPos & JB01_DATA_MASK];
				nTempPos++;
				m_nDataPos++;
				m_nDataUsed++;
			}
		}


		// Write it out
		WriteUserData();

		MonitorCallback();
		if (m_bAbortRequested)
			return JB01_E_ABORT;
	}

	return JB01_E_OK;

} // DecompressLoop()
