//
// jb01_compress.cpp
// (c)2002 Jonathan Bennett (jon@hiddensoft.com)
//
//

#include <stdio.h>
#include <math.h>
#include <windows.h>
#include "jb01_compress.h"


///////////////////////////////////////////////////////////////////////////////
// GetFileSize()
//
// Uses Win32 functions to quickly get the size of a file (rather than using
// fseek/ftell which is slow on large files
//
// Use BEFORE opening a file! :)
//
///////////////////////////////////////////////////////////////////////////////

ULONG JB01_Compress::GetFileSize(const char *szFile)
{
	HANDLE	hFile;
	ULONG	nSize;

	hFile = CreateFile(szFile, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL, NULL);

	if ( hFile == INVALID_HANDLE_VALUE )
		return 0;

	nSize = ::GetFileSize(hFile, NULL);

	CloseHandle(hFile);

	return nSize;

} // GetFileSize()


///////////////////////////////////////////////////////////////////////////////
// SetDefaults()
// Should be called once before first using the Compress() function
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::SetDefaults(void)
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

	SetCompressionLevel(2);

} // SetDefaults()


///////////////////////////////////////////////////////////////////////////////
// Compress()
///////////////////////////////////////////////////////////////////////////////

int JB01_Compress::Compress(void)
{
	int		nRes;
	UCHAR	szAlgID[] = JB01_ALGID;

	// If this is file input open our file for reading
	if (m_nInputType == HS_COMP_FILE)
	{
		m_nDataSize = GetFileSize(m_szSrcFile);

		if ( (m_fSrc = fopen(m_szSrcFile, "rb")) == NULL)
			return JB01_E_READINGSRC;			// Error
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
			return JB01_E_WRITINGDST;			// Error
		}
	}
	else
		m_fDst = NULL;


	// Initialise vars
	m_nUserDataPos			= 0;				// Bytes read from input
	m_nUserCompPos			= 0;				// Bytes written to output
	m_nDataPos				= 0;				// Pos in our internal data buffer
	m_nCompPos				= 0;				// Pos in our internal comp buffer
	m_nLookAheadSize		= 0;
	m_nDataReadPos			= 0;
	m_nCompressedLong		= 0;				// Compressed stream temporary 32bit value
	m_nCompressedBitsFree	= 32;				// Number of bits unused in temporary value
	m_bAbortRequested		= false;


	// Allocate the memory needed for compression
	nRes = AllocMem();
	if (nRes != JB01_E_OK)
	{
		if (m_fSrc)
			fclose(m_fSrc);
		if (m_fDst)
			fclose(m_fDst);

		return nRes;							// Return error code
	}


	// Write out the algorithm ID and revision number
	strcpy((char *)m_bComp, JB01_ALGID);
	m_nCompPos += 4;							// ID must be 4 bytes

	// Write out the uncompressed data size (makes it easier to decompress/allocate buffers
	// if we can ask for the filesize first)
	m_bComp[m_nCompPos++] = (UCHAR)(m_nDataSize >> 24);
	m_bComp[m_nCompPos++] = (UCHAR)(m_nDataSize >> 16);
	m_bComp[m_nCompPos++] = (UCHAR)(m_nDataSize >> 8);
	m_bComp[m_nCompPos++] = (UCHAR)(m_nDataSize);

	WriteUserCompData();						// Commit to mem/disk


	// Initialize pseudo adaptive huffman routines
	HuffmanInit();

	// Initialize our hash table
	HashTableInit();

	// Perform the compression
	CompressLoop();

	// Free memory used by compression
	FreeMem();

	// Close files if required
	if (m_fSrc)
		fclose(m_fSrc);
	if (m_fDst)
		fclose(m_fDst);

	return nRes;								// Return with success message

} // Compress()


///////////////////////////////////////////////////////////////////////////////
// AllocMem()
//
// Allocates our block buffer and hash table memory
//
///////////////////////////////////////////////////////////////////////////////

int JB01_Compress::AllocMem(void)
{
	m_bData				= (UCHAR *)malloc(JB01_DATA_SIZE * sizeof(UCHAR));

	// The most data that will ever be in our compressed stream buffer is when we write the header
	// which is 8 bytes - so lets go nuts and allocate double :)
	m_bComp				= (UCHAR *)malloc(16 * sizeof(UCHAR));

	// Hash table related
	m_HashTable			= (struct JB01_Hash **)malloc(JB01_HASHTABLE_SIZE * sizeof(struct JB01_Hash *));
	m_HashChainCounts	= (UINT *)malloc(JB01_HASHTABLE_SIZE * sizeof(UINT));

	m_nHashEntriesMax	= JB01_MAXWINDOWLENGTH;
	m_HashEntryMemPool	= (void *)malloc(m_nHashEntriesMax * sizeof(struct JB01_Hash));
	m_HashMemAllocTable	= (struct JB01_Hash **)malloc(m_nHashEntriesMax * sizeof(struct JB01_Hash *));


	// Huffman stuff
	// Tree can be 2n-1 elements in size
	// Number of output codes = size of alphabet
	m_HuffmanLiteralTree
		= (JB01_HuffmanNode *)malloc(((2*JB01_HUFF_LITERAL_ALPHABETSIZE)-1) * sizeof (JB01_HuffmanNode));
	m_HuffmanLiteralOutput
		= (JB01_HuffmanOutput *)malloc(JB01_HUFF_LITERAL_ALPHABETSIZE * sizeof (JB01_HuffmanOutput));

	m_HuffmanOffsetTree
		= (JB01_HuffmanNode *)malloc(((2*JB01_HUFF_OFFSET_ALPHABETSIZE)-1) * sizeof (JB01_HuffmanNode));
	m_HuffmanOffsetOutput
		= (JB01_HuffmanOutput *)malloc(JB01_HUFF_OFFSET_ALPHABETSIZE * sizeof (JB01_HuffmanOutput));


	// Make sure that we got all the memory we needed
	if ( (m_bData == NULL) || (m_bComp == NULL)
			|| (m_HashTable == NULL)			|| (m_HashChainCounts == NULL)
			|| (m_HashEntryMemPool == NULL)		|| (m_HashMemAllocTable == NULL)
			|| (m_HuffmanLiteralTree == NULL)	|| (m_HuffmanLiteralOutput == NULL)
			|| (m_HuffmanOffsetTree == NULL)	|| (m_HuffmanOffsetOutput == NULL) )
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
// Frees our block buffer and hash table memory
//
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::FreeMem(void)
{
	HS_COMP_SafeFree(m_bData);
	HS_COMP_SafeFree(m_bComp);

	HS_COMP_SafeFree(m_HashTable);
	HS_COMP_SafeFree(m_HashChainCounts);
	HS_COMP_SafeFree(m_HashEntryMemPool);
	HS_COMP_SafeFree(m_HashMemAllocTable);

	HS_COMP_SafeFree(m_HuffmanLiteralTree);
	HS_COMP_SafeFree(m_HuffmanLiteralOutput);
	HS_COMP_SafeFree(m_HuffmanOffsetTree);
	HS_COMP_SafeFree(m_HuffmanOffsetOutput);

} // FreeMem()


///////////////////////////////////////////////////////////////////////////////
// SetupCompressionLevel()
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::SetCompressionLevel(UINT nCompressionLevel)
{
	switch (nCompressionLevel)
	{
		default:
			m_nHashChainLimit	= 1;	// Fastest
			break;

		case 1:
			m_nHashChainLimit	= 4;	// Fast
			break;

		case 2:
			m_nHashChainLimit	= 16;	// Default / Normal level
			break;

		case 3:
			m_nHashChainLimit	= 64;	// Good
			break;

		case 4:
			m_nHashChainLimit	= 4096;	// Best
			break;

		case 5:
			m_nHashChainLimit	= JB01_MAXWINDOWLENGTH;	// Insane / Max
			break;

	}

} // SetupCompressionLevel()


///////////////////////////////////////////////////////////////////////////////
// MonitorUpdate()
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::MonitorCallback(void)
{
	static	UINT nDelay = 0;
	int		nRes;

	if (m_lpfnMonitor == NULL)
		return;

	if (nDelay > 4096)							// Call function every 4096 loops (~8192 bytes)
	{
		nDelay = 0;

		nRes = m_lpfnMonitor(m_nUserDataPos, m_nUserCompPos,
						(m_nUserDataPos * 100) / m_nDataSize);
		if (!nRes)
			m_bAbortRequested = true;
	}
	else
		++nDelay;

} // MonitorUpdate()


///////////////////////////////////////////////////////////////////////////////
// CompressLoop()
///////////////////////////////////////////////////////////////////////////////

int JB01_Compress::CompressLoop(void)
{
	ULONG	nMaxPos;
	ULONG	nOffset1, nOffset2;
	UINT	nLen1, nLen2;
	UINT	nIncrement;


	// Loop around until there is no more data, stop matching HASHORDER from the
	// end of the block so that we can remove some overrun code in the loop
	// +1 is for the lazy eval
	if (m_nDataSize >= (JB01_HASHORDER+1) )
		nMaxPos	= m_nDataSize - (JB01_HASHORDER+1);
	else
		nMaxPos = 0;

	while ( m_nDataPos < nMaxPos )
	{
		// Read in user data if required
		ReadUserData();

		// Check for a match at the current position
		FindMatches(m_nDataPos, nOffset1, nLen1, 0);	// Search for matches for current position
//		nLen1 = 0;

		// Did we get a match?
		if (nLen1)
		{
			// Do a match at next position to see if it's better?
			FindMatches(m_nDataPos+1, nOffset2, nLen2, nLen1);
			//nLen2 = 0;

			if ( nLen2 > (nLen1+1) )
			{
				// Match at +1 is better, write a literal then this match
				CompressedStreamWriteLiteral(m_bData[m_nDataPos & JB01_DATA_MASK]);	// Literal
				CompressedStreamWriteLen(nLen2 - JB01_MINMATCHLEN);	// Match Len
				CompressedStreamWriteOffset(nOffset2);				// Match offset
				nIncrement = nLen2+1;								// Move forwards matched len

			}
			else
			{
				CompressedStreamWriteLen( nLen1 - JB01_MINMATCHLEN );	// Match Len
				CompressedStreamWriteOffset(nOffset1);// Match offset
				nIncrement = nLen1;				// Move forwards matched len
			}
		}
		else
		{
			// No matches, just store the literal byte
			CompressedStreamWriteLiteral(m_bData[m_nDataPos & JB01_DATA_MASK]);
			nIncrement = 1;						// Move forward 1 literal
		}

		// We have skipped forwards either 1 byte or xxx bytes (if matched) we must now
		// add entries in the hash table for all the entries we've skipped
		HashTableAdd(nIncrement);			// Hashes at CURRENT POSITION for nIncrement bytes, also deletes old hashes

		// Update monitor variables and call user-defined callback
		MonitorCallback();
		if (m_bAbortRequested)
			return JB01_E_ABORT;

	} // End while


	// We will have stopped just short of the end of data because of the way the
	// hashing function/lazy eval needs to work, now output the remaining data as literals
	while ( m_nDataPos < m_nDataSize )
	{
		ReadUserData();

		CompressedStreamWriteLiteral(m_bData[m_nDataPos & JB01_DATA_MASK]);
		++m_nDataPos;
		--m_nLookAheadSize;
	}

	CompressedStreamWriteBitsFlush();		// Make sure all bits written

	return JB01_E_OK;						// Return with success message

} // CompressLoop()


///////////////////////////////////////////////////////////////////////////////
// ReadUserData()
// Fills up the data lookahead buffer
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::ReadUserData(void)
{
	ULONG	nBytes;
	ULONG	nLoop;
	ULONG	nUserDataLeft;

	// Do we need to read in anymore data or are we done?
	if (m_nUserDataPos >= m_nDataSize)
		return;

	nUserDataLeft = m_nDataSize - m_nUserDataPos;	// And how much left to go?

	// Work out how many bytes to read in to keep a full lookahead buffer
	// +1 for lazy eval support and also if we get a max match then the hashing
	// function hashes all those bytes matched so we have to make sure there is
	// data there!
	nBytes = (JB01_MAXMATCHLEN + (JB01_HASHORDER-1) + 1) - m_nLookAheadSize;

	//How much data can/should we read in?
	if (nUserDataLeft < nBytes)
		nBytes = nUserDataLeft;					// Just read in what's left


	nLoop = nBytes;

	// Memory or file as the input?
	if (m_nInputType == HS_COMP_MEM)
	{
		while (nLoop)
		{
			nLoop--;
			m_bData[m_nDataReadPos & JB01_DATA_MASK] = m_bUserData[m_nUserDataPos];
			m_nDataReadPos++;
			m_nUserDataPos++;
			//memcpy(bDest, &m_bUserData[m_nUserDataPos], nBytes);
		}
	}
	else
	{
		while (nLoop)
		{
			nLoop--;
			m_bData[m_nDataReadPos & JB01_DATA_MASK] = fgetc(m_fSrc);
			m_nDataReadPos++;
			m_nUserDataPos++;					// Keep track of this even though using a file
			//fread(bDest, sizeof(UCHAR), nBytes, m_fSrc);
		}
	}

	// Update totals
	m_nLookAheadSize		+= nBytes;					// Update how full the buffer is

} // ReadUserData()


///////////////////////////////////////////////////////////////////////////////
// WriteUserCompData()
// Outputs data to the compressed data stream (file or memory)
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::WriteUserCompData(void)
{
	// As we always start from 0 m_nCompPos is equal to the number of bytes in the buffer

	// Memory or file as the output?
	if (m_nOutputType == HS_COMP_MEM)
		memcpy(&m_bUserCompData[m_nUserCompPos], m_bComp, m_nCompPos);
	else
		fwrite(m_bComp, sizeof(UCHAR), m_nCompPos, m_fDst);

	m_nUserCompPos	+= m_nCompPos;				// Keep track even if writing to file
	m_nCompPos		= 0;

} // WriteUserData()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamWriteBits()
//
// Will write a number of bits (variable) into the compressed data stream
//
// When there are no more bits to send, you should call the function with the
// parameters 0, 0 to make sure that any left over bits are flushed into the
// compressed stream
//
// Note 16 bits is the maximum allowed value for this function!!!!!
// ================================================================
// This equates to a maximum window size of 65535
//
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::CompressedStreamWriteBits(UINT nValue, UINT nNumBits)
{
	while (nNumBits)
	{
		--nNumBits;

		// Make room for another bit (shift left once)
		m_nCompressedLong = m_nCompressedLong << 1;

		// Merge (OR) our value into the temporary long
		m_nCompressedLong = m_nCompressedLong | ((nValue >> nNumBits) & 0x00000001);

		// Update how many bits remain unused (sub 1)
		--m_nCompressedBitsFree;

		// Now check if we have filled our temporary long with bits (32bits)
		if (!m_nCompressedBitsFree)
		{
			// We now need to dump the highest 16 bits into our compressed
			// stream.  Highest order 8 bits first
			m_bComp[m_nCompPos++] = (UCHAR)(m_nCompressedLong >> 24);
			m_bComp[m_nCompPos++] = (UCHAR)(m_nCompressedLong >> 16);

			// We've now written out 16 bits so make more room (16 bits more room :) )
			m_nCompressedBitsFree += 16;

			// Write out the compressed data into the user compressed data stream (or file)
			WriteUserCompData();
		}

	} // End while

} // CompressedStreamWriteBits()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamWriteBitsFlush()
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::CompressedStreamWriteBitsFlush(void)
{
	// We have been asked to finish, flush remaining by using up remaining
	// unused bits with zeros
	// NOTE: Our compressed stream reading functions in the decompression
	// routines read in 16 bit chunks, so we must make sure that when we
	// flush that our ouput is aligned on 16 bits (in the CompressedStreamWriteBits function
	// we always write out in 32bit chunks so we only need to worry about the 16bit thing
	// here).  Otherwise we will overrun input buffers...

	// Shift left by number of Bits remaining so everything is as high up as possible
	m_nCompressedLong = m_nCompressedLong << m_nCompressedBitsFree;

	// Now write out minimum number of 16bit chunks to complete the stream
	while (m_nCompressedBitsFree < 32)
	{
		m_bComp[m_nCompPos++] = (UCHAR)(m_nCompressedLong >> 24);	// Move into char conversion position
		m_bComp[m_nCompPos++] = (UCHAR)(m_nCompressedLong >> 16);	// Move into char conversion position
		m_nCompressedLong = m_nCompressedLong << 16;				// Shift everything up 16 bits
		m_nCompressedBitsFree += 16;								// 16 more bits free
	}

	m_nCompressedBitsFree = 32;					// All done

	// Write out the compressed data into the user compressed data stream (or file)
	WriteUserCompData();

} // CompressedStreamWriteBitsFlush()



///////////////////////////////////////////////////////////////////////////////
// HuffmanInit()
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::HuffmanInit(void)
{
	// Literal and match length tree
	HuffmanZero(m_HuffmanLiteralTree, JB01_HUFF_LITERAL_ALPHABETSIZE);
	HuffmanGenerate(m_HuffmanLiteralTree, m_HuffmanLiteralOutput, JB01_HUFF_LITERAL_ALPHABETSIZE, 0);
	m_bHuffmanLiteralFullyActive = false;
	m_nHuffmanLiteralIncrement = JB01_HUFF_LITERAL_INITIALDELAY;
	m_nHuffmanLiteralsLeft	= m_nHuffmanLiteralIncrement;


	// Offset tree
	HuffmanZero(m_HuffmanOffsetTree, JB01_HUFF_OFFSET_ALPHABETSIZE);
	HuffmanGenerate(m_HuffmanOffsetTree, m_HuffmanOffsetOutput, JB01_HUFF_OFFSET_ALPHABETSIZE, 0);
	m_bHuffmanOffsetFullyActive = false;
	m_nHuffmanOffsetIncrement = JB01_HUFF_OFFSET_INITIALDELAY;
	m_nHuffmanOffsetsLeft	= m_nHuffmanOffsetIncrement;

} // HuffmanInit()


///////////////////////////////////////////////////////////////////////////////
// HuffmanZero()
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::HuffmanZero(JB01_HuffmanNode *HuffTree, UINT nAlphabetSize)
{
	UINT	i;

	// Reset the freqencies for all entries
	// At the start all entries are equally probably for an unknown file
	// A frequency of zero at the start creates a worst case tree with 255 char codes :(
	// SO we never let it go below 1
	for (i=0; i<nAlphabetSize; ++i)
	{
		HuffTree[i].nFrequency	= 1;
		HuffTree[i].nChildLeft	= i;			// The children on the leaf node will ALWAYS
		HuffTree[i].nChildRight	= i;			// equal themselves to indicate a leaf!
	}
}


///////////////////////////////////////////////////////////////////////////////
// HuffmanGenerate()
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::HuffmanGenerate(JB01_HuffmanNode *HuffTree, JB01_HuffmanOutput *HuffOutput, UINT nAlphabetSize, UINT nFreqMod)
{
	UINT	i, j;
	UINT	nNextBlankEntry;
	UINT	nByte1 = 0, nByte2 = 0;
	ULONG	nByte1Freq, nByte2Freq;
	UINT	nParent;
	UINT	nRoot;
	UINT	nEndNode;
	ULONG	nCode, nCodeTemp;

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
			printf("\n\nDamnit - huffman code too long\n\n");
			for (i=0; i<nAlphabetSize; ++i)
				HuffTree[i].nFrequency = (HuffTree[i].nFrequency >> 2) + 1;

			HuffmanGenerate(HuffTree, HuffOutput, nAlphabetSize, nFreqMod);
			return;
		}
	} // End For


	// Now fill our Huffman output table for each byte
	for (i=0; i<nAlphabetSize; ++i)				// nAlphabetSize symbols to code
	{
		nParent		= i;
		j			= 0;
		nCodeTemp	= 0;
		// Number of bits long the code is
		while (nParent != nRoot)
		{
			nCodeTemp = (nCodeTemp<<1) | HuffTree[nParent].cValue;	// cValue will only ever be 1/0
			++j;								// Update number of bits used in code
			nParent = HuffTree[nParent].nParent;
		}

		HuffOutput[i].nNumBits = j;				// Store number of bits

		// Now the code we just got is actually reversed, so we need to get it the right way
		nCode		= 0;
		while (j--)
		{
			nCode		= (nCode<<1) | (nCodeTemp & 0x00000001);
			nCodeTemp	= nCodeTemp >> 1;
		}

		HuffOutput[i].nCode = nCode;			// Write the final code
	}

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
// CompressedStreamWriteLiteral()
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::CompressedStreamWriteLiteral(UINT nChar)
{
	// Write the huffman code for this value
	CompressedStreamWriteBits(m_HuffmanLiteralOutput[nChar].nCode, m_HuffmanLiteralOutput[nChar].nNumBits);

	// Update the frequency of this character
	m_HuffmanLiteralTree[nChar].nFrequency++;
	--m_nHuffmanLiteralsLeft;

	// If we have coded enough literals, then generate/regenerate the huffman tree
	if (!m_nHuffmanLiteralsLeft)
	{
		if (m_bHuffmanLiteralFullyActive)
		{
			m_nHuffmanLiteralsLeft	= JB01_HUFF_LITERAL_DELAY;
			HuffmanGenerate(m_HuffmanLiteralTree, m_HuffmanLiteralOutput, JB01_HUFF_LITERAL_ALPHABETSIZE, JB01_HUFF_LITERAL_FREQMOD);
		}
		else
		{
			m_nHuffmanLiteralIncrement += JB01_HUFF_LITERAL_INITIALDELAY;
			if (m_nHuffmanLiteralIncrement >= JB01_HUFF_LITERAL_DELAY)
				m_bHuffmanLiteralFullyActive = true;

			m_nHuffmanLiteralsLeft	= JB01_HUFF_LITERAL_INITIALDELAY;
			HuffmanGenerate(m_HuffmanLiteralTree, m_HuffmanLiteralOutput, JB01_HUFF_LITERAL_ALPHABETSIZE, 0);
		}
	}

} // CompressedStreamWriteLiteral()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamWriteLen()
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::CompressedStreamWriteLen(const UINT nLength)
{
	UINT	nCode;
	UINT	nExtraBits;
	UINT	nMSBValue;
	UINT	nValue;

	// First 8 bits are unique, then goes in 4 codes per bit
	if (nLength <= 7)
	{
		nValue		= nLength;
		nCode		= nLength + 256;
		nExtraBits	= 0;
	}
	else
	{
		nExtraBits	= 1;						// Base number of bits
		nMSBValue	= 1 << (nExtraBits+2);		// MSB value
		nValue		= nLength - nMSBValue;
		nCode		= 264;						// Base code

		while (nValue >= (nMSBValue))
		{
			++nExtraBits;
			nMSBValue	= nMSBValue << 1;			// MSB value
			nCode		= nCode + 4;				// Base code
			nValue		= nValue - (nMSBValue>>1);
		}

		// Now work out the proper code
		nValue		= nLength - nMSBValue;
		nCode		= (nValue>>nExtraBits) + nCode;
		nValue		= nValue & ((1 << nExtraBits)-1);
	}

	// Write the huffman code for this value
	CompressedStreamWriteBits(m_HuffmanLiteralOutput[nCode].nCode, m_HuffmanLiteralOutput[nCode].nNumBits);

	// Now write out the remaining value and bits
	if (nExtraBits)
		CompressedStreamWriteBits(nValue, nExtraBits);

	// Update the frequency of this character
	m_HuffmanLiteralTree[nCode].nFrequency++;
	--m_nHuffmanLiteralsLeft;

	// If we have coded enough literals, then generate/regenerate the huffman tree
	if (!m_nHuffmanLiteralsLeft)
	{
		if (m_bHuffmanLiteralFullyActive)
		{
			m_nHuffmanLiteralsLeft	= JB01_HUFF_LITERAL_DELAY;
			HuffmanGenerate(m_HuffmanLiteralTree, m_HuffmanLiteralOutput, JB01_HUFF_LITERAL_ALPHABETSIZE, JB01_HUFF_LITERAL_FREQMOD);
		}
		else
		{
			m_nHuffmanLiteralIncrement += JB01_HUFF_LITERAL_INITIALDELAY;
			if (m_nHuffmanLiteralIncrement >= JB01_HUFF_LITERAL_DELAY)
				m_bHuffmanLiteralFullyActive = true;

			m_nHuffmanLiteralsLeft	= JB01_HUFF_LITERAL_INITIALDELAY;
			HuffmanGenerate(m_HuffmanLiteralTree, m_HuffmanLiteralOutput, JB01_HUFF_LITERAL_ALPHABETSIZE, 0);
		}
	}

} // CompressedStreamWriteLen()


///////////////////////////////////////////////////////////////////////////////
// CompressedStreamWriteOffset()
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::CompressedStreamWriteOffset(UINT nOffset)
{
	UINT	nCode;
	UINT	nExtraBits;
	UINT	nMSBValue;
	UINT	nValue;


	// First 4 bits are unique, then goes in 2 codes per bit
	if (nOffset <= 3)
	{
		nValue		= nOffset;
		nCode		= nOffset;
		nExtraBits	= 0;
	}
	else
	{
		nExtraBits	= 1;						// Base number of bits
		nMSBValue	= 1 << (nExtraBits+1);		// MSB value
		nValue		= nOffset - nMSBValue;
		nCode		= 4;						// Base code

		while (nValue >= (nMSBValue))
		{
			nExtraBits++;
			nMSBValue	= nMSBValue << 1;			// MSB value
			nCode		= nCode + 2;				// Base code
			nValue		= nValue - (nMSBValue>>1);
		}

		// Now work out the proper code
		nValue		= nOffset - nMSBValue;
		nCode		= (nValue>>nExtraBits) + nCode;
		nValue		= nValue & ((1 << nExtraBits)-1);
	}

	// Write the huffman code for this value
	CompressedStreamWriteBits(m_HuffmanOffsetOutput[nCode].nCode, m_HuffmanOffsetOutput[nCode].nNumBits);

	// Now write out the remaining value and bits
	if (nExtraBits)
		CompressedStreamWriteBits(nValue, nExtraBits);

	// Update the frequency of this character
	m_HuffmanOffsetTree[nCode].nFrequency++;
	--m_nHuffmanOffsetsLeft;

	// If we have coded enough literals, then generate/regenerate the huffman tree
	if (!m_nHuffmanOffsetsLeft)
	{
		if (m_bHuffmanOffsetFullyActive)
		{
			m_nHuffmanOffsetsLeft	= JB01_HUFF_OFFSET_DELAY;
			HuffmanGenerate(m_HuffmanOffsetTree, m_HuffmanOffsetOutput, JB01_HUFF_OFFSET_ALPHABETSIZE, JB01_HUFF_OFFSET_FREQMOD);
		}
		else
		{
			m_nHuffmanOffsetIncrement += JB01_HUFF_OFFSET_INITIALDELAY;
			if (m_nHuffmanOffsetIncrement >= JB01_HUFF_OFFSET_DELAY)
				m_bHuffmanOffsetFullyActive = true;

			m_nHuffmanOffsetsLeft	= JB01_HUFF_OFFSET_INITIALDELAY;
			HuffmanGenerate(m_HuffmanOffsetTree, m_HuffmanOffsetOutput, JB01_HUFF_OFFSET_ALPHABETSIZE, 0);
		}
	}

} // CompressedStreamWriteOffset()


///////////////////////////////////////////////////////////////////////////////
// FindMatches()
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::FindMatches(ULONG nInitialDataPos, ULONG &nOffset, UINT &nLen, UINT nBestLen)
{
	// m_nDataSize is the same as the end position of our input, so don't go past this boundary...

	ULONG	nBestOffset;
	UINT	nTempLen;
	struct	JB01_Hash *lpTempHash;
	UINT	nHash;
	ULONG	nPos1, nPos2;
	ULONG	nTempWPos, nDPos;
//	ULONG	nTooOldPos;

	// Reset all variables
	nBestOffset = 0;

	// Get our window start position, if the window would take us beyond
	// the start of the file, just use 0
//	if (nInitialDataPos >= JB01_MAXWINDOWLENGTH )
//		nTooOldPos = nInitialDataPos - JB01_MAXWINDOWLENGTH;
//	else
//		nTooOldPos = 0;

	// Generate a hash of the next three chars
	nHash = m_bData[nInitialDataPos & JB01_DATA_MASK];
	nHash = nHash ^ (m_bData[(nInitialDataPos+1) & JB01_DATA_MASK] << 7);
	nHash = nHash ^ (m_bData[(nInitialDataPos+2) & JB01_DATA_MASK] << 11);
	nHash = nHash & 0x0000ffff;


	// Main loop
	lpTempHash = m_HashTable[nHash];			// Get our match from the hash table

	while (lpTempHash && (nBestLen < JB01_MAXMATCHLEN))
	{
		nTempWPos = lpTempHash->nPos;

		// Is this entry too old, any remaining entries will also be too old!
//		if (nTempWPos < nTooOldPos)
//		{
//			printf("Too old entry found - hash table broken!\n");
//			break;
//		}

		// First, check the byte at the current bestlen+1 match, if this doesn't
		// match then let's not waste any time with it.  Best len +1 because if the
		// match is the _same_ as the current bestlen then we wouldn't use it anyway as we
		// favour nearer matches (current bestlen will be nearest)
		if (nBestLen)
		{
			nPos1 = nTempWPos + nBestLen;
			nPos2 = nInitialDataPos + nBestLen;
			if ( (nPos2 < m_nDataSize) && (m_bData[nPos1 & JB01_DATA_MASK] != m_bData[nPos2 & JB01_DATA_MASK]) )
			{
				lpTempHash = lpTempHash->lpNext;
				continue;
			}
		}

		nTempLen	= 0;
		nDPos		= nInitialDataPos;

		// See how many bytes match
		while ( (nDPos < m_nDataSize) && (nTempLen < JB01_MAXMATCHLEN) )
		{
			if (m_bData[nTempWPos & JB01_DATA_MASK] != m_bData[nDPos & JB01_DATA_MASK])
				break;

			++nTempWPos;
			++nDPos;
			++nTempLen;
		}

		// See if this match was better than previous match (favor small offsets)
		// Current best match will be "closer" than the next match of same length
		if (nTempLen > nBestLen)
		{
			nBestLen	= nTempLen;
			nBestOffset = lpTempHash->nPos;
		}

		lpTempHash = lpTempHash->lpNext;

	} // End while


	// Setup our return values of bestoffset and bestlen
	if (nBestLen < JB01_MINMATCHLEN)
		nLen = 0;								// Not good enough, same as no match at all
	else
	{
		nOffset	= nInitialDataPos - nBestOffset;	// Return value
		nLen		= nBestLen;						// Return value
	}

} // FindMatches()


///////////////////////////////////////////////////////////////////////////////
// HashTableInit()
///////////////////////////////////////////////////////////////////////////////

void JB01_Compress::HashTableInit(void)
{
	UINT i;

	// Clear the hash table
	for (i=0; i<JB01_HASHTABLE_SIZE; ++i)
		m_HashTable[i] = NULL;

	for (i=0; i<JB01_HASHTABLE_SIZE; ++i)
		m_HashChainCounts[i] = 0;

	// Pre-allocate all our hash entries
	m_nHashEntriesFree	= m_nHashEntriesMax;
	for (i=0; i<m_nHashEntriesMax; ++i)
		m_HashMemAllocTable[i] = (struct JB01_Hash *)((UCHAR *)m_HashEntryMemPool + (i * sizeof(struct JB01_Hash)));

} // HashTableInit()


///////////////////////////////////////////////////////////////////////////////
// HashTableAdd()
// Adds a hash entry of "nPos" at index "nHash"
///////////////////////////////////////////////////////////////////////////////

inline void JB01_Compress::HashTableAdd(UINT nBytes)
{
	struct JB01_Hash	*lpTempHash, *lpTempLast;
	UINT				nHash;
	ULONG				nOldestPos;

	while (nBytes--)
	{

		// First we delete the hash at the oldest postion (the hashadd function is used
		// just before incrementing m_nDataPos so this entry is about to leave the sliding
		// window, -1 so that a lazy findmatch can never be given an entry too old. (This
		// means that our max offset will actually be MAXWINDOWLENGTH-1 but this is a small
		// price to pay to avoid "is this entry too old" checks
		// Get the start of the window offset (to help remove old hash entries)
		if (m_nDataPos >= (JB01_MAXWINDOWLENGTH-1) )
		{
			nOldestPos = m_nDataPos - (JB01_MAXWINDOWLENGTH-1);

			// Get hash at oldest
			nHash = m_bData[nOldestPos & JB01_DATA_MASK];
			nHash = nHash ^ (m_bData[(nOldestPos+1) & JB01_DATA_MASK] << 7);
			nHash = nHash ^ (m_bData[(nOldestPos+2) & JB01_DATA_MASK] << 11);
			nHash = nHash & 0x0000ffff;

			lpTempHash = m_HashTable[nHash];

			if (lpTempHash)
			{
				if (lpTempHash->lpPrev->nPos <= nOldestPos)
				{
					if (lpTempHash->lpPrev->nPos < nOldestPos)
						printf("Whoops\n");

					if (lpTempHash->lpNext == NULL)
					{
						// Single entry
						//HashEntryFree(lpTempHash);
						m_HashMemAllocTable[m_nHashEntriesFree++] = lpTempHash;
						m_HashTable[nHash] = NULL;
					}
					else
					{
						// More than one entry
						lpTempLast = lpTempHash->lpPrev;
						lpTempHash->lpPrev = lpTempLast->lpPrev;
						lpTempLast->lpPrev->lpNext = NULL;	// New end of list (note, this may be the initial entry!)
						//HashEntryFree(lpTempLast);
						m_HashMemAllocTable[m_nHashEntriesFree++] = lpTempLast;
					}

					m_HashChainCounts[nHash]--;
				}
			}
		}


	//	if (!m_nHashEntriesFree)
	//		printf("FOOBAR\n");		// Our memory allocation is bad


		// Get hash at current position
		nHash = m_bData[m_nDataPos & JB01_DATA_MASK];
		nHash = nHash ^ (m_bData[(m_nDataPos+1) & JB01_DATA_MASK] << 7);
		nHash = nHash ^ (m_bData[(m_nDataPos+2) & JB01_DATA_MASK] << 11);
		nHash = nHash & 0x0000ffff;


		// Get our first entry at this hash index
		lpTempHash = m_HashTable[nHash];

		if (!lpTempHash)
		{
			// Initial entry
			//lpTempHash				= HashEntryAlloc();
			lpTempHash				= m_HashMemAllocTable[--m_nHashEntriesFree];
			lpTempHash->lpNext		= NULL;
			lpTempHash->lpPrev		= lpTempHash;	// Pointer to prev, OR if head of list, points to tail entry
			lpTempHash->nPos		= m_nDataPos;
			m_HashTable[nHash]		= lpTempHash;
			m_HashChainCounts[nHash]= 1;

			++m_nDataPos;						// Advance one byte
			--m_nLookAheadSize;					// Therefore we have one less lookahead byte
			continue;							// Next while iteration
		}

		// If we are here, there is at least one entry already, and a single entry is a special case.

		if (m_nHashChainLimit == 1)
		{
			lpTempHash->nPos = m_nDataPos;

			++m_nDataPos;						// Advance one byte
			--m_nLookAheadSize;					// Therefore we have one less lookahead byte
			continue;							// Next while iteration
		}


		// If at chain limit, or no free memory left for hashing, starting overwriting the
		// oldest entries with the new entries
		if (m_HashChainCounts[nHash] == m_nHashChainLimit)
		{
			// Get the last entry
			lpTempLast = m_HashTable[nHash]->lpPrev;

			// Make next to last entry in list the last entry
			lpTempLast->lpPrev->lpNext = NULL;

			// Use last entry, then move it to the top of the list
			lpTempLast->lpNext	= m_HashTable[nHash];
			lpTempLast->nPos	= m_nDataPos;
			m_HashTable[nHash]	= lpTempLast;

			++m_nDataPos;						// Advance one byte
			--m_nLookAheadSize;					// Therefore we have one less lookahead byte
			continue;							// Next while iteration
		}


		// Add an entry to the top of the list
		// First get a new entry
		lpTempHash = m_HashMemAllocTable[--m_nHashEntriesFree];
		lpTempHash->nPos = m_nDataPos;

		// Get last entry and modify lpNext
		lpTempLast = m_HashTable[nHash]->lpPrev;
		lpTempLast->lpNext = NULL;	// Need this for entries == 1

		// Modify current top of list so that previous is our new entry
		m_HashTable[nHash]->lpPrev = lpTempHash;
		lpTempHash->lpNext = m_HashTable[nHash];
		lpTempHash->lpPrev = lpTempLast;

		// Make new entry top of list
		m_HashTable[nHash] = lpTempHash;
		m_HashChainCounts[nHash]++;

		++m_nDataPos;							// Advance one byte
		--m_nLookAheadSize;						// Therefore we have one less lookahead byte

	} // End While

} // HashTableAdd()

