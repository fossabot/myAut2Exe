#ifndef __JB01_COMPRESS_H
#define __JB01_COMPRESS_H

//
// jb01_compress.h
// (c)2002-2003 Jonathan Bennett (jon@hiddensoft.com)
//
// STANDALONE CLASS
//
// This is an implementation of LZSS compression
//
// Note, this is a "as simple to follow as possible" piece of code, there are many
// more efficient and quicker (and cryptic) methods but this was done as a learning
// exercise.
//

//
// LICENSE
// This code is free to use, but please do me the decency of giving me credit :)
//

// The start of the compression stream will contain "other" items apart from data.
//
// Algorithm ID				4 bytes
// Revision ID				1 byte
// Total Uncompressed Size	4 bytes (1 ULONG)
// Compressed Data			nn bytes (16 bit aligned)
//


// Generic defines (same accross multiple compressors, LZSS, LZP, etc)
#define	HS_COMP_FILE	0						// Source/Dest is a file
#define	HS_COMP_MEM		1						// Source/Dest is memory

#define HS_COMP_SafeFree(x) { if (x) free(x); }



// Alg ID
#define JB01_ALGID			"JB01"				// Algorithm ID (4 bytes)

// Error codes
#define	JB01_E_OK			0					// OK
#define JB01_E_NOTJB01		1					// Not a valid LZSS data stream
#define JB01_E_READINGSRC	2					// Error reading source file
#define JB01_E_WRITINGDST	3					// Error writing dest file
#define JB01_E_ABORT		4					// Operation was aborted
#define JB01_E_MEMALLOC		5					// Couldn't allocate all the mem we needed


// Working block sizes (in bytes)
#define JB01_DATA_SIZE		(128*1024)			// Uncompressed buffer - must be >window+lookahead+1 and a POWER OF 2 (for ANDing)
#define JB01_DATA_MASK		(JB01_DATA_SIZE-1)	// AND mask (data size MUST be a power of 2)


// Define our hash and hash table structure
#define JB01_HASHORDER		3					// Order-3 hashing
#define JB01_HASHTABLE_SIZE	65536				// Our hash function gives values 0-65535

typedef struct JB01_Hash
{
	ULONG	nPos;								// ABSOLUTE (non-ANDed) position in data stream
	struct 	JB01_Hash *lpNext;					// Next entry in linked list or NULL
	struct	JB01_Hash *lpPrev;

} _JB01_Hash;


// Define huffman and related
#define JB01_MINMATCHLEN				3			// Minimum match length

#define	JB01_HUFF_LITERAL_ALPHABETSIZE	(256 + 32)	// Number of characters to represent literals(256) + match lengths (32 match len codes, 0-511)
#define JB01_HUFF_LITERAL_LENSTART		256			// Match lens are elements 256 and above in the literal alphabet
#define JB01_MAXMATCHLEN				(511 + JB01_MINMATCHLEN)	// Maximum match length (511 lengths over 32 codes)

#define	JB01_HUFF_OFFSET_ALPHABETSIZE	32			// Number of characters to represent offset codes (32 codes = 0-65535)
#define JB01_MAXWINDOWLENGTH			65535 		// Sliding window size (must be <=65535)

#define JB01_HUFF_LITERAL_INITIALDELAY	(JB01_HUFF_LITERAL_ALPHABETSIZE / 4)	// Literal trees are initially regenerated after this many codings
#define JB01_HUFF_LITERAL_DELAY			(JB01_HUFF_LITERAL_ALPHABETSIZE * 12)	// Literal trees are thereafter regenerated after this many codings

#define JB01_HUFF_OFFSET_INITIALDELAY	(JB01_HUFF_OFFSET_ALPHABETSIZE / 4)		// Offset trees are initially regenerated after this many codings
#define JB01_HUFF_OFFSET_DELAY			(JB01_HUFF_OFFSET_ALPHABETSIZE * 12)	// Offset trees are thereafter regenerated after this many codings

#define	JB01_HUFF_LITERAL_FREQMOD		1			// Frequency modifier for old values
#define	JB01_HUFF_OFFSET_FREQMOD		1			// Frequency modifier for old values

#define JB01_HUFF_MAXCODEBITS			16			// Maximum number of bits for a huffman code (must be <=16 for bitwriting functions)

typedef struct JB01_HuffmanNode
{
	ULONG	nFrequency;							// Frequency value
	bool	bSearchMe;
	UINT	nParent;
	UINT	nChildLeft;							// Left child node or NULL
	UINT	nChildRight;						// Right child node or NULL
	UCHAR	cValue;

} _JB01_HuffmanNode;

typedef struct JB01_HuffmanOutput
{
	UINT	nNumBits;							// Code length
	ULONG	nCode;								// Code (max 32bits)
} _JB01_HuffmanOutput;



// Monitor function declaration
typedef int (*JB01_MonitorProc)(ULONG nBytesIn, ULONG nBytesOut, UINT nPercentComplete);



class JB01_Compress
{
public:
	// Functions
	JB01_Compress() { SetDefaults(); }			// Constructor

	int		Compress();

	void	SetDefaults(void);
	void	SetInputType(UINT nInput) { m_nInputType = nInput; }
	void	SetOutputType(UINT nOutput) { m_nOutputType = nOutput; }
	void	SetInputFile(const char *szSrc) { strcpy(m_szSrcFile, szSrc); }
	void	SetOutputFile(const char *szDst) { strcpy(m_szDstFile, szDst); }
	void	SetInputBuffer(UCHAR *bBuf, ULONG nSize) { m_bUserData = bBuf; m_nDataSize = nSize; }
	void	SetOutputBuffer(UCHAR *bBuf) { m_bUserCompData = bBuf; }
	void	SetCompressionLevel(UINT nCompressionLevel);	// Set the hash chain limit
	void	SetMonitorCallback(JB01_MonitorProc lpfnMonitor) { m_lpfnMonitor = lpfnMonitor; }

	ULONG	GetCompressedSize(void) { return m_nUserCompPos; }

	// Monitor functions
//	UINT	GetPercentComplete(void) { return ((UINT)(((double)m_nUserDataPos/(double)m_nDataSize) * 100.0)); }
	UINT	GetPercentComplete(void) { return ((m_nUserDataPos / m_nDataSize) * 100); }

private:
	// User supplied buffers and counters
	UCHAR	*m_bUserData;						// When compressing from memory this is the user input buffer
	UCHAR	*m_bUserCompData;					// When compressing to memory this is the user output buffer
	ULONG	m_nUserDataPos;						// Position in user uncompressed stream (also used for info)
	ULONG	m_nUserCompPos;						// Position in user compressed stream (also used for info)

	// Master variables
	ULONG	m_nDataSize;						// TOTAL file uncompressed size

	UINT	m_nInputType, m_nOutputType;		// File or memory output and input

	FILE	*m_fSrc, *m_fDst;					// Source and destination filepointers
	char	m_szSrcFile[_MAX_PATH+1];
	char	m_szDstFile[_MAX_PATH+1];

	UCHAR	*m_bData;
	ULONG	m_nDataPos;							// Position in input stream
	ULONG	m_nLookAheadSize;					// How "lookahead" data is available
	ULONG	m_nDataReadPos;						// Where new data should be read into

	UCHAR	*m_bComp;
	ULONG	m_nCompPos;							// Temp position holder for internal compressed data buffer


	// Misc
	bool	m_bAbortRequested;

	// Temporary variables used for the bit operations
	ULONG	m_nCompressedLong;					// Compressed stream temporary 32bit value
	USHORT	m_nCompressedBitsFree;				// Number of bits unused in temporary value

	// Hash table related variables
	UINT	m_nHashChainLimit;					// The max length of each hash chain
	struct	JB01_Hash **m_HashTable;			// Hash table
	UINT	*m_HashChainCounts;					// Number of entries in each hash chain
	UINT	m_nHashEntriesMax;					// Maximum number of hash entries allowed
	UINT	m_nHashEntriesFree;					// Number of free hash positions we have
	void	*m_HashEntryMemPool;				// Pointer to the single block of memory allocated
	struct	JB01_Hash **m_HashMemAllocTable;	// Our table to keep track of memory allocated


	// Huffman variables
	struct	JB01_HuffmanNode *m_HuffmanLiteralTree;		// The huffman literal/len tree
	struct	JB01_HuffmanOutput *m_HuffmanLiteralOutput;	// Output buffer for chars
	ULONG	m_nHuffmanLiteralsLeft;				// Number of literals before huffman regenerated
	bool	m_bHuffmanLiteralFullyActive;
	ULONG	m_nHuffmanLiteralIncrement;

	struct	JB01_HuffmanNode *m_HuffmanOffsetTree;		// The huffman offset tree
	struct	JB01_HuffmanOutput *m_HuffmanOffsetOutput;	// Output buffer for chars
	ULONG	m_nHuffmanOffsetsLeft;				// Number of literals before huffman regenerated
	bool	m_bHuffmanOffsetFullyActive;
	ULONG	m_nHuffmanOffsetIncrement;

	// Monitor variables and related
	JB01_MonitorProc m_lpfnMonitor;				// The monitor callback function (or NULL)


	// Functions
	ULONG		GetFileSize(const char *szFile);
	int			AllocMem(void);
	void		FreeMem(void);
	inline void	WriteUserCompData(void);		// Write compressed data to file/mem
	inline void	ReadUserData(void);				// Read data from file or memory

	int		CompressLoop(void);					// The main compression loop
	void	FindMatches(ULONG nInitialDataPos, ULONG &nOffset, UINT &nLen, UINT nBestLen);	// Searches for pattern matches


	// Bit operation functions
	inline	void	CompressedStreamWriteBits(UINT nValue, UINT nNumBits);
	void			CompressedStreamWriteBitsFlush(void);
	inline	void	CompressedStreamWriteLiteral(UINT nChar);
	inline	void	CompressedStreamWriteLen(UINT nChar);
	inline	void	CompressedStreamWriteOffset(UINT nOffset);

	// Hash table functions
	void			HashTableInit(void);				// Make hash table ready for first use
	inline	void	HashTableAdd(UINT nBytes);			// Add an entry to the table

	// Huffman functions
	void	HuffmanInit(void);
	void	HuffmanZero(JB01_HuffmanNode *HuffTree, UINT nAlphabetSize);	// Clears the frequencies in the huffman tree
	void	HuffmanGenerate(JB01_HuffmanNode *HuffTree, JB01_HuffmanOutput *HuffOutput, UINT nAlphabetSize, UINT nFreqMod);

	// Monitor functions
	inline	void	MonitorCallback(void);
};


#endif
