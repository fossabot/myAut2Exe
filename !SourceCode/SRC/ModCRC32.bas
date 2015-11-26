Attribute VB_Name = "ModCRC32"
Option Explicit
' Src based on ... http://vb-tec.de/crc.htm
Private pInititialized As Boolean
Private pTable(0 To 255) As Long
Const Reserved& = 0
Private Declare Function Mul Lib "MSVBVM60.DLL" Alias "_allmul" (ByVal dw1 As Long, ByVal Reserved As Long, ByVal dw3 As Long, ByVal Reserved As Long) As Long

Private m_l2Power(0 To 30) As Long
Private m_lOnBits(0 To 30) As Long

Public Sub CRCInit(Optional ByVal Poly As Long = &HEDB88320)
   
' init for Lshift...
  Class_Initialize
  
  Dim CRC As Long
  Dim int8 As Integer
  Dim Bits As Integer
  
  For int8 = 0 To 255
  
    CRC = Mul(int8, 0, &H1000000, 0) '* (2 ^ &H18)
    For Bits = 7 To 0 Step -1
    
      'crc32 & 0x80000000) ?
      If CRC < 0 Then
        'CRC = (CRC << 1) ^ Poly
        CRC = Mul(CRC, 0, 2, 0) Xor Poly
      Else
        'CRC = (CRC << 1)
        CRC = Mul(CRC, 0, 2, 0)
      End If
    
    Next Bits
    pTable(int8) = CRC
  
  Next int8
  pInititialized = True
End Sub


Public Function CRC32(ByRef Bytes() As Byte) As Long
  
  If Not pInititialized Then CRCInit
' BenchStart
  ' CRC berechnen:
  CRC32 = &HFFFFFFFF
  Dim i As Long
  For i = LBound(Bytes) To UBound(Bytes)
  
    'CRC = (CRC << 0x18) ^ pTable[Bytes[i] ^ (CRC >> 0x08)]
'    CRC32 = LShift(CRC32, &H8) Xor pTable(Bytes(i) Xor RShift(CRC32, &H18) And &HFF&)
  
    DoEventsVerySeldom
    CRC32 = Mul(CRC32, 0, &H100, 0) Xor pTable((Bytes(i) Xor _
    RShift(CRC32, &H18) And &HFF&))
  
  Next i
'BenchEnd
End Function




Private Sub Class_Initialize()
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
End Sub


'*******************************************************************************
' LShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' A left shift takes all the set binary bits and moves them left, in-filling
' with zeros in the vacated bits on the right. This function is equivalent to
' the << operator in Java and C++
'*******************************************************************************
Public Function LShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    '
    ' A shift of zero is no shift at all.
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
        
    ' A shift of 31 will result in the right most bit becoming the left most
    ' bit and all other bits being cleared
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
        
    ' A shift of less than zero or more than 31 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    ' If the left most bit that remains will end up in the negative bit
    ' position (&H80000000) we would end up with an overflow if we took the
    ' standard route. We need to strip the left most bit and add it back
    ' afterwards.
    If (lValue And m_l2Power(31 - iShiftBits)) Then
    
        ' (Value And OnBits(31 - (Shift + 1))) chops off the left most bits that
        ' we are shifting into, but also the left most bit we still want as this
        ' is going to end up in the negative bit marker position (&H80000000).
        ' After the multiplication/shift we Or the result with &H80000000 to
        ' turn the negative bit on.
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * _
            m_l2Power(iShiftBits)) Or &H80000000
    
    Else
    
        ' (Value And OnBits(31-Shift)) chops off the left most bits that we are
        ' shifting into so we do not get an overflow error when we do the
        ' multiplication/shift
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * _
            m_l2Power(iShiftBits))
        
    End If
End Function

'*******************************************************************************
' RShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' The right shift of an unsigned long integer involves shifting all the set bits
' to the right and in-filling on the left with zeros. This function is
' equivalent to the >>> operator in Java or the >> operator in C++ when used on
' an unsigned long.
'*******************************************************************************
Public Function RShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    '
    ' A shift of zero is no shift at all
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
        
    ' A shift of 31 will clear all bits and move the left most bit to the right
    ' most bit position
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
        
    ' A shift of less than zero or more than 31 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    ' We do not care about the top most bit or the final bit, the top most bit
    ' will be taken into account in the next stage, the final bit (whether it
    ' is an odd number or not) is being shifted into, so we do not give a jot
    ' about it
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
    ' If the top most bit (&H80000000) was set we need to do things differently
    ' as in a normal VB signed long integer the top most bit is used to indicate
    ' the sign of the number, when it is set it is a negative number, so just
    ' deviding by a factor of 2 as above would not work.
    ' NOTE: (lValue And  &H80000000) is equivalent to (lValue < 0), you could
    ' get a very marginal speed improvement by changing the test to (lValue < 0)
    If (lValue And &H80000000) Then
        ' We take the value computed so far, and then add the left most negative
        ' bit after it has been shifted to the right the appropriate number of
        ' places
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function



