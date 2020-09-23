Attribute VB_Name = "SHA_256_Optimized"

'*******************************************************************************
' MODULE:            SHA256Optimized
' FILENAME:          SHA256Optimized.cls
' AUTHOR:            Phil Fresle
' CREATED:          10-Apr-2001
' COPYRIGHT:       Copyright 2001 Phil Fresle. All Rights Reserved.

' OPTIMIZATION:   DAVID SVAITER
' AS OF:                21-Nov-2001
'
' MODIFICATION HISTORY:
'
' 21-Nov-2001      David Svaiter                Optimized loops and math functions (Math DLL) based
'                                                              on pure ASM code.
'
' 10-Apr-2001       Phil Fresle                     Initial Version
'*******************************************************************************
Option Explicit



Private Declare Function LSHL Lib "MATH.DLL" (ByRef valor As Long, ByRef ToLeft As Long) As Long
Private Declare Function LSHR Lib "MATH.DLL" (ByRef valor As Long, ByRef ToLeft As Long) As Long
Private Declare Function LRL Lib "MATH.DLL" (ByRef valor As Long, ByRef ToLeft As Long) As Long       '  rotaciona "n" vezes
Private Declare Function LRR Lib "MATH.DLL" (ByRef valor As Long, ByRef ToLeft As Long) As Long
Private Declare Function LRLb Lib "MATH.DLL" (ByRef valor As Long, ByRef ToLeft As Long) As Long     '  rotaciona CL
Private Declare Function LRRb Lib "MATH.DLL" (ByRef valor As Long, ByRef ToLeft As Long) As Long
Private Declare Function ADDu Lib "MATH.DLL" (ByRef valor As Long, ByRef valor2 As Long) As Long

Private Declare Function GetInputState Lib "user32" () As Long



Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long
Private K(63)           As Long




Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE

'*******************************************************************************
' Class_Initialize (SUB)
'*******************************************************************************
Private Sub Initialize()

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
    
    ' Just put together the K array once
    K(0) = &H428A2F98
    K(1) = &H71374491
    K(2) = &HB5C0FBCF
    K(3) = &HE9B5DBA5
    K(4) = &H3956C25B
    K(5) = &H59F111F1
    K(6) = &H923F82A4
    K(7) = &HAB1C5ED5
    K(8) = &HD807AA98
    K(9) = &H12835B01
    K(10) = &H243185BE
    K(11) = &H550C7DC3
    K(12) = &H72BE5D74
    K(13) = &H80DEB1FE
    K(14) = &H9BDC06A7
    K(15) = &HC19BF174
    K(16) = &HE49B69C1
    K(17) = &HEFBE4786
    K(18) = &HFC19DC6
    K(19) = &H240CA1CC
    K(20) = &H2DE92C6F
    K(21) = &H4A7484AA
    K(22) = &H5CB0A9DC
    K(23) = &H76F988DA
    K(24) = &H983E5152
    K(25) = &HA831C66D
    K(26) = &HB00327C8
    K(27) = &HBF597FC7
    K(28) = &HC6E00BF3
    K(29) = &HD5A79147
    K(30) = &H6CA6351
    K(31) = &H14292967
    K(32) = &H27B70A85
    K(33) = &H2E1B2138
    K(34) = &H4D2C6DFC
    K(35) = &H53380D13
    K(36) = &H650A7354
    K(37) = &H766A0ABB
    K(38) = &H81C2C92E
    K(39) = &H92722C85
    K(40) = &HA2BFE8A1
    K(41) = &HA81A664B
    K(42) = &HC24B8B70
    K(43) = &HC76C51A3
    K(44) = &HD192E819
    K(45) = &HD6990624
    K(46) = &HF40E3585
    K(47) = &H106AA070
    K(48) = &H19A4C116
    K(49) = &H1E376C08
    K(50) = &H2748774C
    K(51) = &H34B0BCB5
    K(52) = &H391C0CB3
    K(53) = &H4ED8AA4A
    K(54) = &H5B9CCA4F
    K(55) = &H682E6FF3
    K(56) = &H748F82EE
    K(57) = &H78A5636F
    K(58) = &H84C87814
    K(59) = &H8CC70208
    K(60) = &H90BEFFFA
    K(61) = &HA4506CEB
    K(62) = &HBEF9A3F7
    K(63) = &HC67178F2
End Sub


'*******************************************************************************
' SHA256 (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sMessage - String - Message to digest
'
' RETURN VALUE:
' String - The digest
'
' DESCRIPTION:
' Takes a string and uses the SHA-256 digest to produce a signature for it.
'
' NOTE: Due to the way in which the string is processed the routine assumes a
' single byte character set. VB passes unicode (2-byte) character strings, the
' ConvertToWordArray function uses on the first byte for each character. This
' has been done this way for ease of use, to make the routine truely portable
' you could accept a byte array instead, it would then be up to the calling
' routine to make sure that the byte array is generated from their string in
' a manner consistent with the string type.
'*******************************************************************************
Public Function SHA256o(sMessage As String, Optional status As Boolean = False) As String


    Initialize
    
    Dim Hash(7) As Long
    Dim w(63)   As Long
    Dim a       As Long
    Dim b       As Long
    Dim c       As Long
    Dim d       As Long
    Dim e       As Long
    Dim F       As Long
    Dim g       As Long
    Dim h       As Long
    Dim i       As Long
    Dim j       As Long
    Dim t1      As Long
    Dim t2      As Long
    
    
    
    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long
    
    Const MODULUS_BITS      As Long = 512
    Const CONGRUENT_BITS    As Long = 448
    
    Dim NextPercent As Variant, AtualPercent As Variant
    
    
    
    
    
    ' Initial hash values
    Hash(0) = &H6A09E667
    Hash(1) = &HBB67AE85
    Hash(2) = &H3C6EF372
    Hash(3) = &HA54FF53A
    Hash(4) = &H510E527F
    Hash(5) = &H9B05688C
    Hash(6) = &H1F83D9AB
    Hash(7) = &H5BE0CD19
    
    ' Preprocessing. Append padding bits and length and convert to words
    
    lMessageLength = Len(sMessage)
    
    
    On Local Error Resume Next
    
    Dim total As Long, Passo As Long
    
    
    ' Get padded number of words. Message needs to be congruent to 448 bits,
    ' modulo 512 bits. If it is exactly congruent to 448 bits, modulo 512 bits
    ' it must still have another 512 bits added. 512 bits = 64 bytes
    ' (or 16 * 4 byte words), 448 bits = 56 bytes. This means lNumberOfWords must
    ' be a multiple of 16 (i.e. 16 * 4 (bytes) * 8 (bits))
    lNumberOfWords = (((lMessageLength + _
        ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ _
        (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * _
        (MODULUS_BITS \ BITS_TO_A_WORD)
    
    ReDim lWordArray(lNumberOfWords - 1)
    
    ' Combine each block of 4 bytes (ascii code of character) into one long
    ' value and store in the message. The high-order (most significant) bit of
    ' each byte is listed first. However, unlike MD5 we put the high-order
    ' (most significant) byte first in each word.
    lBytePosition = 0
    lByteCount = 0
    
    
    Passo = 1
    total = lMessageLength
    
    
    
    
    
    Do Until lByteCount >= lMessageLength
        ' Each word is 4 bytes
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        
        ' NOTE: This is where we are using just the first byte of each unicode
        ' character, you may want to make the change here, or to the SHA256 method
        ' so it accepts a byte array.
        
        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LSHL(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    
    
       If GetInputState() <> 0 Then DoEvents
      
    Loop







    ' Terminate according to SHA-256 rules with a 1 bit, zeros and the length in
    ' bits stored in the last two words
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    ' Add a terminating 1 bit, all the rest of the bits to the end of the
    ' word array will default to zero
    
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LSHL(&H80, lBytePosition)
    

    ' We put the length of the message in bits into the last two words, to get
    ' the length in bits we need to multiply by 8 (or left shift 3). This left
    ' shifted value is put in the last word. Any bits shifted off the left edge
    ' need to be put in the penultimate word, we can work out which bits by shifting
    ' right the length by 29 bits.
    
    lWordArray(lNumberOfWords - 1) = LSHL(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = LSHR(lMessageLength, 29)
    
    
'    lWordArray(lNumberOfWords - 1) = LSHL(lMessageLength, 3)
'     lWordArray(lNumberOfWords - 2) = LSHR(lMessageLength, 29)
    
    
    
    
    
    
    
    
    ' If BolaStatus = True And Status = True Then BotaBola
      
    
    total = UBound(lWordArray)
    Passo = 1
    
    ' Main loop
    For i = 0 To total Step 16
        a = Hash(0)
        b = Hash(1)
        c = Hash(2)
        d = Hash(3)
        e = Hash(4)
        F = Hash(5)
        g = Hash(6)
        h = Hash(7)
        
        For j = 0 To 63
            If j < 16 Then
                w(j) = lWordArray(j + i)
            Else
                w(j) = ADDu(ADDu(ADDu((LRRb(w(j - 2), 17) Xor LRRb(w(j - 2), 19) Xor LSHR(w(j - 2), 10)), w(j - 7)), (LRRb(w(j - 15), 7) Xor LRRb(w(j - 15), 18) Xor LSHR(w(j - 15), 3))), w(j - 16))
            End If
                
            t1 = ADDu(ADDu(ADDu(ADDu(h, (LRRb(e, 6) Xor LRRb(e, 11) Xor LRRb(e, 25))), ((e And F) Xor ((Not e) And g))), K(j)), w(j))
            t2 = ADDu((LRRb(a, 2) Xor LRRb(a, 13) Xor LRRb(a, 22)), ((a And b) Xor (a And c) Xor (b And c)))
            
            h = g
            g = F
            F = e
            e = ADDu(d, t1)
            d = c
            c = b
            b = a
            a = ADDu(t1, t2)
            
        Next
        
        Hash(0) = ADDu(a, Hash(0))
        Hash(1) = ADDu(b, Hash(1))
        Hash(2) = ADDu(c, Hash(2))
        Hash(3) = ADDu(d, Hash(3))
        Hash(4) = ADDu(e, Hash(4))
        Hash(5) = ADDu(F, Hash(5))
        Hash(6) = ADDu(g, Hash(6))
        Hash(7) = ADDu(h, Hash(7))
    
      If GetInputState() <> 0 Then
            DoEvents
      End If
    
    Next
    
    ' Output the 512 bit digest
    SHA256o = UCase(Right$("00000000" & Hex(Hash(0)), 8) & _
        Right("00000000" & Hex(Hash(1)), 8) & _
        Right("00000000" & Hex(Hash(2)), 8) & _
        Right("00000000" & Hex(Hash(3)), 8) & _
        Right("00000000" & Hex(Hash(4)), 8) & _
        Right("00000000" & Hex(Hash(5)), 8) & _
        Right("00000000" & Hex(Hash(6)), 8) & _
        Right("00000000" & Hex(Hash(7)), 8))
    




End Function



