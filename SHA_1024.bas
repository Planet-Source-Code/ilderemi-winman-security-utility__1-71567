Attribute VB_Name = "SHA_1024"
'*******************************************************************************
' MODULE:            SHA-1024
' AUTHOR:            David Svaiter
' CREATED:          28 July 2002
' COPYRIGHT:       Copyright 2002 Cypher (www.cypher.com.br)
'
'  This code calculates message digest (unique signatures for strings/values), using
'   the SHA algorithm on 1024 bits of resolution (128 bytes for each result/digest).
'
'   We use it here at CYPHER to calculates secure values inside our crypto shareware tools and
'   custom software solutions.  Feel free to use it into your programs and remember to attach
'   the MATH.DLL library, since our modules use it to get faster binary calculations.  This MATH
'   is a CYPHER property (developed code under copyrights), but you can use and distribute it freely.
'
'   We do not distribute codes in our site, but if you want privacy in your computer environment,
'   we invite you to know our cryptographic applications at:     www.cypher.com.br
'
'
'   Enjoy It !

'
'   Independent to use it freely, you must display the notice
'   SHA-1024 by CYPHER (www.cypher.com.br)
'   in your program and manuals.




'*******************************************************************************
' This code is inpired in the Phil Fresle's code (SHA256)
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
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
Private K(160)           As Long




Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE

'*******************************************************************************
' Class_Initialize (SUB)
'*******************************************************************************
Private Sub Initialize()

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
K(1) = &HD728AE22
K(2) = &H71374491
K(3) = &H23EF65CD
K(4) = &HB5C0FBCF
K(5) = &HEC4D3B2F
K(6) = &HE9B5DBA5
K(7) = &H8189DBBC
K(8) = &H3956C25B
K(9) = &HF348B538
K(10) = &H59F111F1
K(11) = &HB605D019
K(12) = &H923F82A4
K(13) = &HAF194F9B
K(14) = &HAB1C5ED5
K(15) = &HDA6D8118
K(16) = &HD807AA98
K(17) = &HA3030242
K(18) = &H12835B01
K(19) = &H45706FBE
K(20) = &H243185BE
K(21) = &H4EE4B28C
K(22) = &H550C7DC3
K(23) = &HD5FFB4E2
K(24) = &H72BE5D74
K(25) = &HF27B896F
K(26) = &H80DEB1FE
K(27) = &H3B1696B1
K(28) = &H9BDC06A7
K(29) = &H25C71235
K(30) = &HC19BF174
K(31) = &HCF692694
K(32) = &HE49B69C1
K(33) = &H9EF14AD2
K(34) = &HEFBE4786
K(35) = &H384F25E3
K(36) = &HFC19DC6
K(37) = &H8B8CD5B5
K(38) = &H240CA1CC
K(39) = &H77AC9C65
K(40) = &H2DE92C6F
K(41) = &H592B0275
K(42) = &H4A7484AA
K(43) = &H6EA6E483
K(44) = &H5CB0A9DC
K(45) = &HBD41FBD4
K(46) = &H76F988DA
K(47) = &H831153B5
K(48) = &H983E5152
K(49) = &HEE66DFAB
K(50) = &HA831C66D
K(51) = &H2DB43210
K(52) = &HB00327C8
K(53) = &H98FB213F
K(54) = &HBF597FC7
K(55) = &HBEEF0EE4
K(56) = &HC6E00BF3
K(57) = &H3DA88FC2
K(58) = &HD5A79147
K(59) = &H930AA725
K(60) = &H6CA6351
K(61) = &HE003826F
K(62) = &H14292967
K(63) = &HA0E6E70
K(64) = &H27B70A85
K(65) = &H46D22FFC
K(66) = &H2E1B2138
K(67) = &H5C26C926
K(68) = &H4D2C6DFC
K(69) = &H5AC42AED
K(70) = &H53380D13
K(71) = &H9D95B3DF
K(72) = &H650A7354
K(73) = &H8BAF63DE
K(74) = &H766A0ABB
K(75) = &H3C77B2A8
K(76) = &H81C2C92E
K(77) = &H47EDAEE6
K(78) = &H92722C85
K(79) = &H1482353B
K(80) = &HA2BFE8A1
K(81) = &H4CF10364
K(82) = &HA81A664B
K(83) = &HBC423001
K(84) = &HC24B8B70
K(85) = &HD0F89791
K(86) = &HC76C51A3
K(87) = &H654BE30
K(88) = &HD192E819
K(89) = &HD6EF5218
K(90) = &HD6990624
K(91) = &H5565A910
K(92) = &HF40E3585
K(93) = &H5771202A
K(94) = &H106AA070
K(95) = &H32BBD1B8
K(96) = &H19A4C116
K(97) = &HB8D2D0C8
K(98) = &H1E376C08
K(99) = &H5141AB53
K(100) = &H2748774C
K(101) = &HDF8EEB99
K(102) = &H34B0BCB5
K(103) = &HE19B48A8
K(104) = &H391C0CB3
K(105) = &HC5C95A63
K(106) = &H4ED8AA4A
K(107) = &HE3418ACB
K(108) = &H5B9CCA4F
K(109) = &H7763E373
K(110) = &H682E6FF3
K(111) = &HD6B2B8A3
K(112) = &H748F82EE
K(113) = &H5DEFB2FC
K(114) = &H78A5636F
K(115) = &H43172F60
K(116) = &H84C87814
K(117) = &HA1F0AB72
K(118) = &H8CC70208
K(119) = &H1A6439EC
K(120) = &H90BEFFFA
K(121) = &H23631E28
K(122) = &HA4506CEB
K(123) = &HDE82BDE9
K(124) = &HBEF9A3F7
K(125) = &HB2C67915
K(126) = &HC67178F2
K(127) = &HE372532B
K(128) = &HCA273ECE
K(129) = &HEA26619C
K(130) = &HD186B8C7
K(131) = &H21C0C207
K(132) = &HEADA7DD6
K(133) = &HCDE0EB1E
K(134) = &HF57D4F7F
K(135) = &HEE6ED178
K(136) = &H6F067AA
K(137) = &H72176FBA
K(138) = &HA637DC5
K(139) = &HA2C898A6
K(140) = &H113F9804
K(141) = &HBEF90DAE
K(142) = &H1B710B35
K(143) = &H131C471B
K(144) = &H28DB77F5
K(145) = &H23047D84
K(146) = &H32CAAB7B
K(147) = &H40C72493
K(148) = &H3C9EBE0A
K(149) = &H15C9BEBC
K(150) = &H431D67C4
K(151) = &H9C100D4C
K(152) = &H4CC5D4BE
K(153) = &HCB3E42B6
K(154) = &H597F299C
K(155) = &HFC657E2A
K(156) = &H5FCB6FAB
K(157) = &H3AD6FAEC
K(158) = &H6C44198C
K(159) = &H4A475817

End Sub


'*******************************************************************************
' SHA1024 (FUNCTION)
'
' PARAMETERS:
' (In/Out) -  String
'
' RETURN VALUE:
' String - The digest
'
' DESCRIPTION:
' Takes a string and uses the SHA-1024 digest to produce a signature for it.
'
'*******************************************************************************
Public Function SHA1024(sMessage As String, Optional status As Boolean = False) As String


    Initialize
      


    Dim Hash(31) As Long
    
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
    
    Dim M As Long
    Dim n As Long
    Dim o As Long
    Dim p As Long
    Dim q As Long
    Dim R As Long
    Dim S As Long
    Dim t As Long
    
    
    
    
    Dim a2       As Long
    Dim b2       As Long
    Dim c2      As Long
    Dim d2       As Long
    Dim e2       As Long
    Dim F2      As Long
    Dim g2       As Long
    Dim h2       As Long
    
    Dim M2 As Long
    Dim n2 As Long
    Dim o2 As Long
    Dim p2 As Long
    Dim q2 As Long
    Dim R2 As Long
    Dim S2 As Long
    Dim t22 As Long
    
    
    
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
    Hash(1) = &H13BCC908
    Hash(2) = &HBB67AE85
    Hash(3) = &H84CAA73B
    Hash(4) = &H3C6EF372
    Hash(5) = &HFE94F82B
    Hash(6) = &HA54FF53A
    Hash(7) = &H5F1D36F1
    Hash(8) = &H510E527F
    Hash(9) = &HADE682D1
    Hash(10) = &H9B05688C
    Hash(11) = &H2B3E6C1F
    Hash(12) = &H1F83D9AB
    Hash(13) = &HFB41BD6B
    Hash(14) = &H5BE0CD19
    Hash(15) = &H137E2179
    
    
    '   since there is no values to SHA-1024 (the algorithm does not exist !)
    '   we implemented random values.  You can change it to customize
    '   your SHA-1024 algorithm.
    
    Hash(16) = &H60E63737
    Hash(17) = &H832F4262
    Hash(18) = &HEBAACFAD
    Hash(19) = &HD5281EF6
    Hash(20) = &H37205732
    Hash(21) = &HA9B8C7D6
    Hash(22) = &H33247728
    Hash(23) = &H28372944
    Hash(24) = &H98765432
    Hash(25) = &H4598E482
    Hash(26) = &H2E769911
    Hash(27) = &HB6F8AA91
    Hash(28) = &H74EE1728
    Hash(29) = &H20573528
    Hash(30) = &H84297657
    Hash(31) = &HF6E99A4E
    
    
    lMessageLength = Len(sMessage)
    
    
    On Local Error Resume Next
    
    Dim total As Long, Passo As Long
    
    
    lNumberOfWords = (((lMessageLength + _
        ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ _
        (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * _
        (MODULUS_BITS \ BITS_TO_A_WORD)
    
    ReDim lWordArray(lNumberOfWords - 1)
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






    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
    
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LSHL(&H80, lBytePosition)
    
    
    lWordArray(lNumberOfWords - 1) = LSHL(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = LSHR(lMessageLength, 29)
    
    
    
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
        
        M = Hash(8)
        n = Hash(9)
        o = Hash(10)
        p = Hash(11)
        q = Hash(12)
        R = Hash(13)
        S = Hash(14)
        t = Hash(15)
        
        
        a2 = Hash(16)
        b2 = Hash(17)
        c2 = Hash(18)
        d2 = Hash(19)
        e2 = Hash(20)
        F2 = Hash(21)
        g2 = Hash(22)
        h2 = Hash(23)
        
        M2 = Hash(24)
        n2 = Hash(25)
        o2 = Hash(26)
        p2 = Hash(27)
        q2 = Hash(28)
        R2 = Hash(29)
        S2 = Hash(30)
        t22 = Hash(31)
        
        
        
        
        For j = 0 To 159
            If j < 16 Then
                w(j) = lWordArray(j + i)
            Else
                w(j) = ADDu(ADDu(ADDu((LRRb(w(j - 2), 17) Xor LRRb(w(j - 2), 19) Xor LSHR(w(j - 2), 10)), w(j - 7)), (LRRb(w(j - 15), 7) Xor LRRb(w(j - 15), 18) Xor LSHR(w(j - 15), 3))), w(j - 16))
            End If
                
            t1 = ADDu(ADDu(ADDu(ADDu(h, (LRRb(e, 6) Xor LRRb(e, 11) Xor LRRb(e, 25))), ((e And F) Xor ((Not e) And g))), K(j)), w(j))
            t2 = ADDu((LRRb(a, 2) Xor LRRb(a, 13) Xor LRRb(a, 22)), ((a And b) Xor (a And c) Xor (b And c)))
            
            
            t22 = ADDu(S, g)
            S2 = ADDu(R, F)
            R2 = ADDu(q, e)
            q2 = ADDu(p, d)
            p2 = ADDu(o, c)
            o2 = ADDu(n, b)
            n2 = ADDu(M2, t2)
            M2 = ADDu(h2, a)
            
            h2 = ADDu(g2, t2)
            g2 = F2
            F2 = e2
            e2 = d2
            d2 = ADDu(c2, t1)
            c2 = b2
            b2 = a2
            a2 = t
            
            t = ADDu(S, g)
            S = ADDu(R, F)
            R = ADDu(q, e)
            q = ADDu(p, d)
            p = ADDu(o, c)
            o = ADDu(n, b)
            n = ADDu(d, t2)
            M = ADDu(h, a)
            
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
    
          
        Hash(8) = ADDu(M, Hash(8))
        Hash(9) = ADDu(n, Hash(9))
        Hash(10) = ADDu(o, Hash(10))
        Hash(11) = ADDu(p, Hash(11))
        Hash(12) = ADDu(q, Hash(12))
        Hash(13) = ADDu(R, Hash(13))
        Hash(14) = ADDu(S, Hash(14))
        Hash(15) = ADDu(t, Hash(15))
      
        Hash(16) = ADDu(a2, Hash(16))
        Hash(17) = ADDu(b2, Hash(17))
        Hash(18) = ADDu(c2, Hash(18))
        Hash(19) = ADDu(d2, Hash(19))
        Hash(20) = ADDu(e2, Hash(20))
        Hash(21) = ADDu(F2, Hash(21))
        Hash(22) = ADDu(g2, Hash(22))
        Hash(23) = ADDu(h2, Hash(23))
        Hash(24) = ADDu(M2, Hash(24))
        Hash(25) = ADDu(n2, Hash(25))
        Hash(26) = ADDu(o2, Hash(26))
        Hash(27) = ADDu(p2, Hash(27))
        Hash(28) = ADDu(q2, Hash(28))
        Hash(29) = ADDu(R2, Hash(29))
        Hash(30) = ADDu(S2, Hash(30))
        Hash(31) = ADDu(t22, Hash(31))
      
      
      If GetInputState() <> 0 Then
            DoEvents
      End If
    
    
    Next
    
    
    
    SHA1024 = UCase(Right$("00000000" & Hex(Hash(0)), 8) & _
        Right("00000000" & Hex(Hash(1)), 8) & _
        Right("00000000" & Hex(Hash(2)), 8) & Right("00000000" & Hex(Hash(3)), 8) & _
        Right("00000000" & Hex(Hash(4)), 8) & Right("00000000" & Hex(Hash(5)), 8) & _
        Right("00000000" & Hex(Hash(6)), 8) & Right("00000000" & Hex(Hash(7)), 8) & _
        Right("00000000" & Hex(Hash(8)), 8) & Right("00000000" & Hex(Hash(9)), 8) & _
        Right("00000000" & Hex(Hash(10)), 8) & Right("00000000" & Hex(Hash(11)), 8) & _
        Right("00000000" & Hex(Hash(12)), 8) & Right("00000000" & Hex(Hash(13)), 8) & _
        Right("00000000" & Hex(Hash(14)), 8) & Right("00000000" & Hex(Hash(15)), 8))
    
    
    SHA1024 = SHA1024 & UCase(Right$("00000000" & Hex(Hash(16)), 8) & _
        Right("00000000" & Hex(Hash(17)), 8) & _
        Right("00000000" & Hex(Hash(18)), 8) & Right("00000000" & Hex(Hash(19)), 8) & _
        Right("00000000" & Hex(Hash(20)), 8) & Right("00000000" & Hex(Hash(21)), 8) & _
        Right("00000000" & Hex(Hash(22)), 8) & Right("00000000" & Hex(Hash(23)), 8) & _
        Right("00000000" & Hex(Hash(24)), 8) & Right("00000000" & Hex(Hash(25)), 8) & _
        Right("00000000" & Hex(Hash(26)), 8) & Right("00000000" & Hex(Hash(27)), 8) & _
        Right("00000000" & Hex(Hash(28)), 8) & Right("00000000" & Hex(Hash(29)), 8) & _
        Right("00000000" & Hex(Hash(30)), 8) & Right("00000000" & Hex(Hash(31)), 8))
    
    
    
    
    



End Function



