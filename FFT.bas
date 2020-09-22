Attribute VB_Name = "FFT"
Option Explicit

' Radix-2 FFT by Murphy McCauley

Private Const PI        As Double = 3.14159265358979
Private Const PI2       As Double = PI * 2

Private m_lngP2(16)     As Long
Private m_sngDLA(16)    As Single
Private m_sngDLB(16)    As Single

Public Sub InitFFT()
    Dim i As Long
    
    ' Lookup Tables for Alpha and Beta
    If m_lngP2(0) = 0 Then
        For i = 0 To 16
            m_lngP2(i) = 2 ^ i
            m_sngDLA(i) = 2 * Sin(0.5 * PI2 / (m_lngP2(i) * 2)) ^ 2
            m_sngDLB(i) = Sin(PI2 / (m_lngP2(i) * 2))
        Next
    End If
End Sub

Public Sub RealFFT( _
    ByVal NumSamples As Long, _
    RealIn() As Single, _
    RealOut() As Single, ImagOut() As Single _
)

    Static rev         As Long, NumBits    As Long

    Static i           As Long, j          As Long
    Static k           As Long, n          As Long
    Static l           As Long

    Static BlockSize   As Long, BlockEnd   As Long

    Static DeltaAr     As Single
    Static Alpha       As Single, Beta     As Single

    Static TR          As Single, TI       As Single
    Static AR          As Single, AI       As Single

    For n = 0& To 16&
        If NumSamples = m_lngP2(n) Then
            NumBits = n
            Exit For
        End If
    Next
    
    For i = 0& To NumSamples - 1&
        rev = 0&
        k = i

        For j = 0& To NumBits - 1&
            rev = (rev * 2&) Or (k And 1&)
            k = k \ 2&
        Next

        RealOut(rev) = RealIn(i)
    Next

    BlockEnd = 1
    BlockSize = 2
    l = 0

    Do While BlockSize <= NumSamples
        Alpha = m_sngDLA(l)
        Beta = m_sngDLB(l)
        l = (l + 1) Mod NumBits

        For i = 0& To NumSamples - 1 Step BlockSize
            AR = 1#
            AI = 0#
            
            j = i
            For n = 0& To BlockEnd - 1&
                k = j + BlockEnd
                TR = AR * RealOut(k) - AI * ImagOut(k)
                TI = AI * RealOut(k) + AR * ImagOut(k)
                RealOut(k) = RealOut(j) - TR
                ImagOut(k) = ImagOut(j) - TI
                RealOut(j) = RealOut(j) + TR
                ImagOut(j) = ImagOut(j) + TI
                DeltaAr = Alpha * AR + Beta * AI
                AI = AI - Alpha * AI + Beta * AR
                AR = AR - DeltaAr
                j = j + 1&
            Next
        Next

        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop
End Sub
