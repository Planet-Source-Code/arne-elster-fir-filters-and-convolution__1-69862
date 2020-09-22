Attribute VB_Name = "SignalFilter"
Option Explicit


' Windowed Sinc Filter and Convolution
' from http://www.dspguide.com/
'
' factor:   Value from 0.0 to 0.5. This is the cutoff.
'           Bsp.: Audiosignal sampled with 22050 Hz,
'                 Cutoff frequency should be 220.5 Hz.
'                 factor = 220.5/22050 = 0.01
'
' Taps:     Quality of the filter. The more taps, the steeper
'           the rolloff band gets. But the more taps, the longer
'           the delay, and the more calculations.



'#################################################################
'#################################################################
'
'
' compiled with MS VC++ 2005
'
' int __stdcall Convolve(float * samples, int nSamples,
'                        float * kernel,  int nKernel,
'                        float * output,  float * overlap)
' {
'     int i, j, k;
'
'     for (i=0; i < nSamples; i++)
'         for (j=0; j < nKernel; j++)
'             output[i+j] += samples[i] * kernel[j];
'
'     k = nSamples >= nKernel ? nKernel : nSamples;
'
'     for (i=0; i < k; i++)
'     {
'         samples[i] += overlap[i];
'         overlap[i]  = output[nSamples+i];
'     }
'
'     return nSamples;
' }

Private Const ASM_CONVOLUTION As String = _
    "8B44240853558B6C241C565785C07E3C8B5C24148BF52BDD894424" & _
    "188B7C242085FF7E1A8B54241C8BCED90433D80A83C20483C1044F" & _
    "D841FCD959FC75EC8B4C241883C60449894C241875D085C07E158B" & _
    "4C24148BD52BD18BF08B3C0A893983C1044E75F58B4C24288B5424" & _
    "1433FF8D7485002BD18B5C24203BC37D028BD83BFB7D18D9040AD8" & _
    "014783C60483C104D95C0AFC8B5EFC8959FCEBDA5F5E5D5BC21800"

Private m_blnFastConv   As Boolean
Private m_udtHook       As HookData
Private m_udtASM        As MachineCode


'#################################################################
'#################################################################


Public Type FilterKernel
    kernel()            As Single
    olap()              As Single
    taps                As Long
End Type

Public Enum FilterType
    FilterHighpass = 0
    FilterLowpass
End Enum

Private Const PI        As Single = 3.14159265358979
Private Const PI2       As Single = PI * 2


' Calculate Windowed Sinc FIR Filter kernel
' from http://www.dspguide.com/
Public Function CreateFilter(ByVal ftp As FilterType, ByVal taps As Long, ByVal factor As Single) As FilterKernel
    Dim omega       As Single
    Dim n           As Single
    Dim sum         As Single
    Dim m           As Long
    Dim i           As Long

    omega = PI2 * factor
    m = taps / 2

    With CreateFilter
        ReDim .kernel(taps - 1) As Single
        ReDim .olap(taps - 1) As Single
        .taps = taps
    
        ' Sinc function
        For i = 0 To taps - 1
            If i - m = 0 Then
                .kernel(i) = omega
            Else
                n = i - m
                .kernel(i) = Sin(omega * n) / n
            End If
            
            ' bell-shaped window to minimze ripple
            .kernel(i) = .kernel(i) * HammingWindow(i, taps)
            sum = sum + .kernel(i)
        Next
        
        If sum = 0 Then sum = 1
        
        For i = 0 To taps - 1
            ' normalize kernel
            .kernel(i) = .kernel(i) / sum
        Next
        
        If ftp = FilterHighpass Then
            ' spectral inversion to convert the lowpass to a highpass
            For i = 0 To taps - 1
                .kernel(i) = -.kernel(i)
            Next
            .kernel(m) = .kernel(m) + 1
        End If
    End With
End Function


' clears overlap of the filter
Public Sub ResetFilter(kernel As FilterKernel)
    Dim i           As Long
    
    For i = 0 To kernel.taps - 1
        kernel.olap(i) = 0
    Next
End Sub


' filter a signal with Overlap-Add method
Public Sub FilterProcess(sngValues() As Single, kernel As FilterKernel)
    Dim sngOut()    As Single
    Dim i           As Long
    Dim n           As Long
    
    n = UBound(sngValues) + 1
    
    If m_blnFastConv Then
        ReDim sngOut(n + kernel.taps - 1) As Single
        FastConvolve sngValues(0), n, _
                     kernel.kernel(0), kernel.taps, _
                     sngOut(0), kernel.olap(0)
    Else
        If n >= kernel.taps - 1 Then
            sngOut = Convolve(sngValues, kernel.kernel)
    
            For i = 0 To kernel.taps - 1
                sngOut(i) = sngOut(i) + kernel.olap(i)
                kernel.olap(i) = sngOut(n + i)
            Next
    
            For i = 0 To n - 1
                sngValues(i) = sngOut(i)
            Next
        End If
    End If
End Sub


' convolution (Input Side method)
Private Function Convolve(a() As Single, B() As Single) As Single()
    Dim c() As Single
    Dim i   As Long
    Dim j   As Long
    
    ReDim c(UBound(a) + UBound(B) + 1) As Single
    
    For i = 0 To UBound(a)
        For j = 0 To UBound(B)
            c(i + j) = c(i + j) + a(i) * B(j)
        Next
    Next
    
    Convolve = c
End Function


Private Function FastConvolve( _
    samples As Single, ByVal nSamples As Long, _
    kernel As Single, ByVal nKernel As Long, _
    output As Single, overlap As Single _
) As Long

    FastConvolve = -1
End Function


Public Sub InitFastConvolution()
    If Not m_blnFastConv Then
        m_udtASM = ASMStringToMemory(ASM_CONVOLUTION)
        m_udtHook = RedirectFunction(AddressOf FastConvolve, True, m_udtASM.pAsm)
        m_blnFastConv = True
    End If
End Sub


Public Sub TerminateFastConvolution()
    If m_blnFastConv Then
        FreeASMMemory m_udtASM
        RestoreFunction m_udtHook
        m_blnFastConv = False
    End If
End Sub


Public Function HammingWindow(ByVal i As Single, ByVal n As Single) As Single
    HammingWindow = 0.54 - 0.46 * Cos(PI2 * i / n)
End Function

