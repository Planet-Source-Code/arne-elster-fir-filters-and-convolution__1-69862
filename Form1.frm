VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Windowed Sinc FIR Filter mit Convolution"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame4 
      Caption         =   "Times:"
      Height          =   1275
      Left            =   7860
      TabIndex        =   26
      Top             =   5070
      Width           =   1935
      Begin VB.Label lblTimeTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   900
         Width           =   405
      End
      Begin VB.Label lblTimeFFT 
         AutoSize        =   -1  'True
         Caption         =   "FFT Avg.:"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   600
         Width           =   705
      End
      Begin VB.Label lblTimeFilter 
         AutoSize        =   -1  'True
         Caption         =   "Filter Avg.:"
         Height          =   195
         Left            =   135
         TabIndex        =   27
         Top             =   300
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Highpass"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   60
      TabIndex        =   16
      Top             =   5040
      Width           =   7695
      Begin VB.PictureBox picHighPass 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   79
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   5415
      End
      Begin VB.PictureBox picHPKernel 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   5760
         ScaleHeight     =   79
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
      End
      Begin VB.PictureBox picSpecHP 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   5385
         TabIndex        =   18
         Top             =   1800
         Width           =   5415
      End
      Begin VB.PictureBox picSpecHPK 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   5760
         ScaleHeight     =   405
         ScaleWidth      =   1785
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblKernel2 
         AutoSize        =   -1  'True
         Caption         =   " Filter kernel"
         Height          =   195
         Left            =   5760
         TabIndex        =   22
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblFreqSpec 
         AutoSize        =   -1  'True
         Caption         =   "Frequency spectrum:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1485
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Lowpass"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   60
      TabIndex        =   9
      Top             =   2700
      Width           =   7695
      Begin VB.PictureBox picLowPass 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   120
         ScaleHeight     =   67
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   300
         Width           =   5415
      End
      Begin VB.PictureBox picLPKernel 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   5760
         ScaleHeight     =   67
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   119
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   300
         Width           =   1815
      End
      Begin VB.PictureBox picSpecLP 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   120
         ScaleHeight     =   405
         ScaleWidth      =   5385
         TabIndex        =   11
         Top             =   1680
         Width           =   5415
      End
      Begin VB.PictureBox picSpecLPK 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   5760
         ScaleHeight     =   405
         ScaleWidth      =   1785
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblKernel 
         AutoSize        =   -1  'True
         Caption         =   " Filter kernel"
         Height          =   195
         Left            =   5760
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblFreqSpec 
         AutoSize        =   -1  'True
         Caption         =   "Frequency spectrum:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdRectangular 
         Caption         =   "Rects"
         Height          =   375
         Left            =   5940
         TabIndex        =   25
         Top             =   1500
         Width           =   1275
      End
      Begin VB.CommandButton cmdSinusoids 
         Caption         =   "Sinusoids"
         Height          =   375
         Left            =   5940
         TabIndex        =   24
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CommandButton cmdSawtooth 
         Caption         =   "Sawtooth"
         Height          =   375
         Left            =   5940
         TabIndex        =   23
         Top             =   660
         Width           =   1275
      End
      Begin VB.PictureBox picStd 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   180
         ScaleHeight     =   67
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   359
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
         Width           =   5415
      End
      Begin VB.CheckBox chkNoise 
         Caption         =   "Noise"
         Height          =   195
         Left            =   5940
         TabIndex        =   6
         Top             =   300
         Width           =   1155
      End
      Begin VB.PictureBox picSpecInp 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   180
         ScaleHeight     =   405
         ScaleWidth      =   5385
         TabIndex        =   5
         Top             =   1680
         Width           =   5415
      End
      Begin VB.Label lblFreqSpec 
         AutoSize        =   -1  'True
         Caption         =   "Frequency spectrum:"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   1440
         Width           =   1485
      End
   End
   Begin VB.VScrollBar scrlTaps 
      Height          =   4095
      LargeChange     =   10
      Left            =   9420
      Max             =   512
      Min             =   2
      TabIndex        =   3
      Top             =   480
      Value           =   90
      Width           =   315
   End
   Begin VB.VScrollBar scrlFactor 
      Height          =   4095
      LargeChange     =   20
      Left            =   8340
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   1
      Top             =   480
      Value           =   80
      Width           =   315
   End
   Begin VB.Label lblTaps 
      AutoSize        =   -1  'True
      Caption         =   "Taps (100):"
      Height          =   195
      Left            =   8940
      TabIndex        =   2
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblFactor 
      AutoSize        =   -1  'True
      Caption         =   "Factor (0.2):"
      Height          =   195
      Left            =   7860
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const samples           As Long = 512   ' power of 2 because of FFT!

Private m_sngInp(samples - 1)   As Single


' Sawtooth function. With a lowpass you can watch the
' Fourier series of this function
Private Function InputSawtooth(ByVal t As Single) As Single
    Const a As Single = 150
    
    InputSawtooth = 1.9 * (t / a - Fix(t / a + 0.5))
End Function


' simple interference of waves of different frequency
Private Function InputSinusoids(ByVal t As Single) As Single
    InputSinusoids = (Sin(2 * t) + Sin(0.3 * t) + Sin(0.1 * t)) * 0.5
End Function


' rectangular input function
Private Function InputRectangular(ByVal t As Single) As Single
    Const a As Single = 50
    Static x As Single
    
    If x = 0 Then x = 0.8
    
    If t Mod a = 0 Then
        x = -x
    End If
    
    InputRectangular = x
End Function


' plot a function/signal scaled in a PictureBox
Private Sub Plot(pb As PictureBox, sngSamples() As Single, ByVal center As Boolean, ByVal normalize As Boolean)
    Dim dy      As Long, dy2        As Long
    Dim n       As Long, i          As Long
    Dim sngMax  As Single, sngVal   As Single
    Dim yL      As Single, yN       As Single
    Dim x       As Single, k        As Single
    Dim st      As Single
    
    dy = pb.ScaleHeight - 1
    dy2 = dy \ 2
    
    n = UBound(sngSamples) + 1
    st = n / pb.ScaleWidth
    
    If normalize Then
        For i = 0 To n - 1
            If Abs(sngSamples(i)) > sngMax Then sngMax = Abs(sngSamples(i))
        Next
    End If
    
    If sngMax = 0 Then sngMax = 1
    sngVal = sngSamples(0) / sngMax
    
    If center Then
        yL = -sngVal * dy2 + dy2
        
        pb.ForeColor = RGB(160, 160, 180)
        pb.Line (0, dy2)-(pb.ScaleWidth, dy2)
        pb.ForeColor = vbBlack
    Else
        yL = dy - sngVal * dy
    End If
    
    k = k + st
    
    Do
        sngVal = sngSamples(Fix(k)) / sngMax
        
        If center Then
            yN = -sngVal * dy2 + dy2
        Else
            yN = dy - sngVal * dy
        End If
        
        pb.Line (x, yL)-(x + 1, yN)
        yL = yN
        x = x + 1
        k = k + st
    Loop While k < n
End Sub


' filter a signal and plot the result
Private Sub Filter(ByVal lngTaps As Long, ByVal sngFactor As Single)
    Dim sngLP()         As Single
    Dim sngHP()         As Single
    Dim udtLP           As FilterKernel
    Dim udtHP           As FilterKernel
    Dim d               As Double
    Dim tmrTotal        As Double
    Dim tmrFilter       As Double
    Dim tmrFFT          As Double
    
    tmrTotal = Timer
    
    udtLP = CreateFilter(FilterLowpass, lngTaps, sngFactor)
    udtHP = CreateFilter(FilterHighpass, lngTaps, sngFactor)
    
    d = Timer
    sngLP = m_sngInp
    FilterProcess sngLP, udtLP
    picLowPass.Cls:     Plot picLowPass, sngLP, True, False
    picLPKernel.Cls:    Plot picLPKernel, udtLP.kernel, True, False
    tmrFilter = Timer - d

    d = Timer
    picSpecLP.Cls:      DisplayFT picSpecLP, sngLP, True, False
    picSpecLPK.Cls:     DisplayFT picSpecLPK, udtLP.kernel, False, True
    tmrFFT = Timer - d
    
    d = Timer
    sngHP = m_sngInp
    FilterProcess sngHP, udtHP
    picHighPass.Cls:    Plot picHighPass, sngHP, True, False
    picHPKernel.Cls:    Plot picHPKernel, udtHP.kernel, True, False
    tmrFilter = (tmrFilter + (Timer - d)) / 2

    d = Timer
    picSpecHP.Cls:      DisplayFT picSpecHP, sngHP, True, False
    picSpecHPK.Cls:     DisplayFT picSpecHPK, udtHP.kernel, False, True
    tmrFFT = (tmrFFT + (Timer - d)) / 4
    
    lblTimeFilter.Caption = "Filter Avg.: " & Round(tmrFilter * 1000) & " ms"
    lblTimeFFT.Caption = "FFT Avg.: " & Round(tmrFFT * 1000) & " ms"
    lblTimeTotal.Caption = "Total: " & Round((Timer - tmrTotal) * 1000) & " ms"
End Sub


' add white noise to the signal
Private Sub chkNoise_Click()
    Dim i As Long
    
    If chkNoise.Value = 1 Then
        For i = 0 To samples - 1
            m_sngInp(i) = (m_sngInp(i) + Rnd() * 0.5) / 1.5
        Next
    Else
        For i = 0 To samples - 1
            m_sngInp(i) = InputSawtooth(i)
        Next
    End If

    UpdateInpDisplay
    UpdateDisplay
End Sub


Private Sub cmdRectangular_Click()
    Dim i As Long

    For i = 1 To samples - 1
        m_sngInp(i) = InputRectangular(i)
    Next
    
    UpdateInpDisplay
    UpdateDisplay
End Sub


Private Sub cmdSawtooth_Click()
    Dim i As Long
    
    For i = 0 To samples - 1
        m_sngInp(i) = InputSawtooth(i)
    Next
    
    UpdateInpDisplay
    UpdateDisplay
End Sub


Private Sub cmdSinusoids_Click()
    Dim i As Long
    
    For i = 0 To samples - 1
        m_sngInp(i) = InputSinusoids(i)
    Next
    
    UpdateInpDisplay
    UpdateDisplay
End Sub


Private Sub Form_Load()
    Dim i   As Long
    
    InitFFT                     ' prepare lookup tables for FFT
    InitFastConvolution         ' use convolution algorithm written in C
    
    scrlTaps.Max = samples
    
    For i = 0 To samples - 1
        m_sngInp(i) = InputSawtooth(i)
    Next
    
    UpdateInpDisplay
    UpdateDisplay
End Sub


' Fourier Transformation of data and plot of result
Private Sub DisplayFT(pb As PictureBox, data() As Single, ByVal window As Boolean, ByVal normalize As Boolean)
    Dim sngRealInp(samples - 1)     As Single
    Dim sngRealOut(samples - 1)     As Single
    Dim sngImagOut(samples - 1)     As Single
    Dim sngCompOut(samples / 2 - 1) As Single
    Dim i                           As Long
    Dim sngDivisor                  As Single
    
    ' Hamming window for less leak
    For i = 0 To UBound(data)
        If window Then
            sngRealInp(i) = data(i) * HammingWindow(i, UBound(data) + 1)
        Else
            sngRealInp(i) = data(i)
        End If
    Next
    
    ' transform data from time- to frequencydomain
    RealFFT samples, sngRealInp, sngRealOut, sngImagOut

    ' scaled magnitude spectrum from complex transformed
    sngDivisor = samples / 8
    For i = 0 To samples / 2 - 1
        sngCompOut(i) = Sqr(sngRealOut(i) * sngRealOut(i) + sngImagOut(i) * sngImagOut(i)) / sngDivisor
    Next

    Plot pb, sngCompOut, False, normalize
End Sub


Private Sub UpdateInpDisplay()
    picStd.Cls
    picSpecInp.Cls
    Plot picStd, m_sngInp, True, False
    DisplayFT picSpecInp, m_sngInp, True, False
End Sub


Private Sub UpdateDisplay()
    Dim sngFactor   As Single
    Dim lngTaps     As Long
    
    sngFactor = scrlFactor.Value / scrlFactor.Max * 0.5
    lngTaps = scrlTaps.Value
    
    lblFactor.Caption = "Faktor (" & Format(sngFactor, "0.00") & "):"
    lblTaps.Caption = "Taps (" & lngTaps & "):"
    Filter lngTaps, sngFactor
End Sub


Private Sub Form_Unload(Cancel As Integer)
    TerminateFastConvolution
End Sub


Private Sub scrlFactor_Change()
    UpdateDisplay
End Sub


Private Sub scrlFactor_Scroll()
    UpdateDisplay
End Sub


Private Sub scrlTaps_Change()
    UpdateDisplay
End Sub


Private Sub scrlTaps_Scroll()
    UpdateDisplay
End Sub

