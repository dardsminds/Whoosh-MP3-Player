Attribute VB_Name = "modSpectrum"
'//////////////Spectrum module//////////
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

Public Const BI_RGB = 0&
Public Const DIB_RGB_COLORS = 0& '  color table in RGBs

Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(256) As RGBQUAD
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long

'NOTE: Using an API timer will sometimes Crash your program
'Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long
'Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
'Public Const TIME_PERIODIC = 1  ' program for continuous periodic event
'Public timing As Long           ' an API timer Handle

Public SPECWIDTH As Long, SPECHEIGHT As Long
Public specmode As Boolean, specpos As Integer  ' spectrum mode (and marker pos for 2nd mode)
Public specbuf() As Byte    'a pointer

Public chan As Long         'stream/music handle

Public bh As BITMAPINFO     'bitmap header


Public Function Sqrt(ByVal num As Double) As Double
    Sqrt = num ^ 0.5
End Function

Public Sub Spectrum1(ByVal uTimerID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
'    Dim x As Long, Y As Long, y1 As Long
'    Dim fft(300) As Single     'get the FFT data
'
'    Call BASS_ChannelGetData(Mp3(1).chan, fft(0), BASS_DATA_FFT512)
'
'
'    'specRCT.Bottom = frmSPECtest.Picture1.Height
'    'specRCT.Left = 0
'    'specRCT.Right = frmSPECtest.Picture1.Width
'    'specRCT.Top = 0
'
'        ReDim specbuf(SPECWIDTH * (SPECHEIGHT + 1)) As Byte 'clear display
'        For x = 0 To (SPECWIDTH) - 45 Step 2
'            Y = Sqrt(fft(x + 1)) * 2 * SPECHEIGHT - 4 ' scale it (sqrt to make low values more visible)
'
'            'Y = fft(X + 1) * 10 * SPECHEIGHT 'scale it (linearly)
'            If (Y > SPECHEIGHT) Then Y = SPECHEIGHT - 2 ' cap it
'
'            While (Y >= 0)
'               specbuf(Y * SPECWIDTH + x) = Y
'               Y = Y - 1
'             Wend
'        Next x
'
'    'display the update
'    Call SetDIBitsToDevice(Player1.analyzer.hDC, 0, 0, SPECWIDTH, SPECHEIGHT, 0, 0, 0, SPECHEIGHT, specbuf(0), bh, 0)
End Sub
Public Sub Spectrum2(ByVal uTimerID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
    Dim x As Long, Y As Long, y1 As Long
    Dim fft(300) As Single     'get the FFT data
    Call BASS_ChannelGetData(Mp3(2).chan, fft(0), BASS_DATA_FFT512)
   
        ReDim specbuf(SPECWIDTH * (SPECHEIGHT + 1)) As Byte 'clear display
        For x = 0 To (SPECWIDTH) - 45 Step 2
            Y = Sqrt(fft(x + 1)) * 2 * SPECHEIGHT - 4 ' scale it (sqrt to make low values more visible)
            'Y = fft(X + 1) * 10 * SPECHEIGHT 'scale it (linearly)
            If (Y > SPECHEIGHT) Then Y = SPECHEIGHT - 2 ' cap it
                
            While (Y >= 0)
               specbuf(Y * SPECWIDTH + x) = Y
               Y = Y - 1
             Wend
        Next x
   'display the update
    Call SetDIBitsToDevice(Player2.analyzer.hDC, 0, 0, SPECWIDTH, SPECHEIGHT, 0, 0, 0, SPECHEIGHT, specbuf(0), bh, 0)
End Sub

Public Sub SetupBitmap1()
    'create bitmap to draw spectrum in - 8 bit for easy updating :)
    With bh.bmiHeader
        .biBitCount = 8
        .biPlanes = 1
        .biSize = Len(bh.bmiHeader)
        .biWidth = SPECWIDTH
        .biHeight = SPECHEIGHT  'upside down (line 0=bottom)
        .biClrUsed = 256
        .biClrImportant = 256
    End With
    
    Dim a As Byte
    Dim i As Byte
                 
    'setup palette
    For a = 1 To 80 Step 2
        bh.bmiColors(a).rgbGreen = 255
    Next a
    For a = 80 To 160
        bh.bmiColors(a).rgbRed = a * 1.5
    Next a
    For a = 160 To 254
        bh.bmiColors(a).rgbBlue = a
    Next a
End Sub

'Public Function StartTimer()
'    TimerA = KillTimer(0, TimerA)
'    TimerA = SetTimer(0, TimerA, 20, AddressOf ADeckTimer)
'End Function'
'------------------------------
'Public Function StopTimer()
'    TimerA = KillTimer(0, TimerA)
'End Function

