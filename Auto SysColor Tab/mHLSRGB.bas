Attribute VB_Name = "mHLSRGB"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_MENU = 4
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_WINDOWTEXT = 8
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_INFOTEXT = 23
Public Const COLOR_INFOBK = 24

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type HSLColor
    Hue As Long
    Sat As Long
    Lum As Long
End Type

Const HSLMAX As Long = 255
'H, S and L values can be 0 - HSLMAX. 240 matches what is used by MS Win;
'any number less than 1 byte is OK; works best if it is evenly divisible by 6
Const RGBMAX As Long = 255
'R, G, and B value can be 0 - RGBMAX
Const UNDEFINED As Long = 0
'Hue is undefined if Saturation = 0 (greyscale)

Public Sub DrawHLSBox( _
      ByRef picThis As PictureBox _
   )
Dim H As Single
Dim S As Single
Dim R As Long, G As Long, B As Long
Dim lHDC As Long
Dim tR As RECT
Dim lColor As Long
Dim hBr As Long

   picThis.Cls
   lHDC = picThis.hdc
   picThis.Width = 240 * Screen.TwipsPerPixelX
   picThis.Height = 256 * Screen.TwipsPerPixelY
   tR.Right = 1
   tR.Bottom = 2
   For H = -40 To 200
      For S = 128 To 0 Step -1
         HLSToRGB H / 40, S / 128, 0.5, R, G, B
         lColor = RGB(R, G, B)
         hBr = CreateSolidBrush(lColor)
         FillRect lHDC, tR, hBr
         DeleteObject hBr
         tR.Top = tR.Top + 2
         tR.Bottom = tR.Top + 2
      Next S
      tR.Left = tR.Left + 1
      tR.Right = tR.Left + 1
      tR.Top = 0
      tR.Bottom = 2
   Next H
End Sub

Public Sub DrawLuminanceBox( _
      ByRef picThis As PictureBox, _
      ByVal H As Single, _
      ByVal S As Single _
   )
Dim R As Long, G As Long, B As Long
Dim lHDC As Long
Dim tR As RECT
Dim lColor As Long
Dim hBr As Long
Dim L As Long
   
   lHDC = picThis.hdc
   tR.Right = picThis.ScaleWidth \ Screen.TwipsPerPixelX
   tR.Bottom = 1
   For L = 255 To 0 Step -1
      HLSToRGB H, S, L / 255, R, G, B
      hBr = CreateSolidBrush(RGB(R, G, B))
      FillRect lHDC, tR, hBr
      DeleteObject hBr
      tR.Top = tR.Top + 1
      tR.Bottom = tR.Bottom + 1
   Next L
   picThis.Refresh
End Sub

Public Sub RGBToHLS( _
      ByVal R As Long, ByVal G As Long, ByVal B As Long, _
      H As Single, S As Single, L As Single _
   )
Dim Max As Single
Dim Min As Single
Dim delta As Single
Dim rR As Single, rG As Single, rB As Single

   rR = R / 255: rG = G / 255: rB = B / 255

'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
        Max = Maximum(rR, rG, rB)
        Min = Minimum(rR, rG, rB)
        L = (Max + Min) / 2    '{This is the lightness}
        '{Next calculate saturation}
        If Max = Min Then
            'begin {Acrhomatic case}
            S = 0
            H = 0
           'end {Acrhomatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If L <= 0.5 Then
               S = (Max - Min) / (Max + Min)
           Else
               S = (Max - Min) / (2 - Max - Min)
            End If
            '{Next calculate the hue.}
            delta = Max - Min
           If rR = Max Then
                H = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = Max Then
                H = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
           ElseIf rB = Max Then
                H = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
            End If
            'Debug.Print h
            'h = h * 60
           'If h < 0# Then
           '     h = h + 360            '{Make degrees be nonnegative}
           'End If
        'end {Chromatic Case}
      End If
'end {RGB_to_HLS}
End Sub

Private Function iMax(a As Long, B As Long) As Long
    'Return the Larger of two values
    iMax = IIf(a > B, a, B)
End Function


Private Function iMin(a As Long, B As Long) As Long
    'Return the smaller of two values
    iMin = IIf(a < B, a, B)
End Function

Public Function GetHSL(iColor As Long) As HSLColor
    'Returns an HSLCol datatype containing H ue, Luminescence
    'and Saturation; given an RGB Color value
    Dim R As Long, G As Long, B As Long
    Dim cMax As Long, cMin As Long
    Dim RDelta As Double, GDelta As Double, BDelta As Double
    Dim H As Double, S As Double, L As Double
    Dim cMinus As Long, cPlus As Long
    
    R = GetRed(iColor)
    G = GetGreen(iColor)
    B = GetBlue(iColor)
    
    cMax = iMax(iMax(R, G), B) 'Highest and lowest
    cMin = iMin(iMin(R, G), B) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin 'calculations somewhat.
    
    'Calculate luminescence (lightness)
    L = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    


    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        S = 0 'Saturation 0 for greyscale
        H = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation


        If L <= (HSLMAX / 2) Then
            S = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            S = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
        
        'Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus
        


        Select Case cMax
            Case CLng(R)
            H = BDelta - GDelta
            Case CLng(G)
            H = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(B)
            H = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
    
    If H < 0 Then H = H + HSLMAX
End If

GetHSL.Hue = CLng(H)
GetHSL.Lum = CLng(L)
GetHSL.Sat = CLng(S)
End Function

Public Sub HLSToRGB( _
      ByVal H As Single, ByVal S As Single, ByVal L As Single, _
      R As Long, G As Long, B As Long _
   )
Dim rR As Single, rG As Single, rB As Single
Dim Min As Single, Max As Single

   If S = 0 Then
      ' Achromatic case:
      rR = L: rG = L: rB = L
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If L <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         Min = L * (1 - S)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         Min = L - S * (1 - L)
      End If
      ' Get the Max value:
      Max = 2 * L - Min
      
      ' Now depending on sector we can evaluate the h,l,s:
      If (H < 1) Then
         rR = Max
         If (H < 0) Then
            rG = Min
            rB = rG - H * (Max - Min)
         Else
            rB = Min
            rG = H * (Max - Min) + rB
         End If
      ElseIf (H < 3) Then
         rG = Max
         If (H < 2) Then
            rB = Min
            rR = rB - (H - 2) * (Max - Min)
         Else
            rR = Min
            rB = (H - 2) * (Max - Min) + rR
         End If
      Else
         rB = Max
         If (H < 4) Then
            rR = Min
            rG = rR - (H - 4) * (Max - Min)
         Else
            rG = Min
            rR = (H - 4) * (Max - Min) + rG
         End If
         
      End If
            
   End If
   R = rR * 255: G = rG * 255: B = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then
         Maximum = rR
      Else
         Maximum = rB
      End If
   Else
      If (rB > rG) Then
         Maximum = rB
      Else
         Maximum = rG
      End If
   End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then
         Minimum = rR
      Else
         Minimum = rB
      End If
   Else
      If (rB < rG) Then
         Minimum = rB
      Else
         Minimum = rG
      End If
   End If
End Function

Public Function GetRed(iColor As Long) As Integer
    GetRed = iColor Mod 256
End Function

Public Function GetGreen(iColor As Long) As Integer
    GetGreen = ((iColor And &HFF00FF00) / 256&)
End Function

Public Function GetBlue(iColor As Long) As Integer
    GetBlue = ((iColor And &HFF0000) / 65536) '/ (256& * 256&)
End Function

