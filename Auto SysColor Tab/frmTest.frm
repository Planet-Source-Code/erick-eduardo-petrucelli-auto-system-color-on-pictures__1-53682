VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AutoColor Tab Media Player 9 Setup"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmTest.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picItem 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2040
      Index           =   1
      Left            =   225
      ScaleHeight     =   2040
      ScaleWidth      =   6540
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1725
      Visible         =   0   'False
      Width           =   6540
      Begin VB.PictureBox picBrasil 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   75
         ScaleHeight     =   765
         ScaleWidth      =   840
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Programadores Brasileiros na veia!"
         Top             =   75
         Width           =   840
      End
      Begin VB.Label lblCredit 
         BackStyle       =   0  'Transparent
         Caption         =   "Credits"
         Height          =   1395
         Left            =   975
         TabIndex        =   21
         Top             =   150
         Width           =   5190
      End
      Begin VB.Label lblMail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You can give more info contacting me: erickpetru@zipmail.com.br"
         Height          =   240
         Left            =   225
         TabIndex        =   20
         Top             =   1650
         Width           =   6090
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Concluir"
      Height          =   465
      Left            =   3450
      TabIndex        =   9
      Top             =   4275
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   465
      Left            =   5100
      TabIndex        =   8
      Top             =   4275
      Width           =   1515
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Avançar >>"
      Default         =   -1  'True
      Height          =   465
      Left            =   3450
      TabIndex        =   7
      Top             =   4275
      Width           =   1515
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< &Voltar"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1875
      TabIndex        =   6
      Top             =   4275
      Width           =   1515
   End
   Begin VB.Timer tmrTheme 
      Interval        =   100
      Left            =   6300
      Top             =   75
   End
   Begin VB.PictureBox picTabC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   1650
      Picture         =   "frmTest.frx":000C
      ScaleHeight     =   450
      ScaleWidth      =   2565
      TabIndex        =   4
      Top             =   990
      Width           =   2565
      Begin VB.Label lblTabC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         Height          =   195
         Index           =   0
         Left            =   1095
         TabIndex        =   10
         Top             =   150
         Width           =   795
      End
   End
   Begin VB.PictureBox picTabC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   4065
      Picture         =   "frmTest.frx":3CC6
      ScaleHeight     =   450
      ScaleWidth      =   2565
      TabIndex        =   5
      Top             =   990
      Visible         =   0   'False
      Width           =   2565
      Begin VB.Label lblTabC 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credits"
         Height          =   195
         Index           =   1
         Left            =   1245
         TabIndex        =   11
         Top             =   150
         Width           =   495
      End
   End
   Begin VB.PictureBox picTabN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   0
      Left            =   1650
      Picture         =   "frmTest.frx":7980
      ScaleHeight     =   450
      ScaleWidth      =   2565
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   2565
      Begin VB.Label lblTabN 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         Height          =   195
         Index           =   0
         Left            =   1050
         TabIndex        =   12
         Top             =   150
         Width           =   885
      End
   End
   Begin VB.PictureBox picTabN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Index           =   1
      Left            =   4065
      Picture         =   "frmTest.frx":B63A
      ScaleHeight     =   450
      ScaleWidth      =   2565
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   990
      Width           =   2565
      Begin VB.Label lblTabN 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credits"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   13
         Top             =   150
         Width           =   585
      End
   End
   Begin VB.PictureBox picTop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   0
      Picture         =   "frmTest.frx":F2F4
      ScaleHeight     =   1440
      ScaleWidth      =   6840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   6840
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTest.frx":3E546
         ForeColor       =   &H8000000E&
         Height          =   420
         Left            =   1575
         TabIndex        =   15
         Top             =   450
         Width           =   5025
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Colorization Media Player 9 Tabs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   1575
         TabIndex        =   14
         Top             =   120
         Width           =   4575
      End
      Begin VB.Shape shpBorderTop 
         BorderColor     =   &H80000001&
         Height          =   1515
         Left            =   -75
         Top             =   -75
         Width           =   7140
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   -60
      ScaleHeight     =   2655
      ScaleWidth      =   6960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1430
      Width           =   6960
      Begin VB.PictureBox picItem 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1965
         Index           =   0
         Left            =   225
         ScaleHeight     =   1965
         ScaleWidth      =   6540
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   225
         Width           =   6540
         Begin VB.Label lblStep2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "I make this code in Windows XP. If you find any bug, report it to me. Thank you..."
            Height          =   240
            Left            =   225
            TabIndex        =   22
            Top             =   1575
            Width           =   6090
         End
         Begin VB.Label lblStep1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmTest.frx":3E5CF
            Height          =   465
            Left            =   225
            TabIndex        =   18
            Top             =   975
            Width           =   6090
         End
         Begin VB.Label lblInfo 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frmTest.frx":3E671
            Height          =   645
            Left            =   225
            TabIndex        =   17
            Top             =   225
            Width           =   6090
         End
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H80000001&
         Height          =   2640
         Left            =   0
         Top             =   0
         Width           =   6945
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "COMCTL32" ()

Private Type IMAGELISTDRAWPARAMS
    cbSize As Long
    hIml As Long
    i As Long
    hdcDst As Long
    x As Long
    y As Long
    cX As Long
    cY As Long
    xBitmap As Long
    yBitmap As Long
    rgbBk As Long
    rgbFg As Long
    fStyle As Long
    dwRop As Long
    fState As Long
    Frame As Long
    crEffect As Long
End Type

Public Enum ImageListStateConstants
   ILS_NORMAL = &H0& 'The image state is not modified.
   ILS_GLOW = &H1& ' Adds a glow effect to the icon, which causes the icon to appear to glow with a given color around the edges. The color for the glow effect is passed to the IImageList::Draw method in the crEffect member of IMAGELISTDRAWPARAMS.
   ILS_SHADOW = &H2& 'Adds a drop shadow effect to the icon. The color for the drop shadow effect is passed to the IImageList::Draw method in the crEffect member of IMAGELISTDRAWPARAMS.
   ILS_SATURATE = &H4& ' Saturates the icon by increasing each color component of the RGB triplet for each pixel in the icon. The amount to increase is indicated by the frame member in the IMAGELISTDRAWPARAMS method.
   ILS_ALPHA = &H8& ' Alpha blends the icon. Alpha blending controls the transparency level of an icon, according to the value of its alpha channel. The value of the alpha channel is indicated by the frame member in the IMAGELISTDRAWP
End Enum

Private Declare Function ImageList_DrawIndirect Lib "COMCTL32.DLL" ( _
    pimldp As IMAGELISTDRAWPARAMS) As Long
    
Private Const ILD_IMAGE = &H20&
Private Const ILD_PRESERVEALPHA = &H1000&
Private Const CLR_NONE = -1

Private m_cIml As cVBALImageList

Dim Caminho As String, i As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private m_cMask As New cAlphaDibSection
Private m_cImage As New cAlphaDibSection
Private m_cAlphaImage As New cAlphaDibSection

Private m_cSourceImage As New cDIBSection
Private m_cColouriseImage As New cDIBSection
Private m_cColourise As New cColourise
Dim sPic As StdPicture
Dim OldTheme As Integer, OldColor As Long

Private Sub createAlphaImage()
    ' Load picture:
    m_cImage.CreateFromPicture LoadPicture(Caminho & "PaintColor.bmp")
    
    ' Load alpha channel:
    m_cMask.CreateFromPicture LoadPicture(Caminho & "PaintMask.bmp")
    
    ' Create a new image, which is just a copy of the picture
    ' in m_cImage to build the alpha version in.  Note if
    ' we didn't want to display the image without alpha later,
    ' we could just work on m_cImage directly instead.
    m_cAlphaImage.Create m_cImage.Width, m_cImage.Height
    m_cImage.PaintPicture m_cAlphaImage.hdc
    
    ' Point byte arrays at the image bits for ease of
    ' manipulation of the data:
    Dim tMask As SAFEARRAY2D
    Dim bMask() As Byte
    Dim tImage As SAFEARRAY2D
    Dim bImage() As Byte
     
     With tMask
         .cbElements = 1
         .cDims = 2
         .Bounds(0).lLbound = 0
         .Bounds(0).cElements = m_cMask.Height
         .Bounds(1).lLbound = 0
         .Bounds(1).cElements = m_cMask.BytesPerScanLine()
         .pvData = m_cMask.DIBSectionBitsPtr
     End With
     CopyMemory ByVal VarPtrArray(bMask()), VarPtr(tMask), 4
     
     With tImage
         .cbElements = 1
         .cDims = 2
         .Bounds(0).lLbound = 0
         .Bounds(0).cElements = m_cAlphaImage.Height
         .Bounds(1).lLbound = 0
         .Bounds(1).cElements = m_cAlphaImage.BytesPerScanLine()
         .pvData = m_cAlphaImage.DIBSectionBitsPtr
     End With
     CopyMemory ByVal VarPtrArray(bImage()), VarPtr(tImage), 4
    
    Dim x As Long, y As Long
    Dim bAlpha As Long
    For y = 0 To m_cAlphaImage.Height - 1
       For x = 0 To m_cAlphaImage.BytesPerScanLine - 4 Step 4 ' each item has 4 bytes: R,G,B,A
          ' Get the red value from the mask to use as the alpha
          ' value:
          bAlpha = bMask(x, y)
          ' Set the alpha in the alpha image..
          bImage(x + 3, y) = bAlpha
          ' Now premultiply the r/g/b values by the alpha divided
          ' by 255.  This is required for the AlphaBlend GDI function,
          ' see MSDN/Platform SDK/GDI/BLENDFUNCTION for more
          ' details:
          bImage(x, y) = bImage(x, y) * bAlpha \ 255
          bImage(x + 1, y) = bImage(x + 1, y) * bAlpha \ 255
          bImage(x + 2, y) = bImage(x + 2, y) * bAlpha \ 255
       Next x
    Next y
    
    ' Clear up the temporary array descriptors.  You
    ' only need to do this on NT but best to be safe.
    CopyMemory ByVal VarPtrArray(bMask), 0&, 4
    CopyMemory ByVal VarPtrArray(bImage), 0&, 4
    
    DrawImage
End Sub

Private Sub DrawImage()
    Dim x As Long
    Dim y As Long
      
    Dim i As Long, lCol As Long
    
    lCol = vbButtonFace
         
    ' Draw normally, so only the drop shadow uses the alpha effect:
    m_cAlphaImage.AlphaPaintPicture picTop.hdc, x, y
          
    x = x + m_cAlphaImage.Width + 2
       
    picTop.Refresh
End Sub

Private Sub Process()
    Dim Color As HSLColor
    
    If m_iTheme = 1 Then
        picTabN(0).Cls
        picTabN(0).Picture = LoadPicture(Caminho & "Tab1.bmp")
        picTabN(1).Cls
        picTabN(1).Picture = LoadPicture(Caminho & "Tab1.bmp")
        picTabC(0).Cls
        picTabC(0).Picture = LoadPicture(Caminho & "Tab2-1.bmp")
        picTabC(1).Cls
        picTabC(1).Picture = LoadPicture(Caminho & "Tab2-2.bmp")
        picTop.Cls
        picTop.Picture = LoadPicture(Caminho & "BackTop.bmp")
        
        lBackColor = GetPixel(picTabN(0).hdc, 50, 0)
        shpBorderTop.BorderColor = lBackColor
        shpBorder.BorderColor = lBackColor
        
        OldTheme = m_iTheme
        OldColor = GetSysColor(COLOR_BTNFACE)
        Exit Sub
    ElseIf m_iTheme = 2 Then
        Color = GetHSL(RGB(124, 124, 148))
        Color.Sat = Color.Sat + 25
    ElseIf m_iTheme = 3 Then
        Color = GetHSL(RGB(165, 188, 132))
    Else
        Color = GetHSL(BlendColor(vbHighlight, vb3DHighlight, 77))
    End If
        
    Set sPic = LoadPicture(Caminho & "Tab1.bmp")
    m_cSourceImage.CreateFromPicture sPic
    m_cColouriseImage.Create m_cSourceImage.Width, m_cSourceImage.Height

    m_cColourise.Hue = Color.Hue + 40
    m_cColourise.Saturation = Color.Sat - 5
    m_cColourise.Process m_cSourceImage, m_cColouriseImage
    
    m_cColouriseImage.PaintPicture picTabN(0).hdc
    m_cColouriseImage.PaintPicture picTabN(1).hdc
    
    
    Set sPic = LoadPicture(Caminho & "Tab2-1.bmp")
    m_cSourceImage.CreateFromPicture sPic
    m_cColouriseImage.Create m_cSourceImage.Width, m_cSourceImage.Height
    m_cColourise.Process m_cSourceImage, m_cColouriseImage
    
    m_cColouriseImage.PaintPicture picTabC(0).hdc
    
    
    Set sPic = LoadPicture(Caminho & "Tab2-2.bmp")
    m_cSourceImage.CreateFromPicture sPic
    m_cColouriseImage.Create m_cSourceImage.Width, m_cSourceImage.Height
    m_cColourise.Process m_cSourceImage, m_cColouriseImage
    
    m_cColouriseImage.PaintPicture picTabC(1).hdc
    
    Set sPic = LoadPicture(Caminho & "BackTop.bmp")
    m_cSourceImage.CreateFromPicture sPic
    m_cColouriseImage.Create m_cSourceImage.Width, m_cSourceImage.Height
    m_cColourise.Process m_cSourceImage, m_cColouriseImage
    
    m_cColouriseImage.PaintPicture picTop.hdc
    
    lBackColor = GetPixel(picTabN(0).hdc, 50, 0)
    shpBorderTop.BorderColor = lBackColor
    shpBorder.BorderColor = lBackColor
    
    OldTheme = m_iTheme
    OldColor = GetSysColor(COLOR_BTNFACE)
End Sub

Private Sub cmdBack_Click()
    Call picTabN_Click(0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    End
End Sub

Private Sub cmdDone_Click()
    Unload Me
    End
End Sub

Private Sub cmdNext_Click()
    Call picTabN_Click(1)
End Sub

Private Sub Form_Load()
    SetIcon Me.hwnd, "LOGO"

    Caminho = App.Path
    If (Right$(Caminho, 1) <> "\") Then Caminho = Caminho & "\"
    
    Set m_cIml = New cVBALImageList
    m_cIml.IconSizeX = 48
    m_cIml.IconSizeY = 48
    m_cIml.ColourDepth = ILC_COLOR32
    m_cIml.Create
    
    m_cIml.AddFromFile Caminho & "Brasil.ico", IMAGE_ICON
    
    DrawIcon
    
    picTabN(0).Refresh
    picTabN(1).Refresh
    picTabC(0).Refresh
    picTabC(1).Refresh
    picTop.Refresh
    
    InitTheme GetDesktopWindow()
    
    Process
    
    createAlphaImage
    
    lblCredit.Caption = "Developed by: Erick Eduardo Petrucelli" + vbCrLf + "e-Rex Interactive 2004  - Taquaritinga - SP / Brasil" + vbCrLf + vbCrLf + "Thanks for: http://www.vbaccelerator.com/" + vbCrLf + "                    http://www.planetsourcecode.com/" + vbCrLf + "                    http://www.vbmania.com.br/" + vbCrLf + "                    http://www.vbweb.com.br/"
End Sub

Private Sub DrawIcon()
Dim i As Long
Dim idp As IMAGELISTDRAWPARAMS

    picBrasil.Cls
    
    idp.cbSize = Len(idp)
    idp.hIml = m_cIml.hIml
    idp.hdcDst = Me.hdc
    idp.rgbBk = CLR_NONE
    idp.fState = ILD_PRESERVEALPHA Or ILD_IMAGE

    ' draw standard:
    m_cIml.DrawImage 1, picBrasil.hdc, 2, 1
    
    idp.x = (1 - 1) * (m_cIml.IconSizeX + 4) + 4
    idp.y = 4
    idp.i = 1 - 1
End Sub

Private Sub lblTabN_Click(Index As Integer)
    Call picTabN_Click(Index)
End Sub

Private Sub picTabC_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        If Index < 1 Then
            Call picTabN_Click(Index + 1)
        End If
    ElseIf KeyCode = vbKeyLeft Then
        If Index > 0 Then
            Call picTabN_Click(Index - 1)
        End If
    End If
End Sub

Private Sub picTabN_Click(Index As Integer)
    If Index = 0 Then
        picTabN(0).Visible = False
        picTabC(0).Visible = True
        picTabN(1).Visible = True
        picTabC(1).Visible = False
        Call picTabC(0).ZOrder(0)
        
        cmdCancel.Enabled = True
        cmdNext.Visible = True
        cmdNext.Enabled = True
        cmdBack.Enabled = False
        cmdDone.Visible = False
        cmdNext.Default = True
    Else
        picTabN(0).Visible = True
        picTabC(0).Visible = False
        picTabN(1).Visible = False
        picTabC(1).Visible = True
        Call picTabC(1).ZOrder(0)
        
        cmdCancel.Enabled = False
        cmdNext.Enabled = False
        cmdNext.Visible = False
        cmdBack.Enabled = True
        cmdDone.Visible = True
        cmdDone.Default = True
    End If
    
    picItem(0).Visible = False
    picItem(1).Visible = False
    picItem(Index).Visible = True
    picTabC(Index).SetFocus
End Sub

Private Sub picTabC_GotFocus(Index As Integer)
    lblTabC(Index).FontBold = True
End Sub

Private Sub picTabC_LostFocus(Index As Integer)
    lblTabC(Index).FontBold = False
End Sub

Private Sub tmrTheme_Timer()
    InitTheme GetDesktopWindow()
   
    If OldTheme = m_iTheme Then
        If OldColor = GetSysColor(COLOR_BTNFACE) Then
            Exit Sub
        End If
    Else
        If OldColor = GetSysColor(COLOR_BTNFACE) Then
            Exit Sub
        End If
    End If
    picTabN(0).Refresh
    picTabN(1).Refresh
    picTabC(0).Refresh
    picTabC(1).Refresh
    picTop.Refresh
    
    Process
    
    createAlphaImage
End Sub
