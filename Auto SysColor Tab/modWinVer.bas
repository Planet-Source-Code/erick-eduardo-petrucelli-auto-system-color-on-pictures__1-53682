Attribute VB_Name = "modWinVer"
Option Explicit

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Const CLR_INVALID = -1
Public Const CLR_NONE = CLR_INVALID

Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" ( _
    ByVal pszThemeFileName As Long, _
    ByVal dwMaxNameChars As Long, _
    ByVal pszColorBuff As Long, _
    ByVal cchMaxColorChars As Long, _
    ByVal pszSizeBuff As Long, _
    ByVal cchMaxSizeChars As Long _
   ) As Long
Private Declare Function GetThemeFilename Lib "uxtheme.dll" _
   (ByVal hTheme As Long, _
    ByVal iPartId As Long, _
    ByVal iStateId As Long, _
    ByVal iPropId As Long, _
    pszThemeFileName As Long, _
    ByVal cchMaxBuffChars As Long _
   ) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Private m_bIsXp As Boolean
Private m_bIsNt As Boolean
Private m_bIs2000OrAbove As Boolean
Private m_bHasGradientAndTransparency As Boolean

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Private Type OSVERSIONINFO
   dwVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion(0 To 127) As Byte
End Type

Private Const VER_PLATFORM_WIN32_NT = 2
Global m_iTheme As Integer
Global Caminho As String
Global lBackColor As Long

Public Function IsWinXP() As Boolean
   Dim tOSV As OSVERSIONINFO
   tOSV.dwVersionInfoSize = Len(tOSV)
   GetVersionEx tOSV
   
   m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
   If (tOSV.dwMajorVersion > 5) Then
      m_bHasGradientAndTransparency = True
      m_bIsXp = True
      m_bIs2000OrAbove = True
   ElseIf (tOSV.dwMajorVersion = 5) Then
      m_bHasGradientAndTransparency = True
      m_bIs2000OrAbove = True
      If (tOSV.dwMinorVersion >= 1) Then
         m_bIsXp = True
      End If
   ElseIf (tOSV.dwMajorVersion = 4) Then ' NT4 or 9x/ME/SE
      If (tOSV.dwMinorVersion >= 10) Then
         m_bHasGradientAndTransparency = True
      End If
   Else ' Too old
   End If
   
   IsWinXP = m_bIsXp
End Function

Public Sub InitTheme(ByVal hwnd As Long)
Dim hTheme As Long
Dim lPtrColorName As Long
Dim lPtrThemeFile As Long
Dim sThemeFile As String
Dim sColorName As String
Dim sShellStyle As String
Dim hRes As Long
Dim iPos As Long
Dim lhWndD As Long
Dim lhDCC As Long
Dim lBitsPixel As Long

   If IsWinXP = True Then
      On Error Resume Next
      hTheme = OpenThemeData(hwnd, StrPtr("ExplorerBar"))
      If Not (hTheme = 0) Then
         
         ReDim bThemeFile(0 To 260 * 2) As Byte
         lPtrThemeFile = VarPtr(bThemeFile(0))
         ReDim bColorName(0 To 260 * 2) As Byte
         lPtrColorName = VarPtr(bColorName(0))
         hRes = GetCurrentThemeName(lPtrThemeFile, 260, lPtrColorName, 260, 0, 0)
         
         sThemeFile = bThemeFile
         iPos = InStr(sThemeFile, vbNullChar)
         If (iPos > 1) Then sThemeFile = Left(sThemeFile, iPos - 1)
         sColorName = bColorName
         iPos = InStr(sColorName, vbNullChar)
         If (iPos > 1) Then sColorName = Left(sColorName, iPos - 1)
         
         Select Case LCase(sColorName)
         Case "normalcolor"
            m_iTheme = 1
         Case "metallic"
            m_iTheme = 2
         Case "homestead"
            m_iTheme = 3
         Case Else
            m_iTheme = 0
         End Select
         
         CloseThemeData hTheme
      Else
         m_iTheme = 0
      End If
   End If
End Sub

Public Sub GetPath()
    Caminho = App.Path
    If Right$(Caminho, 1) <> "\" Then
        Caminho = Caminho & "\"
    End If
End Sub

Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal Alpha As Long = 128 _
   ) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
      )
End Property


