VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColourise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

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
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long


Private m_lHue As Long
Private m_lSaturation As Long

Public Property Get Hue() As Long
   Hue = m_lHue
End Property
Public Property Let Hue(ByVal value As Long)
   m_lHue = value
End Property
Public Property Get Saturation() As Long
   Saturation = m_lSaturation
End Property
Public Property Let Saturation(ByVal value As Long)
   m_lSaturation = value
End Property

Public Sub Process( _
      cSrc As cDIBSection, _
      cDst As cDIBSection _
   )
      Dim bDib() As Byte
      Dim bDibDst() As Byte
      Dim tSA As SAFEARRAY2D
      Dim tSADst As SAFEARRAY2D
    
      ' Get the bits in the from DIB section:
      With tSA
          .cbElements = 1
          .cDims = 2
          .Bounds(0).lLbound = 0
          .Bounds(0).cElements = cSrc.Height
          .Bounds(1).lLbound = 0
          .Bounds(1).cElements = cSrc.BytesPerScanLine
          .pvData = cSrc.DIBSectionBitsPtr
      End With
      CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
   
      ' Get the bits in the from DIB section:
      With tSADst
          .cbElements = 1
          .cDims = 2
          .Bounds(0).lLbound = 0
          .Bounds(0).cElements = cDst.Height
          .Bounds(1).lLbound = 0
          .Bounds(1).cElements = cDst.BytesPerScanLine()
          .pvData = cDst.DIBSectionBitsPtr
      End With
      CopyMemory ByVal VarPtrArray(bDibDst()), VarPtr(tSADst), 4
      
      Dim x As Long
      Dim y As Long
      Dim xEnd As Long
      Dim yEnd As Long
      Dim H As Single
      Dim S As Single
      Dim L As Single
      Dim lR As Long
      Dim lG As Long
      Dim lB As Long
      Dim hDN As Single
      Dim sDN As Single
      
      ' Calculate denormalized Hue & Saturation:
      hDN = ((m_lHue * 6#) / 255#) - 1#
      sDN = (m_lSaturation / 255#)
      
      xEnd = cSrc.BytesPerScanLine() - 3
      yEnd = cSrc.Height - 1
           
      For x = 0 To xEnd Step 3
         For y = 0 To yEnd
            ' Obtain the luminance:
            RGBToHLS bDib(x + 2, y), bDib(x + 1, y), bDib(x, y), H, S, L
            ' Now get the new colour using the input hue and saturation
            HLSToRGB hDN, sDN, L, lR, lG, lB
            bDibDst(x + 2, y) = lR
            bDibDst(x + 1, y) = lG
            bDibDst(x, y) = lB
         Next y
      Next x
      
      CopyMemory ByVal VarPtrArray(bDibDst), 0&, 4
      CopyMemory ByVal VarPtrArray(bDib), 0&, 4
End Sub

