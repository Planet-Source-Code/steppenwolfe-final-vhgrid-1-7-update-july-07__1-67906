VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/* alpabelnd flags
Private Const AC_SRC_OVER                       As Long = &H0
Private Const AC_SRC_ALPHA                      As Long = &H1
Private Const AC_SRC_NO_PREMULT_ALPHA           As Long = &H1
Private Const AC_SRC_NO_ALPHA                   As Long = &H2
Private Const C_DST_NO_PREMULT_ALPHA            As Long = &H10
Private Const AC_DST_NO_ALPHA                   As Long = &H20
'/* pen styles
Private Const PS_SOLID                          As Long = 0
Private Const PS_DASH                           As Long = 1
Private Const PS_DOT                            As Long = 2
Private Const PS_DASHDOT                        As Long = 3
Private Const PS_DASHDOTDOT                     As Long = 4
Private Const PS_NULL                           As Long = 5
Private Const PS_INSIDEFRAME                    As Long = 6
Private Const PS_USERSTYLE                      As Long = 7
Private Const PS_ALTERNATE                      As Long = 8
Private Const PS_STYLE_MASK                     As Long = &HF
Private Const PS_ENDCAP_ROUND                   As Long = &H0
Private Const PS_ENDCAP_SQUARE                  As Long = &H100
Private Const PS_ENDCAP_FLAT                    As Long = &H200
Private Const PS_ENDCAP_MASK                    As Long = &HF00
Private Const PS_JOIN_ROUND                     As Long = &H0
Private Const PS_JOIN_BEVEL                     As Long = &H1000
Private Const PS_JOIN_MITER                     As Long = &H2000
Private Const PS_JOIN_MASK                      As Long = &HF000
Private Const PS_COSMETIC                       As Long = &H0
Private Const PS_GEOMETRIC                      As Long = &H10000
Private Const PS_TYPE_MASK                      As Long = &HF0000
'/* StretchBlt Modes
Private Const BLACKONWHITE                      As Long = 1
Private Const WHITEONBLACK                      As Long = 2
Private Const COLORONCOLOR                      As Long = 3
Private Const HALFTONE                          As Long = 4
Private Const MAXSTRETCHBLTMODE                 As Long = 4
'/* New StretchBlt Modes
Private Const STRETCH_ANDSCANS                  As Long = BLACKONWHITE
Private Const STRETCH_ORSCANS                   As Long = WHITEONBLACK
Private Const STRETCH_DELETESCANS               As Long = COLORONCOLOR
Private Const STRETCH_HALFTONE                  As Long = HALFTONE


Public Enum eStretchBltMode
    sbmColorOnColor = COLORONCOLOR
    sbmHalftone = HALFTONE
End Enum

Public Enum PEN_STYLE
    epsSolid = PS_SOLID
    epsDash = PS_DASH
    epsDot = PS_DOT
    epsDashDot = PS_DASHDOT
    epsDashDotDot = PS_DASHDOTDOT
    epsInsideFrame = PS_INSIDEFRAME
    epsUserStyle = PS_USERSTYLE
    epsAlternate = PS_ALTERNATE
    epsStyleMask = PS_STYLE_MASK
    epsEndcapRound = PS_ENDCAP_ROUND
    epsEndcapSquare = PS_ENDCAP_SQUARE
    epsEndcapFlat = PS_ENDCAP_FLAT
    epsEndcapMask = PS_ENDCAP_MASK
    epsJoinRound = PS_JOIN_MASK
    epsJoinBevel = PS_JOIN_BEVEL
    epsJoinMiter = PS_JOIN_MITER
    epsJoinMask = PS_JOIN_MASK
    epsCosmetic = PS_COSMETIC
    epsGeometric = PS_GEOMETRIC
    epsTypeMask = PS_TYPE_MASK
End Enum

Public Enum BRUSH_STYLE
    BS_SOLID = 0
    BS_NULL = 1
    BS_HOLLOW = BS_NULL
    BS_HATCHED = 2
    BS_PATTERN = 3
    BS_INDEXED = 4
    BS_DIBPATTERN = 5
    BS_DIBPATTERNPT = 6
    BS_PATTERN8X8 = 7
    BS_DIBPATTERN8X8 = 8
    BS_MONOPATTERN = 9
End Enum

Public Enum HATCH_STYLE
    HS_HORIZONTAL = 0
    HS_VERTICAL = 1
    HS_FDIAGONAL = 2
    HS_BDIAGONAL = 3
    HS_CROSS = 4
    HS_DIAGCROSS = 5
End Enum

Public Enum RAST_FLAGS
    SRCCOPY = &HCC0020
    SRCPAINT = &HEE0086
    SRCAND = &H8800C6
    SRCINVERT = &H660046
    SRCERASE = &H440328
    NOTSRCCOPY = &H330008
    NOTSRCERASE = &H1100A6
    MERGECOPY = &HC000CA
    MERGEPAINT = &HBB0226
    PATCOPYX = &HF00021
    PATPAINT = &HFB0A09
    PATINVERT = &H5A0049
    DSTINVERT = &H550009
    BLACKNESS = &H42
    WHITENESS = &HFF0062
End Enum

Public Enum GRADIENT_DIRECTION
    Fill_None = -1
    Fill_Horizontal = 0
    Fill_Vertical = 1
End Enum


Private Type VERSIONINFO
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
End Type

Private Type TRIVERTEX
    x                                           As Long
    y                                           As Long
    Red                                         As Integer
    Green                                       As Integer
    Blue                                        As Integer
    alpha                                       As Integer
End Type

Private Type BLENDFUNCTION
    BlendOp                                     As Byte
    BlendFlags                                  As Byte
    SourceConstantAlpha                         As Byte
    AlphaFormat                                 As Byte
End Type

Private Type PictDesc
    cbSize                                      As Long
    pictType                                    As Long
    hicon                                       As Long
    hpal                                        As Long
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Private Declare Function AlphaBlend Lib "Msimg32.dll" (ByVal DstHdc As Long, _
                                                       ByVal DstX As Long, _
                                                       ByVal DstY As Long, _
                                                       ByVal DstWidth As Long, _
                                                       ByVal DstHeight As Long, _
                                                       ByVal SrcHdc As Long, _
                                                       ByVal SrcX As Long, _
                                                       ByVal SrcY As Long, _
                                                       ByVal SrcWidth As Long, _
                                                       ByVal ScrHeight As Long, _
                                                       ByVal lpBlend As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal DstHdc As Long, _
                                             ByVal DstX As Long, _
                                             ByVal DstY As Long, _
                                             ByVal DstWidth As Long, _
                                             ByVal DstHeight As Long, _
                                             ByVal SrcHdc As Long, _
                                             ByVal SrcX As Long, _
                                             ByVal SrcY As Long, _
                                             ByVal dwRop As RAST_FLAGS) As Long

Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal DstHdc As Long, _
                                                     ByVal DstX As Long, _
                                                     ByVal DstY As Long, _
                                                     ByVal DstWidth As Long, _
                                                     ByVal DstHeight As Long, _
                                                     ByVal SrcHdc As Long, _
                                                     ByVal SrcX As Long, _
                                                     ByVal SrcY As Long, _
                                                     ByVal nSrcWidth As Long, _
                                                     ByVal nSrcHeight As Long, _
                                                     ByVal dwRop As RAST_FLAGS) As Long

Private Declare Function TransparentBlt Lib "Msimg32.dll" (ByVal DstHdc As Long, _
                                                           ByVal DstX As Long, _
                                                           ByVal DstY As Long, _
                                                           ByVal DstWidth As Long, _
                                                           ByVal DstHeight As Long, _
                                                           ByVal SrcHdc As Long, _
                                                           ByVal SrcX As Long, _
                                                           ByVal SrcY As Long, _
                                                           ByVal SrcWidth As Long, _
                                                           ByVal ScrHeight As Long, _
                                                           ByVal crTransparent As Long) As Boolean

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, _
                                                        ByVal nStretchMode As Long) As Long

Private Declare Function GradientFill Lib "Msimg32.dll" (ByVal hdc As Long, _
                                                         pVertex As TRIVERTEX, _
                                                         ByVal dwNumVertex As Long, _
                                                         pMesh As GRADIENT_RECT, _
                                                         ByVal dwNumMesh As Long, _
                                                         ByVal dwMode As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, _
                                                           ByVal HPALETTE As Long, _
                                                           ColorRef As Long) As Long

Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nXOrg As Long, _
                                                    ByVal nYOrg As Long, _
                                                    lppt As Any) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As VERSIONINFO) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, _
                                                                      riid As Any, _
                                                                      ByVal fOwn As Long, _
                                                                      ipic As IPicture) As Long

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal X1 As Long, _
                                                ByVal y1 As Long, _
                                                ByVal x2 As Long, _
                                                ByVal y2 As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long
                                                     
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long
                                                 
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, _
                                                       ByVal crColor As Long) As Long

Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long


Private m_bWin32        As Boolean
Private m_lMaskColor    As Long


Private Sub Class_Initialize()

    '/* os check & default mask
    m_bWin32 = Compatability_Check
    m_lMaskColor = &HFF00FF

End Sub

Public Property Get Mask() As Long
    Mask = m_lMaskColor
End Property

Public Property Let Mask(PropVal As Long)
    m_lMaskColor = PropVal
End Property

Public Sub AlphaBlit(ByVal lDstDc As Long, _
                     ByVal lDstX As Long, _
                     ByVal lDstY As Long, _
                     ByVal lDstWidth As Long, _
                     ByVal lDstHeight As Long, _
                     ByVal lSrcDc As Long, _
                     ByVal lSrcX As Long, _
                     ByVal lSrcY As Long, _
                     ByVal lSrcWidth As Long, _
                     ByVal lSrcHeight As Long, _
                     ByVal lTrsIdx As Byte)

Dim lBfPtr     As Long
Dim tBlend     As BLENDFUNCTION

    With tBlend
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0&
        .SourceConstantAlpha = lTrsIdx
        .AlphaFormat = 0&
    End With
    
    CopyMemory lBfPtr, tBlend, 4&
    AlphaBlend lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, lBfPtr
    
End Sub

Public Sub AlphaPixel(ByVal lDstDc As Long, _
                      ByVal lDstX As Long, _
                      ByVal lDstY As Long, _
                      ByVal lDstWidth As Long, _
                      ByVal lDstHeight As Long, _
                      ByVal lSrcDc As Long, _
                      ByVal lSrcX As Long, _
                      ByVal lSrcY As Long, _
                      ByVal lSrcWidth As Long, _
                      ByVal lSrcHeight As Long, _
                      ByVal lTrsIdx As Long)

 Dim tBlend As BLENDFUNCTION
 Dim lBfPtr As Long

    If (lTrsIdx = 0) Or (lTrsIdx > 255) Then
        lTrsIdx = 200
    End If
    With tBlend
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0&
        .SourceConstantAlpha = lTrsIdx
        .AlphaFormat = AC_SRC_ALPHA
    End With
    
    CopyMemory lBfPtr, tBlend, 4
    AlphaBlend lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, lBfPtr

End Sub

Public Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR, _
                               Optional ByVal lAlpha As Long = 128) As Long

Dim lCFrom As Long
Dim lCTo   As Long
Dim lSrcR  As Long
Dim lSrcG  As Long
Dim lSrcB  As Long
Dim lDstR  As Long
Dim lDstG  As Long
Dim lDstB  As Long

    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000

    BlendColor = RGB(((lSrcR * lAlpha) / 255) + ((lDstR * (255 - lAlpha)) / 255), ((lSrcG * lAlpha) / 255) + ((lDstG * (255 - lAlpha)) / 255), ((lSrcB * lAlpha) / 255) + ((lDstB * (255 - lAlpha)) / 255))

End Property

Public Sub Blit(ByVal lDstDc As Long, _
                ByVal lDstX As Long, _
                ByVal lDstY As Long, _
                ByVal lDstWidth As Long, _
                ByVal lDstHeight As Long, _
                ByVal lSrcDc As Long, _
                ByVal lSrcX As Long, _
                ByVal lSrcY As Long, _
                ByVal eFlags As RAST_FLAGS)


    BitBlt lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, eFlags

End Sub

Public Function IconToPicture(ByVal lhIcon As Long) As Picture

Dim Pic             As PictDesc
Dim Guid(0 To 3)    As Long

    With Pic
        .cbSize = Len(Pic)
        .pictType = vbPicTypeIcon
        .hicon = lhIcon
    End With
    Guid(0) = &H7BF80980
    Guid(1) = &H101ABF32
    Guid(2) = &HAA00BB8B
    Guid(3) = &HAB0C3000
    OleCreatePictureIndirect Pic, Guid(0), True, IconToPicture

End Function

Public Sub Stretch(ByVal lDstDc As Long, _
                   ByVal lDstX As Long, _
                   ByVal lDstY As Long, _
                   ByVal lDstWidth As Long, _
                   ByVal lDstHeight As Long, _
                   ByVal lSrcDc As Long, _
                   ByVal lSrcX As Long, _
                   ByVal lSrcY As Long, _
                   ByVal lSrcWidth As Long, _
                   ByVal lSrcHeight As Long, _
                   ByVal eFlags As RAST_FLAGS, _
                   Optional ByVal eStretchMode As eStretchBltMode = sbmColorOnColor)

Dim lMode As Long

    lMode = SetStretchBltMode(lDstDc, eStretchMode)
    StretchBlt lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, eFlags
    SetStretchBltMode lDstDc, lMode
    
End Sub

Public Sub Transparent(ByVal lDstDc As Long, _
                       ByVal lDstX As Long, _
                       ByVal lDstY As Long, _
                       ByVal lDstWidth As Long, _
                       ByVal lDstHeight As Long, _
                       ByVal lSrcDc As Long, _
                       ByVal lSrcX As Long, _
                       ByVal lSrcY As Long, _
                       ByVal lSrcWidth As Long, _
                       ByVal lSrcHeight As Long, _
                       ByVal lTrsColor As Long, _
                       Optional ByVal bRemapTrans As Boolean)


    If bRemapTrans Then
        lTrsColor = RemapTrans(lSrcDc, lSrcX, lSrcY)
    End If
    If m_bWin32 Then
        TransparentBlt lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, lTrsColor
    Else
        TransMask lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, lTrsColor
    End If

End Sub

Public Function TransMask(ByVal lDstDc As Long, _
                          ByVal lDstX As Long, _
                          ByVal lDstY As Long, _
                          ByVal lDstWidth As Long, _
                          ByVal lDstHeight As Long, _
                          ByVal lSrcDc As Long, _
                          ByVal lSrcX As Long, _
                          ByVal lSrcY As Long, _
                          ByVal lSrcWidth As Long, _
                          ByVal lSrcHeight As Long, _
                          ByVal lTrsColor As Long) As Boolean

Dim lColor              As Long
Dim lTempMaskDc         As Long
Dim lMaskDc             As Long
Dim lTempDc             As Long
Dim lTrsDc              As Long
Dim lMaskBmp            As Long
Dim lMaskBmpOld         As Long
Dim lTmpMaskBmp         As Long
Dim lTmpMaskBmpOld      As Long
Dim lTmpBmp             As Long
Dim lTmpBmpOld          As Long
Dim lTrsBmp             As Long
Dim lTrsBmpOld          As Long
    
    If Not CreateDc(False, lTempMaskDc, lMaskBmp, lMaskBmpOld, lSrcWidth, lSrcHeight) Then
        If Not CreateDc(True, lMaskDc, lTmpMaskBmp, lTmpMaskBmpOld, lSrcWidth, lSrcHeight) Then
            If Not CreateDc(False, lTempDc, lTmpBmp, lTmpBmpOld, lDstWidth, lDstHeight) Then
                BitBlt lTempMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, lSrcDc, lSrcX, lSrcY, SRCCOPY
                If lTrsColor = &HFFFFFF Then
                    BitBlt lMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, lTempMaskDc, 0&, 0&, WHITENESS
                    BitBlt lMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, lTempMaskDc, 0&, 0&, SRCINVERT
                Else
                    If CreateDc(True, lTrsDc, lTrsBmp, lTrsBmpOld, lSrcWidth, lSrcHeight) Then
                        lColor = SetBkColor(lTempMaskDc, lTrsColor)
                        BitBlt lTrsDc, 0&, 0&, lSrcWidth, lSrcHeight, lTempMaskDc, 0&, 0&, SRCCOPY
                        SetBkColor lTempMaskDc, lColor
                        BitBlt lTempMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, lTrsDc, 0&, 0&, SRCPAINT
                        BitBlt lMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, lTempMaskDc, 0&, 0&, WHITENESS
                        BitBlt lMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, lTrsDc, 0&, 0&, SRCINVERT
                        SelectObject lTrsDc, lTrsBmpOld
                        DeleteObject lTrsBmp
                        lTrsBmp = 0
                        DeleteObject lTrsDc
                        lTrsDc = 0
                    End If
                End If
                BitBlt lTempDc, 0&, 0&, lDstWidth, lDstHeight, lDstDc, lDstX, lDstY, SRCCOPY
                If Not lSrcWidth = lDstWidth Or Not lSrcHeight = lDstHeight Then
                    StretchBlt lTempDc, 0&, 0&, lDstWidth, lDstHeight, lMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, SRCPAINT
                    StretchBlt lTempDc, 0&, 0&, lDstWidth, lDstHeight, lTempMaskDc, 0&, 0&, lSrcWidth, lSrcHeight, SRCAND
                Else
                    BitBlt lTempDc, 0&, 0&, lSrcWidth, lSrcHeight, lMaskDc, 0&, 0&, SRCPAINT
                    BitBlt lTempDc, 0&, 0&, lSrcWidth, lSrcHeight, lTempMaskDc, 0&, 0&, SRCAND
                End If
                
                If Not BitBlt(lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lTempDc, 0&, 0&, SRCCOPY) = 0 Then
                    TransMask = True
                End If
            End If
        End If
    End If

    If Not (lTempDc = 0) Then
        SelectObject lTempDc, lTmpBmpOld
        DeleteObject lTmpBmp
        DeleteObject lTempDc
    End If
    
    If Not (lMaskDc = 0) Then
        SelectObject lMaskDc, lTmpMaskBmpOld
        DeleteObject lTmpMaskBmp
        DeleteObject lMaskDc
    End If
    
    If Not (lTempMaskDc = 0) Then
        SelectObject lTempMaskDc, lMaskBmpOld
        DeleteObject lMaskBmp
        DeleteObject lTempMaskDc
    End If
    
End Function

Public Sub Tile(ByVal lDstDc As Long, _
                ByVal lDstX As Long, _
                ByVal lDstY As Long, _
                ByVal lDstWidth As Long, _
                ByVal lDstHeight As Long, _
                ByVal lSrcDc As Long, _
                ByVal lSrcWidth As Long, _
                ByVal lSrcHeight As Long, _
                Optional ByVal lOffsetX As Long, _
                Optional ByVal lOffsetY As Long)

Dim SrcWidth            As Long
Dim SrcHeight           As Long
Dim lSrcX               As Long
Dim lSrcY               As Long
Dim lSrcStartX          As Long
Dim lSrcStartY          As Long
Dim lSrcStartWidth      As Long
Dim lSrcStartHeight     As Long
Dim lNDestX             As Long
Dim lNDestY             As Long
Dim lNDestWidth         As Long
Dim lNDestHeight        As Long

    SrcWidth = lSrcWidth
    SrcHeight = lSrcHeight
    lSrcStartX = ((lDstX + lOffsetX) Mod SrcWidth)
    lSrcStartY = ((lDstY + lOffsetY) Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    lNDestY = lDstY
    lNDestHeight = lSrcStartHeight

    Do While lNDestY < (lDstY + lDstHeight)
        If (lNDestY + lNDestHeight) > (lDstY + lDstHeight) Then
            lNDestHeight = lDstY + lDstHeight - lNDestY
        End If
        lNDestWidth = lSrcStartWidth
        lNDestX = lDstX
        lSrcX = lSrcStartX
        Do While lNDestX < (lDstX + lDstWidth)
            If (lNDestX + lNDestWidth) > (lDstX + lDstWidth) Then
                lNDestWidth = lDstX + lDstWidth - lNDestX
                If lNDestWidth = 0 Then
                    lNDestWidth = 4
                End If
            End If
            Blit lDstDc, lNDestX, lNDestY, lNDestWidth, lNDestHeight, lSrcDc, lSrcX, lSrcY, vbSrcCopy
            lNDestX = lNDestX + lNDestWidth
            lSrcX = 0
            lNDestWidth = SrcWidth
        Loop
        lNDestY = lNDestY + lNDestHeight
        lSrcY = 0
        lNDestHeight = SrcHeight
    Loop

End Sub

Private Function CreateDc(ByVal bMono As Boolean, _
                          ByRef lHdc As Long, _
                          ByRef lBitmap As Long, _
                          ByRef lBitmapOld As Long, _
                          ByVal lWidth As Long, _
                          ByVal lHeight As Long) As Boolean
  
Dim lWndDc  As Long
Dim lHwnd   As Long

    lHdc = 0
    lBitmap = 0
    lBitmapOld = 0

    If bMono Then
        lWndDc = 0
    Else
        lHwnd = GetDesktopWindow
        lWndDc = GetDC(lHwnd)
    End If
    lHdc = CreateCompatibleDC(lWndDc)
    If bMono Then
        lWndDc = lHdc
    End If
    
    If Not lHdc = 0 Then
        lBitmap = CreateCompatibleBitmap(lWndDc, lWidth, lHeight)
        If Not lBitmap = 0 Then
            lBitmapOld = SelectObject(lHdc, lBitmap)
            CreateDc = True
        End If
    End If

    If Not bMono Then
        ReleaseDC lHwnd, lWndDc
    End If

End Function

Public Sub Gradient(ByVal lHdc As Long, _
                    ByVal lLeft As Long, _
                    ByVal lWidth As Long, _
                    ByVal lTop As Long, _
                    ByVal lHeight As Long, _
                    ByVal lStartColor As Long, _
                    ByVal lEndColor As Long, _
                    ByVal eDirection As GRADIENT_DIRECTION, _
                    Optional ByVal bJuxtapose As Boolean)

Dim btClrs(3)           As Byte
Dim btVert(7)           As Byte
Dim tGradRect           As GRADIENT_RECT
Dim tVert(1)            As TRIVERTEX

    '/* Init vertices : Set Position : Define Size
    With tVert(0)
        If bJuxtapose Then
            .x = lLeft
            .y = lTop
        Else
            .x = lLeft + lWidth
            .y = lTop + lHeight
        End If
    End With
    With tVert(1)
        If bJuxtapose Then
            .x = lLeft + lWidth
            .y = lTop + lHeight
        Else
            .x = lLeft
            .y = lTop
        End If
    End With
        
    '/* Init vertices :colors, initial
    CopyMemory btClrs(0), lEndColor, &H4
    '/* Red
    btVert(1) = btClrs(0)
    '/* Green
    btVert(3) = btClrs(1)
    '/* Blue
    btVert(5) = btClrs(2)
    CopyMemory tVert(0).Red, btVert(0), &H8
    '/* Init vertices :colors, final
    CopyMemory btClrs(0), lStartColor, &H4
    '/* Red
    btVert(1) = btClrs(0)
    '/* Green
    btVert(3) = btClrs(1)
    '/* Blue
    btVert(5) = btClrs(2)
    CopyMemory tVert(1).Red, btVert(0), &H8
    '/* Init gradient rect
    With tGradRect
        .UpperLeft = 0&
        .LowerRight = 1&
    End With
    '/* Fill the DC
    GradientFill lHdc, tVert(0), 2, tGradRect, 1, eDirection

End Sub

Public Sub DrawRectangle(ByVal lHdc As Long, _
                         ByVal lX As Long, _
                         ByVal lY As Long, _
                         ByVal lRight As Long, _
                         ByVal lBottom As Long, _
                         Optional ByVal lFrameColor As Long = -1, _
                         Optional ByVal lFillColor As Long = -1, _
                         Optional ByVal lLineWidth As Long = 1, _
                         Optional ByVal ePenStyle As PEN_STYLE = epsSolid)

Dim lhBrush     As Long
Dim lhPen       As Long
Dim lhPenOld    As Long
Dim tRect       As RECT

    If Not (lFrameColor = -1) Then
        lhPen = CreatePen(ePenStyle, lLineWidth, lFrameColor)
        lhPenOld = SelectObject(lHdc, lhPen)
        Rectangle lHdc, lX, lY, lRight, lBottom
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        lhPen = 0
    Else
        Rectangle lHdc, lX, lY, lRight, lBottom
    End If
    If Not (lFillColor = -1) Then
        With tRect
            .left = lX
            .Right = lRight
            .top = lY
            .Bottom = lBottom
            InflateRect tRect, -1, -1
            .left = .left + 1
        End With
        lhBrush = CreateSolidBrush(lFillColor)
        FillRect lHdc, tRect, lhBrush
        DeleteObject lhBrush
        lhBrush = 0
    End If
    
End Sub

Public Sub DrawPatternRect(ByVal lHdc As Long, _
                           ByVal lX As Long, _
                           ByVal lY As Long, _
                           ByVal lRight As Long, _
                           ByVal lBottom As Long, _
                           Optional ByVal lhBitmap As Long)

Dim lhBrush As Long
Dim tRect   As RECT

    With tRect
        .left = lX
        .top = lY
        .Right = lRight
        .Bottom = lBottom
    End With
    lhBrush = CreatePatternBrush(lhBitmap)
    FillRect lHdc, tRect, lhBrush
    DeleteObject lhBrush
    lhBrush = 0

End Sub

Public Sub DrawHatchRect(ByVal lHdc As Long, _
                         ByVal lX As Long, _
                         ByVal lY As Long, _
                         ByVal lRight As Long, _
                         ByVal lBottom As Long, _
                         Optional ByVal lBrushColor As Long = &H0, _
                         Optional ByVal ePattern As HATCH_STYLE = HS_VERTICAL)

Dim lhBrush     As Long
Dim tRect       As RECT

    With tRect
        .left = lX
        .top = lY
        .Right = lRight
        .Bottom = lBottom
    End With
    lhBrush = CreateHatchBrush(ePattern, lBrushColor)
    FillRect lHdc, tRect, lhBrush
    DeleteObject lhBrush
    lhBrush = 0

End Sub

Public Sub FrameRectangle(ByVal lHdc As Long, _
                          ByVal lX As Long, _
                          ByVal lY As Long, _
                          ByVal lRight As Long, _
                          ByVal lBottom As Long, _
                          Optional ByVal lFrameColor As Long = &H0)

Dim lhBrush As Long
Dim tRect   As RECT

    With tRect
        .left = lX
        .top = lY
        .Right = lRight
        .Bottom = lBottom
    End With
    lhBrush = CreateSolidBrush(lFrameColor)
    FrameRect lHdc, tRect, lhBrush
    DeleteObject lhBrush
    lhBrush = 0
    
End Sub

Public Sub FramePattern(ByVal lHdc As Long, _
                        ByVal lX As Long, _
                        ByVal lY As Long, _
                        ByVal lRight As Long, _
                        ByVal lBottom As Long, _
                        ByVal lDepth As Long, _
                        Optional ByVal ePenStyle As PEN_STYLE = epsSolid, _
                        Optional ByVal lFrameColor As Long = &H0)

Dim lCt         As Long
Dim lhPen       As Long
Dim lhPenOld    As Long

    lhPen = CreatePen(ePenStyle, 1&, lFrameColor)
    lhPenOld = SelectObject(lHdc, lhPen)
    If (lDepth > 1) Then
        For lCt = 0 To lDepth - 1
            Rectangle lHdc, lX + lCt, lY + lCt, lRight - lCt, lBottom - lCt
        Next lCt
    Else
        Rectangle lHdc, lX, lY, lRight, lBottom
    End If
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen
    lhPen = 0

End Sub

Public Function XPShift(ByVal lColor As Long, _
                        Optional ByVal Base As Long = &HB0) As Long

Dim lRed        As Long
Dim lBlue       As Long
Dim lGreen      As Long
Dim lDelta      As Long

    lBlue = ((lColor \ &H10000) Mod &H100)
    lGreen = ((lColor \ &H100) Mod &H100)
    lRed = (lColor And &HFF)
    lDelta = &HFF - Base
    lBlue = Base + lBlue * lDelta \ &HFF
    lGreen = Base + lGreen * lDelta \ &HFF
    lRed = Base + lRed * lDelta \ &HFF
    If lRed > 255 Then lRed = 255
    If lGreen > 255 Then lGreen = 255
    If lBlue > 255 Then lBlue = 255

    XPShift = lRed + 256& * lGreen + 65536 * lBlue

End Function

Private Function LightenColor(ByVal RValue As Integer, _
                              ByVal GValue As Integer, _
                              ByVal BValue As Integer, _
                              Optional ByVal Offset As Long = 1) As Long

Dim intR As Integer
Dim intG As Integer
Dim intB As Integer

    intR = Abs(RValue + Offset)
    intG = Abs(GValue + Offset)
    intB = Abs(BValue + Offset)

    LightenColor = RGB(intR, intG, intB)

End Function

Private Function DarkenColor(ByVal RValue As Integer, _
                             ByVal GValue As Integer, _
                             ByVal BValue As Integer, _
                             Optional ByVal Offset As Long = 1) As Long

Dim intR As Integer
Dim intG As Integer
Dim intB As Integer

    intR = Abs(RValue - Offset)
    intG = Abs(GValue - Offset)
    intB = Abs(BValue - Offset)
    DarkenColor = RGB(intR, intG, intB)

End Function

Public Function TranslateColor(ByVal oColor As OLE_COLOR, _
                               Optional ByVal lhPal As Long = 0) As Long
                        
    If OleTranslateColor(oColor, lhPal, TranslateColor) Then
        TranslateColor = -1
    End If
    
End Function

Private Function RemapTrans(ByVal lHdc As Long, _
                            ByVal lSrcX As Long, _
                            ByVal lSrcY As Long) As Long

    RemapTrans = GetPixel(lHdc, lSrcX, lSrcY)
    
End Function

Public Sub SetBrushOrigin(ByVal lHdc As Long, _
                          ByVal lX As Long, _
                          ByVal lY As Long, _
                          ByVal lWidth As Long, _
                          ByVal lHeight As Long)

   SetBrushOrgEx lHdc, -lX And lWidth, -lY And lHeight, ByVal 0&
   
End Sub

Public Function GetBlendVal(ByVal sTransparency As Single) As Long

    If (sTransparency < 0) Then
        sTransparency = 0
    ElseIf (sTransparency > 250) Then
        sTransparency = 255
    End If
    GetBlendVal = CLng(sTransparency / 100 * 255) * &H10000
  
End Function

Private Function Compatability_Check() As Boolean

Dim tVer  As VERSIONINFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        Compatability_Check = True
    End If

End Function
