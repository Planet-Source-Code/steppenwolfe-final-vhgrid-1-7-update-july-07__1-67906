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


Private Const VER_PLATFORM_WIN32s               As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS        As Long = 1
Private Const VER_PLATFORM_WIN32_NT             As Long = 2
Private Const ICC_USEREX_CLASSES                As Long = &H200
Private Const AC_SRC_OVER                       As Long = &H0
Private Const AC_SRC_ALPHA                      As Long = &H1

Private Const AC_SRC_NO_PREMULT_ALPHA           As Long = &H1
Private Const AC_SRC_NO_ALPHA                   As Long = &H2
Private Const C_DST_NO_PREMULT_ALPHA            As Long = &H10
Private Const AC_DST_NO_ALPHA                   As Long = &H20

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

Private Type SAFEARRAYBOUND
    cElements                                   As Long
    lLbound                                     As Long
End Type

Private Type SAFEARRAY1D
    cDims                                       As Integer
    fFeatures                                   As Integer
    cbElements                                  As Long
    cLocks                                      As Long
    pvData                                      As Long
    Bounds                                      As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims                                       As Integer
    fFeatures                                   As Integer
    cbElements                                  As Long
    cLocks                                      As Long
    pvData                                      As Long
    Bounds(0 To 1)                              As SAFEARRAYBOUND
End Type

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

Private Type BITMAP
    bmType                                      As Long
    bmWidth                                     As Long
    bmHeight                                    As Long
    bmWidthBytes                                As Long
    bmPlanes                                    As Integer
    bmBitsPixel                                 As Integer
    bmBits                                      As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                                      As Long
    biWidth                                     As Long
    biHeight                                    As Long
    biPlanes                                    As Integer
    biBitCount                                  As Integer
    biCompression                               As Long
    biSizeImage                                 As Long
    biXPelsPerMeter                             As Long
    biYPelsPerMeter                             As Long
    biClrUsed                                   As Long
    biClrImportant                              As Long
End Type

Private Type RECT
    lLeft                                       As Long
    lTop                                        As Long
    lRight                                      As Long
    lBottom                                     As Long
End Type

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

Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Enum GRADIENT_DIRECTION
    [Fill_None] = -1
    [Fill_Horizontal] = 0
    [Fill_Vertical] = 1
End Enum

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

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

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

Private Declare Function GradientFill Lib "Msimg32.dll" (ByVal hdc As Long, _
                                                         pVertex As TRIVERTEX, _
                                                         ByVal dwNumVertex As Long, _
                                                         pMesh As GRADIENT_RECT, _
                                                         ByVal dwNumMesh As Long, _
                                                         ByVal dwMode As Long) As Long


Private Declare Function StretchDIBits Lib "gdi32" (ByVal DstHdc As Long, _
                                                    ByVal DstX As Long, _
                                                    ByVal DstY As Long, _
                                                    ByVal dX As Long, _
                                                    ByVal dy As Long, _
                                                    ByVal SrcX As Long, _
                                                    ByVal SrcY As Long, _
                                                    ByVal wSrcWidth As Long, _
                                                    ByVal wSrcHeight As Long, _
                                                    lpBits As Any, _
                                                    lpBitsInfo As BITMAPINFOHEADER, _
                                                    ByVal wUsage As Long, _
                                                    ByVal dwRop As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long
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

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long



Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long


Private m_bWin32                                As Boolean
Private m_lXOriginOffset                        As Long
Private m_lYOriginOffset                        As Long
Private m_lMaskColor                            As Long


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

 
 Dim tBlend     As BLENDFUNCTION
 Dim lBfPtr     As Long

    If (lTrsIdx = 0) Or (lTrsIdx > 255) Then
        lTrsIdx = 200
    End If
    With tBlend
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = lTrsIdx
        .AlphaFormat = 0
    End With
    
    CopyMemory lBfPtr, tBlend, 4
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

 Dim tBlend     As BLENDFUNCTION
 Dim lBfPtr     As Long

    If (lTrsIdx = 0) Or (lTrsIdx > 255) Then
        lTrsIdx = 200
    End If
    With tBlend
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = lTrsIdx
        .AlphaFormat = AC_SRC_ALPHA
    End With
    
    CopyMemory lBfPtr, tBlend, 4
    AlphaBlend lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, lBfPtr

End Sub

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
                   Optional ByVal bPrsClr As Boolean)

    If Not bPrsClr Then
        StretchBlt lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, eFlags
    Else
        StretchBlt lDstDc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDc, lSrcX, lSrcY, lSrcWidth, lSrcHeight, eFlags
    End If

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
    
    If Not CreateDc(False, lTempMaskDc, lMaskBmp, lMaskBmpOld, lSrcWidth, lSrcHeight) Then GoTo Handler
    If Not CreateDc(True, lMaskDc, lTmpMaskBmp, lTmpMaskBmpOld, lSrcWidth, lSrcHeight) Then GoTo Handler
    If Not CreateDc(False, lTempDc, lTmpBmp, lTmpBmpOld, lDstWidth, lDstHeight) Then GoTo Handler
    
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
                  
On Error GoTo 0

Handler:
    If Not lTempDc = 0 Then
        SelectObject lTempDc, lTmpBmpOld
        DeleteObject lTmpBmp
        DeleteObject lTempDc
    End If
    
    If Not lMaskDc = 0 Then
        SelectObject lMaskDc, lTmpMaskBmpOld
        DeleteObject lTmpMaskBmp
        DeleteObject lMaskDc
    End If
    
    If Not lTempMaskDc = 0 Then
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
                    ByVal lRight As Long, _
                    ByVal lTop As Long, _
                    ByVal lBottom As Long, _
                    ByVal lStartColor As Long, _
                    ByVal lEndColor As Long, _
                    ByVal eDirection As GRADIENT_DIRECTION, _
                    Optional ByVal bJuxtapose As Boolean)

Dim btClrs(3)           As Byte
Dim btVert(7)           As Byte
Dim tGradRect           As GRADIENT_RECT
Dim tVert(1)            As TRIVERTEX

On Error GoTo Handler
    
    '/* Check If the Fill is From Left to Right
    If bJuxtapose Then
        '/* Init vertices : Set Position : Define Size
        tVert(0).x = lLeft: tVert(1).x = lLeft + lRight
        tVert(0).y = lTop: tVert(1).y = lTop + lBottom
    Else
        '/* Init vertices : Set Position : Define Size
        tVert(0).x = lLeft + lRight: tVert(1).x = lLeft
        tVert(0).y = lTop + lBottom: tVert(1).y = lTop
    End If
        
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
        .UpperLeft = 0
        .LowerRight = 1
    End With
    '/* Fill the DC
    GradientFill lHdc, tVert(0), 2, tGradRect, 1, eDirection
    
Handler:
    On Error GoTo 0

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

Public Sub SetBrushOrigin(ByVal hdc As Long, _
                          ByVal xPixels As Long, _
                          ByVal yPixels As Long, _
                          ByVal lWidth As Long, _
                          ByVal lHeight As Long)

   SetBrushOrgEx hdc, -xPixels And lWidth, -yPixels And lHeight, ByVal 0&
   
End Sub

Public Function GetBlendVal(ByVal sTransparency As Single) As Long

  If sTransparency < 0 Then sTransparency = 0
  If sTransparency > 250 Then sTransparency = 255
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
