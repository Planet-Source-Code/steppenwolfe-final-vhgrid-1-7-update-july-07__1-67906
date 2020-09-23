VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreDc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const DIB_RGB_COLORS                    As Long = &H0
Private Const BI_RGB                            As Long = &H0

Public Enum eGrayScaleFormulas
    gsclCCIR709 = 0
    gsclNTSCPAL = 1
    gsclSimpleAvg = 2
End Enum

Private Type SAFEARRAYBOUND
    cElements                                   As Long
    lLbound                                     As Long
End Type

Private Type SAFEARRAYID
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
    biClrImportant  As Long
End Type

Private Type RGBQUAD
    rgbBlue                                     As Byte
    rgbGreen                                    As Byte
    rgbRed                                      As Byte
    rgbReserved                                 As Byte
End Type

Private Type BITMAPINFO
    bmiHeader                                   As BITMAPINFOHEADER
    bmiColors                                   As RGBQUAD
End Type

Private Type Guid
    Data1                                       As Long
    Data2                                       As Integer
    Data3                                       As Integer
    Data4(7)                                    As Byte
End Type


Private Type PICTUREINFO
    Size                                        As Long
    type                                        As Long
    hBmp                                        As Long
    hPal                                        As Long
    reserved                                    As Long
End Type


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, _
                                                                     ByVal Length As Long)

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, _
                                                       pBitmapInfo As Any, _
                                                       ByVal un As Long, _
                                                       lpVoid As Any, _
                                                       ByVal handle As Long, _
                                                       ByVal dw As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTUREINFO, _
                                                                      RefIID As Guid, _
                                                                      ByVal fPictureOwnsHandle As Long, _
                                                                      iPic As IPicture) As Long

Private m_bUseAlpha                             As Boolean
Private m_bInit                                 As Boolean
Private m_bMono                                 As Boolean
Private m_lHdc                                  As Long
Private m_hBmpOld                               As Long
Private m_hBmp                                  As Long
Private m_lHandle                               As Long
Private m_hDIb                                  As Long
Private m_lpBits                                As Long
Private m_lWidth                                As Long
Private m_lHeight                               As Long
Private m_lSizeX                                As Long
Private m_lSizeY                                As Long
Private m_lDibDC                                As Long
Private m_hDibOld                               As Long
Private m_lPtr                                  As Long
Private m_tBI                                   As BITMAPINFO
Private m_tBIH                                  As BITMAPINFOHEADER


'/* use 32bit image
Public Property Get UseAlpha() As Boolean
    UseAlpha = m_bUseAlpha
End Property

Public Property Let UseAlpha(PropVal As Boolean)
    m_bUseAlpha = PropVal
End Property

Public Property Get hdc() As Long
    hdc = m_lHdc
End Property

Public Property Get handle() As Long
    handle = m_hBmp
End Property

Public Property Get Handle32() As Long
    Handle32 = m_lHandle
End Property

Public Property Get Bits() As Long
    Bits = m_lpBits
End Property

Public Property Let Bits(ByVal PropVal As Long)
    m_lpBits = PropVal
End Property

Public Property Get BytesPerScanLine() As Long
    '/* scans must align on dword boundaries m_lDibDC
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get hDib() As Long
    hDib = m_hDIb
End Property

Public Property Get DibDc() As Long
    DibDc = m_lDibDC
End Property

Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Let Height(ByVal lHeight As Long)
    If (lHeight > m_lHeight) Then
        ImageCreate m_lWidth, lHeight
    End If
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Let Width(ByVal lWidth As Long)
    If (lWidth > m_lWidth) Then
        ImageCreate lWidth, m_lHeight
    End If
End Property

Public Property Get Mono() As Boolean
    Mono = m_bMono
End Property

Public Property Let Mono(ByVal bState As Boolean)

    If Not (m_bMono = bState) Then
        m_bInit = True
    End If
    m_bMono = bState
    
End Property

Public Function ColorizeImage(ByVal lColor As Long, _
                              Optional ByVal lStOveride As Single, _
                              Optional ByVal eGreyScale As eGrayScaleFormulas = gsclSimpleAvg)

Dim lRed    As Long
Dim lGreen  As Long
Dim lBlue   As Long
Dim lHue    As Single
Dim lSat    As Single
Dim lLum    As Single

    If CreateDIBDc(Width, Height) Then
        If IsGreyScale(lColor) Then
            ColorGreyScale eGreyScale
        Else
            LongToRgb lColor, lRed, lGreen, lBlue
            RGBToHLS lRed, lGreen, lBlue, lHue, lSat, lLum
            If Not lStOveride = 0 Then
                Colourise lHue, lStOveride
            Else
                Colourise lHue, lSat
            End If
        End If
        BitBlt m_lHdc, 0&, 0&, Width, Height, m_lDibDC, 0&, 0&, &HCC0020
    End If
    
End Function

Private Function IsGreyScale(ByVal lColor As Long) As Boolean

    If (lColor And &HFF) = ((lColor \ &H100&) And &HFF) Then
        If (lColor And &HFF) = ((lColor \ &H10000) And &HFF) Then
            IsGreyScale = True
        End If
    End If

End Function

Private Sub ColorGreyScale(ByVal eFormula As eGrayScaleFormulas)

Dim bDib()      As Byte
Dim x           As Long
Dim y           As Long
Dim xMax        As Long
Dim yMax        As Long
Dim lB          As Long
Dim lG          As Long
Dim lR          As Long
Dim r           As Single
Dim g           As Single
Dim b           As Single
Dim tSA         As SAFEARRAY2D
 
    Select Case eFormula
    Case gsclSimpleAvg
        r = 0.333: g = 0.334: b = r
    Case gsclNTSCPAL
        r = 0.299: g = 0.587: b = 0.114
    Case Else
        r = 0.213: g = 0.715: b = 0.072
    End Select
   
    '/* fHue runs from -1 to 5
    '/* have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1&
        .cDims = 2&
        .Bounds(0).lLbound = 0&
        .Bounds(0).cElements = Height
        .Bounds(1).lLbound = 0&
        .Bounds(1).cElements = BytesPerScanLine
        .pvData = DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4&
        
    yMax = Height - 1
    xMax = Width - 1

    For y = 0& To yMax
        For x = 0& To xMax * 3& Step 3&
            bDib(x, y) = ((0& + (bDib(x, y) * b) + (bDib(x + 1, y) * g) + (bDib(x + 2, y) * r)) * 255&) \ 255
            bDib(x + 1, y) = bDib(x, y)
            bDib(x + 2, y) = bDib(x, y)
        Next x
    Next y
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4

End Sub

Private Sub Colourise(ByVal fHue As Single, _
                      ByVal fSaturation As Single)

Dim bDib()      As Byte
Dim x           As Long
Dim y           As Long
Dim xMax        As Long
Dim yMax        As Long
Dim lB          As Long
Dim lG          As Long
Dim lR          As Long
Dim h           As Single
Dim s           As Single
Dim l           As Single
Dim tSA         As SAFEARRAY2D
    
    '/* fHue runs from -1 to 5
    '/* have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1&
        .cDims = 2&
        .Bounds(0).lLbound = 0&
        .Bounds(0).cElements = Height
        .Bounds(1).lLbound = 0&
        .Bounds(1).cElements = BytesPerScanLine
        .pvData = DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4&
        
    yMax = Height - 1
    xMax = Width - 1

    For x = 0 To (xMax * 3) Step 3
        For y = 0 To yMax
            RGBToHLS bDib(x + 2, y), bDib(x + 1, y), bDib(x, y), h, s, l
            s = fSaturation
            h = fHue
            HLSToRGB h, s, l, lR, lG, lB
            bDib(x, y) = lB
            bDib(x + 1, y) = lG
            bDib(x + 2, y) = lR
        Next y
    Next x
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4&
    
End Sub

Private Function ConvertTo32(ByVal lHsource As Long, _
                             ByRef tBMP As BITMAP, _
                             ByRef gAlpha As Byte) As Long

Dim aSBits()    As Byte
Dim i           As Long
Dim lHdc        As Long
Dim lhDIB       As Long
Dim hDIBold     As Long
Dim thDC        As Long
Dim tOldBmp     As Long
Dim uBIH        As BITMAPINFOHEADER
Dim uSSA        As SAFEARRAYID

    With uBIH
        .biBitCount = 32&
        .biHeight = tBMP.bmHeight
        .biWidth = tBMP.bmWidth
        .biPlanes = 1&
        .biSize = Len(uBIH)
    End With
    
    '/* create dib section
    With tBMP
        .bmWidthBytes = 4 * .bmWidth
    End With
    lHdc = CreateCompatibleDC(0&)
    If lHdc = 0 Then Exit Function
    lhDIB = CreateDIBSection(lHdc, uBIH, DIB_RGB_COLORS, tBMP.bmBits, 0&, 0&)
    If lhDIB = 0 Then Exit Function
    m_lpBits = tBMP.bmBits
    
    thDC = CreateCompatibleDC(0&)
    If thDC = 0 Then
        DeleteObject lhDIB
        lhDIB = 0
    Else
        hDIBold = SelectObject(lHdc, lhDIB)
        tOldBmp = SelectObject(thDC, lHsource)
        With tBMP
            BitBlt lHdc, 0&, 0&, .bmWidth, .bmHeight, thDC, 0&, 0&, vbSrcCopy
        End With
        '/* clean up
        SelectObject lHdc, hDIBold
        SelectObject thDC, tOldBmp
        DeleteDC thDC
        '/* build to 32bit
        If Not (tBMP.bmBitsPixel = 32) Then
            With tBMP
                MapDIBits uSSA, aSBits(), .bmBits, .bmWidthBytes * .bmHeight
            End With
            For i = 3 To UBound(aSBits) Step 4
                aSBits(i) = gAlpha
            Next
            Call UnmapDIBits(aSBits)
            gAlpha = 255
        End If
    End If

    DeleteDC lHdc
    ConvertTo32 = lhDIB
    
End Function

Public Function CreateBitmap(ByVal hBmp As Long, _
                             Optional ByVal hPal As Long = 0) As Picture

Dim tPic        As PICTUREINFO
Dim iPic        As IPicture
Dim tIDispatch  As Guid

    '/* fill GUID info
    With tIDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    '/* fill picture info
    With tPic
        .Size = Len(tPic) ' Length of structure
        .type = vbPicTypeBitmap ' Type of Picture (bitmap)
        .hBmp = hBmp ' Handle to bitmap
        .hPal = hPal ' Handle to palette (may be null)
    End With

    '/* create the picture
    OleCreatePictureIndirect tPic, tIDispatch, 1&, iPic
    '/* return the new picture
    Set CreateBitmap = iPic

End Function

Public Function CreateDIBDc(ByVal lWidth As Long, _
                            ByVal lHeight As Long) As Boolean
    
    m_lDibDC = CreateCompatibleDC(0&)
    If Not (m_lDibDC = 0) Then
        If (CreateDIB(m_lHdc, lWidth, lHeight, m_hDIb)) Then
            m_hDibOld = SelectObject(m_lDibDC, m_hDIb)
            BitBlt m_lDibDC, 0&, 0&, Width, Height, m_lHdc, 0&, 0&, &HCC0020
            CreateDIBDc = True
        Else
            DeleteObject m_lDibDC
            m_lDibDC = 0
        End If
    End If

End Function

Private Function CreateDIB(ByVal lHdc As Long, _
                           ByVal lWidth As Long, _
                           ByVal lHeight As Long, _
                           ByRef hDib As Long) As Boolean
    
    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1&
        .biBitCount = 24&
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight
    End With
    
    hDib = CreateDIBSection(lHdc, m_tBI, DIB_RGB_COLORS, m_lPtr, 0&, 0&)
    CreateDIB = Not (hDib = 0)

End Function

Public Sub CreateFromPicture(oPicture As StdPicture)

Dim lhDCC       As Long
Dim lHdc        As Long
Dim lhBmpOld    As Long
Dim tBMP        As BITMAP

    If oPicture Is Nothing Then Exit Sub
    GetObjectAPI oPicture.handle, Len(tBMP), tBMP
    
    With tBMP
        Width = .bmWidth
        Height = .bmHeight
        lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        lHdc = CreateCompatibleDC(lhDCC)
        '/* use bitmap with alpha channel
        If m_bUseAlpha Then
            m_lHandle = ConvertTo32(oPicture.handle, tBMP, 255)
            If Not m_lHandle = 0 Then
                lhBmpOld = SelectObject(lHdc, m_lHandle)
            '/* default on failure
            Else
                lhBmpOld = SelectObject(lHdc, oPicture.handle)
            End If
        Else
            lhBmpOld = SelectObject(lHdc, oPicture.handle)
        End If
        '/* blit the image into dc
        BitBlt hdc, 0&, 0&, .bmWidth, .bmHeight, lHdc, 0&, 0&, vbSrcCopy
    End With
    
    SelectObject lHdc, lhBmpOld
    DeleteDC lHdc
    DeleteDC lhDCC

End Sub

Public Function ExtractHandle() As Long

    If Not (m_hBmp = 0) Then
        ExtractHandle = m_hBmp
    End If

End Function

Public Function ExtractBitmap() As Long

    If Not (m_hBmpOld = 0) Then
        SelectObject m_lHdc, m_hBmpOld
        m_hBmpOld = 0
    End If
    If Not (m_lHdc = 0) Then
        DeleteDC m_lHdc
        m_lHdc = 0
    End If
    ExtractBitmap = m_hBmp
    m_hBmp = 0
    m_bInit = True

End Function

Private Sub ImageCreate(ByVal lW As Long, _
                        ByVal lH As Long)

Dim lHdc As Long

    ImageDestroy
    lHdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    m_lHdc = CreateCompatibleDC(lHdc)
    m_hBmp = CreateCompatibleBitmap(lHdc, lW, lH)
    m_hBmpOld = SelectObject(m_lHdc, m_hBmp)
    If (m_hBmpOld = 0) Then
        ImageDestroy
    Else
        m_lWidth = lW
        m_lHeight = lH
    End If
    DeleteDC lHdc

End Sub

Public Sub InjectBitmap(ByVal lhBmp As Long)

Dim tBMP As BITMAP

    ImageDestroy
    GetObjectAPI lhBmp, Len(tBMP), tBMP
    With tBMP
        Width = .bmWidth
        Height = .bmHeight
    End With
    If m_bUseAlpha Then
        m_lHandle = ConvertTo32(lhBmp, tBMP, 255&)
    End If
    Init lhBmp
    m_bInit = False

End Sub

Private Function Init(Optional ByVal hBmp As Long = 0) As Boolean

Dim hDCDisp As Long

    If m_bMono Then
        If (m_lHdc = 0) Then
            m_lHdc = CreateCompatibleDC(0&)
        End If
    Else
        hDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        If Not (hDCDisp = 0) Then
            If m_lHdc = 0 Then
                m_lHdc = CreateCompatibleDC(hDCDisp)
                If (m_lHdc = 0) Then
                    Exit Function
                End If
            End If
        Else
            Exit Function
        End If
    End If

    If Not (m_lHdc = 0) Then
        If Not (m_hBmpOld = 0) Then
            SelectObject m_lHdc, m_hBmpOld
            m_hBmpOld = 0
        End If
        If Not (m_hBmp = 0) Then
            DeleteObject m_hBmp
            m_hBmp = 0
        End If
        If (hBmp = 0) Then
            If m_bMono Then
                m_hBmp = CreateCompatibleBitmap(m_lHdc, m_lSizeX, m_lSizeY)
            Else
                m_hBmp = CreateCompatibleBitmap(hDCDisp, m_lSizeX, m_lSizeY)
            End If
        Else
            m_hBmp = hBmp
        End If
        If (m_hBmp = 0) Then
            DeleteDC hDCDisp
            hDCDisp = 0
        Else
            m_hBmpOld = SelectObject(m_lHdc, m_hBmp)
            Init = True
        End If
    Else
        DeleteDC hDCDisp
        hDCDisp = 0
    End If

    If Not (hDCDisp = 0) Then
        DeleteDC hDCDisp
    End If

End Function

Private Sub MapDIBits(uSA As SAFEARRAYID, _
                      aBits() As Byte, _
                      ByVal lpData As Long, _
                      ByVal lSize As Long)

On Error GoTo Handler

    With uSA
        .cbElements = 1
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = lSize
        .pvData = lpData
    End With
    CopyMemory ByVal VarPtrArray(aBits()), VarPtr(uSA), 4&

Handler:
    On Error GoTo 0

End Sub

Private Sub UnmapDIBits(aBits() As Byte)

On Error GoTo Handler

    CopyMemory ByVal VarPtrArray(aBits()), 0&, 4&

Handler:
    On Error GoTo 0

End Sub

Public Sub LongToRgb(ByVal lColor As Long, _
                     ByRef lRed As Long, _
                     ByRef lGreen As Long, _
                     ByRef lBlue As Long)

'/* Converts Long Color to RGB

    lRed = lColor And &HFF
    lGreen = (lColor And &HFF00&) \ &H100&
    lBlue = (lColor And &HFF0000) \ &H10000

End Sub

Public Sub RGBToHLS(ByVal r As Long, _
                    ByVal g As Long, _
                    ByVal b As Long, _
                    h As Single, _
                    s As Single, _
                    l As Single)

Dim Max     As Single
Dim Min     As Single
Dim delta   As Single
Dim rR      As Single
Dim rG      As Single
Dim rB      As Single

    rR = r / 255: rG = g / 255: rB = b / 255
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    l = (Max + Min) / 2
    If Max = Min Then
        s = 0
        h = 0
    Else
        If l <= 0.5 Then
            s = (Max - Min) / (Max + Min)
        Else
            s = (Max - Min) / (2 - Max - Min)
        End If
        delta = Max - Min
        If rR = Max Then
            h = (rG - rB) / delta
        ElseIf rG = Max Then
            h = 2 + (rB - rR) / delta
        ElseIf rB = Max Then
            h = 4 + (rR - rG) / delta
        End If
    End If

End Sub

Public Sub HLSToRGB(ByVal h As Single, _
                    ByVal s As Single, _
                    ByVal l As Single, _
                    r As Long, _
                    g As Long, _
                    b As Long)

Dim rR      As Single
Dim rG      As Single
Dim rB      As Single
Dim Min     As Single
Dim Max     As Single

    If s = 0 Then
        rR = l: rG = l: rB = l
    Else
        If l <= 0.5 Then
            Min = l * (1 - s)
        Else
            Min = l - s * (1 - l)
        End If
        Max = 2 * l - Min
        If (h < 1) Then
            rR = Max
            If (h < 0) Then
                rG = Min
                rB = rG - h * (Max - Min)
            Else
                rB = Min
                rG = h * (Max - Min) + rB
            End If
        ElseIf (h < 3) Then
            rG = Max
            If (h < 2) Then
                rB = Min
                rR = rB - (h - 2) * (Max - Min)
            Else
                rR = Min
                rB = (h - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (h < 4) Then
                rR = Min
                rG = rR - (h - 4) * (Max - Min)
            Else
                rG = Min
                rR = (h - 4) * (Max - Min) + rG
            End If
        End If
    End If
    r = rR * 255: g = rG * 255: b = rB * 255

End Sub

Private Function Maximum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single
   
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

Private Function Minimum(rR As Single, _
                         rG As Single, _
                         rB As Single) As Single

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

Private Sub ImageDestroy()

    If Not (m_hDIb = 0) Then
        SelectObject m_lDibDC, m_hDibOld
        DeleteObject m_hDIb
        DeleteDC m_lDibDC
    End If
    If Not (m_hBmpOld = 0) Then
        SelectObject m_lHdc, m_hBmpOld
        m_hBmpOld = 0
    End If
    If Not (m_hBmp = 0) Then
        DeleteObject m_hBmp
        m_hBmp = 0
    End If
    If Not (m_lHdc = 0) Then
        DeleteDC m_lHdc
        m_lHdc = 0
    End If
    If Not (m_lHandle = 0) Then
        DeleteObject m_lHandle
    End If
    If Not (m_tBIH.biBitCount = 0) Then
        ZeroMemory m_tBIH, Len(m_tBIH)
    End If
    m_lpBits = 0
    m_lWidth = 0
    m_lHeight = 0
    m_lHdc = 0
    m_hDIb = 0
    m_hDibOld = 0
    m_lPtr = 0
    
End Sub

Private Sub Class_Terminate()
    ImageDestroy
End Sub

