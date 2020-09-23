VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'/* based on the vbaccelerator imagelist control, with rewrites and additions:
'/* http://www.vbaccelerator.com/home/VB/Code/Controls/ImageList/vbAccelerator_Image_List_Control/article.asp


Private Const BI_RGB                            As Long = 0

Private Const BITSPIXEL                         As Long = 12

Private Const CLR_INVALID                       As Long = -1
Private Const CLR_NONE                          As Long = -1
Private Const COLOR_WINDOW                      As Long = &H5

Private Const DIB_RGB_COLORS                    As Long = 0

Private Const DSS_DISABLED                      As Long = &H20

Private Const DST_ICON                          As Long = &H3

Private Const FILE_ATTRIBUTE_NORMAL             As Long = &H80

Private Const ILCF_MOVE                         As Long = &H0
Private Const ILCF_SWAP                         As Long = &H1

Private Const ILD_NORMAL                        As Long = &H0
Private Const ILD_TRANSPARENT                   As Long = &H1
Private Const ILD_BLEND25                       As Long = &H2
Private Const ILD_SELECTED                      As Long = &H4
Private Const ILD_FOCUS                         As Long = &H8
Private Const ILD_MASK                          As Long = &H10
Private Const ILD_IMAGE                         As Long = &H20
Private Const ILD_ROP                           As Long = &H40

Private Const LR_LOADMAP3DCOLORS                As Long = &H1000
Private Const LR_LOADFROMFILE                   As Long = &H10
Private Const LR_COPYRETURNORG                  As Long = &H4

Private Const MAX_PATH                          As Long = 260

Private Const SHGFI_ICON                        As Long = &H100
Private Const SHGFI_SYSICONINDEX                As Long = &H4000
Private Const SHGFI_LARGEICON                   As Long = &H0
Private Const SHGFI_SMALLICON                   As Long = &H1
Private Const SHGFI_OPENICON                    As Long = &H2
Private Const SHGFI_SHELLICONSIZE               As Long = &H4
Private Const SHGFI_USEFILEATTRIBUTES           As Long = &H10

Private Const ICON_FLAGS                        As Long = _
    SHGFI_USEFILEATTRIBUTES Or SHGFI_SYSICONINDEX

Private Const VER_PLATFORM_WIN32_NT             As Long = 2


Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Private Enum PixelFormat
    PixelFormatIndexed = &H10000
    PixelFormatGDI = &H20000
    PixelFormatAlpha = &H40000
    PixelFormatPAlpha = &H80000
    PixelFormatExtended = &H100000
    PixelFormatCanonical = &H200000
    PixelFormatUndefined = 0
    PixelFormatDontCare = 0
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    PixelFormat8bppIndexed = &H30803
    PixelFormat16bppGreyScale = &H101004
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bppRGB = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HE200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
    PixelFormatMax = 15
End Enum

Private Enum PaletteFlags
    PaletteFlagsHasAlpha = &H1
    PaletteFlagsGrayScale = &H2
    PaletteFlagsHalftone = &H4
End Enum

Public Enum EISIconSize
    eisLargeIcon = SHGFI_LARGEICON
    eisOpenIcon = SHGFI_OPENICON
    eisShellIcon = SHGFI_SHELLICONSIZE
    eisSmallIcon = SHGFI_SMALLICON
End Enum

Public Enum EILDrawState
    ildNormal = 0
    ildDisabled = 1
    ildSelected = 2
    ildColored = 3
    ildCutDisabled = 4
End Enum

Public Enum EILImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
End Enum

Public Enum EILColourDepth
    ILC_COLOR = &H0
    ILC_MASK = &H1
    ILC_COLOR4 = &H4
    ILC_COLOR8 = &H8
    ILC_COLOR16 = &H10
    ILC_COLOR24 = &H18
    ILC_COLOR32 = &H20
End Enum

Public Enum EILSwapTypes
    ilsCopy = ILCF_MOVE
    ilsSwap = ILCF_SWAP
End Enum


Private Type GdiplusStartupInput
   GdiplusVersion                               As Long
   DebugEventCallback                           As Long
   SuppressBackgroundThread                     As Long
   SuppressExternalCodecs                       As Long
End Type

Private Type BITMAP
    bmType                                      As Long
    bmWidth                                     As Long
    bmHeight                                    As Long
    bmWidthBytes                                As Long
    bmPlanes                                    As Long
    bmBitsPixel                                 As Long
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

Private Type RGBQUAD
    rgbBlue                                     As Byte
    rgbGreen                                    As Byte
    rgbRed                                      As Byte
    rgbReserved                                 As Byte
End Type

Private Type BITMAPDATA
    Width                                       As Long
    Height                                      As Long
    stride                                      As Long
    PixelFormat                                 As Long
    scan0                                       As Long
    Reserved                                    As Long
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type IMAGEINFO
    hBitmapImage                                As Long
    hBitmapMask                                 As Long
    cPlanes                                     As Long
    cBitsPerPixel                               As Long
    rcImage                                     As RECT
End Type

Private Type PictDesc
    cbSizeofStruct                              As Long
    picType                                     As Long
    hImage                                      As Long
    xExt                                        As Long
    yExt                                        As Long
End Type

Private Type Guid
    Data1                                       As Long
    Data2                                       As Integer
    Data3                                       As Integer
    Data4(0 To 7)                               As Byte
End Type

Private Type ICONINFO
    fIcon                                       As Long
    xHotspot                                    As Long
    yHotspot                                    As Long
    hBmMask                                     As Long
    hbmColor                                    As Long
End Type

Private Type BITMAPINFO_1BPP
    bmiHeader                                   As BITMAPINFOHEADER
    bmiColors(0 To 1)                           As RGBQUAD
End Type

Private Type BITMAPINFO_ABOVE8
    bmiHeader                                   As BITMAPINFOHEADER
End Type

Private Type tagINITCOMMONCONTROLSEX
    dwSize                                      As Long
    dwICC                                       As Long
End Type

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type

Private Type SHFILEINFOA
    hicon                                       As Long
    iIcon                                       As Long
    dwAttributes                                As Long
    szDisplayName                               As String * MAX_PATH
    szTypeName                                  As String * 80
End Type

Private Type SHFILEINFOW
    hicon                                       As Long
    iIcon                                       As Long
    dwAttributes                                As Long
    szDisplayName(MAX_PATH)                     As Byte
    szTypeName(80)                              As Byte
End Type



Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function SHGetFileInfoA Lib "shell32.dll" (ByVal pszPath As String, _
                                                           ByVal dwAttributes As Long, _
                                                           psfi As SHFILEINFOA, _
                                                           ByVal cbSizeFileInfo As Long, _
                                                           ByVal uFlags As Long) As Long

Private Declare Function SHGetFileInfoW Lib "shell32.dll" (ByVal pszPath As Long, _
                                                           ByVal dwAttributes As Long, _
                                                           psfi As SHFILEINFOW, _
                                                           ByVal cbSizeFileInfo As Long, _
                                                           ByVal uFlags As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetIconInfo Lib "user32" (ByVal hicon As Long, _
                                                   piconinfo As ICONINFO) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long

Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long

Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, _
                                                ByVal hBitmap As Long, _
                                                ByVal nStartScan As Long, _
                                                ByVal nNumScans As Long, _
                                                lpBits As Any, _
                                                lpBI As Any, _
                                                ByVal wUsage As Long) As Long

Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal hBitmap As Long, _
                                                ByVal nStartScan As Long, _
                                                ByVal nNumScans As Long, _
                                                lpBits As Any, _
                                                lpBI As Any, _
                                                ByVal wUsage As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, _
                                                  lpvSource As Any, _
                                                  ByVal cbCopy As Long)

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hicon As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, _
                                                                    ByVal lpsz As String, _
                                                                    ByVal un1 As Long, _
                                                                    ByVal n1 As Long, _
                                                                    ByVal n2 As Long, _
                                                                    ByVal un2 As Long) As Long

Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, _
                                                                        ByVal lpsz As Long, _
                                                                        ByVal un1 As Long, _
                                                                        ByVal n1 As Long, _
                                                                        ByVal n2 As Long, _
                                                                        ByVal un2 As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                    ByVal hBrush As Long, _
                                                                    ByVal lpDrawStateProc As Long, _
                                                                    ByVal lParam As Long, _
                                                                    ByVal wParam As Long, _
                                                                    ByVal x As Long, _
                                                                    ByVal y As Long, _
                                                                    ByVal cx As Long, _
                                                                    ByVal cy As Long, _
                                                                    ByVal fuFlags As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, _
                                                               ByVal HPALETTE As Long, _
                                                               pccolorref As Long) As Long

Private Declare Function ImageList_GetBkColor Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal hImageList As Long, _
                                                               ByVal I As Long, _
                                                               ByVal hicon As Long) As Long

Private Declare Function ImageList_Convert Lib "comctl32" Alias "ImageList_Draw" (ByVal hImageList As Long, _
                                                                                  ByVal ImgIndex As Long, _
                                                                                  ByVal hDCDest As Long, _
                                                                                  ByVal x As Long, _
                                                                                  ByVal y As Long, _
                                                                                  ByVal flags As Long) As Long

Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_Replace Lib "comctl32" (ByVal hImageList As Long, _
                                                           ByVal ImgIndex As Long, _
                                                           ByVal hbmImage As Long, _
                                                           ByVal hBmMask As Long) As Long

Private Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, _
                                                       ByVal hbmImage As Long, _
                                                       hBmMask As Long) As Long

Private Declare Function ImageList_Remove Lib "comctl32" (ByVal hImageList As Long, _
                                                          ByVal ImgIndex As Long) As Long

Private Declare Function ImageList_GetImageInfo Lib "comctl32" (ByVal hIml As Long, _
                                                                ByVal I As Long, _
                                                                pImageInfo As IMAGEINFO) As Long

Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hIml As Long, _
                                                           ByVal hicon As Long) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, _
                                                           ByVal ImgIndex As Long, _
                                                           ByVal fuFlags As Long) As Long

Private Declare Function ImageList_SetImageCount Lib "comctl32" (ByVal hImageList As Long, _
                                                                 uNewCount As Long)

Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function ImageList_SetIconSize Lib "comctl32" (ByVal hImageList As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, _
                                                        ByVal I As Long, _
                                                        ByVal hdcDst As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long, _
                                                        ByVal fStyle As Long) As Long

Private Declare Function ImageList_DrawEx Lib "comctl32" (ByVal hIml As Long, _
                                                          ByVal I As Long, _
                                                          ByVal hdcDst As Long, _
                                                          ByVal x As Long, _
                                                          ByVal y As Long, _
                                                          ByVal dx As Long, _
                                                          ByVal dy As Long, _
                                                          ByVal rgbBk As Long, _
                                                          ByVal rgbFg As Long, _
                                                          ByVal fStyle As Long) As Long

Private Declare Function ImageList_Copy Lib "comctl32" (ByVal himlDst As Long, _
                                                        ByVal iDst As Long, _
                                                        ByVal himlSrc As Long, _
                                                        ByVal iSrc As Long, _
                                                        ByVal uFlags As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, _
                                                                      riid As Guid, _
                                                                      ByVal fPictureOwnsHandle As Long, _
                                                                      ipic As IPicture) As Long

Private Declare Function InitCommonControls Lib "comctl32" () As Boolean

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As OSVERSIONINFO) As Long

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, _
                                                       inputbuf As GdiplusStartupInput, _
                                                       Optional ByVal outputbuf As Long = 0) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, _
                                                          graphics As Long) As Long

Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, _
                                                          Width As Long) As Long

Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, _
                                                           Height As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long

Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, _
                                                                    ByVal hpal As Long, _
                                                                    bmap As Long) As Long

Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal graphics As Long, _
                                                      ByVal image As Long, _
                                                      ByVal x As Single, _
                                                      ByVal y As Single) As Long

Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal lImage As Long, _
                                                                lFormat As Long) As Long

Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal BITMAP As Long, _
                                                           rc As RECT, _
                                                           ByVal flags As ImageLockMode, _
                                                           ByVal PixelFormat As Long, _
                                                           lockedBitmapData As BITMAPDATA) As Long

Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, _
                                                                  ByVal Height As Long, _
                                                                  ByVal stride As Long, _
                                                                  ByVal PixelFormat As Long, _
                                                                  scan0 As Any, BITMAP As Long) As Long

Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal BITMAP As Long, _
                                                             lockedBitmapData As BITMAPDATA) As Long

Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, _
                                                           ByVal x As Long, _
                                                           ByVal y As Long, _
                                                           color As Long) As Long


Private m_bIsXp                                 As Boolean
Private m_bIsNt                                 As Boolean
Private m_bGdiPlusLoaded                        As Boolean
Private m_bUseGdiPlus                           As Boolean
Private m_hIml                                  As Long
Private m_lIconSizeX                            As Long
Private m_lIconSizeY                            As Long
Private m_lOwnerDc                              As Long
Private m_lGdiPlusToken                         As Long
Private m_eColourDepth                          As EILColourDepth
Private m_sImlName                              As String
Private m_cKeys                                 As Collection


Private Sub Class_Initialize()

    m_lIconSizeX = 16
    m_lIconSizeY = 16
    m_eColourDepth = ILC_COLOR
    Set m_cKeys = New Collection
    InitCommonControls
    VersionCheck
    
End Sub

Private Sub VersionCheck()
'/* operating system check

Dim tVer As OSVERSIONINFO

    With tVer
        .dwVersionInfoSize = LenB(tVer)
        GetVersionEx tVer
        m_bIsNt = ((.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
        If (.dwMajorVersion >= 5) Then
            m_bIsXp = True
        End If
    End With

End Sub

Private Function LibraryExists(ByVal sModule As String) As Boolean
'/* test for library support

Dim lhMod   As Long
Dim bLoad   As Boolean

    lhMod = GetModuleHandleA(sModule)
    If (lhMod = 0) Then
        lhMod = LoadLibraryA(sModule)
        bLoad = True
    End If
    If Not (lhMod = 0) Then
        LibraryExists = True
        If bLoad Then
            FreeLibrary lhMod
        End If
    End If

End Function

Public Property Get GdiPlusLoaded() As Boolean
    GdiPlusLoaded = m_bGdiPlusLoaded
End Property

Public Property Get UseGdiPlus() As Boolean
    UseGdiPlus = m_bUseGdiPlus
End Property

Public Property Let UseGdiPlus(ByVal PropVal As Boolean)
    If PropVal Then
        If LibraryExists("gdiplus") Then
            LoadGdiPlus
        End If
    Else
        GdiUnload
    End If
    m_bUseGdiPlus = PropVal
End Property

Private Function LoadGdiPlus() As Boolean
'/* load gdi library

Dim tGPInput As GdiplusStartupInput

    GdiUnload
    tGPInput.GdiplusVersion = 1
    If (GdiplusStartup(m_lGdiPlusToken, tGPInput) = 0) Then
        LoadGdiPlus = True
        m_bGdiPlusLoaded = True
    End If
  
End Function

Private Sub GdiUnload()
'/* unload gdi library

    If Not (m_lGdiPlusToken = 0) Then
        GdiplusShutdown m_lGdiPlusToken
        m_lGdiPlusToken = 0
        m_bGdiPlusLoaded = False
    End If
    
End Sub

Private Sub DrawAlphaIcon(ByVal lIndex As Long, _
                          ByVal lDestDc As Long, _
                          ByVal lX As Long, _
                          ByVal lY As Long)

'/* could use variation of this to extract and
'/* draw 32b alpha bitmaps/png also..

Dim lGraphics   As Long
Dim lpBitmap    As Long
Dim lhIcon      As Long
Dim lHeight     As Long
Dim lWidth      As Long
Dim lpBmpNew    As Long
Dim lFormat     As Long
Dim tBmpData    As BITMAPDATA
Dim tIcnInfo    As ICONINFO
Dim tRect       As RECT

    '/* fetch icon handle
    lhIcon = ImageList_GetIcon(m_hIml, lIndex, 1&)
    If Not lhIcon = 0 Then
        '/* icon data structure
        GetIconInfo lhIcon, tIcnInfo
        '/* load graphics
        GdipCreateFromHDC lDestDc, lGraphics
        '/* copy icon bmp to a new image
        GdipCreateBitmapFromHBITMAP tIcnInfo.hbmColor, 0&, lpBitmap
        '/* dispose of resource
        DeleteObject (tIcnInfo.hbmColor)
        DeleteObject (tIcnInfo.hBmMask)
        '/* get the format
        GdipGetImagePixelFormat lpBitmap, lFormat
    End If
    '/* not alpha
    If Not (lFormat < PixelFormat32bppRGB) Then
        '/* image dimensions
        GdipGetImageHeight lpBitmap, lHeight
        GdipGetImageWidth lpBitmap, lWidth
        With tRect
            .Bottom = lHeight
            .left = 0
            .Right = lWidth
            .top = 0
        End With
        '/* create a new 32b bmp
        GdipCreateBitmapFromScan0 lWidth, lHeight, 0&, PixelFormat32bppARGB, ByVal 0&, lpBmpNew
        '/* copy 32b data structure
        GdipBitmapLockBits lpBitmap, tRect, ImageLockModeRead, lFormat, tBmpData
        GdipBitmapLockBits lpBmpNew, tRect, ImageLockModeWrite Or ImageLockModeUserInputBuf, PixelFormat32bppARGB, tBmpData
        '/* unlock
        GdipBitmapUnlockBits lpBmpNew, tBmpData
        GdipBitmapUnlockBits lpBitmap, tBmpData
        '/* test for alpha channel
        If IsAlphaBitmap(lpBmpNew, lWidth, lHeight) Then
            '/* draw the alpha image
            GdipDrawImage lGraphics, lpBmpNew, lX, lY
        Else
            '/* draw flat image
            ImageList_DrawEx m_hIml, lIndex, lDestDc, lX, lY, 0&, 0&, CLR_NONE, 0&, ILD_TRANSPARENT
        End If
    Else
        '/* draw flat image
        ImageList_DrawEx m_hIml, lIndex, lDestDc, lX, lY, 0&, 0&, CLR_NONE, 0&, ILD_TRANSPARENT
    End If

    '/* cleanup
    DestroyIcon lhIcon
    GdipDisposeImage lpBitmap
    GdipDisposeImage lpBmpNew
    GdipDeleteGraphics lGraphics
   
End Sub

Private Function IsAlphaBitmap(ByVal lBitmap As Long, _
                               ByVal lWidth As Long, _
                               ByVal lHeight As Long) As Boolean

'/* if there is an alpha channel
'/* colors are right shifted

Dim lY      As Long
Dim lX      As Long
Dim lArgb   As Long

    For lY = 0 To lHeight - 1
        For lX = 0 To lWidth - 1
            GdipBitmapGetPixel lBitmap, lX, lY, lArgb
            If (lArgb > &HFF000000) Then
                If (lArgb < &HFFFFFFFF) Then
                    IsAlphaBitmap = True
                    Exit For
                End If
            End If
        Next lX
        If IsAlphaBitmap Then
            Exit For
        End If
    Next lY

End Function

'**********************************************************************
'*                              PROPERTIES
'**********************************************************************

Public Property Get ColourDepth() As EILColourDepth
'/* [get] color depth
    ColourDepth = m_eColourDepth
End Property

Public Property Let ColourDepth(ByVal PropVal As EILColourDepth)
'/* [let] color depth
    If (PropVal > SystemColourDepth) Then
        PropVal = SystemColourDepth
    End If
    If Not (PropVal = m_eColourDepth) Then
        m_eColourDepth = PropVal
        Create
    End If
    m_eColourDepth = PropVal
End Property

Private Property Get SystemColourDepth() As EILColourDepth
'/* [get] system color depth
Dim lHdc As Long

    lHdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    SystemColourDepth = GetDeviceCaps(lHdc, BITSPIXEL)
    DeleteDC lHdc

End Property

Public Property Get ImageCount() As Long
'/* [get] image count
    If Not (m_hIml = 0) Then
        ImageCount = ImageList_GetImageCount(m_hIml)
    End If
End Property

Public Property Let ImageCount(ByVal PropVal As Long)
'/* [let] dummy

End Property

Public Property Get ItemCopyOfIcon(ByVal PropVal As Variant) As Long
'/* [get] copy of icon
Dim lIndex As Long

    lIndex = ItemIndex(PropVal)
    If Not (lIndex = -1) Then
        ItemCopyOfIcon = ImageList_GetIcon(m_hIml, lIndex, ILD_TRANSPARENT)
    End If

End Property

Public Property Get ItemIndex(ByVal PropVal As Variant) As Long
'/* [get] icon index

Dim lResult As Long

    ItemIndex = -1
    If IsNumeric(PropVal) Then
        lResult = CLng(PropVal)
        If (PropVal > 0) Then
            If (PropVal <= ImageCount) Then
                ItemIndex = PropVal - 1
            End If
        End If
    Else
        ItemIndex = IndexFromKey(CStr(PropVal))
    End If

End Property

Public Property Get ItemKey(ByVal lIndex As Long) As String
'/* [get] item key
    ItemKey = ""
    If (lIndex > 0) Then
        If (lIndex <= ImageCount) Then
            ItemKey = KeyFromIndex(lIndex)
        End If
    End If
End Property

Public Property Let ItemKey(ByVal lIndex As Long, _
                            ByVal PropVal As String)
'/* [let] icon key
    If (lIndex > 0) Then
        If (lIndex <= ImageCount) Then
            AddKey lIndex, PropVal
        End If
    End If

End Property

Public Property Get ItemPicture(ByVal PropVal As Variant) As IPicture
'/* [get] item picture
Dim lIndex  As Long
Dim lhIcon  As Long

    lIndex = ItemIndex(PropVal)
    If Not (lIndex = -1) Then
        lhIcon = ImageList_GetIcon(m_hIml, lIndex, ILD_TRANSPARENT)
        If Not (lhIcon = 0) Then
            Set ItemPicture = IconToPicture(lhIcon)
        End If
    End If

End Property

Public Property Get KeyExists(ByVal PropVal As String) As Boolean
'/* [get] test item key
    If (ImageCount > 0) Then
        If Not (IndexFromKey(PropVal) = -1) Then
            KeyExists = True
        End If
    End If
End Property

Public Property Get hIml() As Long
'/* [get] iml handle
    hIml = m_hIml
End Property

Public Property Get IconSizeX() As Long
'/* [get] icon width
    IconSizeX = m_lIconSizeX
End Property

Public Property Let IconSizeX(ByVal PropVal As Long)
'/* [let] icon width
    If Not (PropVal = m_lIconSizeX) Then
        m_lIconSizeX = PropVal
        Create
    End If
    m_lIconSizeX = PropVal
End Property

Public Property Get IconSizeY() As Long
'/* [get] icon height
    IconSizeY = m_lIconSizeY
End Property

Public Property Let IconSizeY(ByVal PropVal As Long)
'/* [let] icon height
    If Not (PropVal = m_lIconSizeY) Then
        m_lIconSizeY = PropVal
        Create
    End If
    m_lIconSizeY = PropVal
End Property

Public Property Get ImlName() As String
'/* [get] iml name
    ImlName = m_sImlName
End Property

Public Property Let ImlName(ByVal PropVal As String)
'/* [let] iml name
    m_sImlName = PropVal
End Property

Public Property Let OwnerHDC(ByVal PropVal As Long)
'/* [let] owner handle
    m_lOwnerDc = PropVal
End Property


'**********************************************************************
'*                              SUPPORT
'**********************************************************************

Public Function AddFromFile(ByVal sFileName As String, _
                            ByVal eType As EILImageTypes, _
                            Optional ByVal sKey As String, _
                            Optional ByVal bMapSysColors As Boolean = False, _
                            Optional ByVal lBackColor As OLE_COLOR = -1) As Long

Dim lhImage As Long
Dim lFlags  As Long

    AddFromFile = -1
    If Not (m_hIml = 0) Then
        lFlags = LR_LOADFROMFILE
        If bMapSysColors Then
            lFlags = lFlags Or LR_LOADMAP3DCOLORS
        End If
        lhImage = LoadImage(App.hInstance, sFileName, eType, 0&, 0&, lFlags)
        AddFromFile = AddFromHandle(lhImage, eType, sKey, lBackColor)
        Select Case eType
        Case IMAGE_ICON
            DestroyIcon lhImage
        Case IMAGE_CURSOR
            DestroyCursor lhImage
        Case IMAGE_BITMAP
            DeleteObject lhImage
        End Select
    End If

End Function

Public Function AddFromHandle(ByVal lhImage As Long, _
                              ByVal eType As EILImageTypes, _
                              Optional ByVal sKey As String, _
                              Optional ByVal lBackColor As Long = -1) As Boolean
    
    If Not (m_hIml = 0) Then
        If Not (lhImage = 0) Then
            If (eType = IMAGE_BITMAP) Then
                If (lBackColor = -1) Then
                    lBackColor = GetImageBackColor(lhImage)
                End If
                ImageList_AddMasked m_hIml, lhImage, lBackColor
            ElseIf (eType = IMAGE_ICON) Or (eType = IMAGE_CURSOR) Then
                ImageList_AddIcon m_hIml, lhImage
            End If
            If Not (Len(sKey) = 0) Then
                AddKey ImageCount, sKey
            End If
        End If
    End If

End Function

Public Function AddFromResourceID(ByVal lID As Long, _
                                  ByVal lhInst As Long, _
                                  ByVal eType As EILImageTypes, _
                                  Optional ByVal sKey As String, _
                                  Optional ByVal bMapSysColors As Boolean = False, _
                                  Optional ByVal lBackColor As Long = -1) As Long

Dim lhImage As Long
Dim lFlags    As Long
Dim lX     As Long
Dim lY     As Long

    AddFromResourceID = -1
    If Not (m_hIml = 0) Then
        If bMapSysColors Then
            lFlags = LR_LOADMAP3DCOLORS
        End If
        If Not (eType = IMAGE_BITMAP) Then
            lX = m_lIconSizeX
            lY = m_lIconSizeY
        End If
        If (lhInst = 0) Then
            lFlags = lFlags Or LR_COPYRETURNORG
        End If
        lhImage = LoadImageLong(lhInst, lID, eType, lX, lY, lFlags)
        AddFromResourceID = AddFromHandle(lhImage, eType, sKey, lBackColor)
        Select Case eType
        Case IMAGE_ICON
            DestroyIcon lhImage
        Case IMAGE_CURSOR
            DestroyCursor lhImage
        Case IMAGE_BITMAP
            DeleteObject lhImage
        End Select
    End If

End Function

Public Function BitmapToPicture(ByVal lhBmp As Long) As IPicture

Dim NewPic   As Picture
Dim tPicConv As PictDesc
Dim IGuid    As Guid

    If Not (lhBmp = 0) Then
        With tPicConv
            .cbSizeofStruct = Len(tPicConv)
            .picType = vbPicTypeBitmap
            .hImage = lhBmp
        End With
        With IGuid
            .Data1 = &H20400
            .Data4(0) = &HC0
            .Data4(7) = &H46
        End With
        OleCreatePictureIndirect tPicConv, IGuid, True, NewPic
        Set BitmapToPicture = NewPic
    End If

End Function

Public Sub Clear()
    Create
    Set m_cKeys = Nothing
    Set m_cKeys = New Collection
End Sub

Public Function Create() As Boolean

    Destroy
    m_hIml = ImageList_Create(m_lIconSizeX, m_lIconSizeY, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    If Not (m_hIml = 0) Then
        If Not (m_hIml = -1) Then
            Create = True
        End If
    End If

End Function

Public Sub Destroy()

    If Not (m_hIml = 0) Then
        ImageList_Destroy m_hIml
        m_hIml = 0
    End If

End Sub

Public Sub DrawImage(ByVal lHdc As Long, _
                     ByVal lIndex As Long, _
                     ByVal lX As Long, _
                     ByVal lY As Long, _
                     Optional ByVal eState As EILDrawState = ildNormal, _
                     Optional ByVal lDither As Long = vbWindowBackground)

Dim lhIcon     As Long
Dim lFlags    As Long
Dim lColor    As Long

    If (lIndex > -1) Then
        If Not (m_hIml = 0) Then
            lFlags = ILD_TRANSPARENT
            Select Case eState
            Case ildColored
                lFlags = lFlags Or ILD_BLEND25
                lColor = TranslateColor(lDither)
                ImageList_DrawEx m_hIml, lIndex, lHdc, lX, lY, 0&, 0&, CLR_NONE, lColor, lFlags
            Case ildCutDisabled
                lColor = GetSysColor(COLOR_WINDOW)
                lFlags = lFlags Or ILD_SELECTED
                ImageList_DrawEx m_hIml, lIndex, lHdc, lX, lY, 0&, 0&, CLR_NONE, lColor, lFlags
            Case ildDisabled
                lhIcon = ImageList_GetIcon(m_hIml, lIndex, 0&)
                DrawState lHdc, 0&, 0&, lhIcon, 0&, lX, lY, m_lIconSizeX, m_lIconSizeY, DST_ICON Or DSS_DISABLED
                DestroyIcon lhIcon
            Case ildSelected
                lFlags = lFlags Or ILD_SELECTED
                ImageList_Draw m_hIml, lIndex, lHdc, lX, lY, lFlags '<- looks better
               ' ImageList_DrawEx m_hIml, lIndex, lHdc, lX, lY, 0&, 0&, CLR_NONE, CLR_NONE, lFlags
            Case Else
                If (m_eColourDepth = ILC_COLOR32) Then
                    If m_bGdiPlusLoaded Then
                        DrawAlphaIcon lIndex, lHdc, lX, lY
                    Else
                        ImageList_DrawEx m_hIml, lIndex, lHdc, lX, lY, 0&, 0&, CLR_NONE, 0&, ILD_TRANSPARENT
                    End If
                Else
                    ImageList_DrawEx m_hIml, lIndex, lHdc, lX, lY, 0&, 0&, CLR_NONE, 0&, ILD_TRANSPARENT
                End If
            End Select
        End If
    End If

End Sub

Public Function IconToPicture(ByVal lhIcon As Long) As IPicture

Dim NewPic  As Picture
Dim PicConv As PictDesc
Dim IGuid   As Guid

    If Not (lhIcon = 0) Then
        With PicConv
            .cbSizeofStruct = Len(PicConv)
            .picType = vbPicTypeIcon
            .hImage = lhIcon
        End With
        With IGuid
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(2) = &H0
            .Data4(3) = &HAA
            .Data4(4) = &H0
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With
        OleCreatePictureIndirect PicConv, IGuid, True, NewPic
        Set IconToPicture = NewPic
    End If

End Function

Public Sub RemoveImage(ByVal vKey As Variant)

Dim lIndex  As Long

    If Not (m_hIml = 0) Then
        If IsNumeric(vKey) Then
            lIndex = vKey
            If (lIndex > -1) Then
                ImageList_Remove m_hIml, lIndex
                RemoveKey lIndex
            End If
        Else
            lIndex = IndexFromKey(CStr(vKey))
            If (lIndex > -1) Then
                ImageList_Remove m_hIml, lIndex
                RemoveKey lIndex
            End If
        End If
    End If

End Sub

Private Sub AddKey(ByVal lIndex As Long, _
                   ByVal sKey As String)

On Error Resume Next

    With m_cKeys
        If (.Item(lIndex)) Then
            .Remove (lIndex)
        End If
        .Add lIndex, sKey
    End With
    
On Error GoTo 0

End Sub

Public Function IndexFromKey(ByVal sKey As String) As Long

On Error Resume Next

    IndexFromKey = m_cKeys.Item(sKey)
    If Not (Err.Number = 0) Then
        IndexFromKey = -1
    End If
    
On Error GoTo 0

End Function

Public Function KeyFromIndex(ByVal lIndex As Long) As String

On Error Resume Next

    KeyFromIndex = FetchItemKey(lIndex, m_cKeys)
    
On Error GoTo 0

End Function

Private Sub RemoveKey(ByVal lIndex As Long)

On Error Resume Next

    m_cKeys.Remove lIndex
    
On Error GoTo 0

End Sub

Public Sub SwapOrCopyImage(ByVal lSrcIcon As Long, _
                           ByVal lDstIcon As Long, _
                           Optional ByVal eSwap As EILSwapTypes = ilsSwap)

Dim lDst    As Long
Dim lSrc    As Long
Dim sKeyDst As String
Dim sKeySrc As String

    If Not (m_hIml = 0) Then
        lDst = ItemIndex(lSrcIcon)
        If lDst > -1 Then
            lSrc = ItemIndex(lDstIcon)
            If lSrc > -1 Then
                ImageList_Copy m_hIml, lDst, m_hIml, lSrc, eSwap
                sKeySrc = KeyFromIndex(lSrcIcon)
                sKeyDst = KeyFromIndex(lDstIcon)
                If Not (Len(sKeySrc) = 0) Then
                    AddKey lDstIcon, sKeySrc
                End If
                If Not (Len(sKeyDst) = 0) Then
                    AddKey lSrcIcon, sKeyDst
                End If
            End If
        End If
    End If

End Sub

Private Function FetchItemKey(ByVal lIndex As Long, _
                              ByRef cCol As Collection) As String

Dim lCt     As Long
Dim lPtr    As Long
Dim sKey    As String

    If Not (cCol Is Nothing) Then
        If Not (lIndex < 1) Then
            If Not (lIndex > cCol.Count) Then
                Select Case lIndex
                Case Is <= cCol.Count / 2
                    RtlMoveMemory lPtr, ByVal ObjPtr(cCol) + 24, 4&
                    For lCt = 2 To lIndex
                        RtlMoveMemory lPtr, ByVal lPtr + 24, 4&
                    Next lCt
                Case Else
                    RtlMoveMemory lPtr, ByVal ObjPtr(cCol) + 28, 4&
                    For lCt = cCol.Count - 1 To lIndex Step -1
                        RtlMoveMemory lPtr, ByVal lPtr + 20, 4&
                    Next lCt
                End Select
                lCt = StrPtr(sKey)
                RtlMoveMemory ByVal VarPtr(sKey), ByVal lPtr + 16, 4&
                FetchItemKey = sKey
                RtlMoveMemory ByVal VarPtr(sKey), lCt, 4&
            End If
        End If
    End If

End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                                Optional hpal As Long = 0) As Long

    If OleTranslateColor(clr, hpal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function

Private Function GetImageBackColor(ByVal lhImage As Long) As Long

Dim lHdc    As Long
Dim lTmpDc  As Long
Dim lhBmp   As Long

    lTmpDc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    If Not (lTmpDc = 0) Then
        lHdc = CreateCompatibleDC(lTmpDc)
        DeleteDC lTmpDc
        If Not (lHdc = 0) Then
            lhBmp = SelectObject(lHdc, lhImage)
            If Not (lhBmp = 0) Then
                GetImageBackColor = GetPixel(lHdc, 0&, 0&)
                SelectObject lHdc, lhBmp
            End If
            DeleteObject lHdc
        End If
    End If
                        
End Function


'**********************************************************************
'*                              SYSTEM IML
'**********************************************************************

Public Property Get SystemImlHandle(ByVal sFile As String, _
                                    ByVal eIconSize As EISIconSize) As Long

Dim lHandle As Long
Dim tFI     As SHFILEINFOA

    SystemImlHandle = -1
    If Not (Len(sFile) = 0) Then
        lHandle = SHGetFileInfoA(sFile, FILE_ATTRIBUTE_NORMAL, tFI, LenB(tFI), ICON_FLAGS Or eIconSize)
        If Not (lHandle = 0) Then
            SystemImlHandle = lHandle
        End If
    End If

End Property

Public Function SystemIconIndex(ByVal sFile As String, _
                                ByVal eIconSize As EISIconSize) As Long

Dim lFlags  As Long
Dim lResult As Long

    SystemIconIndex = -1
    lFlags = SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES Or eIconSize
    
    If m_bIsNt Then
        If Not (LenB(sFile) = 0) Then
            Dim tFW As SHFILEINFOW
            lResult = SHGetFileInfoW(StrPtr(sFile), FILE_ATTRIBUTE_NORMAL, tFW, Len(tFW), lFlags)
            If Not (lResult = 0) Then
                SystemIconIndex = tFW.iIcon
            End If
        End If
    Else
        If Not (Len(sFile) = 0) Then
            Dim tFA As SHFILEINFOA
            lResult = SHGetFileInfoA(sFile, FILE_ATTRIBUTE_NORMAL, tFA, Len(tFA), lFlags)
            If Not (lResult = 0) Then
                SystemIconIndex = tFA.iIcon
            End If
        End If
    End If

End Function

Public Function SystemIconHandle(ByVal sFile As String, _
                                 ByVal eIconSize As EISIconSize) As Long

Dim lFlags  As Long
Dim lResult As Long

    SystemIconHandle = -1
    lFlags = SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES Or eIconSize
    
    If m_bIsNt Then
        If Not (Len(sFile) = 0) Then
            Dim tFW As SHFILEINFOW
            lResult = SHGetFileInfoW(StrPtr(sFile), FILE_ATTRIBUTE_NORMAL, tFW, Len(tFW), lFlags)
            If Not (lResult = 0) Then
                SystemIconHandle = tFW.hicon
            End If
        End If
    Else
        If Not (Len(sFile) = 0) Then
            Dim tFA As SHFILEINFOA
            lResult = SHGetFileInfoA(sFile, FILE_ATTRIBUTE_NORMAL, tFA, Len(tFA), lFlags)
            If Not (lResult = 0) Then
                SystemIconHandle = tFA.hicon
            End If
        End If
    End If
   
End Function


'**********************************************************************
'*                              SAVE/RESTORE
'**********************************************************************

Private Sub SaveKeys(ByRef bKeys() As Byte)

Dim lItem As Long
Dim sKeys As String

    If (m_cKeys.Count > 0) Then
        For lItem = 1 To m_cKeys.Count
            sKeys = sKeys & lItem & ItemKey(lItem) & Chr$(30)
        Next lItem
        sKeys = left$(sKeys, Len(sKeys) - 1)
        bKeys = sKeys
    End If
    
End Sub

Private Sub LoadKeys(ByRef bKeys() As Byte)

Dim lItem   As Long
Dim sKeys   As String
Dim aKeys() As String

    If (UBound(bKeys) > 0) Then
        sKeys = bKeys
        aKeys = Split(sKeys, Chr$(30))
        For lItem = 0 To UBound(aKeys)
            m_cKeys.Add left$(aKeys(lItem), 1), Mid$(aKeys(lItem), 2)
        Next lItem
    End If
    
End Sub

Public Sub SaveIcons(ByRef bKeys() As Byte, _
                     ByRef bIcons() As Byte)

Dim lCt     As Long
Dim lSize   As Long
Dim lStart  As Long
Dim lhIcon  As Long
Dim lHdc    As Long

    If (m_hIml > 0) Then
        If (ImageCount > -1) Then
            If (IconSizeX > 48) Then '<- 72x 32b icons can be past 16k boundary causing hang and gpf
                ReDim bIcons(0 To 32768 * ImageCount) As Byte
            Else
                ReDim bIcons(0 To 16384& * ImageCount) As Byte
            End If
            lHdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            For lCt = 1 To ImageCount
                lhIcon = ImageList_GetIcon(m_hIml, lCt - 1, 0&)
                If Not (lhIcon = 0) Then
                    If Not (lhIcon = -1) Then
                        SerialiseIcon lHdc, lhIcon, bIcons, lStart, lSize
                        DestroyIcon lhIcon
                    End If
                End If
                lStart = lStart + lSize
            Next lCt
            DeleteDC lHdc
        End If
    End If
    SaveKeys bKeys

End Sub

Public Function SerialiseIcon(ByVal lHdc As Long, _
                              ByVal lhIcon As Long, _
                              ByRef b() As Byte, _
                              ByVal lByteStart As Long, _
                              ByRef lArraySize As Long) As Boolean

Dim lR              As Long
Dim lMonoSize       As Long
Dim lColourSize     As Long
Dim tII             As ICONINFO

    lR = GetIconInfo(lhIcon, tII)
    If Not (lR = 0) Then
        ' store fIcon, xHotspot, yHotspot:
        RtlMoveMemory b(lByteStart), tII, 12
        ' store the colour bitmap:
        lByteStart = lByteStart + 12
        With tII
            If (SerialiseBitmap(lHdc, .hbmColor, False, b(), lByteStart, lColourSize)) Then
                lByteStart = lByteStart + lColourSize
                If (SerialiseBitmap(lHdc, .hBmMask, True, b(), lByteStart, lMonoSize)) Then
                    lByteStart = lByteStart + lMonoSize
                    lArraySize = lColourSize + lMonoSize + 12
                    SerialiseIcon = True
                End If
            End If
            DeleteObject .hbmColor
            DeleteObject .hBmMask
        End With
    End If

End Function

Private Function SerialiseBitmap(ByVal lHdc As Long, _
                                 ByVal hbm As Long, _
                                 ByVal bMono As Boolean, _
                                 ByRef b() As Byte, _
                                 ByVal lByteStart As Long, _
                                 ByRef lByteSize As Long) As Boolean

Dim lSize   As Long
Dim lR      As Long
Dim tBM     As BITMAP
Dim tBI1    As BITMAPINFO_1BPP
Dim tBI     As BITMAPINFO_ABOVE8

    lR = GetObjectAPI(hbm, Len(tBM), tBM)
    If Not (lR = 0) Then
        RtlMoveMemory b(lByteStart), tBM, Len(tBM)
        If (bMono) Then
            With tBI1.bmiHeader
                .biSize = Len(tBI1.bmiHeader)
                .biWidth = tBM.bmWidth
                .biHeight = tBM.bmHeight
                .biPlanes = 1
                .biBitCount = 1
                .biCompression = BI_RGB
                lSize = (.biWidth + 7) / 8
                lSize = ((lSize + 3) \ 4) * 4
                lSize = lSize * .biHeight
            End With
            lR = GetDIBits(lHdc, hbm, 0&, tBM.bmHeight, b(lByteStart + Len(tBM)), tBI1, DIB_RGB_COLORS)
          Else
            With tBI.bmiHeader
                .biSize = Len(tBI.bmiHeader)
                .biWidth = tBM.bmWidth
                .biHeight = tBM.bmHeight
                .biPlanes = 1
                .biBitCount = 24
                .biCompression = BI_RGB
                .biBitCount = 32
                lSize = .biWidth * .biHeight * 4
            End With
            lR = GetDIBits(lHdc, hbm, 0, tBM.bmHeight, b(lByteStart + Len(tBM)), tBI, DIB_RGB_COLORS)
        End If
        If Not (lR = 0) Then
            lByteSize = lSize + Len(tBM)
            SerialiseBitmap = True
        End If
    End If

End Function

Public Function RestoreIcons(ByRef bKeys() As Byte, _
                             ByRef bIcons() As Byte)

Dim lSize       As Long
Dim lStart      As Long
Dim lItemSize   As Long
Dim lhIcon      As Long
Dim lHdc        As Long

    lSize = UBound(bIcons)
    If (lSize > 0) Then
        lHdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        Do While lStart < lSize
            If (DeSerialiseIcon(lHdc, lhIcon, bIcons, lStart, lItemSize)) Then
                ImageList_AddIcon m_hIml, lhIcon
                DestroyIcon lhIcon
            End If
            lStart = lStart + lItemSize
        Loop
        DeleteDC lHdc
        Erase bIcons
    End If
    LoadKeys bKeys
    
End Function

Public Function DeSerialiseIcon(ByVal lHdc As Long, _
                                ByRef lhIcon As Long, _
                                ByRef b() As Byte, _
                                ByVal lByteStart As Long, _
                                ByRef lArraySize As Long)

Dim lColourSize     As Long
Dim lMonoSize       As Long
Dim tII             As ICONINFO

    lhIcon = 0
    ' get fIcon, xHotspot, yHotspot:
    RtlMoveMemory tII, b(lByteStart), 12
    With tII
        .fIcon = 1
        lByteStart = lByteStart + 12
        ' get the colour bitmap:
        If (DeSerialiseBitmap(lHdc, .hbmColor, False, b(), lByteStart, lColourSize)) Then
            lByteStart = lByteStart + lColourSize
            ' get the mono bitmap:
            If (DeSerialiseBitmap(lHdc, .hBmMask, True, b(), lByteStart, lMonoSize)) Then
                ' Set the size:
                lArraySize = lColourSize + lMonoSize + 12
                ' Create the icon from the structure:
                lhIcon = CreateIconIndirect(tII)
                DeSerialiseIcon = (lhIcon <> 0)
                DeleteObject .hbmColor
                DeleteObject .hBmMask
            Else
                DeleteObject .hbmColor
            End If
        End If
    End With

End Function

Private Function DeSerialiseBitmap(ByVal lHdc As Long, _
                                   ByRef hbm As Long, _
                                   ByVal bMono As Boolean, _
                                   ByRef b() As Byte, _
                                   ByVal lByteStart As Long, _
                                   ByRef lByteSize As Long) As Boolean

Dim lSize   As Long
Dim lR      As Long
Dim tBM     As BITMAP
Dim tBI1    As BITMAPINFO_1BPP
Dim tBI     As BITMAPINFO_ABOVE8

    RtlMoveMemory tBM, b(lByteStart), Len(tBM)
    If Not (bMono) Then
        hbm = CreateCompatibleBitmap(lHdc, tBM.bmWidth, tBM.bmHeight)
      Else
        hbm = CreateBitmapIndirect(tBM)
    End If
    If Not (hbm = 0) Then
        If (bMono) Then
            With tBI1.bmiHeader
                .biSize = Len(tBI1.bmiHeader)
                .biWidth = tBM.bmWidth
                .biHeight = tBM.bmHeight
                .biPlanes = 1
                .biBitCount = 1
                .biCompression = BI_RGB
            End With
            With tBI1
                lSize = (.bmiHeader.biWidth + 7) / 8
                lSize = ((lSize + 3) \ 4) * 4
                lSize = lSize * .bmiHeader.biHeight
                .bmiColors(1).rgbBlue = 255
                .bmiColors(1).rgbGreen = 255
                .bmiColors(1).rgbRed = 255
            End With
            lR = SetDIBits(lHdc, hbm, 0, tBM.bmHeight, b(lByteStart + Len(tBM)), tBI1, DIB_RGB_COLORS)
        Else
            With tBI.bmiHeader
                .biSize = Len(tBI.bmiHeader)
                .biWidth = tBM.bmWidth
                .biHeight = tBM.bmHeight
                .biPlanes = 1
                .biBitCount = 24
                .biCompression = BI_RGB
                .biBitCount = 32
                lSize = .biWidth * .biHeight * 4
            End With
            lR = SetDIBits(lHdc, hbm, 0, tBM.bmHeight, b(lByteStart + Len(tBM)), tBI, DIB_RGB_COLORS)
        End If

        lByteSize = lSize + Len(tBM)
        If Not (lR = 0) Then
            DeSerialiseBitmap = True
          Else
            DeleteObject hbm
        End If
    End If

End Function

Private Sub Class_Terminate()
    GdiUnload
    Destroy
End Sub

