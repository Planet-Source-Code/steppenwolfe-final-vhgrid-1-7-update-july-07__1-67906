Attribute VB_Name = "mSupport"
Option Explicit

Private Const CF_PRINTERFONTS                   As Long = &H2
Private Const CF_SCREENFONTS                    As Long = &H1
Private Const CF_BOTH                           As Long = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_EFFECTS                        As Long = &H100&
Private Const CF_FORCEFONTEXIST                 As Long = &H10000
Private Const CF_INITTOLOGFONTSTRUCT            As Long = &H40&
Private Const CF_LIMITSIZE                      As Long = &H2000&

Private Const CLIP_DEFAULT_PRECIS               As Long = 0

Private Const DEFAULT_CHARSET                   As Long = 1
Private Const DEFAULT_QUALITY                   As Long = 0
Private Const DEFAULT_PITCH                     As Long = 0

Private Const FF_ROMAN                          As Long = 16

Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const GMEM_MOVEABLE                     As Long = &H2
Private Const GMEM_ZEROINIT                     As Long = &H40

Private Const LOGPIXELSY                        As Long = 90

Private Const MAX_PATH                          As Long = 260

Private Const OUT_DEFAULT_PRECIS                As Long = 0

Private Const REGULAR_FONTTYPE                  As Long = &H400

Public Enum EColorFlag
    CC_RGBINIT = &H1
    CC_FULLOPEN = &H2
    CC_PREVENTFULLOPEN = &H4
    CC_SHOWHELP = &H8
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
    CC_SolidColor = &H80
    CC_ANYCOLOR = &H100
End Enum


Private Type SHFILEINFOA
    hIcon                                       As Long
    iIcon                                       As Long
    dwAttributes                                As Long
    szDisplayName                               As String * MAX_PATH
    szTypeName                                  As String * 80
End Type

Private Type SHFILEINFOW
    hIcon                                       As Long
    iIcon                                       As Long
    dwAttributes                                As Long
    szDisplayName(MAX_PATH)                     As Byte
    szTypeName(80)                              As Byte
End Type

Private Type LOGFONT
    lfHeight                                    As Long
    lfWidth                                     As Long
    lfEscapement                                As Long
    lfOrientation                               As Long
    lfWeight                                    As Long
    lfItalic                                    As Byte
    lfUnderline                                 As Byte
    lfStrikeOut                                 As Byte
    lfCharSet                                   As Byte
    lfOutPrecision                              As Byte
    lfClipPrecision                             As Byte
    lfQuality                                   As Byte
    lfPitchAndFamily                            As Byte
    lfFaceName(32)                              As Byte
End Type

Private Type NEWTEXTMETRIC
    tmHeight                                    As Long
    tmAscent                                    As Long
    tmDescent                                   As Long
    tmInternalLeading                           As Long
    tmExternalLeading                           As Long
    tmAveCharWidth                              As Long
    tmMaxCharWidth                              As Long
    tmWeight                                    As Long
    tmOverhang                                  As Long
    tmDigitizedAspectX                          As Long
    tmDigitizedAspectY                          As Long
    tmFirstChar                                 As Byte
    tmLastChar                                  As Byte
    tmDefaultChar                               As Byte
    tmBreakChar                                 As Byte
    tmItalic                                    As Byte
    tmUnderlined                                As Byte
    tmStruckOut                                 As Byte
    tmPitchAndFamily                            As Byte
    tmCharSet                                   As Byte
    ntmFlags                                    As Long
    ntmSizeEM                                   As Long
    ntmCellHeight                               As Long
    ntmAveWidth                                 As Long
End Type

Private Type CHOOSECOLOR
    lStructSize                                 As Long
    hwndOwner                                   As Long
    hInstance                                   As Long
    rgbResult                                   As Long
    lpCustColors                                As Long
    flags                                       As Long
    lCustData                                   As Long
    lpfnHook                                    As Long
    lpTemplateName                              As String
End Type

Private Type CHOOSEFONT
    lStructSize                                 As Long
    hwndOwner                                   As Long
    hdc                                         As Long
    lpLogFont                                   As Long
    iPointSize                                  As Long
    flags                                       As Long
    rgbColors                                   As Long
    lCustData                                   As Long
    lpfnHook                                    As Long
    lpTemplateName                              As Long
    hInstance                                   As Long
    lpszStyle                                   As Long
    nFontType                                   As Integer
    MISSING_ALIGNMENT                           As Integer
    nSizeMin                                    As Long
    nSizeMax                                    As Long
End Type

Private Type VERSIONINFO
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
End Type

Private Declare Function EnumFontFamiliesA Lib "gdi32" (ByVal hdc As Long, _
                                                        ByVal lpszFamily As String, _
                                                        ByVal lpEnumFontFamProc As Long, _
                                                        lParam As Any) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

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

Private Declare Function ChooseColorA Lib "comdlg32.dll" (pColor As CHOOSECOLOR) As Long

Private Declare Function ChooseFontA Lib "comdlg32.dll" (pChoosefont As CHOOSEFONT) As Long

Private Declare Function ChooseFontW Lib "comdlg32.dll" (pChoosefont As CHOOSEFONT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hDest As Any, _
                                                                     hSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                     ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As VERSIONINFO) As Long


Private m_lFontCount As Long


Private Function IsNT() As Boolean

Dim tVer  As VERSIONINFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        IsNT = True
    End If

End Function

Public Function EnumSystemFonts(ByVal lHdc As Long) As Variant

Dim vFonts As Variant

    m_lFontCount = 0
    ReDim vFonts(1, 0)
    EnumFontFamiliesA lHdc, vbNullString, AddressOf EnumFontFamProc, vFonts
    EnumSystemFonts = vFonts
    Erase vFonts

End Function

Public Function ShowColorDialog(ByVal lOwnerHwnd As Long, _
                                ByVal lDfltClr As Long, _
                                ByRef lCustomClr() As Long, _
                                Optional ByVal ShowMode As Integer = 0) As Long

Dim tCD As CHOOSECOLOR

On Error GoTo Handler

    With tCD
        .lStructSize = Len(tCD)
        .hwndOwner = lOwnerHwnd
        .hInstance = App.hInstance
        .flags = CC_ANYCOLOR
        Select Case ShowMode
        Case 1
            .flags = .flags Or CC_FULLOPEN
        Case 2
            .flags = .flags Or CC_PREVENTFULLOPEN
        End Select
        .flags = .flags Or CC_RGBINIT
        .rgbResult = lDfltClr
        .lpCustColors = VarPtr(lCustomClr(0))
        If ChooseColorA(tCD) = 1 Then
            ShowColorDialog = .rgbResult
        Else
            ShowColorDialog = -1
        End If
    End With

Handler:

End Function

Public Function ShowFontDialog(ByVal lOwnerHwnd As Long) As StdFont

Dim lhMem       As Long
Dim lPtr        As Long
Dim lRet        As Long
Dim lChar       As Long
Dim lHdc        As Long
Dim sDftFnt     As String
Dim bteFont()   As Byte
Dim tCF         As CHOOSEFONT
Dim tFont       As LOGFONT
Dim oStdFnt     As StdFont

On Error GoTo Handler

    sDftFnt = "MS Sans Serif" & Chr$(0)
    With tFont
        bteFont = StrConv(sDftFnt, vbFromUnicode)
        For lChar = 0 To UBound(bteFont)
            .lfFaceName(lChar) = bteFont(lChar)
        Next lChar
        .lfHeight = 0
        .lfWidth = 0
        .lfEscapement = 0
        .lfOrientation = 0
        .lfWeight = FW_NORMAL
        .lfCharSet = DEFAULT_CHARSET
        .lfOutPrecision = OUT_DEFAULT_PRECIS
        .lfClipPrecision = CLIP_DEFAULT_PRECIS
        .lfQuality = DEFAULT_QUALITY
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    End With

    lhMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(tFont))
    lPtr = GlobalLock(lhMem)
    CopyMemory ByVal lPtr, tFont, Len(tFont)
    lHdc = GetDC(lOwnerHwnd)
    
    With tCF
        .lStructSize = Len(tCF)
        .hwndOwner = lOwnerHwnd
        .hdc = lHdc
        .lpLogFont = lPtr
        .iPointSize = 120
        .flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .rgbColors = RGB(0, 0, 0)
        .nFontType = REGULAR_FONTTYPE
        .nSizeMin = 10
        .nSizeMax = 72
    End With
    
    lRet = ChooseFontA(tCF)
    If Not (lRet = 0) Then
        CopyMemory tFont, ByVal lPtr, Len(tFont)
        Set oStdFnt = New StdFont
        With oStdFnt
            .Bold = (tFont.lfWeight >= FW_BOLD)
            .Charset = tFont.lfCharSet
            .Italic = CBool(tFont.lfItalic)
            .Name = StrConv(tFont.lfFaceName, vbUnicode)
            .Size = HeightToPoints(tFont.lfHeight)
            .Strikethrough = tFont.lfStrikeOut
            .Underline = tFont.lfUnderline
            .Weight = tFont.lfWeight
        End With
        Set ShowFontDialog = oStdFnt
        Set oStdFnt = Nothing
    End If

Handler:
    If Not (lHdc = 0) Then
        ReleaseDC lOwnerHwnd, lHdc
    End If
    If Not (lhMem = 0) Then
        GlobalUnlock lhMem
        GlobalFree lhMem
    End If

End Function

Private Function HeightToPoints(ByVal lNum As Long) As Single
   HeightToPoints = (-72 * lNum) / PixelsPerInchY
End Function

Private Function PixelsPerInchY() As Long

Dim lHwnd   As Long
Dim lHdc    As Long
   
   lHwnd = GetDesktopWindow()
   lHdc = GetDC(lHwnd)
   PixelsPerInchY = GetDeviceCaps(lHdc, LOGPIXELSY)
   ReleaseDC lHwnd, lHdc

End Function

Private Function EnumFontFamProc(lpLF As LOGFONT, _
                                 lpTM As NEWTEXTMETRIC, _
                                 ByVal lFontType As Long, _
                                 lParam As Variant) As Long

Dim sName As String

    sName = StrConv(lpLF.lfFaceName, vbUnicode)
    ReDim Preserve lParam(1, 0 To m_lFontCount)
    lParam(0, m_lFontCount) = left$(sName, InStr(sName, vbNullChar) - 1)
    lParam(1, m_lFontCount) = lFontType
    m_lFontCount = m_lFontCount + 1
    EnumFontFamProc = 1

End Function

