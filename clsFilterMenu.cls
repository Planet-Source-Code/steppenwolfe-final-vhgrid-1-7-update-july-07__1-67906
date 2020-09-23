VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilterMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements GXISubclass

Private Const BM_TRANSPARENT                    As Long = &H1

Private Const DT_CENTER                         As Long = &H1&
Private Const DT_VCENTER                        As Long = &H4&
Private Const DT_SINGLELINE                     As Long = &H20&
Private Const DT_RTLREADING                     As Long = &H20000

Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const HWND_TOP                          As Long = 0

Private Const LF_ANTIALIASED_QUALITY            As Long = 4

Private Const LOGPIXELSY                        As Long = 90

Private Const MOUSEEVENTF_MOVE                  As Long = &H1
Private Const MOUSEEVENTF_LEFTDOWN              As Long = &H2
Private Const MOUSEEVENTF_LEFTUP                As Long = &H4

Private Const QS_ALLINPUT                       As Double = _
    (&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H20 Or &H40 Or &H80)

Private Const SS_OWNERDRAW                      As Long = &HD

Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_NOOWNERZORDER                 As Long = &H200

Private Const WS_TABSTOP                        As Long = &H10000
Private Const WS_CLIPCHILDREN                   As Long = &H2000000
Private Const WS_CLIPSIBLINGS                   As Long = &H4000000
Private Const WS_CHILD                          As Long = &H40000000

Private Const WS_EX_TOOLWINDOW                  As Long = &H80

Public Enum EFTTransition
    eftNone = 0&
    eftSlide = 1&
    eftInflate = 2&
End Enum

Public Enum EFSPosition
    RightBottom = 0&
    RightCenter = 1&
    RightTop = 2&
    LeftBottom = 3&
    LeftCenter = 4&
    LeftTop = 5&
End Enum

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Type VERSIONINFO
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
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

Private Type DRAWITEMSTRUCT
    CtlType                                     As Long
    CtlID                                       As Long
    itemID                                      As Long
    itemAction                                  As Long
    itemState                                   As Long
    hwndItem                                    As Long
    hdc                                         As Long
    rcItem                                      As RECT
    itemData                                    As Long
End Type


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As VERSIONINFO) As Long

Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As String, _
                                                       ByVal lpWindowName As String, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As Long, _
                                                       ByVal lpWindowName As Long, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As String, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As Long, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function CreateDc Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 lpDeviceName As Any, _
                                                                 lpOutput As Any, _
                                                                 lpInitData As Any) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, _
                                                                 pHandles As Long, _
                                                                 ByVal fWaitAll As Long, _
                                                                 ByVal dwMilliseconds As Long, _
                                                                 ByVal dwWakeMask As Long) As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
                                              ByVal dx As Long, _
                                              ByVal dy As Long, _
                                              ByVal cButtons As Long, _
                                              ByVal dwExtraInfo As Long)

Private Declare Function GetMessageExtraInfo Lib "user32" () As Long


Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Public Event FilterIndex(ByVal lIndex As Long)
Public Event DestroyMe()

Private m_bIsNt                                 As Boolean
Private m_bIsActive                             As Boolean
Private m_bGradient                             As Boolean
Private m_bXPColors                             As Boolean
Private m_bShadow                               As Boolean
Private m_bShowing                              As Boolean
Private m_bDestroy                              As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bFontRightLeading                     As Boolean
Private m_bFilterHideExact                      As Boolean
Private m_lTitleColor                           As Long
Private m_lTransparency                         As Long
Private m_lSafeTimer                            As Long
Private m_lTransition                           As Long
Private m_lFilterMenuHwnd                       As Long
Private m_lColorOffset                          As Long
Private m_lhTitleFont                           As Long
Private m_lhTextFont                            As Long
Private m_lPosition                             As Long
Private m_lBackColor                            As Long
Private m_lForeColor                            As Long
Private m_lParentHwnd                           As Long
Private m_lCtrlHwnd                             As Long
Private m_lWidth                                As Long
Private m_lHeight                               As Long
Private m_lControlColor                         As Long
Private m_lControlForeColor                     As Long
Private m_lThemeIndex                           As Long
Private m_sTitle                                As String
Private m_sListItems()                          As String
Private m_tRWnd                                 As RECT
Private m_oTextFont                             As StdFont
Private m_oTitleFont                            As StdFont
Private m_cList                                 As clsODControl
Private WithEvents m_cFilter                    As clsODControl
Attribute m_cFilter.VB_VarHelpID = -1
Private WithEvents m_cClose                     As clsODControl
Attribute m_cClose.VB_VarHelpID = -1
Private WithEvents m_cExact                     As clsODControl
Attribute m_cExact.VB_VarHelpID = -1
Private m_cFilterMenuDc                         As clsStoreDc
Private m_cRender                               As clsRender
Private m_SBSubclass                            As GXMSubclass


Private Sub Class_Initialize()

    m_bIsNt = CompatabilityCheck
    Set m_SBSubclass = New GXMSubclass
    Set m_cRender = New clsRender
    Set m_cFilterMenuDc = New clsStoreDc
    m_lWidth = 150
    m_lBackColor = GetSysColor(&H80000018 And &H1F)
    m_lForeColor = &H343434
    m_lTitleColor = &HFEFEFE
    m_lColorOffset = &HCACACA
    Set m_cList = New clsODControl
    Set m_cFilter = New clsODControl
    Set m_cClose = New clsODControl
    Set m_cExact = New clsODControl
    ReDim m_sListItems(0)
    m_lTransparency = -1
    m_lControlColor = -1

End Sub

Private Sub m_cClose_Click()
    m_bDestroy = True
End Sub

Private Sub m_cFilter_Click()

Dim lCt As Long

    With m_cList
        For lCt = 0 To (.ListCount - 1)
            If .SelectedItem(lCt) Then
                RaiseEvent FilterIndex(lCt)
            End If
        Next lCt
    End With
    
End Sub

Public Property Get BackColor() As Long
    BackColor = m_lBackColor
End Property

Public Property Let BackColor(ByVal PropVal As Long)
    If m_bXPColors Then
        m_lBackColor = m_cRender.XPShift(PropVal)
    Else
        m_lBackColor = PropVal
    End If
End Property

Public Property Get Checked() As Boolean
    Checked = m_cExact.Checked
End Property

Public Property Get ColorOffset() As Long
    ColorOffset = m_lColorOffset
End Property

Public Property Let ColorOffset(ByVal PropVal As Long)
    If m_bXPColors Then
        m_lColorOffset = m_cRender.XPShift(PropVal)
    Else
        m_lColorOffset = PropVal
    End If
End Property

Public Property Get ControlColor() As Long
    ControlColor = m_lControlColor
End Property

Public Property Let ControlColor(ByVal PropVal As Long)
    m_lControlColor = PropVal
End Property

Public Property Get ControlForeColor() As Long
    ControlForeColor = m_lControlForeColor
End Property

Public Property Let ControlForeColor(ByVal PropVal As Long)
    m_lControlForeColor = PropVal
End Property

Public Property Get FilterHideExact() As Boolean
    FilterHideExact = m_bFilterHideExact
End Property

Public Property Let FilterHideExact(ByVal PropVal As Boolean)
    m_bFilterHideExact = PropVal
End Property

Public Property Get FontRightLeading() As Boolean
'/* [get] right align fonts
    FontRightLeading = m_bFontRightLeading
End Property

Public Property Let FontRightLeading(ByVal PropvVal As Boolean)
'/* [let] right align fonts
    m_bFontRightLeading = PropvVal
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_lForeColor
End Property

Public Property Let ForeColor(ByVal PropVal As Long)
    m_lForeColor = PropVal
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_bGradient
End Property

Public Property Let Gradient(ByVal PropVal As Boolean)
    m_bGradient = PropVal
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Let Height(ByVal PropVal As Long)
    m_lHeight = PropVal
End Property

Public Property Get hwnd() As Long
    hwnd = m_lFilterMenuHwnd
End Property

Public Property Get Position() As EFSPosition
    Position = m_lPosition
End Property

Public Property Let Position(ByVal PropVal As EFSPosition)
    m_lPosition = PropVal
End Property

Public Property Get Shadow() As Boolean
    Shadow = m_bShadow
End Property

Public Property Let Shadow(ByVal PropVal As Boolean)
    m_bShadow = PropVal
End Property

Public Property Get Title() As String
    Title = m_sTitle
End Property

Public Property Let Title(ByVal PropVal As String)
    m_sTitle = PropVal
End Property

Public Property Get TitleColor() As Long
    TitleColor = m_lTitleColor
End Property

Public Property Let TitleColor(ByVal PropVal As Long)
    m_lTitleColor = PropVal
End Property

Public Property Get TextFont() As StdFont
    Set TextFont = m_oTextFont
End Property

Public Property Set TextFont(ByVal PropVal As StdFont)
    Set m_oTextFont = PropVal
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = m_oTitleFont
End Property

Public Property Set TitleFont(ByVal PropVal As StdFont)
    Set m_oTitleFont = PropVal
End Property

Public Property Get ThemeIndex() As Long
    ThemeIndex = m_lThemeIndex
End Property

Public Property Let ThemeIndex(ByVal PropVal As Long)
    m_lThemeIndex = PropVal
End Property

Public Property Get Transparency() As Long
    Transparency = m_lTransparency
End Property

Public Property Let Transparency(ByVal PropVal As Long)

    If (PropVal < 70) Then
        PropVal = 70
    ElseIf (PropVal > 255) Then
        PropVal = 254
    End If
    m_lTransparency = PropVal
    
End Property

Public Property Get Transition() As EFTTransition
    Transition = m_lTransition
End Property

Public Property Let Transition(ByVal PropVal As EFTTransition)
    m_lTransition = PropVal
End Property

Public Property Get UseUnicode() As Boolean
    UseUnicode = m_bUseUnicode
End Property

Public Property Let UseUnicode(ByVal PropVal As Boolean)
    m_bUseUnicode = PropVal
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Let Width(ByVal PropVal As Long)
    m_lWidth = PropVal
End Property

Public Property Get XPColors() As Boolean
    XPColors = m_bXPColors
End Property

Public Property Let XPColors(ByVal PropVal As Boolean)
    m_bXPColors = PropVal
End Property


Private Sub ActivateList()

Dim tPcd    As POINTAPI
Dim tRect   As RECT

    GetWindowRect m_cList.hwnd, tRect
    CopyMemory tPcd, tRect, Len(tPcd)
    ScreenToClient m_lFilterMenuHwnd, tPcd
    With tPcd
        mouse_event MOUSEEVENTF_MOVE, (.x + 7), (.y + 7), 0&, GetMessageExtraInfo()
    End With
    mouse_event MOUSEEVENTF_LEFTDOWN, 0&, 0&, 0&, GetMessageExtraInfo()
    mouse_event MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, GetMessageExtraInfo()
    
End Sub

Public Sub AddItem(ByVal lIndex As Long, _
                   ByVal sItem As String)

On Error GoTo Handler

    If UBound(m_sListItems) < lIndex Then
        ReDim Preserve m_sListItems(0 To lIndex)
    End If
    m_sListItems(lIndex) = sItem

Handler:

End Sub

Private Function CompatabilityCheck() As Boolean

Dim tVer As VERSIONINFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        CompatabilityCheck = True
    End If

End Function

Private Sub CreateControls()

Dim lBtm As Long

    With m_cList
        If Not (m_lControlColor = -1) Then
            .BackColor = m_lControlColor
        Else
            .BackColor = m_lBackColor
        End If
        If Not (m_lControlForeColor = -1) Then
            .ForeColor = m_lControlForeColor
        Else
            .ForeColor = m_lForeColor
        End If
        .UseUnicode = m_bUseUnicode
        .BorderStyle ecbsThin
        If Not m_bFilterHideExact Then
            lBtm = 49
        Else
            lBtm = 12
        End If
        .Create m_lFilterMenuHwnd, 5, 19, (m_tRWnd.Right - 12), (m_tRWnd.Bottom - lBtm), ecsListBox
        .HFont = m_lhTextFont
        .BorderStyle ecbsThin
    End With
    
    With m_cClose
        If Not (m_lControlColor = -1) Then
            .ThemeColor = m_lBackColor
        End If
        If Not (m_lControlForeColor = -1) Then
            .ForeColor = m_lControlForeColor
        Else
            .ForeColor = m_lForeColor
        End If
        .Name = "Close"
        .UseUnicode = m_bUseUnicode
        .Create m_lFilterMenuHwnd, (m_tRWnd.Right - 19), 4, 12, 12, ecsCommandButton
        .HFont = m_lhTextFont
        .Text = "-"
    End With
    
    If Not m_bFilterHideExact Then
        With m_cExact
            If Not (m_lControlColor = -1) Then
                .ThemeColor = m_lBackColor
            End If
            If Not (m_lControlColor = -1) Then
                .BackColor = m_lControlColor
            Else
                .BackColor = m_lBackColor
            End If
            If Not (m_lControlForeColor = -1) Then
                .ForeColor = m_lControlForeColor
            Else
                .ForeColor = m_lForeColor
            End If
            .Name = "Exact"
            .UseUnicode = m_bUseUnicode
            .Create m_lFilterMenuHwnd, 5, (m_tRWnd.Bottom - 25), 50, 14, ecsCheckBox
            .HFont = m_lhTextFont
            .AutoBackColor = True
            .Text = "Exact"
        End With
    End If
    
    With m_cFilter
        If Not (m_lControlColor = -1) Then
            .ThemeColor = m_lBackColor
        End If
        If Not (m_lControlColor = -1) Then
            .BackColor = m_lControlColor
        Else
            .BackColor = m_lBackColor
        End If
        If Not (m_lControlForeColor = -1) Then
            .ForeColor = m_lControlForeColor
        Else
            .ForeColor = m_lForeColor
        End If
        .HiliteColor = &HDCDCDC
        .Name = "Filter"
        .UseUnicode = m_bUseUnicode
        .Create m_lFilterMenuHwnd, (m_tRWnd.Right - 45), (m_tRWnd.Bottom - 25), 40, 18, ecsCommandButton
        .HFont = m_lhTextFont
        .Text = "Filter"
    End With

End Sub

Private Function CreateFont(ByVal oFont As StdFont, _
                            Optional ByVal bTitle As Boolean) As Long
'*/ change list font

Dim lChar       As Long
Dim lHdc        As Long
Dim tLF         As LOGFONT
Dim bteFont()   As Byte

On Error GoTo Handler
    
    lHdc = CreateDc("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    With tLF
        bteFont = StrConv(oFont.Name, vbFromUnicode)
        For lChar = 0 To UBound(bteFont)
            .lfFaceName(lChar) = bteFont(lChar)
        Next lChar
        If bTitle Then
            oFont.Bold = True
            oFont.Size = 9
        End If
        .lfHeight = -MulDiv(oFont.Size, GetDeviceCaps(lHdc, LOGPIXELSY), 72)
        .lfItalic = oFont.Italic
        .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
        .lfUnderline = oFont.Underline
        .lfStrikeOut = oFont.Strikethrough
        .lfCharSet = 4
        .lfQuality = LF_ANTIALIASED_QUALITY
    End With
    DeleteDC lHdc
    
    If m_bUseUnicode Then
        CreateFont = CreateFontIndirectW(tLF)
    Else
        CreateFont = CreateFontIndirectA(tLF)
    End If

On Error GoTo 0
Exit Function

Handler:

End Function

Private Sub CreateSortMenu()

Dim bOnDesktop  As Boolean
Dim lTTStyle    As Long

    m_lParentHwnd = GetParent(m_lCtrlHwnd)
    bOnDesktop = (m_lParentHwnd = GetDesktopWindow())
    lTTStyle = WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_TABSTOP Or SS_OWNERDRAW

    If m_bIsNt Then
        m_lFilterMenuHwnd = CreateWindowExW(-bOnDesktop * WS_EX_TOOLWINDOW, StrPtr("static"), StrPtr(""), lTTStyle, _
            0&, 0&, 0&, 0&, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
    Else
        m_lFilterMenuHwnd = CreateWindowExA(-bOnDesktop * WS_EX_TOOLWINDOW, "static", "", lTTStyle, _
            0&, 0&, 0&, 0&, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
    End If
    
End Sub

Private Sub DefaultFont(ByRef oFont As StdFont, _
                        Optional ByVal bTitle As Boolean)

    Set oFont = New StdFont
    With oFont
        .Charset = 3
        .Name = "MS Sans Serif"
        .Weight = 400
        .Size = 8
        If bTitle Then
            .Bold = True
        End If
    End With

End Sub

Private Sub DestroyFonts()

    If Not (m_lhTitleFont = 0) Then
        DeleteObject m_lhTitleFont
        m_lhTitleFont = 0
    End If
    If Not (m_oTitleFont Is Nothing) Then
        Set m_oTitleFont = Nothing
    End If
    If Not (m_lhTextFont = 0) Then
        DeleteObject m_lhTextFont
        m_lhTitleFont = 0
    End If
    If Not (m_oTextFont Is Nothing) Then
        Set m_oTextFont = Nothing
    End If
    
End Sub

Public Sub DrawMenu()

    CreateSortMenu
    MessageAttach
    Show
    LoadFonts
    CreateControls
    LoadList
    SetFocus m_cList.hwnd
    ActivateList
    m_bShowing = True
    StatusLoop

End Sub

Private Sub DrawText(ByVal lHdc As Long, _
                     ByVal sText As String, _
                     ByRef tRect As RECT)

Dim lFlags As Long

    SetBkMode lHdc, BM_TRANSPARENT
    SetTextColor lHdc, m_lTitleColor
    lFlags = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    sText = sText & Chr$(0)
    If m_bIsNt Then
        If m_bFontRightLeading Then
            lFlags = lFlags Or DT_RTLREADING
        End If
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
    Else
        DrawTextA lHdc, sText, -1, tRect, lFlags
    End If
    
End Sub

Private Sub FrameMenu(ByVal lHdc As Long, _
                      ByRef tRect As RECT)

Dim lhPen       As Long
Dim lhPenOld    As Long
Dim tPnt        As POINTAPI

    With tRect
        '/* outer
        lhPen = CreatePen(0, 1, &H808080)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, (.left - 1), (.Bottom - 1), tPnt
        LineTo lHdc, (.Right - 1), (.Bottom - 1)
        LineTo lHdc, (.Right - 1), .top
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        '/* inner
        lhPen = CreatePen(0, 1, &H999999)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, .left, (.Bottom - 2), tPnt
        LineTo lHdc, (.Right - 2), (.Bottom - 2)
        LineTo lHdc, (.Right - 2), .top
        LineTo lHdc, .left, .top
        LineTo lHdc, .left, (.Bottom - 2)
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        .left = 1
        .top = 1
        .Right = .Right - 2
        .Bottom = .Bottom - 2
    End With

End Sub

Private Sub LoadFonts()
        
    DestroyFonts
    If (m_oTextFont Is Nothing) Then
        DefaultFont m_oTextFont, False
    End If
    m_lhTextFont = CreateFont(m_oTextFont, False)
    If (m_oTitleFont Is Nothing) Then
        DefaultFont m_oTitleFont, True
    End If
    m_lhTitleFont = CreateFont(m_oTitleFont, True)
    
End Sub

Private Sub LoadList()

Dim lCt As Long

On Error GoTo Handler

    For lCt = 0 To UBound(m_sListItems)
        m_cList.InsertItem m_sListItems(lCt), lCt
    Next lCt
    
Handler:
    On Error GoTo 0
    
End Sub

Private Sub PaintMenu(ByVal lHdc As Long, _
                     ByRef tRect As RECT)

Dim lhBrush As Long

    lhBrush = CreateSolidBrush(m_lBackColor)
    FillRect lHdc, tRect, lhBrush
    DeleteObject lhBrush

End Sub

Private Sub MsgTimer(ByVal lInterval As Long)
'/* rough timer

Dim lTick   As Long
Dim lCount  As Long

On Error Resume Next

    If Not lInterval = -1 Then
        lTick = GetTickCount()
        lTick = lTick + lInterval
        If lTick > 0& Then
            lTick = ((lTick + &H80000000) + lInterval) + &H80000000
        Else
            lTick = ((lTick - &H80000000) + lInterval) - &H80000000
        End If
        Do
            If Err.Number = 0 Then Exit Sub
            lCount = GetTickCount()
            lCount = lTick - lCount
            If lTick > 0& Then
                lCount = ((lTick + &H80000000) - (lCount - &H80000000))
            Else
                lCount = ((lTick - &H80000000) - (lCount + &H80000000))
            End If
            If IIf((lCount Xor lInterval) > 0&, lCount > lInterval, lCount < 0&) Then
                Exit Sub
            End If
            MsgWaitForMultipleObjects 0&, 0&, 0&, lCount, QS_ALLINPUT
            DoEvents
        Loop
    End If
    
On Error GoTo 0

End Sub

Private Sub StatusLoop()

On Error GoTo Handler

    Do Until (TestActive = True)
        DoEvents
        If m_bDestroy Then
            Exit Do
        End If
        MsgTimer 50
    Loop

On Error GoTo 0

Handler:
    Destroy
    RaiseEvent DestroyMe
    
End Sub

Private Sub RenderMenu(ByVal lHdc As Long, _
                       ByRef tRect As RECT)

Dim lDrawDc     As Long
Dim lFntOld     As Long
Dim lHzOfst     As Long
Dim tRText      As RECT
Dim tRcpy       As RECT

    With m_cFilterMenuDc
        .Width = tRect.Right
        .Height = tRect.Bottom
        lDrawDc = .hdc
    End With
    
    CopyRect tRcpy, tRect
    FrameMenu lDrawDc, tRcpy
    
    With tRcpy
        If m_bGradient Then
            lHzOfst = (.Bottom / 3)
            m_cRender.Gradient lDrawDc, .left, .Right, .top, lHzOfst, m_lBackColor, m_lColorOffset, Fill_Vertical
            m_cRender.Gradient lDrawDc, .left, .Right, lHzOfst, (.Bottom - lHzOfst), m_lBackColor, m_lColorOffset, Fill_Vertical, True
        Else
            PaintMenu lDrawDc, tRect
        End If
    End With

    lFntOld = SelectObject(lDrawDc, m_lhTitleFont)
    CopyRect tRText, tRect
    tRText.Bottom = 20
    InflateRect tRText, 0, -2
    DrawText lDrawDc, m_sTitle, tRText
    SelectObject lDrawDc, lFntOld
    
    With tRect
        If (m_lTransparency = -1) Then
            m_cRender.Blit lHdc, 0, 0, .Right, .Bottom, lDrawDc, 0, 0, SRCCOPY
        Else
            m_cRender.AlphaBlit lHdc, 0, 0, .Right, .Bottom, lDrawDc, 0, 0, .Right, .Bottom, m_lTransparency
        End If
    End With
    
    SelectObject lDrawDc, lFntOld
    
End Sub

Public Sub SetPosition(ByVal lX As Long, _
                       ByVal lY As Long)
    
    If (m_lFilterMenuHwnd = 0) Then Exit Sub
    SetWindowPos m_lFilterMenuHwnd, HWND_TOP, lX, lY, 0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER
    
End Sub

Public Sub SetSize(ByVal lWidth As Long, _
                   ByVal lHeight As Long)

    If (m_lFilterMenuHwnd = 0) Then Exit Sub
    SetWindowPos m_lFilterMenuHwnd, HWND_TOP, 0&, 0&, lWidth, lHeight, SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOOWNERZORDER

End Sub

Private Sub Show()
    
    If Not (m_lFilterMenuHwnd = 0) Then
        With m_tRWnd
            SetWindowPos m_lFilterMenuHwnd, 0&, .left, .top, .Right, .Bottom, SWP_NOOWNERZORDER Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
        End With
    End If
    
End Sub

Public Sub ShowMenu(ByVal lHwnd As Long, _
                    ByVal sTitle As String, _
                    ByVal lX As Long, _
                    ByVal lY As Long, _
                    ByVal lWidth As Long, _
                    ByVal lHeight As Long)

    If Not (lHwnd = 0) Then
        m_lCtrlHwnd = lHwnd
        Title = sTitle
    End If
    
    With m_tRWnd
        .Bottom = lHeight
        .left = lX
        .Right = lWidth
        .top = lY
    End With
    
    If m_bIsActive Then
        Destroy
    End If
    DrawMenu
    
End Sub


'> Subclass
'>>>>>>>>>>>>>>>>
Private Sub MessageAttach()

    If Not m_SBSubclass Is Nothing Then
        If Not m_bIsActive Then
            With m_SBSubclass
                .Subclass m_lParentHwnd, Me
                .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
            End With
            m_bIsActive = True
        End If
    End If
    
End Sub

Private Sub MessageDetach()

    If Not m_SBSubclass Is Nothing Then
        If m_bIsActive Then
            With m_SBSubclass
                .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                .UnSubclass m_lParentHwnd
            End With
            m_bIsActive = False
        End If
    End If
    
End Sub

Private Function TestActive() As Boolean

Dim tPnt    As POINTAPI
Dim tRect   As RECT

    GetCursorPos tPnt
    GetWindowRect m_lFilterMenuHwnd, tRect
    If (PtInRect(tRect, tPnt.x, tPnt.y) = 0) Then
        m_lSafeTimer = m_lSafeTimer + 1
    Else
        m_lSafeTimer = 0
    End If
    If (m_lSafeTimer > 10) Then
        TestActive = True
    End If
    
End Function

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Dim tDIstc As DRAWITEMSTRUCT

    Select Case uMsg
    Case WM_DRAWITEM
        CopyMemory tDIstc, ByVal lParam, LenB(tDIstc)
        With tDIstc
            RenderMenu .hdc, .rcItem
        End With
    End Select
    
End Sub

'> Cleanup
'>>>>>>>>>>>>>>>>
Private Sub DestroyMenu()

    If Not (m_lFilterMenuHwnd = 0) Then
        MessageDetach
        DestroyWindow m_lFilterMenuHwnd
        m_bShowing = False
    End If

End Sub

Public Sub Destroy()

    If Not (m_lFilterMenuHwnd = 0) Then
        MessageDetach
        DestroyFonts
        If Not m_cList Is Nothing Then Set m_cList = Nothing
        If Not m_cFilter Is Nothing Then Set m_cFilter = Nothing
        If Not m_cClose Is Nothing Then Set m_cClose = Nothing
        If Not m_cExact Is Nothing Then Set m_cExact = Nothing
        If Not m_oTextFont Is Nothing Then Set m_oTextFont = Nothing
        If Not m_oTitleFont Is Nothing Then Set m_oTitleFont = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cFilterMenuDc Is Nothing Then Set m_cFilterMenuDc = Nothing
        If Not m_SBSubclass Is Nothing Then Set m_SBSubclass = Nothing
        DestroyMenu
        m_lFilterMenuHwnd = 0
    End If

End Sub

Private Sub Class_Terminate()
    Destroy
End Sub


