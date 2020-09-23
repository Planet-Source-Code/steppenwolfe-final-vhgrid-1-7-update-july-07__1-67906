VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdvancedEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements GXISubclass


Private Const HDS_FLAT                          As Long = &H200

Private Const NM_FIRST                          As Long = &HFFFF + 1

Private Const SWP_SHOWWINDOW                    As Long = &H40

Private Const WS_OVERLAPPED                     As Long = &H0
Private Const WS_CLIPSIBLINGS                   As Long = &H4000000
Private Const WS_CAPTION                        As Long = &HC00000
Private Const WS_SYSMENU                        As Long = &H80000
Private Const WS_THICKFRAME                     As Long = &H40000
Private Const WS_MINIMIZEBOX                    As Long = &H20000
Private Const WS_MAXIMIZEBOX                    As Long = &H10000
Private Const WS_CLIPCHILDREN                   As Long = &H2000000

Private Const WS_EX_TOOLWINDOW                  As Long = &H80
Private Const WS_EX_RTLREADING                  As Long = &H2000

Private Const CLRYELLOW = "&H00FFFF &H00F0F0 &H00E1E1 &H00D2D2 &H00C3C3 &H00B4B4 &H00A5A5 &H009696 &H008787 &H007878 &H006969" & _
                        " &H9CFFFF &H8DF0F0 &H7EE1E1 &H6FD2D2 &H60C3C3 &H51B4B4 &H42A5A5 &H339696 &H248787 &H157878 &H066969" & _
                        " &HD2FFFF &HC3F0F0 &HB4E1E1 &HA5D2D2 &H96C3C3 &H87B4B4 &H78A5A5 &H699696 &H5A8787 &H4B7878 &H3C6969" & _
                        " &HEBFFFF &HDCF0F0 &HCDE1E1 &HBED2D2 &HAFC3C3 &HA0B4B4 &H91A5A5 &H829696 &H738787 &H647878 &H556969"
Private Const CLRMAGENTA = "&HFF00FF &HF000F0 &HE100E1 &HD200D2 &HC300C3 &HB400B4 &HA500A5 &H960096 &H870087 &H780078 &H690069" & _
                        " &HFF9CFF &HF08DF0 &HE17EE1 &HD26FD2 &HC360C3 &HB451B4 &HA542A5 &H963396 &H872487 &H781578 &H690669" & _
                        " &HFFD2FF &HF0C3F0 &HE1B4E1 &HD2A5D2 &HC396C3 &HB487B4 &HA578A5 &H966996 &H875A87 &H784B78 &H693C69" & _
                        " &HFFEBFF &HF0DCF0 &HE1CDE1 &HD2BED2 &HC3AFC3 &HB4A0B4 &HA591A5 &H968296 &H877387 &H786478 &H695569"
Private Const CLRCYAN = "&HFFFF00 &HF0F000 &HE1E100 &HD2D200 &HC3C300 &HB4B400 &HA5A500 &H969600 &H878700 &H787800 &H696900" & _
                        " &HFFFF9C &HF0F08D &HE1E17E &HD2D26F &HC3C360 &HB4B451 &HA5A542 &H969633 &H878724 &H787815 &H696906" & _
                        " &HFFFFD2 &HF0F0C3 &HE1E1B4 &HD2D2A5 &HC3C396 &HB4B487 &HA5A578 &H969669 &H87875A &H78784B &H69693C" & _
                        " &HFFFFEB &HF0F0DC &HE1E1CD &HD2D2BE &HC3C3AF &HB4B4A0 &HA5A591 &H969682 &H878773 &H787864 &H696955"
Private Const CLRBLUE = "&HFF0000 &HF00000 &HE10000 &HD20000 &HC30000 &HB40000 &HA50000 &H960000 &H870000 &H780000 &H690000" & _
                        " &HFF9C9C &HF08D8D &HE17E7E &HD26F6F &HC36060 &HB45151 &HA54242 &H963333 &H872424 &H781515 &H690606" & _
                        " &HFFD2D2 &HF0C3C3 &HE1B4B4 &HD2A5A5 &HC39696 &HB48787 &HA57878 &H966969 &H875A5A &H784B4B &H963C3C" & _
                        " &HFFEBEB &HF0DCDC &HE1CDCD &HD2BEBE &HC3AFAF &HB4A0A0 &HA59191 &H968282 &H877373 &H786464 &H695555"
Private Const CLRRED = "&H0000FF &H0000F0 &H0000E1 &H0000D2 &H0000C3 &H0000B4 &H0000A5 &H000096 &H000087 &H000078 &H000069" & _
                        " &H9C9CFF &H8D8DF0 &H7E7EE1 &H6F6FD2 &H6060C3 &H5151B4 &H4242A5 &H333396 &H242487 &H151578 &H060669" & _
                        " &HD2D2FF &HC3C3F0 &HB4B4E1 &HA5A5D2 &H9696C3 &H8787B4 &H7878A5 &H696996 &H5A5A87 &H4B4B78 &H3C3C69" & _
                        " &HEBEBFF &HDCDCF0 &HCDCDE1 &HBEBED2 &HAFAFC3 &HA0A0B4 &H9191A5 &H828296 &H737387 &H646478 &H555569"
Private Const CLRGREEN = "&H00FF00 &H00F000 &H00E100 &H00D200 &H00C300 &H00B400 &H00A500 &H009600 &H008700 &H007800 &H006900" & _
                        " &H9CFF9C &H8DF08D &H7EE17E &H6FD26F &H60C360 &H51B451 &H42A542 &H339633 &H248724 &H157815 &H066906" & _
                        " &HD2FFD2 &HC3F0C3 &HB4E1B4 &HA5D2A5 &H96C396 &H87B487 &H78A578 &H699669 &H5A875A &H4B784B &H3C693C" & _
                        " &HEBFFEB &HDCF0DC &HCDE1CD &HBED2BE &HAFC3AF &HA0B4A0 &H91A591 &H829682 &H738773 &H647864 &H556955"
Private Const CLRGREY = "&HD2D2D2 &HC3C3C3 &HB4B4B4 &HA5A5A5 &H969696 &H878787 &H787878 &H696969 &H5A5A5A &H4B4B4B &H3C3C3C" & _
                        " &HEBEBEB &HDCDCDC &HCDCDCD &HBEBEBE &HAFAFAF &HA0A0A0 &H919191 &H828282 &H737373 &H646464 &H555555" & _
                        " &HFAFAFA &HEBEBEB &HDCDCDC &HCDCDCD &HBEBEBE &HAFAFAF &HA0A0A0 &H919191 &H828282 &H737373 &H646464"



Public Enum ECAThemeStyle
    ecaAzure = 0&
    ecaClassic = 1&
    ecaGloss = 2&
    ecaMetal = 3&
    ecaXp = 4&
End Enum

Public Enum EIAImageType
    eiaBitmap = 0&
    eiaIcon = 1&
    eiaCursor = 2&
    eiaMetafile = 3&
End Enum

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type VERSIONINFO
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
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

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, _
                                                           ByVal ImgIndex As Long, _
                                                           ByVal fuFlags As Long) As Long

Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Event ReturnData(ByVal sText As String, ByVal oFont As StdFont, ByVal lIcon As Long, ByVal lForeColor As Long, ByVal lBackColor As Long)
Public Event DestroyMe()


Private m_bIsNt                                 As Boolean
Private m_bShowing                              As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bFontRightLeading                     As Boolean
Private m_lParentHwnd                           As Long
Private m_lHostHwnd                             As Long
Private m_lWidth                                As Long
Private m_lHeight                               As Long
Private m_lImlHwnd                              As Long
Private m_lForeColor                            As Long
Private m_lBackColor                            As Long
Private m_lIconIndex                            As Long
Private m_lhIcon                                As Long
Private m_lThemeColor                           As Long
Private m_lOffsetColor                          As Long
Private m_sEditText                             As String
Private m_eImageType                            As EIAImageType
Private m_oFont                                 As StdFont
Private m_eThemeStyle                           As ECAThemeStyle
Private m_tRWnd                                 As RECT
Private m_cTxEditBox                            As clsODControl
Private m_cTxSize                               As clsODControl
Private m_cPbDisplay                            As clsODControl
Private m_cLbIcon                               As clsODControl
Private m_cLbFontSelect                         As clsODControl
Private m_cLbFontColor                          As clsODControl
Private m_cLbBackColor                          As clsODControl
Private m_cLbSize                               As clsODControl
Private WithEvents m_cCbBackColor               As clsODControl
Attribute m_cCbBackColor.VB_VarHelpID = -1
Private WithEvents m_cCbFontSelect              As clsODControl
Attribute m_cCbFontSelect.VB_VarHelpID = -1
Private WithEvents m_cCbFontColor               As clsODControl
Attribute m_cCbFontColor.VB_VarHelpID = -1
Private WithEvents m_cBtScrollUp                As clsODControl
Attribute m_cBtScrollUp.VB_VarHelpID = -1
Private WithEvents m_cBtScrollDwn               As clsODControl
Attribute m_cBtScrollDwn.VB_VarHelpID = -1
Private WithEvents m_cBtClose                   As clsODControl
Attribute m_cBtClose.VB_VarHelpID = -1
Private WithEvents m_cBtSave                    As clsODControl
Attribute m_cBtSave.VB_VarHelpID = -1
Private WithEvents m_cBtFontBold                As clsODControl
Attribute m_cBtFontBold.VB_VarHelpID = -1
Private WithEvents m_cBtFontItalic              As clsODControl
Attribute m_cBtFontItalic.VB_VarHelpID = -1
Private WithEvents m_cBtFontStrike              As clsODControl
Attribute m_cBtFontStrike.VB_VarHelpID = -1
Private WithEvents m_cBtFontUnderline           As clsODControl
Attribute m_cBtFontUnderline.VB_VarHelpID = -1
Private m_cHostSubclass                         As GXMSubclass


Private Sub Class_Initialize()

    m_bIsNt = CompatabilityCheck
    m_lForeColor = -1
    m_lBackColor = -1
    m_lIconIndex = -1
    
End Sub

Private Sub m_cBtClose_Click()
'/* signal close
    DestroyHost
    RaiseEvent DestroyMe
End Sub

Private Sub m_cBtSave_Click()
'/* return changes to parent

    If Not m_oFont Is Nothing Then
        With m_oFont
            .Bold = m_cBtFontBold.CommandPushed
            .Italic = m_cBtFontItalic.CommandPushed
            .Strikethrough = m_cBtFontStrike.CommandPushed
            .Underline = m_cBtFontUnderline.CommandPushed
        End With
    End If
    '/* send data
    RaiseEvent ReturnData(m_cTxEditBox.Text, m_oFont, m_lIconIndex, m_lForeColor, m_lBackColor)

End Sub

Private Sub m_cCbFontSelect_ItemChange(ByVal lItem As Long)
'/* font selection change

    With m_cCbFontSelect
        If (.ListIndex = 0) Then
            '/* open font dialog
            ComboGetFont
        Else
            '/* update local font
            If Not (.ListIndex = -1) Then
                Set m_oFont = New StdFont
                m_oFont.Name = .ListText(.ListIndex)
            End If
        End If
    End With

End Sub

Private Sub ComboGetFont()
'/* open font dialog

On Error Resume Next

    Set m_oFont = ShowFontDialog(m_lParentHwnd)

On Error GoTo 0

End Sub

Private Sub m_cBtScrollDwn_Click()
'/* scroll down icon list

    m_lIconIndex = (m_lIconIndex - 1)
    EditLoadPicture m_lIconIndex

End Sub

Private Sub m_cBtScrollUp_Click()
'/* scroll up icon list

    m_lIconIndex = (m_lIconIndex + 1)
    EditLoadPicture m_lIconIndex
    
End Sub

Private Sub m_cCbBackColor_ItemChange(ByVal lItem As Long)
'/* backcolor change

    With m_cCbBackColor
        If (.ListIndex = 0) Then
            '/* open color dialog
            m_lBackColor = ComboExtendedColors
            .ComboIndexColor = m_lBackColor
        Else
            If Not (.ListIndex = -1) Then
                m_lBackColor = CLng(.ListText(.ListIndex))
            End If
        End If
    End With

End Sub

Private Sub m_cCbFontColor_ItemChange(ByVal lItem As Long)
'/* forecolor change

    With m_cCbFontColor
        If (.ListIndex = 0) Then
            '/* open color dialog
            m_lForeColor = ComboExtendedColors
            .ComboIndexColor = m_lForeColor
        Else
            If Not (.ListIndex = -1) Then
                m_lForeColor = CLng(.ListText(.ListIndex))
            End If
        End If
    End With

End Sub

Public Property Get FontRightLeading() As Boolean
'/* [get] right align fonts
    FontRightLeading = m_bFontRightLeading
End Property

Public Property Let FontRightLeading(ByVal PropvVal As Boolean)
'/* [let] right align fonts
    m_bFontRightLeading = PropvVal
End Property

Public Property Get HostHwnd() As Long
    HostHwnd = m_lHostHwnd
End Property

Public Property Let HostHwnd(ByVal PropVal As Long)
    m_lHostHwnd = PropVal
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

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Let Height(ByVal PropVal As Long)
    m_lHeight = PropVal
End Property

Public Sub CreateEditBox(ByVal lOwnerHwnd As Long, _
                         ByVal lX As Long, _
                         ByVal lY As Long, _
                         ByVal eThemeStyle As ECAThemeStyle, _
                         Optional ByVal lThemecolor As Long = -1, _
                         Optional ByVal lThemeOffsetColor As Long = -1, _
                         Optional ByVal sText As String, _
                         Optional ByVal lImlHwnd As Long = -1, _
                         Optional ByVal lIcnIndex As Long = -1, _
                         Optional eImageType As EIAImageType)

'/* create edit window

    If Not (lOwnerHwnd = 0) Then
        m_lParentHwnd = lOwnerHwnd
        '/* create host window
        CreateWindow
        '/* subclass
        Attach
        With m_tRWnd
            .left = lX
            .Right = 362
            .top = lY
            .Bottom = 282
        End With
        '/* show
        SetPosition m_lHostHwnd, m_tRWnd
        '/* store image data
        If (lImlHwnd > -1) Then
            m_lImlHwnd = lImlHwnd
            m_eImageType = eImageType
            m_lIconIndex = lIcnIndex
        End If
        m_lThemeColor = lThemecolor
        m_lOffsetColor = lThemeOffsetColor
        m_sEditText = sText
        m_eThemeStyle = eThemeStyle
        '/* create support controls
        CreateControls lImlHwnd, lIcnIndex
    End If
    
End Sub

Private Sub CreateControls(ByVal lImlHwnd As Long, _
                           Optional ByVal lIcnIndex As Long = -1)

'/* create support controls

Dim oFont As StdFont

    Set m_cTxEditBox = New clsODControl
    With m_cTxEditBox
        '/* control name
        .Name = "txtEdit"
        '/* borderstyle
        .BorderStyle ecbsThin
        .UseUnicode = m_bUseUnicode
        .FontRightLeading = m_bFontRightLeading
        '/* create control window
        .Create m_lHostHwnd, 9, 6, 337, 118, ecsTextBox
        '/* add text
        .Text = m_sEditText
    End With
    
    Set m_cLbFontColor = New clsODControl
    With m_cLbFontColor
        .Name = "lblColor"
        .BorderStyle ecbsNone
        .AutoBackColor = True
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 9, 129, 44, 13, ecsLabel
        .Text = "ForeColor"
        .AutoSize = True
    End With
    
    Set m_cCbFontColor = New clsODControl
    With m_cCbFontColor
        .Name = "cbForeColor"
        .BorderStyle ecbsThin
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 9, 141, 160, 120, ecsImageCombo
        .AddItem "More.."
    End With
    
    Set m_cLbFontSelect = New clsODControl
    With m_cLbFontSelect
        .Name = "lblCellfont"
        .BorderStyle ecbsNone
        .AutoBackColor = True
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 9, 170, 41, 13, ecsLabel
        .Text = "Cell Font"
        .AutoSize = True
    End With
    
    Set m_cCbFontSelect = New clsODControl
    With m_cCbFontSelect
        .Name = "cbFontSelect"
        .BorderStyle ecbsThin
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 9, 182, 160, 120, ecsImageCombo
        .AddItem "More.."
    End With
    
    Set m_cBtFontBold = New clsODControl
    With m_cBtFontBold
        .Name = "cmdBold"
        .HiliteColor = &HCCCCCC
        Set oFont = New StdFont
        With oFont
            .Name = "ARIAL"
            .Size = 8
            .Bold = True
        End With
        Set .Font = oFont
        .CommandPushButton = True
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 9, 212, 22, 22, ecsCommandButton
        .Text = "B"
        Set oFont = Nothing
    End With
    
    Set m_cBtFontItalic = New clsODControl
    With m_cBtFontItalic
        .Name = "cmdItalic"
        .HiliteColor = &HCCCCCC
        Set oFont = New StdFont
        With oFont
            .Name = "ARIAL"
            .Size = 8
            .Bold = True
            .Italic = True
        End With
        Set .Font = oFont
        .CommandPushButton = True
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 35, 212, 22, 22, ecsCommandButton
        .Text = "I"
        Set oFont = Nothing
    End With
    
    Set m_cBtFontStrike = New clsODControl
    With m_cBtFontStrike
        .Name = "cmdStrike"
        .HiliteColor = &HCCCCCC
        Set oFont = New StdFont
        With oFont
            .Name = "ARIAL"
            .Size = 8
            .Bold = True
            .Strikethrough = True
        End With
        Set .Font = oFont
        .CommandPushButton = True
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 61, 212, 22, 22, ecsCommandButton
        .Text = "S"
        Set oFont = Nothing
    End With
    
    Set m_cBtFontUnderline = New clsODControl
    With m_cBtFontUnderline
        .Name = "cmdUnderline"
        .HiliteColor = &HCCCCCC
        Set oFont = New StdFont
        With oFont
            .Name = "Arial"
            .Size = 8
            .Bold = True
            .Underline = True
        End With
        Set .Font = oFont
        .CommandPushButton = True
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 87, 212, 22, 22, ecsCommandButton
        .Text = "U"
        Set oFont = Nothing
    End With

    Set m_cLbSize = New clsODControl
        With m_cLbSize
        .Name = "lblSize"
        .BorderStyle ecbsNone
        .AutoBackColor = True
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 137, 222, 10, 11, ecsLabel
        Set oFont = New StdFont
        With oFont
            .Name = "Small Fonts"
            .Size = 7
            .Bold = True
        End With
        Set .Font = oFont
        .Text = "Size"
        .AutoSize = True
    End With
    
    Set m_cTxSize = New clsODControl
    With m_cTxSize
        .Name = "txtSize"
        .BorderStyle ecbsThin
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 117, 214, 18, 19, ecsTextBox
        .Text = "8"
    End With
    
    Set m_cLbBackColor = New clsODControl
    With m_cLbBackColor
        .Name = "lblBackColor"
        .BorderStyle ecbsNone
        .AutoBackColor = True
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 186, 180, 48, 13, ecsLabel
        .Text = "BackColor"
        .AutoSize = True
    End With
    
    Set m_cCbBackColor = New clsODControl
    With m_cCbBackColor
        .Name = "cbBackColor"
        .BorderStyle ecbsThin
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 186, 192, 160, 120, ecsImageCombo
        .AddItem "More.."
    End With

    Set m_cBtSave = New clsODControl
    With m_cBtSave
        .Name = "cmdSave"
        .HiliteColor = &HCCCCCC
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 210, 225, 64, 22, ecsCommandButton
        .Text = "Save"
    End With
    
    Set m_cBtClose = New clsODControl
    With m_cBtClose
        .Name = "cmdClose"
        .HiliteColor = &HCCCCCC
        .ThemeStyle = m_eThemeStyle
        If (m_lThemeColor > -1) Then
            .ThemeColor = m_lThemeColor
        End If
        If (m_lOffsetColor > -1) Then
            .HiliteColor = m_lOffsetColor
        End If
        .FontRightLeading = m_bFontRightLeading
        .UseUnicode = m_bUseUnicode
        .Create m_lHostHwnd, 282, 225, 64, 22, ecsCommandButton
        .Text = "Close"
    End With

    If (lImlHwnd > -1) Then
        Set m_cLbIcon = New clsODControl
        With m_cLbIcon
            .Name = "lblCellIcon"
            .BorderStyle ecbsNone
            .AutoBackColor = True
            .FontRightLeading = m_bFontRightLeading
            .UseUnicode = m_bUseUnicode
            .Create m_lHostHwnd, 189, 129, 21, 13, ecsLabel
            .Text = "Cell Icon"
            .AutoSize = True
        End With
        
        Set m_cPbDisplay = New clsODControl
        With m_cPbDisplay
            .Name = ""
            .BorderStyle ecbsThin
            .BackColor = vbWhite
            .UseUnicode = m_bUseUnicode
            .Create m_lHostHwnd, 189, 141, 40, 34, ecsPictureBox
            '/* load current icon
            If (lIcnIndex > -1) Then
                EditLoadPicture lIcnIndex
            End If
        End With

        Set m_cBtScrollUp = New clsODControl
        With m_cBtScrollUp
            .Name = "cmdScrollUp"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_eThemeStyle
            Set oFont = Nothing
            Set oFont = New StdFont
            With oFont
                .Name = "Arial"
                .Size = 8
                .Bold = True
            End With
            Set .Font = oFont
            If (m_lThemeColor > -1) Then
                .ThemeColor = m_lThemeColor
            End If
            If (m_lOffsetColor > -1) Then
                .HiliteColor = m_lOffsetColor
            End If
            .UseUnicode = m_bUseUnicode
            .Create m_lHostHwnd, 234, 141, 16, 16, ecsCommandButton
            .Text = Chr$(43)
        End With
        
        Set m_cBtScrollDwn = New clsODControl
        With m_cBtScrollDwn
            .Name = "cmdScrollDown"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_eThemeStyle
            Set .Font = oFont
            If (m_lThemeColor > -1) Then
                .ThemeColor = m_lThemeColor
            End If
            If (m_lOffsetColor > -1) Then
                .HiliteColor = m_lOffsetColor
            End If
            .UseUnicode = m_bUseUnicode
            .Create m_lHostHwnd, 234, 160, 16, 16, ecsCommandButton
            .Text = Chr$(45)
        End With
    End If
    
    Set oFont = Nothing
    ComboAddColors
    ComboAddFonts
    
End Sub

Private Sub EditLoadPicture(ByVal lIcnIndex As Long)
'/* load icon into picturebox

Dim lWidth      As Long
Dim lHeight     As Long
Dim lImgCount   As Long

    If Not (m_lImlHwnd = -1) Then
        '/* get image count
        lImgCount = ImageList_GetImageCount(m_lImlHwnd)
        If (lIcnIndex > (lImgCount - 1)) Then
            m_lIconIndex = (lImgCount - 1)
        ElseIf m_lIconIndex < 0 Then
            m_lIconIndex = 0
        Else
            m_lIconIndex = lIcnIndex
        End If
        '/* clean up
        If Not (m_lhIcon = 0) Then
            EditDestroyIcon
        End If
        '/* create the icon copy
        m_lhIcon = ImageList_GetIcon(m_lImlHwnd, m_lIconIndex, 0&)
        If (m_lhIcon > 0) Then
            '/* get icon size
            ImageList_GetIconSize m_lImlHwnd, lWidth, lHeight
            '/* load to picturebox
            m_cPbDisplay.PictureBoxLoadImage m_lhIcon, m_eImageType, lWidth, lHeight
        End If
    End If

End Sub

Private Sub EditDestroyIcon()
'/* destroy icon copy

    If Not (m_lhIcon = 0) Then
        DestroyIcon m_lhIcon
        m_lhIcon = 0
    End If

End Sub

Private Sub ComboAddFonts()
'/* add font list and icons

Dim lCt     As Long
Dim lHdc    As Long
Dim lTtHnd  As Long
Dim lRtHnd  As Long
Dim vFont   As Variant
Dim cIml    As clsImageList

    '/* get system icon handles
    Set cIml = New clsImageList
    lRtHnd = cIml.SystemIconHandle(".FON", eisSmallIcon)
    lTtHnd = cIml.SystemIconHandle(".TTF", eisSmallIcon)
    With m_cCbFontSelect
        '/* init ods imagelist
        .InitListBoxIml 14, 14
        '/* add the icons
        .ImlListBoxAddIcon lRtHnd
        .ImlListBoxAddIcon lTtHnd
    End With
    '/* get the system fonts list
    lHdc = GetDC(m_lParentHwnd)
    vFont = EnumSystemFonts(lHdc)
    ReleaseDC m_lParentHwnd, lHdc
    '/* add font list to combo
    For lCt = 0 To UBound(vFont, 2)
        Select Case vFont(1, lCt)
        Case 0, 1
            m_cCbFontSelect.AddItem vFont(0, lCt), 0
        Case Else
            m_cCbFontSelect.AddItem vFont(0, lCt), 1
        End Select
    Next lCt
    Set cIml = Nothing
    
End Sub

Private Sub ComboAddColors()
'/* load combo color list

Dim lCt     As Long
Dim sColor  As String
Dim sClr()  As String

On Error Resume Next

    '/* split color const
    sColor = CLRYELLOW & CLRMAGENTA & CLRCYAN & CLRBLUE & CLRRED & CLRGREEN & CLRGREY
    sClr = Split(sColor, Chr$(32))
    '/* add to lists
    For lCt = 0 To UBound(sClr)
        m_cCbBackColor.AddItem sClr(lCt), , CLng(sClr(lCt))
        m_cCbFontColor.AddItem sClr(lCt), , CLng(sClr(lCt))
    Next lCt

On Error GoTo 0

End Sub

Private Function ComboExtendedColors() As Long
'/* launch color dialog

Dim lRet        As Long
Dim lCust()    As Long

    ReDim lCust(15)
    lCust(0) = &HFFFFFF
    lRet = ShowColorDialog(m_lParentHwnd, &HFFFFFF, lCust, 1)
    If Not (lRet = -1) Then
        ComboExtendedColors = lRet
    End If

End Function

Private Sub CreateWindow()
'/* create api window

Dim lTTStyle As Long
Dim lExStyle As Long
Dim sTitle   As String

    '/* style constants
    lTTStyle = WS_CLIPSIBLINGS Or WS_SYSMENU Or HDS_FLAT Or WS_CLIPCHILDREN
    lExStyle = WS_EX_TOOLWINDOW
    If m_bFontRightLeading Then
        lExStyle = lExStyle Or WS_EX_RTLREADING
    End If
    sTitle = "Advanced Edit"
    '/* create tool/header window
    If m_bIsNt Then
        m_lHostHwnd = CreateWindowExW(lExStyle, StrPtr("SysHeader32"), StrPtr(sTitle), lTTStyle, _
            0&, 0&, 0&, 0&, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
    Else
        m_lHostHwnd = CreateWindowExA(lExStyle, "SysHeader32", sTitle, lTTStyle, _
            0&, 0&, 0&, 0&, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
    End If

End Sub

Private Function CompatabilityCheck() As Boolean
'/* nt version check

Dim tVer As VERSIONINFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        CompatabilityCheck = True
    End If

End Function

Private Sub SetPosition(ByVal lHwnd As Long, _
                        ByRef tRect As RECT)

'/* show window

    If Not (m_lHostHwnd = 0) Then
        With tRect
            SetWindowPos lHwnd, 0&, .left, .top, .Right, .Bottom, SWP_SHOWWINDOW
        End With
    End If

End Sub

Private Sub Attach()
'/* attach subclasser

    If Not (m_lHostHwnd = 0) Then
        Set m_cHostSubclass = New GXMSubclass
        With m_cHostSubclass
            .Subclass m_lHostHwnd, Me
            .AddMessage m_lHostHwnd, WM_CLOSE, MSG_AFTER
        End With
    End If
    
End Sub

Private Sub Detach()
'/* detach subclasser

    If Not m_cHostSubclass Is Nothing Then
        With m_cHostSubclass
            .DeleteMessage m_lHostHwnd, WM_CLOSE, MSG_AFTER
            .UnSubclass m_lHostHwnd
        End With
        Set m_cHostSubclass = Nothing
    End If
    
End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)
    
    '/* signal window termination to parent
    If (uMsg = WM_CLOSE) Then
        RaiseEvent DestroyMe
    End If
    
End Sub


'> Cleanup
'>>>>>>>>>>>>>>>>
Private Sub DestroyHost()
'/* destroy host window

    '/* detach subclasser
    Detach
    '/* destroy window
    If Not m_lHostHwnd = 0 Then
        DestroyWindow m_lHostHwnd
        m_lHostHwnd = 0
        m_bShowing = False
    End If
    
End Sub

Public Sub DestroyEditBox()
'/* cleanup
    
    EditDestroyIcon
    If Not m_cTxEditBox Is Nothing Then Set m_cTxEditBox = Nothing
    If Not m_cTxSize Is Nothing Then Set m_cTxSize = Nothing
    
    If Not m_cPbDisplay Is Nothing Then Set m_cPbDisplay = Nothing
    
    If Not m_cCbBackColor Is Nothing Then Set m_cCbBackColor = Nothing
    If Not m_cCbFontSelect Is Nothing Then Set m_cCbFontSelect = Nothing
    If Not m_cCbFontColor Is Nothing Then Set m_cCbFontColor = Nothing
    
    If Not m_cBtFontBold Is Nothing Then Set m_cBtFontBold = Nothing
    If Not m_cBtFontItalic Is Nothing Then Set m_cBtFontItalic = Nothing
    If Not m_cBtFontStrike Is Nothing Then Set m_cBtFontStrike = Nothing
    If Not m_cBtFontUnderline Is Nothing Then Set m_cBtFontUnderline = Nothing
    If Not m_cBtClose Is Nothing Then Set m_cBtClose = Nothing
    If Not m_cBtSave Is Nothing Then Set m_cBtSave = Nothing
    If Not m_cBtScrollUp Is Nothing Then Set m_cBtScrollUp = Nothing
    If Not m_cBtScrollDwn Is Nothing Then Set m_cBtScrollDwn = Nothing
    
    If Not m_cLbIcon Is Nothing Then Set m_cLbIcon = Nothing
    If Not m_cLbFontSelect Is Nothing Then Set m_cLbFontSelect = Nothing
    If Not m_cLbFontColor Is Nothing Then Set m_cLbFontColor = Nothing
    If Not m_cLbBackColor Is Nothing Then Set m_cLbBackColor = Nothing
    If Not m_cLbSize Is Nothing Then Set m_cLbSize = Nothing
    If Not m_oFont Is Nothing Then Set m_oFont = Nothing
    DestroyHost

End Sub

Private Sub Class_Terminate()
    DestroyEditBox
End Sub
