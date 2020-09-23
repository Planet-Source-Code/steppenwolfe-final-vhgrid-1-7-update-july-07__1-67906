VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsODControl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'/~ 15 custom controls in one small class..

Implements GXISubclass

Private Const BCM_FIRST                         As Long = &H1600
Private Const BCM_GETIDEALSIZE                  As Long = (BCM_FIRST + &H1)
Private Const BCM_GETIMAGELIST                  As Long = (BCM_FIRST + &H3)
Private Const BCM_GETTEXTMARGIN                 As Long = (BCM_FIRST + &H5)
Private Const BCM_SETIMAGELIST                  As Long = (BCM_FIRST + &H2)
Private Const BCM_SETTEXTMARGIN                 As Long = (BCM_FIRST + &H4)

Private Const BM_TRANSPARENT                    As Long = &H1

Private Const BM_GETCHECK                       As Long = &HF0
Private Const BM_SETCHECK                       As Long = &HF1
Private Const BM_GETSTATE                       As Long = &HF2
Private Const BM_SETSTATE                       As Long = &HF3
Private Const BM_SETSTYLE                       As Long = &HF4
Private Const BM_CLICK                          As Long = &HF5
Private Const BM_GETIMAGE                       As Long = &HF6
Private Const BM_SETIMAGE                       As Long = &HF7

Private Const BN_CLICKED                        As Long = &H0
Private Const BN_PAINT                          As Long = &H1
Private Const BN_HILITE                         As Long = &H2
Private Const BN_UNHILITE                       As Long = &H3
Private Const BN_DISABLE                        As Long = &H4
Private Const BN_DOUBLECLICKED                  As Long = &H5
Private Const BN_SETFOCUS                       As Long = &H6
Private Const BN_KILLFOCUS                      As Long = &H7

Private Const BST_UNCHECKED                     As Long = &H0
Private Const BST_CHECKED                       As Long = &H1
Private Const BST_INDETERMINATE                 As Long = &H2
Private Const BST_PUSHED                        As Long = &H4
Private Const BST_FOCUS                         As Long = &H8

Private Const BS_PUSHBUTTON                     As Long = &H0
Private Const BS_DEFPUSHBUTTON                  As Long = &H1
Private Const BS_CHECKBOX                       As Long = &H2
Private Const BS_AUTOCHECKBOX                   As Long = &H3
Private Const BS_RADIOBUTTON                    As Long = &H4
Private Const BS_3STATE                         As Long = &H5
Private Const BS_AUTO3STATE                     As Long = &H6
Private Const BS_GROUPBOX                       As Long = &H7
Private Const BS_USERBUTTON                     As Long = &H8
Private Const BS_AUTORADIOBUTTON                As Long = &H9
Private Const BS_OWNERDRAW                      As Long = &HB
Private Const BS_LEFTTEXT                       As Long = &H20
Private Const BS_TEXT                           As Long = &H0
Private Const BS_ICON                           As Long = &H40
Private Const BS_BITMAP                         As Long = &H80
Private Const BS_LEFT                           As Long = &H100
Private Const BS_RIGHT                          As Long = &H200
Private Const BS_CENTER                         As Long = &H300
Private Const BS_TOP                            As Long = &H400
Private Const BS_BOTTOM                         As Long = &H800
Private Const BS_VCENTER                        As Long = &HC00
Private Const BS_PUSHLIKE                       As Long = &H1000
Private Const BS_MULTILINE                      As Long = &H2000
Private Const BS_NOTIFY                         As Long = &H4000
Private Const BS_FLAT                           As Long = &H8000

Private Const CB_GETEDITSEL                     As Long = &H140
Private Const CB_LIMITTEXT                      As Long = &H141
Private Const CB_SETEDITSEL                     As Long = &H142
Private Const CB_ADDSTRING                      As Long = &H143
Private Const CB_DELETESTRING                   As Long = &H144
Private Const CB_DIR                            As Long = &H145
Private Const CB_GETCOUNT                       As Long = &H146
Private Const CB_GETCURSEL                      As Long = &H147
Private Const CB_GETLBTEXT                      As Long = &H148
Private Const CB_GETLBTEXTLEN                   As Long = &H149
Private Const CB_INSERTSTRING                   As Long = &H14A
Private Const CB_RESETCONTENT                   As Long = &H14B
Private Const CB_FINDSTRING                     As Long = &H14C
Private Const CB_SELECTSTRING                   As Long = &H14D
Private Const CB_SETCURSEL                      As Long = &H14E
Private Const CB_SHOWDROPDOWN                   As Long = &H14F
Private Const CB_GETITEMDATA                    As Long = &H150
Private Const CB_SETITEMDATA                    As Long = &H151
Private Const CB_GETDROPPEDCONTROLRECT          As Long = &H152
Private Const CB_SETITEMHEIGHT                  As Long = &H153
Private Const CB_GETITEMHEIGHT                  As Long = &H154
Private Const CB_SETEXTENDEDUI                  As Long = &H155
Private Const CB_GETEXTENDEDUI                  As Long = &H156
Private Const CB_GETDROPPEDSTATE                As Long = &H157
Private Const CB_FINDSTRINGEXACT                As Long = &H158
Private Const CB_SETLOCALE                      As Long = &H159
Private Const CB_GETLOCALE                      As Long = &H15A
Private Const CB_GETTOPINDEX                    As Long = &H15B
Private Const CB_SETTOPINDEX                    As Long = &H15C
Private Const CB_GETHORIZONTALEXTENT            As Long = &H15D
Private Const CB_SETHORIZONTALEXTENT            As Long = &H15E
Private Const CB_GETDROPPEDWIDTH                As Long = &H15F
Private Const CB_SETDROPPEDWIDTH                As Long = &H160
Private Const CB_INITSTORAGE                    As Long = &H161

Private Const CLR_NONE                          As Long = -1
Private Const CBN_SELCHANGE                     As Long = &H1
Private Const CBN_DBLCLK                        As Long = &H2
Private Const CBN_SETFOCUS                      As Long = &H3
Private Const CBN_KILLFOCUS                     As Long = &H4
Private Const CBN_EDITCHANGE                    As Long = &H5
Private Const CBN_EDITUPDATE                    As Long = &H6
Private Const CBN_DROPDOWN                      As Long = &H7
Private Const CBN_CLOSEUP                       As Long = &H8
Private Const CBN_SELENDCANCEL                  As Long = &HA

Private Const CBS_SIMPLE                        As Long = &H1
Private Const CBS_DROPDOWN                      As Long = &H2
Private Const CBS_DROPDOWNLIST                  As Long = &H3
Private Const CBS_OWNERDRAWFIXED                As Long = &H10
Private Const CBS_OWNERDRAWVARIABLE             As Long = &H20
Private Const CBS_AUTOHSCROLL                   As Long = &H40
Private Const CBS_SORT                          As Long = &H100
Private Const CBS_HASSTRINGS                    As Long = &H200

Private Const CCM_FIRST                         As Long = &H2000
Private Const CCM_SETUNICODEFORMAT              As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT              As Long = (CCM_FIRST + 6)

Private Const DT_TOP                            As Long = &H0&
Private Const DT_LEFT                           As Long = &H0&
Private Const DT_CENTER                         As Long = &H1&
Private Const DT_RIGHT                          As Long = &H2&
Private Const DT_VCENTER                        As Long = &H4&
Private Const DT_BOTTOM                         As Long = &H8&
Private Const DT_WORDBREAK                      As Long = &H10&
Private Const DT_SINGLELINE                     As Long = &H20&
Private Const DT_EXPANDTABS                     As Long = &H40&
Private Const DT_TABSTOP                        As Long = &H80&
Private Const DT_NOCLIP                         As Long = &H100&
Private Const DT_EXTERNALLEADING                As Long = &H200&
Private Const DT_CALCRECT                       As Long = &H400&
Private Const DT_NOPREFIX                       As Long = &H800&
Private Const DT_INTERNAL                       As Long = &H1000&
Private Const DT_EDITCONTROL                    As Long = &H2000&
Private Const DT_PATH_ELLIPSIS                  As Long = &H4000&
Private Const DT_END_ELLIPSIS                   As Long = &H8000&
Private Const DT_MODIFYSTRING                   As Long = &H10000
Private Const DT_RTLREADING                     As Long = &H20000
Private Const DT_WORD_ELLIPSIS                  As Long = &H40000

Private Const EM_SETREADONLY                    As Long = &HCF

Private Const ES_LEFT                           As Long = &H0
Private Const ES_CENTER                         As Long = &H1
Private Const ES_RIGHT                          As Long = &H2
Private Const ES_MULTILINE                      As Long = &H4
Private Const ES_UPPERCASE                      As Long = &H8
Private Const ES_LOWERCASE                      As Long = &H10
Private Const ES_PASSWORD                       As Long = &H20
Private Const ES_AUTOVSCROLL                    As Long = &H40
Private Const ES_AUTOHSCROLL                    As Long = &H80
Private Const ES_NOHIDESEL                      As Long = &H100
Private Const ES_OEMCONVERT                     As Long = &H400
Private Const ES_READONLY                       As Long = &H800
Private Const ES_WANTRETURN                     As Long = &H1000
Private Const ES_NUMBER                         As Long = &H2000
Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const GW_CHILD                          As Long = &H5

Private Const GWL_STYLE                         As Long = (-16)
Private Const GWL_EXSTYLE                       As Long = (-20)

Private Const HWND_TOP                          As Long = 0
Private Const HWND_BOTTOM                       As Long = 1
Private Const HWND_TOPMOST                      As Long = -1
Private Const HWND_NOTOPMOST                    As Long = -2

Private Const ILD_NORMAL                        As Long = &H0
Private Const ILD_TRANSPARENT                   As Long = &H1
Private Const ILD_BLEND25                       As Long = &H2
Private Const ILD_FOCUS                         As Long = &H4
Private Const ILD_SELECTED                      As Long = &H8

Private Const ILC_COLOR                         As Long = &H0
Private Const ILC_MASK                          As Long = &H1
Private Const ILC_COLOR4                        As Long = &H4
Private Const ILC_COLOR8                        As Long = &H8
Private Const ILC_COLOR16                       As Long = &H10
Private Const ILC_COLOR24                       As Long = &H18
Private Const ILC_COLOR32                       As Long = &H20
Private Const ILC_COLORDDB                      As Long = &HFE

Private Const IMAGE_BITMAP                      As Long = &H0
Private Const IMAGE_ICON                        As Long = &H1
Private Const IMAGE_CURSOR                      As Long = &H2

Private Const LB_ADDSTRING                      As Long = &H180
Private Const LB_INSERTSTRING                   As Long = &H181
Private Const LB_DELETESTRING                   As Long = &H182
Private Const LB_SELITEMRANGEEX                 As Long = &H183
Private Const LB_RESETCONTENT                   As Long = &H184
Private Const LB_SETSEL                         As Long = &H185
Private Const LB_SETCURSEL                      As Long = &H186
Private Const LB_GETSEL                         As Long = &H187
Private Const LB_GETCURSEL                      As Long = &H188
Private Const LB_GETTEXT                        As Long = &H189
Private Const LB_GETTEXTLEN                     As Long = &H18A
Private Const LB_GETCOUNT                       As Long = &H18B
Private Const LB_SELECTSTRING                   As Long = &H18C
Private Const LB_DIR                            As Long = &H18D
Private Const LB_GETTOPINDEX                    As Long = &H18E
Private Const LB_FINDSTRING                     As Long = &H18F
Private Const LB_GETSELCOUNT                    As Long = &H190
Private Const LB_GETSELITEMS                    As Long = &H191
Private Const LB_SETTABSTOPS                    As Long = &H192
Private Const LB_GETHORIZONTALEXTENT            As Long = &H193
Private Const LB_SETHORIZONTALEXTENT            As Long = &H194
Private Const LB_SETCOLUMNWIDTH                 As Long = &H195
Private Const LB_ADDFILE                        As Long = &H196
Private Const LB_SETTOPINDEX                    As Long = &H197
Private Const LB_GETITEMRECT                    As Long = &H198
Private Const LB_GETITEMDATA                    As Long = &H199
Private Const LB_SETITEMDATA                    As Long = &H19A
Private Const LB_SELITEMRANGE                   As Long = &H19B
Private Const LB_SETANCHORINDEX                 As Long = &H19C
Private Const LB_GETANCHORINDEX                 As Long = &H19D
Private Const LB_SETCARETINDEX                  As Long = &H19E
Private Const LB_GETCARETINDEX                  As Long = &H19F
Private Const LB_SETITEMHEIGHT                  As Long = &H1A0
Private Const LB_GETITEMHEIGHT                  As Long = &H1A1
Private Const LB_FINDSTRINGEXACT                As Long = &H1A2
Private Const LB_SETLOCALE                      As Long = &H1A5
Private Const LB_GETLOCALE                      As Long = &H1A6
Private Const LB_SETCOUNT                       As Long = &H1A7

Private Const LBN_SELCHANGE                     As Long = &H1
Private Const LBN_DBLCLK                        As Long = &H2
Private Const LBN_SELCANCEL                     As Long = &H3
Private Const LBN_SETFOCUS                      As Long = &H4
Private Const LBN_KILLFOCUS                     As Long = &H5

Private Const LBS_NOTIFY                        As Long = &H1
Private Const LBS_SORT                          As Long = &H2
Private Const LBS_NOREDRAW                      As Long = &H4
Private Const LBS_MULTIPLESEL                   As Long = &H8
Private Const LBS_OWNERDRAWFIXED                As Long = &H10
Private Const LBS_OWNERDRAWVARIABLE             As Long = &H20
Private Const LBS_HASSTRINGS                    As Long = &H40
Private Const LBS_USETABSTOPS                   As Long = &H80
Private Const LBS_NOINTEGRALHEIGHT              As Long = &H100
Private Const LBS_MULTICOLUMN                   As Long = &H200
Private Const LBS_WANTKEYBOARDINPUT             As Long = &H400
Private Const LBS_EXTENDEDSEL                   As Long = &H800
Private Const LBS_DISABLENOSCROLL               As Long = &H1000
Private Const LBS_NODATA                        As Long = &H2000
Private Const LBS_NOSEL                         As Long = &H4000

Private Const LF_ANTIALIASED_QUALITY            As Long = 4
Private Const LOGPIXELSY                        As Long = 90

Private Const H_MAX                             As Long = &HFFFF + 1
Private Const NM_FIRST                          As Long = H_MAX
Private Const NM_CLICK                          As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK                         As Long = (NM_FIRST - 3)
Private Const NM_RETURN                         As Long = (NM_FIRST - 4)
Private Const NM_RCLICK                         As Long = (NM_FIRST - 5)
Private Const NM_KILLFOCUS                      As Long = (NM_FIRST - 8)
Private Const NM_CUSTOMDRAW                     As Long = (NM_FIRST - 12)
Private Const ODS_CHECKED                       As Long = &H8
Private Const ODS_DISABLED                      As Long = &H4
Private Const ODS_GRAYED                        As Long = &H2
Private Const ODS_SELECTED                      As Long = &H1
Private Const ODS_FOCUS                         As Long = &H10
Private Const ODS_COMBOBOXCLICK                 As Long = &H1011
Private Const ODS_COMBOBOXEDIT                  As Long = &H1000

Private Const PS_SOLID                          As Long = 0

Private Const SS_LEFT                           As Long = &H0
Private Const SS_CENTER                         As Long = &H1
Private Const SS_RIGHT                          As Long = &H2
Private Const SS_ICON                           As Long = &H3
Private Const SS_BLACKRECT                      As Long = &H4
Private Const SS_GRAYRECT                       As Long = &H5
Private Const SS_WHITERECT                      As Long = &H6
Private Const SS_BLACKFRAME                     As Long = &H7
Private Const SS_GRAYFRAME                      As Long = &H8
Private Const SS_WHITEFRAME                     As Long = &H9
Private Const SS_USERITEM                       As Long = &HA
Private Const SS_SIMPLE                         As Long = &HB
Private Const SS_LEFTNOWORDWRAP                 As Long = &HC
Private Const SS_OWNERDRAW                      As Long = &HD
Private Const SS_BITMAP                         As Long = &HE
Private Const SS_ENHMETAFILE                    As Long = &HF
Private Const SS_ETCHEDHORZ                     As Long = &H10
Private Const SS_ETCHEDVERT                     As Long = &H11
Private Const SS_ETCHEDFRAME                    As Long = &H12
Private Const SS_TYPEMASK                       As Long = &H1F
Private Const SS_NOPREFIX                       As Long = &H80
Private Const SS_NOTIFY                         As Long = &H100
Private Const SS_CENTERIMAGE                    As Long = &H200
Private Const SS_RIGHTJUST                      As Long = &H400
Private Const SS_REALSIZEIMAGE                  As Long = &H800
Private Const SS_SUNKEN                         As Long = &H1000
Private Const SS_ENDELLIPSIS                    As Long = &H4000
Private Const SS_PATHELLIPSIS                   As Long = &H8000
Private Const SS_WORDELLIPSIS                   As Long = &HC000
Private Const SS_ELLIPSISMASK                   As Long = &HC000

Private Const STM_SETICON                       As Long = &H170
Private Const STM_GETICON                       As Long = &H171
Private Const STM_SETIMAGE                      As Long = &H172
Private Const STM_GETIMAGE                      As Long = &H173

Private Const STN_CLICKED                       As Long = 0
Private Const STN_DBLCLK                        As Long = 1
Private Const STN_ENABLE                        As Long = 2
Private Const STN_DISABLE                       As Long = 3

Private Const SW_HIDE                           As Long = &H0
Private Const SW_NORMAL                         As Long = &H1
Private Const SW_SHOW                           As Long = &H5

Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_NOREDRAW                      As Long = &H8
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_HIDEWINDOW                    As Long = &H80
Private Const SWP_NOCOPYBITS                    As Long = &H100
Private Const SWP_NOOWNERZORDER                 As Long = &H200
Private Const SWP_NOSENDCHANGING                As Long = &H400

Private Const WS_OVERLAPPED                     As Long = &H0
Private Const WS_MAXIMIZEBOX                    As Long = &H10000
Private Const WS_TABSTOP                        As Long = &H10000
Private Const WS_MINIMIZEBOX                    As Long = &H20000
Private Const WS_GROUP                          As Long = &H20000
Private Const WS_THICKFRAME                     As Long = &H40000
Private Const WS_SYSMENU                        As Long = &H80000
Private Const WS_HSCROLL                        As Long = &H100000
Private Const WS_VSCROLL                        As Long = &H200000
Private Const WS_DLGFRAME                       As Long = &H400000
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_CAPTION                        As Long = &HC00000
Private Const WS_MAXIMIZE                       As Long = &H1000000
Private Const WS_CLIPCHILDREN                   As Long = &H2000000
Private Const WS_CLIPSIBLINGS                   As Long = &H4000000
Private Const WS_DISABLED                       As Long = &H8000000
Private Const WS_VISIBLE                        As Long = &H10000000
Private Const WS_MINIMIZE                       As Long = &H20000000
Private Const WS_CHILD                          As Long = &H40000000
Private Const WS_POPUP                          As Long = &H80000000

Private Const WS_EX_LTRREADING                  As Long = &H0
Private Const WS_EX_LEFT                        As Long = &H0
Private Const WS_EX_RIGHTSCROLLBAR              As Long = &H0
Private Const WS_EX_DLGMODALFRAME               As Long = &H1
Private Const WS_EX_NOPARENTNOTIFY              As Long = &H4
Private Const WS_EX_TOPMOST                     As Long = &H8
Private Const WS_EX_ACCEPTFILES                 As Long = &H10
Private Const WS_EX_TRANSPARENT                 As Long = &H20
Private Const WS_EX_MDICHILD                    As Long = &H40
Private Const WS_EX_TOOLWINDOW                  As Long = &H80
Private Const WS_EX_WINDOWEDGE                  As Long = &H100
Private Const WS_EX_CLIENTEDGE                  As Long = &H200
Private Const WS_EX_CONTEXTHELP                 As Long = &H400
Private Const WS_EX_RIGHT                       As Long = &H1000
Private Const WS_EX_RTLREADING                  As Long = &H2000
Private Const WS_EX_LEFTSCROLLBAR               As Long = &H4000
Private Const WS_EX_CONTROLPARENT               As Long = &H10000
Private Const WS_EX_STATICEDGE                  As Long = &H20000
Private Const WS_EX_APPWINDOW                   As Long = &H40000

Private Const VK_LBUTTON                        As Long = &H1
Private Const VK_RBUTTON                        As Long = &H2
Private Const VK_TAB                            As Long = &H9
Private Const VK_ENTER                          As Long = &HD
Private Const VK_CONTROL                        As Long = &H11
Private Const VK_LEFT                           As Long = &H25
Private Const VK_RIGHT                          As Long = &H27
Private Const VK_ESCAPE                         As Long = &H1B
Private Const VK_SPACE                          As Long = &H20
Private Const VK_UP                             As Long = &H26
Private Const VK_DOWN                           As Long = &H28
Private Const VK_UCASEA                         As Long = &H41
Private Const VK_LCASEA                         As Long = &H61

Private Const VER_PLATFORM_WIN32_NT             As Long = 2


Public Enum EITImageType
    eitBitmap = 0&
    eitIcon = 1&
    eitCursor = 2&
    eitMetafile = 3&
End Enum

Public Enum ECCThemeStyle
    eccAzure = 0&
    eccClassic = 1&
    eccGloss = 2&
    eccMetal = 3&
    eccXp = 4&
End Enum

Public Enum ECBSBorderStyle
    ecbsNone = 0&
    ecbsLine = 1&
    ecbsThin = 2&
    ecbsThick = 3&
    ecbs3D = 4&
End Enum

Public Enum ECSControlStyle
    ecsCheckBox = 1&
    ecsComboDropDown = 2&
    ecsComboDropList = 3&
    ecsComboSimple = 4&
    ecsCommandButton = 5&
    ecsImageCombo = 6&
    ecsImageListBox = 7&
    ecsLabel = 8&
    ecsListBox = 9&
    ecsListBoxExtended = 10&
    ecsListBoxMultiSelect = 11&
    ecsOptionButton = 12&
    ecsPictureBox = 13&
    ecsTextBox = 14&
End Enum

Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum


Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                      As Long
    dwFlags                                     As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
End Type

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
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

Private Type COMBOBOXINFO
   cbSize                                       As Long
   rcItem                                       As RECT
   rcButton                                     As RECT
   stateButton                                  As Long
   hwndCombo                                    As Long
   hwndEdit                                     As Long
   hwndList                                     As Long
End Type

Private Type PAINTSTRUCT
    hdc                                         As Long
    fErase                                      As Boolean
    rcPaint                                     As RECT
    fRestore                                    As Boolean
    fIncUpdate                                  As Boolean
    rgbReserved(32)                             As Byte
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


Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

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

Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal fEnable As Long) As Long

Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long


Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, _
                                                       CBInfo As COMBOBOXINFO) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal wCmd As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, _
                                                       ByVal hBitmap As Long, _
                                                       ByVal hBitmapMask As Long) As Long

Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, _
                                                            ByVal I As Long, _
                                                            ByVal hdcDst As Long, _
                                                            ByVal x As Long, _
                                                            ByVal y As Long, _
                                                            ByVal fStyle As Long) As Long

Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, _
                                                           ByVal hicon As Long) As Long

Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal hIml As Long, _
                                                              ByVal I As Long) As Long

Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

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

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, _
                                                 ByVal nCount As Long, _
                                                 lpObject As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateDc Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 lpDeviceName As Any, _
                                                                 lpOutput As Any, _
                                                                 lpInitData As Any) As Long

Private Declare Function GetWindowTextA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal lpString As String, _
                                                      ByVal cch As Long) As Long

Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal lpString As Long, _
                                                      ByVal cch As Long) As Long

Private Declare Function GetWindowTextLengthA Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowTextLengthW Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowTextA Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As String) As Long

Private Declare Function SetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, _
                                                          ByVal lpString As Long) As Long

Private Declare Function EraseRect Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                        lpRect As RECT, _
                                                                        ByVal bErase As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    ByVal lParam As Long) As Long

Private Declare Function PostMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    ByVal lParam As Long) As Long

Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                           ByVal pszSubAppName As Long, _
                                                           ByVal pszSubIdList As Long) As Long

Private Declare Function GetFocus Lib "user32" () As Long

Public Event Click()
Public Event LostFocus()
Public Event ItemChange(ByVal lItem As Long)

Private m_bTrackUser32                          As Boolean
Private m_bPainting                             As Boolean
Private m_bIsNt                                 As Boolean
Private m_bIsXp                                 As Boolean
Private m_bAutoBackColor                        As Boolean
Private m_bLocked                               As Boolean
Private m_bChecked                              As Boolean
Private m_bMouseOver                            As Boolean
Private m_bAutoSize                             As Boolean
Private m_bLabelTransparent                     As Boolean
Private m_bCommandPushButton                    As Boolean
Private m_bCommandPushed                        As Boolean
Private m_bEnabled                              As Boolean
Private m_bVisible                              As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bNoScrollbar                          As Boolean
Private m_bFontRightLeading                     As Boolean
Private m_ImlCommandHnd                         As Long
Private m_lItemIndex                            As Long
Private m_ImlListBoxHnd                         As Long
Private m_lhButtonBrush                         As Long
Private m_lParentHwnd                           As Long
Private m_lCntlHwnd                             As Long
Private m_lEditHwnd                             As Long
Private m_ImlChkBoxHnd                          As Long
Private m_lHFont                                As Long
Private m_lCmdHeight                            As Long
Private m_lCmdWidth                             As Long
Private m_lHdc()                                As Long
Private m_lBmp()                                As Long
Private m_lBmpOld()                             As Long
Private m_lBackColor                            As Long
Private m_lForeColor                            As Long
Private m_lThemeColor                           As Long
Private m_lBorderStyle                          As Long
Private m_lHiliteColor                          As Long
Private m_lListHwnd                             As Long
Private m_sName                                 As String
Private m_eImageType                            As EITImageType
Private m_lImgHnd                               As Long
Private m_lItemImage()                          As Long
Private m_lBoxColor()                           As Long
Private m_tRect                                 As RECT
Private m_oFont                                 As StdFont
Private m_pControlImg                           As StdPicture
Private m_eControlStyle                         As ECSControlStyle
Private m_eBorderStyle                          As ECBSBorderStyle
Private m_eThemeStyle                           As ECCThemeStyle
Private m_cRender                               As clsRender
Private m_cCntrlDc                              As clsStoreDc
Private m_cCntlSubclass                         As GXMSubclass


Private Sub Class_Initialize()

    InitCommonControls
    VersionCheck
    m_bTrackUser32 = FunctionExported("TrackMouseEvent", "User32")
    Set m_cCntlSubclass = New GXMSubclass
    m_lBackColor = &HFFFFFF
    m_lThemeColor = -1
    m_lHiliteColor = -1
    m_bVisible = True
    
End Sub

Private Function VersionCheck() As Boolean
'/* nt version chck

Dim tVer    As OSVERSIONINFO

    With tVer
        .dwVersionInfoSize = LenB(tVer)
        GetVersionEx tVer
        m_bIsNt = ((.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
        If (.dwMajorVersion >= 5) Then
            m_bIsXp = True
        End If
    End With
    VersionCheck = m_bIsNt

End Function


'> Properties
'>>>>>>>>>>>>>>>>>>>>>>
Public Property Get AutoBackColor() As Boolean
    AutoBackColor = m_bAutoBackColor
End Property

Public Property Let AutoBackColor(ByVal PropVal As Boolean)

Dim lColor As Long

    If Not (m_lCntlHwnd = 0) Then
        If PropVal Then
            lColor = GetBackColor
            If Not (lColor < 0) Then
                m_lBackColor = lColor
            End If
        End If
    End If
    m_bAutoBackColor = PropVal
    
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = m_bAutoSize
End Property

Public Property Let AutoSize(ByVal PropVal As Boolean)

    If Not (m_lCntlHwnd = 0) Then
        If PropVal Then
            LabelSize Text
        End If
    End If
    m_bAutoSize = PropVal
    
End Property

Public Property Get BackColor() As Long
    BackColor = m_lBackColor
End Property

Public Property Let BackColor(ByVal PropVal As Long)
    m_lBackColor = PropVal
End Property

Public Property Get Checked() As Boolean
    Checked = m_bChecked
End Property

Public Property Let Checked(ByVal PropVal As Boolean)
    m_bChecked = PropVal
    If Not (m_lCntlHwnd = 0) Then
        UpdateWindow m_lCntlHwnd
    End If
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal PropVal As Boolean)

    If Not (m_lCntlHwnd = 0) Then
        If PropVal Then
            EnableWindow m_lCntlHwnd, 1&
        Else
            EnableWindow m_lCntlHwnd, 0&
        End If
    End If
    m_bEnabled = PropVal

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

Public Property Get HiliteColor() As Long
    HiliteColor = m_lHiliteColor
End Property

Public Property Let HiliteColor(ByVal PropVal As Long)
    m_lHiliteColor = PropVal
End Property

Public Property Get HFont() As Long
    HFont = m_lHFont
End Property

Public Property Let HFont(ByVal PropVal As Long)

    If Not (m_lCntlHwnd = 0) Then
        If Not (PropVal = 0) Then
            m_lHFont = PropVal
            If Not (m_eControlStyle = ecsCommandButton) Then
                If m_bIsNt Then
                    SendMessageLongW m_lCntlHwnd, WM_SETFONT, m_lHFont, 1&
                Else
                    SendMessageLongA m_lCntlHwnd, WM_SETFONT, m_lHFont, 1&
                End If
            End If
        End If
    End If

End Property

Public Property Get hwnd() As Long
    hwnd = m_lCntlHwnd
End Property

Public Property Get LabelTransparent() As Boolean
    LabelTransparent = m_bLabelTransparent
End Property

Public Property Let LabelTransparent(ByVal PropVal As Boolean)
    m_bLabelTransparent = PropVal
End Property

Public Property Get Locked() As Boolean
    Locked = m_bLocked
End Property

Public Property Let Locked(ByVal PropVal As Boolean)

    If Not (m_lCntlHwnd = 0) Then
        SendMessageLongA m_lCntlHwnd, EM_SETREADONLY, Abs(PropVal), 0&
        m_bLocked = PropVal
    End If
    
End Property

Public Property Get Name() As String
    Name = m_sName
End Property

Public Property Let Name(ByVal PropVal As String)
    m_sName = PropVal
End Property

Public Property Get Text() As String

Dim lLen    As Long
Dim sText   As String

    
    Select Case m_eControlStyle
    Case ecsCheckBox, ecsLabel, ecsOptionButton, ecsCommandButton, ecsPictureBox, ecsTextBox
        If Not (m_lCntlHwnd = 0) Then
            If m_bIsNt Then
                lLen = GetWindowTextLengthW(m_lCntlHwnd) + 1
                sText = String(lLen, Chr$(0))
                GetWindowTextW m_lCntlHwnd, StrPtr(sText), lLen
                Text = left$(sText, (lLen - 1))
            Else
                lLen = GetWindowTextLengthA(m_lCntlHwnd) + 1
                sText = String(lLen, Chr$(0))
                GetWindowTextA m_lCntlHwnd, sText, lLen
                Text = sText
            End If
        End If
    Case ecsComboDropDown, ecsImageCombo
        If Not (m_lEditHwnd = 0) Then
            If m_bIsNt Then
                lLen = GetWindowTextLengthW(m_lEditHwnd) + 1
                sText = String(lLen, Chr$(0))
                GetWindowTextW m_lEditHwnd, StrPtr(sText), lLen
                Text = left$(sText, (lLen - 1))
            Else
                lLen = GetWindowTextLengthA(m_lEditHwnd) + 1
                sText = String(lLen, Chr$(0))
                GetWindowTextA m_lEditHwnd, sText, lLen
                Text = sText
            End If
            End If
    End Select

End Property

Public Property Let Text(ByVal PropVal As String)

    
    Select Case m_eControlStyle
    Case ecsCheckBox, ecsLabel, ecsOptionButton, ecsCommandButton, ecsPictureBox, ecsTextBox
        If Not (m_lCntlHwnd = 0) Then
            If m_bIsNt Then
                SetWindowTextW m_lCntlHwnd, StrPtr(PropVal)
            Else
                SetWindowTextA m_lCntlHwnd, PropVal
            End If
        End If
    Case ecsComboDropDown, ecsImageCombo
        If Not (m_lEditHwnd = 0) Then
            If m_bIsNt Then
                SetWindowTextW m_lEditHwnd, StrPtr(PropVal)
            Else
                SetWindowTextA m_lEditHwnd, PropVal
            End If
        End If
    End Select

End Property

Public Property Get ThemeColor() As Long
    ThemeColor = m_lThemeColor
End Property

Public Property Let ThemeColor(ByVal PropVal As Long)
    m_lThemeColor = PropVal
End Property

Public Property Get ThemeStyle() As ECCThemeStyle
    ThemeStyle = m_eThemeStyle
End Property

Public Property Let ThemeStyle(ByVal PropVal As ECCThemeStyle)
    m_eThemeStyle = PropVal
End Property

Public Property Get UseUnicode() As Boolean
    UseUnicode = m_bUseUnicode
End Property

Public Property Let UseUnicode(ByVal PropVal As Boolean)
    m_bUseUnicode = PropVal
End Property


'> Constructors
'>>>>>>>>>>>>>>>>>>>>>>
Public Sub Create(ByVal lParentHwnd As Long, _
                  ByVal lX As Long, _
                  ByVal lY As Long, _
                  ByVal lWidth As Long, _
                  ByVal lHeight As Long, _
                  ByVal eCtrlStyle As ECSControlStyle, _
                  Optional ByVal lForeColor As Long = -1, _
                  Optional ByVal lBackColor As Long = -1, _
                  Optional ByVal sCaption As String, _
                  Optional ByVal oFont As StdFont)

'/* create control window

    If Not (lParentHwnd = 0) Then
        m_lParentHwnd = lParentHwnd
        m_eControlStyle = eCtrlStyle
        If Not (lForeColor = -1) Then
            m_lForeColor = lForeColor
        End If
        If Not (lBackColor = -1) Then
            m_lBackColor = lBackColor
        End If
        
        If Not (oFont Is Nothing) Then
            Set m_oFont = oFont
        End If
        With m_tRect
            .left = lX
            .Right = lX + lWidth
            .top = lY
            .Bottom = lY + lHeight
        End With
        Initialize
        If (LenB(Text) = 0) Then
            If (LenB(sCaption) > 0) Then
                Text = sCaption
            End If
        End If
    End If

End Sub

Private Sub Initialize()
'/* initialize api window control

Dim lwStyle As Long
Dim sStyle  As String
Dim tCBInfo As COMBOBOXINFO

    '/* window styles
    Select Case m_eControlStyle
    Case ecsCheckBox
        lwStyle = WS_CHILD Or BS_OWNERDRAW
        sStyle = "BUTTON"
    Case ecsComboDropDown
        sStyle = "COMBOBOX"
        lwStyle = WS_CHILD Or WS_VSCROLL Or CBS_HASSTRINGS Or CBS_DROPDOWN
    Case ecsComboDropList
        sStyle = "COMBOBOX"
        lwStyle = WS_CHILD Or WS_VSCROLL Or CBS_HASSTRINGS Or CBS_DROPDOWNLIST
    Case ecsComboSimple
        sStyle = "COMBOBOX"
        lwStyle = WS_CHILD Or WS_VSCROLL Or CBS_HASSTRINGS Or CBS_SIMPLE
    Case ecsCommandButton
        lwStyle = WS_CHILD Or BS_PUSHBUTTON Or BS_OWNERDRAW
        sStyle = "BUTTON"
    Case ecsImageCombo
        sStyle = "COMBOBOX"
        lwStyle = WS_CHILD Or WS_VSCROLL Or CBS_HASSTRINGS Or CBS_DROPDOWN Or CBS_OWNERDRAWVARIABLE
    Case ecsImageListBox
        lwStyle = WS_CHILD Or WS_VSCROLL Or LBS_NOINTEGRALHEIGHT Or LBS_HASSTRINGS Or LBS_USETABSTOPS Or LBS_OWNERDRAWFIXED
        sStyle = "LISTBOX"
    Case ecsLabel
        sStyle = "STATIC"
        lwStyle = WS_CHILD Or SS_LEFTNOWORDWRAP Or SS_NOTIFY Or SS_OWNERDRAW
    Case ecsListBox
        lwStyle = WS_CHILD Or WS_VSCROLL Or LBS_HASSTRINGS Or LBS_NOINTEGRALHEIGHT Or LBS_NOTIFY Or LBS_OWNERDRAWVARIABLE
        sStyle = "LISTBOX"
    Case ecsListBoxExtended
        lwStyle = WS_CHILD Or WS_VSCROLL Or LBS_HASSTRINGS Or LBS_NOINTEGRALHEIGHT Or LBS_EXTENDEDSEL Or LBS_NOTIFY Or LBS_OWNERDRAWVARIABLE
        sStyle = "LISTBOX"
    Case ecsListBoxMultiSelect
        lwStyle = WS_CHILD Or WS_VSCROLL Or LBS_HASSTRINGS Or LBS_NOINTEGRALHEIGHT Or LBS_MULTIPLESEL Or LBS_NOTIFY Or LBS_OWNERDRAWVARIABLE
        sStyle = "LISTBOX"
    Case ecsOptionButton
        lwStyle = WS_CHILD Or BS_OWNERDRAW
        sStyle = "BUTTON"
    Case ecsPictureBox
        sStyle = "STATIC"
        lwStyle = WS_CHILD Or SS_NOTIFY Or SS_CENTERIMAGE Or SS_WHITEFRAME
    Case ecsTextBox
        sStyle = "EDIT"
        lwStyle = WS_CHILD Or WS_VSCROLL Or ES_MULTILINE Or ES_AUTOVSCROLL Or ES_LEFT
    End Select
    If m_bNoScrollbar Then
        lwStyle = lwStyle And Not WS_VSCROLL
    End If
    
    Select Case m_eBorderStyle
    Case ecbs3D, ecbsLine
        lwStyle = lwStyle Or WS_BORDER
    Case Else
        lwStyle = lwStyle And Not WS_BORDER
    End Select
    
    Select Case m_eControlStyle
    Case ecsComboDropDown, ecsComboDropList, ecsImageCombo
        Select Case m_eBorderStyle
        Case ecbs3D, ecbsThick, ecbsLine
            BorderStyle ecbsThin
        End Select
    End Select
    If m_bFontRightLeading Then
        m_lBorderStyle = m_lBorderStyle Or WS_EX_RTLREADING
    End If
    '/* create the window
    With m_tRect
        If m_bIsNt Then
            m_lCntlHwnd = CreateWindowExW(m_lBorderStyle, StrPtr(sStyle), StrPtr(m_sName), lwStyle, .left, .top, (.Right - .left), (.Bottom - .top), m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
        Else
            m_lCntlHwnd = CreateWindowExA(m_lBorderStyle, sStyle, m_sName, lwStyle, .left, .top, (.Right - .left), (.Bottom - .top), m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
        End If
    End With
    '/* remove theming
    If m_bIsXp Then
        Select Case m_eControlStyle
        Case ecsCheckBox, ecsCommandButton, ecsLabel, ecsOptionButton, ecsPictureBox, ecsTextBox
            SetWindowTheme m_lCntlHwnd, StrPtr(" "), StrPtr(" ")
        End Select
    End If
    Select Case m_eControlStyle
        Case ecsComboDropDown, ecsImageCombo
        m_lEditHwnd = GetWindow(m_lCntlHwnd, GW_CHILD)
    End Select

    Select Case m_eControlStyle
    Case ecsImageListBox
        InitListBoxIml
        DefineTabStop
    Case ecsImageCombo
        InitListBoxIml
    Case ecsLabel
        If m_bAutoSize Then
            LabelSize Text
        End If
    End Select

    InitSkin
    AttachMessages
    If m_bVisible Then
        Show
    End If
    
    Select Case m_eControlStyle
    Case ecsImageCombo
        MoveEditBox
        ComboInfo m_lCntlHwnd, tCBInfo
        m_lListHwnd = tCBInfo.hwndList
        AttachList
    Case ecsComboDropDown
        ComboInfo m_lCntlHwnd, tCBInfo
        m_lListHwnd = tCBInfo.hwndList
        AttachList
    End Select
    If m_bUseUnicode Then
        SetUnicode True
    End If
    
End Sub

Private Sub AttachList()

    With m_cCntlSubclass
        If Not (m_lListHwnd = 0) Then
            .Subclass m_lListHwnd, Me
            .AddMessage m_lListHwnd, WM_LBUTTONUP, MSG_BEFORE
        End If
    End With
                
End Sub

Private Sub InitSkin()

On Error GoTo Handler

    DefaultFont
    Select Case m_eControlStyle
    Case ecsCheckBox, ecsOptionButton
        CheckBoxLoadImages m_eThemeStyle

    Case ecsComboDropDown, ecsImageCombo, ecsComboDropList
        ComboLoadImage m_eThemeStyle
        
    Case ecsCommandButton
        ButtonLoadImage m_eThemeStyle
        CreateBackbuffer
    End Select
    
Handler:

End Sub

'> Global
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub DefineTabStop()

Dim lTab() As Long

    ReDim lTab(0)
    lTab(0) = 10
    If Not (m_lCntlHwnd = 0) Then
        SendMessageA m_lCntlHwnd, LB_SETTABSTOPS, 1&, lTab(0)
    End If

End Sub

Public Sub AddItem(ByVal sItem As String, _
                   Optional ByVal lImageIdx As Long = -1, _
                   Optional ByVal lBoxColor As Long = -1)

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsComboDropDown, ecsComboDropList, ecsComboSimple
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, CB_ADDSTRING, 0&, StrPtr(sItem)
            Else
                SendMessageA m_lCntlHwnd, CB_ADDSTRING, 0&, sItem
            End If
        Case ecsImageCombo
            ReDim Preserve m_lBoxColor(0 To m_lItemIndex)
            m_lBoxColor(m_lItemIndex) = lBoxColor
            ReDim Preserve m_lItemImage(0 To m_lItemIndex)
            m_lItemImage(m_lItemIndex) = lImageIdx
            m_lItemIndex = m_lItemIndex + 1
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, CB_ADDSTRING, 0&, StrPtr(sItem)
            Else
                SendMessageA m_lCntlHwnd, CB_ADDSTRING, 0&, sItem
            End If
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, LB_ADDSTRING, 0&, StrPtr(sItem)
            Else
                SendMessageA m_lCntlHwnd, LB_ADDSTRING, 0&, sItem
            End If
        Case ecsImageListBox
            ReDim Preserve m_lItemImage(0 To m_lItemIndex)
            m_lItemImage(m_lItemIndex) = lImageIdx
            m_lItemIndex = m_lItemIndex + 1
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, LB_ADDSTRING, 0&, StrPtr(sItem)
            Else
                SendMessageA m_lCntlHwnd, LB_ADDSTRING, 0&, sItem
            End If
        End Select
    End If

End Sub

Public Sub Clear()

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            SendMessageLongA m_lCntlHwnd, LB_RESETCONTENT, 0&, 0&
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            SendMessageLongA m_lCntlHwnd, CB_RESETCONTENT, 0&, 0&
        End Select
    End If

End Sub

Public Sub AddToGroup(ByVal bAddItem As Boolean)

    If Not (m_lCntlHwnd = 0) Then
        If bAddItem Then
            SetStyle WS_GROUP Or WS_TABSTOP, 0
        Else
            SetStyle 0, WS_GROUP Or WS_TABSTOP
        End If
    End If

End Sub

Public Sub InsertItem(ByVal sItem As String, _
                      ByVal lIndex As Long)

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, LB_INSERTSTRING, lIndex, StrPtr(sItem)
            Else
                SendMessageA m_lCntlHwnd, LB_INSERTSTRING, lIndex, sItem
            End If
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, CB_INSERTSTRING, lIndex, StrPtr(sItem)
            Else
                SendMessageA m_lCntlHwnd, CB_INSERTSTRING, lIndex, sItem
            End If
        End Select
    End If

End Sub

Public Sub ItemHeight(ByVal lIndex As Long, _
                      ByVal lHeight As Long)

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            SendMessageLongA m_lCntlHwnd, LB_SETITEMHEIGHT, lIndex, lHeight
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            SendMessageLongA m_lCntlHwnd, CB_SETITEMHEIGHT, lIndex, lHeight
        End Select
    End If

End Sub

Public Property Get ListCount() As Long

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            ListCount = SendMessageLongA(m_lCntlHwnd, LB_GETCOUNT, 0&, 0&)
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            ListCount = SendMessageLongA(m_lCntlHwnd, CB_GETCOUNT, 0&, 0&)
        End Select
    End If

End Property

Public Property Get ListIndex() As Long

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            ListIndex = SendMessageLongA(m_lCntlHwnd, LB_GETCURSEL, 0&, 0&)
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            ListIndex = SendMessageLongA(m_lCntlHwnd, CB_GETCURSEL, 0&, 0&)
        End Select
    End If

End Property

Public Property Let ListIndex(ByVal lIndex As Long)

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            SendMessageLongA m_lCntlHwnd, LB_SETCURSEL, lIndex, 0&
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            SendMessageLongA m_lCntlHwnd, CB_SETCURSEL, lIndex, 0&
        End Select
    End If

End Property

Public Function ListItem(ByVal lItem As Long) As Long

Dim lPtr As Long

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            lPtr = SendMessageLongA(m_lCntlHwnd, LB_GETITEMDATA, lItem, 0&)
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            lPtr = SendMessageLongA(m_lCntlHwnd, CB_GETITEMDATA, lItem, 0&)
        End Select
    End If
    
    ListItem = lPtr
    
End Function

Public Sub RemoveItem(ByVal lIndex As Long)

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            SendMessageLongA m_lCntlHwnd, LB_DELETESTRING, lIndex, 0&
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            SendMessageLongA m_lCntlHwnd, CB_DELETESTRING, lIndex, 0&
        End Select
    End If

End Sub

Public Sub Sorted(ByVal bSorted As Boolean)

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            If bSorted Then
                SetStyle LBS_SORT, 0
            Else
                SetStyle 0, LBS_SORT
            End If
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            If bSorted Then
                SetStyle CBS_SORT, 0
            Else
                SetStyle 0, CBS_SORT
            End If
        End Select
    End If

End Sub

Public Property Get SelectedItem(ByVal lIndex As Long) As Boolean
    
    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            SelectedItem = SendMessageLongA(m_lCntlHwnd, LB_GETSEL, lIndex, 0&)
        Case ecsComboDropList
            Dim tCBInfo     As COMBOBOXINFO
            ComboInfo m_lCntlHwnd, tCBInfo
            If tCBInfo.hwndList > 0 Then
                SelectedItem = SendMessageLongA(tCBInfo.hwndList, LB_GETSEL, lIndex, 0&)
            End If
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo
            SelectedItem = SendMessageLongA(m_lCntlHwnd, CB_GETCURSEL, 0&, 0&)
        End Select
    End If

End Property

Public Property Let SelectedItem(ByVal lIndex As Long, _
                                 ByVal bSelected As Boolean)

Dim lSelect As Long

    If Not (m_lCntlHwnd = 0) Then
        lSelect = (bSelected * -1)
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            SendMessageLongA m_lCntlHwnd, LB_SETSEL, lSelect, lIndex
        Case ecsComboDropList
            Dim tCBInfo     As COMBOBOXINFO
            ComboInfo m_lCntlHwnd, tCBInfo
            If (tCBInfo.hwndList > 0) Then
                SendMessageLongA tCBInfo.hwndList, LB_SETSEL, lSelect, lIndex
            End If
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo
            SendMessageLongA m_lCntlHwnd, CB_SETCURSEL, lSelect, 0&
        End Select
    End If

End Property


'> Command Button
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub ButtonLoadImage(ByVal lIndex As Long)

Dim tBmp As BITMAP

    Select Case lIndex
    Case 0
        Set m_pControlImg = LoadResPicture("AZURE-COMMAND", vbResBitmap)
    Case 1
        Set m_pControlImg = LoadResPicture("CLASSIC-COMMAND", vbResBitmap)
    Case 2
        Set m_pControlImg = LoadResPicture("GLOSS-COMMAND", vbResBitmap)
    Case 3
        Set m_pControlImg = LoadResPicture("METAL-COMMAND", vbResBitmap)
    Case 4
        Set m_pControlImg = LoadResPicture("XPS-COMMAND", vbResBitmap)
    Case 5
        Set m_pControlImg = LoadResPicture("XPB-COMMAND", vbResBitmap)
    Case 6
        Set m_pControlImg = LoadResPicture("XPG-COMMAND", vbResBitmap)
    Case 7
        Set m_pControlImg = LoadResPicture("VISTA-COMMAND", vbResBitmap)
    End Select

    GetObjectA m_pControlImg.handle, Len(tBmp), tBmp
    With tBmp
        m_lCmdWidth = (.bmWidth / 5)
        m_lCmdHeight = .bmHeight
    End With
    
    Set m_cRender = New clsRender
    Set m_cCntrlDc = New clsStoreDc
    m_cCntrlDc.CreateFromPicture m_pControlImg
    If (m_lThemeColor > -1) Then
        m_cCntrlDc.ColorizeImage m_lThemeColor, 0.3
    End If
    Set m_pControlImg = Nothing

End Sub

Private Function CreateButtonBrush() As Long

    If Not (m_lhButtonBrush = 0) Then
        DestroyButtonBrush
    End If
    m_lhButtonBrush = CreateSolidBrush(m_lBackColor)
    CreateButtonBrush = m_lhButtonBrush

End Function

Private Function GetBackColor() As Long

Dim lHdc    As Long
Dim tRect   As RECT
Dim tPcd    As POINTAPI

    lHdc = GetDC(m_lParentHwnd)
    GetWindowRect m_lCntlHwnd, tRect
    CopyMemory tPcd, tRect, Len(tPcd)
    ScreenToClient m_lParentHwnd, tPcd
    With tPcd
        GetBackColor = GetPixel(lHdc, (.x - 1), (.y - 1))
    End With
    ReleaseDC m_lParentHwnd, lHdc

End Function

Public Property Get CommandPushButton() As Boolean
    CommandPushButton = m_bCommandPushButton
End Property

Public Property Let CommandPushButton(ByVal PropVal As Boolean)
    m_bCommandPushButton = PropVal
End Property

Public Property Get CommandPushed() As Boolean
    CommandPushed = m_bCommandPushed
End Property

Public Property Let CommandPushed(ByVal PropVal As Boolean)
    m_bCommandPushed = PropVal
End Property

Public Property Get NoScrollbar() As Boolean
    NoScrollbar = m_bNoScrollbar
End Property

Public Property Let NoScrollbar(ByVal PropVal As Boolean)
    m_bNoScrollbar = PropVal
End Property

Private Sub CommandRender(ByVal lHwnd As Long, _
                          ByVal lHdc As Long, _
                          ByVal lState As Long)

Dim lOldBmp     As Long
Dim lNewBmp     As Long
Dim lDrawDc     As Long
Dim lOffset     As Long
Dim lFlags      As Long
Dim lY          As Long
Dim tRect       As RECT

    If IsWindowEnabled(lHwnd) = 0 Then
        lOffset = 4
    ElseIf m_bCommandPushButton Then
        If lState = 17 Then
            m_bCommandPushed = Not m_bCommandPushed
        End If
        If m_bCommandPushed Then
            lOffset = 1
        Else
            lOffset = 0
        End If
    Else
        Select Case lState
        Case 0
            lOffset = 0
        Case 4
            lOffset = 2
        Case 15
            lOffset = 3
        Case 16
            lOffset = 0
        Case 17
            lOffset = 1
        Case 18
            lOffset = 4
        Case Else
            lOffset = 0
        End Select
    End If
    
    GetClientRect lHwnd, tRect
    lDrawDc = CreateCompatibleDC(lHdc)
    With tRect
        lNewBmp = CreateCompatibleBitmap(lHdc, .Right, .Bottom)
    End With
    lOldBmp = SelectObject(lDrawDc, lNewBmp)

    With tRect
        '/* left side
        m_cRender.Stretch lDrawDc, .left, .top, 3, .Bottom, m_lHdc(lOffset), 0, 0, 3, m_lCmdHeight, SRCCOPY
        '/* middle
        m_cRender.Stretch lDrawDc, .left + 3, 0, (.Right - 6), .Bottom, m_lHdc(lOffset), 3, 0, (m_lCmdWidth - 6), m_lCmdHeight, SRCCOPY
        '/* right side
        m_cRender.Stretch lDrawDc, .Right - 3, 0, 3, .Bottom, m_lHdc(lOffset), (m_lCmdWidth - 3), 0, 3, m_lCmdHeight, SRCCOPY
        '/* draw
        m_cRender.Blit lHdc, 0, 0, .Right, .Bottom, lDrawDc, 0, 0, SRCCOPY
    End With
    
    If Not (m_ImlCommandHnd = 0) Then
        lY = ((tRect.Bottom - 16) / 2) + 1
        ImageList_Draw m_ImlCommandHnd, 0, lHdc, 2, lY, ILD_TRANSPARENT
        tRect.left = 20
    End If
    
    lFlags = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
    If (m_lHiliteColor > -1) Then
        If (lOffset = 1) Then
            SetTextColor lHdc, m_lHiliteColor
        Else
            SetTextColor lHdc, m_lForeColor
        End If
    End If
    DrawText lHdc, lFlags, Text, tRect
    
    '/* cleanup
    SelectObject lDrawDc, lOldBmp
    DeleteObject lNewBmp
    DeleteDC lDrawDc

End Sub

Private Sub DestroyButtonBrush()

    If Not (m_lhButtonBrush = 0) Then
        DeleteObject m_lhButtonBrush
        m_lhButtonBrush = 0
    End If
    
End Sub


'> ComboBox
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Property Get ComboDropped() As Boolean

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsComboDropDown, ecsComboDropList, ecsComboSimple, ecsImageCombo
            ComboDropped = (SendMessageLongA(m_lCntlHwnd, CB_GETDROPPEDSTATE, 0&, 0&) <> 0)
        End Select
    End If

End Property

Private Sub ComboInfo(ByVal lHwnd As Long, _
                      ByRef tCBInfo As COMBOBOXINFO)

    tCBInfo.cbSize = Len(tCBInfo)
    GetComboBoxInfo lHwnd, tCBInfo

End Sub

Private Sub ComboLoadImage(ByVal lIndex As Long)

    Select Case lIndex
    Case 0
        Set m_pControlImg = LoadResPicture("AZURE-COMBO", vbResBitmap)
    Case 1
        Set m_pControlImg = LoadResPicture("CLASSIC-COMBO", vbResBitmap)
    Case 2
        Set m_pControlImg = LoadResPicture("GLOSS-COMBO", vbResBitmap)
    Case 3
        Set m_pControlImg = LoadResPicture("METAL-COMBO", vbResBitmap)
    Case 4
        Set m_pControlImg = LoadResPicture("XPS-COMBO", vbResBitmap)
    Case 5
        Set m_pControlImg = LoadResPicture("XPB-COMBO", vbResBitmap)
    Case 6
        Set m_pControlImg = LoadResPicture("XPG-COMBO", vbResBitmap)
    Case 7
        Set m_pControlImg = LoadResPicture("VISTA-COMBO", vbResBitmap)
    End Select

    Set m_cRender = New clsRender
    Set m_cCntrlDc = New clsStoreDc
    m_cCntrlDc.CreateFromPicture m_pControlImg
    If (m_lThemeColor > -1) Then
        m_cCntrlDc.ColorizeImage m_lThemeColor, 0.3
    End If
    Set m_pControlImg = Nothing
        
End Sub

Public Function ExtendedUI(ByVal bExtended As Boolean)

    If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsComboDropDown, ecsComboDropList, ecsComboSimple, ecsImageCombo
            SendMessageLongA m_lCntlHwnd, CB_SETEXTENDEDUI, Abs(bExtended), 0&
        End Select
    End If

End Function

Private Sub ComboRender(ByVal lHwnd As Long, _
                        ByVal lHdc As Long)

Dim lOffset     As Long
Dim lhBrush     As Long
Dim tRect       As RECT
Dim tCBInfo     As COMBOBOXINFO

    ComboInfo lHwnd, tCBInfo
    If (tCBInfo.stateButton > 0) Then
        lOffset = 1
    End If

    With tCBInfo.rcButton
        Select Case m_eBorderStyle
        Case ecbsNone, ecbsLine
            InflateRect tCBInfo.rcButton, 1, 1
        Case ecbsThin, ecbsThick, ecbs3D
            InflateRect tCBInfo.rcButton, -1, -1
            .top = .top - 2
            .left = .left - 2
            If (m_eControlStyle = ecsComboDropList) Then
                .left = .left + 1
            End If
        End Select
        m_cRender.Stretch lHdc, .left, .top, (.Right - .left), (.Bottom - .top), m_cCntrlDc.hdc, ((m_cCntrlDc.Width / 2) * lOffset), 0, (m_cCntrlDc.Width / 2), m_cCntrlDc.Height, SRCCOPY
    End With

    GetClientRect lHwnd, tRect
    Select Case m_eControlStyle
    Case ecsComboDropList
        lhBrush = CreateSolidBrush(&H333333)
        FrameRect lHdc, tRect, lhBrush
        DeleteObject lhBrush
    End Select
    
    lhBrush = CreateSolidBrush(&HFFFFFF)
    tRect.Right = tCBInfo.rcButton.left
    FillRect lHdc, tRect, lhBrush
    DeleteObject lhBrush
    lhBrush = 0

On Error GoTo 0

End Sub

Private Sub ImageComboRender(ByVal lHdc As Long, _
                             ByVal lItemId As Long, _
                             ByRef tRect As RECT)

Dim lFlags  As Long
Dim lIndex  As Long
Dim sText   As String


On Error GoTo Handler

    lIndex = m_lItemImage(lItemId)
    With tRect
        .left = 2
        If (m_lItemImage(lItemId) > -1) Then
            ImageList_Draw m_ImlListBoxHnd, lIndex, lHdc, .left, .top, ILD_TRANSPARENT
            .left = 22
        ElseIf (m_lBoxColor(lItemId) > -1) Then
            DrawColorBox lHdc, lItemId, tRect
            .left = 20
        End If
        lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        sText = ListText(lItemId)
        DrawText lHdc, lFlags, sText, tRect
    End With

Handler:

End Sub

Private Sub ComboEditBoxRender()

Dim lHdc    As Long
Dim lItemId As Long
Dim lhBrush As Long
Dim tRect   As RECT

    lItemId = ListIndex
    If (lItemId > -1) Then
        If (m_lItemImage(lItemId) > -1) Then
            With tRect
                .left = 3
                .Right = 19
                .top = 2
                .Bottom = 18
            End With
            lhBrush = CreateSolidBrush(&HFFFFFF)
            lHdc = GetDC(m_lCntlHwnd)
            FillRect lHdc, tRect, lhBrush
            DeleteObject lhBrush
            ImageList_Draw m_ImlListBoxHnd, m_lItemImage(lItemId), lHdc, 2, 2, ILD_TRANSPARENT
            ReleaseDC m_lCntlHwnd, lHdc
        ElseIf (m_lBoxColor(lItemId) > -1) Then
            With tRect
                .left = 3
                .Right = 19
                .top = 2
                .Bottom = 18
            End With
            lhBrush = CreateSolidBrush(&HFFFFFF)
            lHdc = GetDC(m_lCntlHwnd)
            FillRect lHdc, tRect, lhBrush
            DeleteObject lhBrush
            DrawColorBox lHdc, lItemId, tRect
            ReleaseDC m_lCntlHwnd, lHdc
        End If
    End If
    
End Sub

Public Property Let ComboIndexColor(ByVal PropVal As Long)

On Error GoTo Handler

    m_lBoxColor(0) = PropVal
    SetWindowTextA m_lEditHwnd, "&H" & Hex(PropVal)
    ComboEditBoxRender
    
Handler:

End Property

Private Sub MoveEditBox()

Dim tRect As RECT

    If Not (m_lEditHwnd = 0) Then
        GetClientRect m_lEditHwnd, tRect
        InflateRect tRect, -2, -2
        With tRect
            SetWindowPos m_lEditHwnd, 0&, 25&, 4&, (.Right - 25), .Bottom, SWP_NOOWNERZORDER Or SWP_NOACTIVATE
        End With
    End If

End Sub

Private Sub DrawColorBox(ByVal lHdc As Long, _
                         ByVal lItemId As Long, _
                         ByRef tRect As RECT)

Dim lhPen       As Long
Dim lhPenOld    As Long
Dim lhBrush     As Long
Dim tPnt        As POINTAPI
Dim tRcpy       As RECT

On Error GoTo Handler

    CopyRect tRcpy, tRect
    With tRcpy
        .left = 2
        .Right = 14
        .top = .top + 2
        .Bottom = .top + 12
    End With
    With tRcpy
        MoveToEx lHdc, .left, .top, tPnt
        lhPen = CreatePen(PS_SOLID, 1, &H0)
        lhPenOld = SelectObject(lHdc, lhPen)
        LineTo lHdc, .Right - 1, .top
        LineTo lHdc, .Right - 1, .Bottom - 1
        LineTo lHdc, .left, .Bottom - 1
        LineTo lHdc, .left, .top
    End With
    
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen
    lhBrush = CreateSolidBrush(m_lBoxColor(lItemId))
    
    InflateRect tRcpy, -1, -1
    FillRect lHdc, tRcpy, lhBrush
    DeleteObject lhBrush

Handler:
    On Error GoTo 0

End Sub

'> Listbox
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Function ListText(ByVal lIndex As Long) As String

Dim lLen        As Long
Dim sBuffer     As String

   If Not (m_lCntlHwnd = 0) Then
        Select Case m_eControlStyle
        Case ecsListBox, ecsListBoxExtended, ecsListBoxMultiSelect, ecsImageListBox
            If m_bIsNt Then
                lLen = SendMessageLongW(m_lCntlHwnd, LB_GETTEXTLEN, lIndex, 0&)
            Else
                lLen = SendMessageLongA(m_lCntlHwnd, LB_GETTEXTLEN, lIndex, 0&)
            End If
            If (lLen < 1) Then
                Exit Function
            Else
                sBuffer = String$((lLen), 0)
            End If
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, LB_GETTEXT, lIndex, StrPtr(sBuffer)
            Else
                SendMessageA m_lCntlHwnd, LB_GETTEXT, lIndex, sBuffer
            End If
        Case ecsComboDropDown, ecsComboSimple, ecsImageCombo, ecsComboDropList
            If m_bIsNt Then
                lLen = SendMessageLongW(m_lCntlHwnd, CB_GETLBTEXTLEN, lIndex, 0&)
            Else
                lLen = SendMessageLongA(m_lCntlHwnd, CB_GETLBTEXTLEN, lIndex, 0&)
            End If
            If (lLen < 1) Then
                Exit Function
            Else
                sBuffer = String$((lLen), 0)
            End If
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, CB_GETLBTEXT, lIndex, StrPtr(sBuffer)
            Else
                SendMessageA m_lCntlHwnd, CB_GETLBTEXT, lIndex, sBuffer
            End If
        End Select
    End If

    ListText = sBuffer
    
End Function

Private Sub ListBoxRender(ByVal lHdc As Long, _
                          ByVal lItemId As Long, _
                          ByVal lState As Long, _
                          ByRef tRect As RECT)

Dim lFlags  As Long
Dim lIndex  As Long
Dim lhBrush As Long
Dim sText   As String
Dim tRSlt   As RECT

    With tRect
        lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        sText = ListText(lItemId)
        CopyRect tRSlt, tRect
        InflateRect tRSlt, -1, -1
        If (lState = 17) Then
            lhBrush = CreateSolidBrush(GetSysColor(vbHighlight And &H1F&))
            FillRect lHdc, tRSlt, lhBrush
            DeleteObject lhBrush
            SetTextColor lHdc, &HFFFFFF
        Else
            lhBrush = CreateSolidBrush(m_lBackColor)
            FillRect lHdc, tRect, lhBrush
            DeleteObject lhBrush
            SetTextColor lHdc, m_lForeColor
        End If
        If (m_eControlStyle = ecsImageListBox) Then
            On Error Resume Next
            lIndex = m_lItemImage(lItemId)
            On Error GoTo 0
            If (m_lItemImage(lItemId) > -1) Then
                ImageList_Draw m_ImlListBoxHnd, lIndex, lHdc, 1&, .top, ILD_TRANSPARENT
            End If
            .left = 20
            DrawText lHdc, lFlags, sText, tRect
        Else
            .left = 2
            DrawText lHdc, lFlags, sText, tRect
        End If
    End With
    
End Sub


'> CheckBox
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub CheckBoxLoadImages(ByVal lIndex As Long)

Dim lMask       As Long
Dim picEmpty    As StdPicture
Dim picPushed   As StdPicture
Dim picDisabled As StdPicture

    InitChkBoxIml
    Set m_cRender = New clsRender
    If (m_eControlStyle = ecsCheckBox) Then
        Select Case lIndex
        Case 0
            Set picEmpty = LoadResPicture("AZURE-CHKEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("AZURE-CHKPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("AZURE-CHKDISABLED", vbResBitmap)
        Case 1
            Set picEmpty = LoadResPicture("CLASSIC-CHKEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("CLASSIC-CHKPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("CLASSIC-CHKDISABLED", vbResBitmap)
        Case 2
            Set picEmpty = LoadResPicture("GLOSS-CHKEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("GLOSS-CHKPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("GLOSS-CHKDISABLED", vbResBitmap)
        Case 3
            Set picEmpty = LoadResPicture("METAL-CHKEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("METAL-CHKPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("METAL-CHKDISABLED", vbResBitmap)
        Case 4
            Set picEmpty = LoadResPicture("XP-CHKEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("XP-CHKPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("XP-CHKDISABLED", vbResBitmap)
        End Select
    ElseIf (m_eControlStyle = ecsOptionButton) Then
        Select Case lIndex
        Case 0
            Set picEmpty = LoadResPicture("AZURE-OPTEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("AZURE-OPTPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("AZURE-OPTDISABLED", vbResBitmap)
        Case 1
            Set picEmpty = LoadResPicture("CLASSIC-OPTEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("CLASSIC-OPTPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("CLASSIC-OPTDISABLED", vbResBitmap)
        Case 2
            Set picEmpty = LoadResPicture("GLOSS-OPTEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("GLOSS-OPTPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("GLOSS-OPTDISABLED", vbResBitmap)
        Case 3
            Set picEmpty = LoadResPicture("METAL-OPTEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("METAL-OPTPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("METAL-OPTDISABLED", vbResBitmap)
        Case 4
            Set picEmpty = LoadResPicture("XP-OPTEMPTY", vbResBitmap)
            Set picPushed = LoadResPicture("XP-OPTPUSHED", vbResBitmap)
            Set picDisabled = LoadResPicture("XP-OPTDISABLED", vbResBitmap)
        End Select
    End If
    
    Set m_cCntrlDc = New clsStoreDc
    m_cCntrlDc.CreateFromPicture picEmpty
    If Not (m_lThemeColor = -1) Then
        m_cCntrlDc.ColorizeImage m_lThemeColor, 0.3
    End If
    lMask = GetPixel(m_cCntrlDc.hdc, 0&, 0&)
    ImlChkBoxAddBmp m_cCntrlDc.ExtractBitmap, lMask
    Set m_cCntrlDc = Nothing
    Set picEmpty = Nothing
    
    Set m_cCntrlDc = New clsStoreDc
    m_cCntrlDc.CreateFromPicture picPushed
    If Not (m_lThemeColor = -1) Then
        m_cCntrlDc.ColorizeImage m_lThemeColor, 0.3
    End If
    lMask = GetPixel(m_cCntrlDc.hdc, 0&, 0&)
    ImlChkBoxAddBmp m_cCntrlDc.ExtractBitmap, lMask
    Set m_cCntrlDc = Nothing
    Set picPushed = Nothing
    
    Set m_cCntrlDc = New clsStoreDc
    m_cCntrlDc.CreateFromPicture picDisabled
    If Not (m_lThemeColor = -1) Then
        m_cCntrlDc.ColorizeImage m_lThemeColor, 0.3
    End If
    lMask = GetPixel(m_cCntrlDc.hdc, 0&, 0&)
    ImlChkBoxAddBmp m_cCntrlDc.ExtractBitmap, lMask
    Set m_cCntrlDc = Nothing
    Set picDisabled = Nothing

End Sub

Private Sub CheckBoxRender(ByVal lHwnd As Long, _
                           ByVal lHdc As Long)

Dim lY      As Long
Dim lChk    As Long
Dim lFlags  As Long
Dim tRect   As RECT

    If (IsWindowEnabled(lHwnd) = 0) Then
        lChk = 2
    ElseIf m_bChecked Then
        lChk = 1
    Else
        lChk = 0
    End If
    
    GetClientRect lHwnd, tRect
    lY = ((tRect.Bottom - 16) / 2)
    FillRect lHdc, tRect, m_lhButtonBrush
    
    Select Case lChk
    Case BST_UNCHECKED
        ImageList_Draw m_ImlChkBoxHnd, 0, lHdc, 0, lY, ILD_TRANSPARENT
        
    Case BST_CHECKED
        ImageList_Draw m_ImlChkBoxHnd, 1, lHdc, 0, lY, ILD_TRANSPARENT

    Case BST_INDETERMINATE
        ImageList_Draw m_ImlChkBoxHnd, 2, lHdc, 0, lY, ILD_TRANSPARENT
    End Select
    
    With tRect
        .left = .left + 18
    End With
    
    lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
    DrawText lHdc, lFlags, Text, tRect

End Sub


'> Label
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub LabelSize(ByVal sText As String)

Dim lHdc    As Long
Dim lFlags  As Long
Dim tRect   As RECT

    lHdc = GetDC(m_lCntlHwnd)
    lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    If m_bIsNt Then
        sText = sText & vbNullChar
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags Or DT_CALCRECT
    Else
        DrawTextA lHdc, sText, -1, tRect, lFlags Or DT_CALCRECT
    End If
    
    InflateRect tRect, 1, 3
    With tRect
        Size tRect.Right, tRect.Bottom
    End With
    
    ReleaseDC m_lCntlHwnd, lHdc
    
End Sub

Public Sub LabelCaption(ByVal sText As String)
    
    If (LenB(sText) > 0) Then
        SetWindowText sText
    End If
    
End Sub

Private Sub SetWindowText(ByVal sText As String)

    If Not (m_lCntlHwnd = 0) Then
        If m_bIsNt Then
            SetWindowTextW m_lCntlHwnd, StrPtr(sText)
        Else
            SetWindowTextA m_lCntlHwnd, sText
        End If
    End If
    
End Sub

'> PictureBox
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Public Sub PictureBoxLoadImage(ByVal lImgHnd As Long, _
                               ByVal eImageType As EITImageType, _
                               ByVal lWidth As Long, _
                               ByVal lHeight As Long)

    If Not (m_lCntlHwnd = 0) Then
        If Not (lImgHnd = 0) Then
            If (eImageType < 4) Then
                Select Case eImageType
                Case eitBitmap
                    SetStyle SS_BITMAP, SS_ICON Or SS_ENHMETAFILE
                Case eitCursor
                    SetStyle SS_ICON, SS_BITMAP Or SS_ENHMETAFILE
                Case eitIcon
                    SetStyle SS_ICON, SS_BITMAP Or SS_ENHMETAFILE
                Case eitMetafile
                    SetStyle SS_ENHMETAFILE, SS_ICON Or SS_BITMAP
                End Select
                m_eImageType = eImageType
                m_lImgHnd = lImgHnd
                If m_bIsNt Then
                    SendMessageLongW m_lCntlHwnd, STM_SETIMAGE, eImageType, lImgHnd
                Else
                    SendMessageLongA m_lCntlHwnd, STM_SETIMAGE, eImageType, lImgHnd
                End If
            End If
        End If
    End If
    
    If m_bAutoSize Then
        Select Case m_eBorderStyle
        Case ecbsLine, ecbsThin
            lWidth = lWidth + 2
            lHeight = lHeight + 2
        Case ecbsThick, ecbs3D
            lWidth = lWidth + 4
            lHeight = lHeight + 4
        Case Else
            lWidth = lWidth + 1
            lHeight = lHeight + 1
        End Select
        Size lWidth, lHeight
    End If
    
End Sub

Public Sub PictureBoxRefresh()

    If Not (m_lCntlHwnd = 0) Then
        If Not (m_lImgHnd = 0) Then
            If m_bIsNt Then
                SendMessageLongW m_lCntlHwnd, STM_SETIMAGE, m_eImageType, m_lImgHnd
            Else
                SendMessageLongA m_lCntlHwnd, STM_SETIMAGE, m_eImageType, m_lImgHnd
            End If
        End If
    End If
    
End Sub

Public Sub PictureBoxClear()
    
    If Not (m_lCntlHwnd = 0) Then
        SendMessageLongA m_lCntlHwnd, STM_SETIMAGE, 0&, 0&
    End If
    
End Sub

Public Sub Refresh()

Dim tRect As RECT

    If Not (m_lCntlHwnd = 0) Then
        GetClientRect m_lCntlHwnd, tRect
        EraseRect m_lCntlHwnd, tRect, 0&
        SendMessageLongA m_lCntlHwnd, WM_PAINT, 0&, 0&
    End If
    
End Sub

Private Function SetUnicode(ByVal bEnable As Boolean) As Boolean
'/* enable/disable unicode processing

Dim lRet As Long

    If Not (m_lCntlHwnd = 0) Then
        If m_bIsNt Then
            If bEnable Then
                If Not UnicodeState Then
                    lRet = SendMessageLongW(m_lCntlHwnd, CCM_SETUNICODEFORMAT, 1&, 0&)
                End If
            Else
                If UnicodeState Then
                    lRet = SendMessageLongW(m_lCntlHwnd, CCM_SETUNICODEFORMAT, 0&, 0&)
                End If
            End If
        End If
    Else
        lRet = -1
    End If
    SetUnicode = (lRet = 0)

End Function

Private Function UnicodeState() As Boolean
'/* get control unicode readiness

    If Not (m_lCntlHwnd = 0) Then
        UnicodeState = SendMessageLongW(m_lCntlHwnd, CCM_GETUNICODEFORMAT, 0&, 0&) <> 0
    End If

End Function

'> Support
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Public Sub BorderStyle(ByVal eStyle As ECBSBorderStyle)

    m_lBorderStyle = 0
    Select Case eStyle
    Case ecbs3D
        m_lBorderStyle = WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
    Case ecbsLine
        m_lBorderStyle = WS_EX_WINDOWEDGE
    Case ecbsNone
        m_lBorderStyle = 0&
    Case ecbsThick
        m_lBorderStyle = WS_EX_CLIENTEDGE
    Case ecbsThin
        m_lBorderStyle = WS_EX_STATICEDGE
    End Select
    m_eBorderStyle = eStyle
    
End Sub

Private Sub CreateBackbuffer()

Dim lCt         As Long
Dim lTmpDc      As Long

On Error GoTo Handler

    ReDim m_lHdc(4)
    ReDim m_lBmp(4)
    ReDim m_lBmpOld(4)

    lTmpDc = m_cCntrlDc.hdc

    For lCt = 0 To 4
        m_lHdc(lCt) = CreateCompatibleDC(lTmpDc)
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lCmdWidth, m_lCmdHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lCmdWidth, m_lCmdHeight, m_cCntrlDc.hdc, (m_lCmdWidth * lCt), 0, SRCCOPY
    Next lCt

Handler:
    On Error GoTo 0

End Sub

Private Sub DestroyBackbuffer()

Dim lCt     As Long

On Error GoTo Handler

    For lCt = 0 To 4
        SelectObject m_lHdc(lCt), m_lBmpOld(lCt)
        DeleteObject m_lBmp(lCt)
        DeleteDC m_lHdc(lCt)
        m_lHdc(lCt) = 0
        m_lBmp(lCt) = 0
        m_lBmpOld(lCt) = 0
    Next lCt

Handler:
    On Error GoTo 0

End Sub

Private Sub DefaultFont()

    If (m_lHFont = 0) Then
        If (m_oFont Is Nothing) Then
            Set m_oFont = New StdFont
            With m_oFont
                .Charset = 3
                .Name = "MS Sans Serif"
                .Weight = 400
                .Size = 8
            End With
            Set Font = m_oFont
        End If
    End If

End Sub

Private Sub DestroyFont()
'*/ font cleanup

    If Not (m_lHFont = 0) Then
        DeleteObject m_lHFont
        m_lHFont = 0
    End If
    If Not (m_oFont Is Nothing) Then
        Set m_oFont = Nothing
    End If

End Sub

Private Sub DrawText(ByVal lHdc As Long, _
                     ByVal lFlags As Long, _
                     ByVal sText As String, _
                     ByRef tRect As RECT)

Dim lFontOld As Long

    If Not (LenB(sText) = 0) Then
        SetBkMode lHdc, BM_TRANSPARENT
        lFontOld = SelectObject(lHdc, m_lHFont)
        sText = sText & Chr$(0)
        If m_bIsNt Then
            If m_bFontRightLeading Then
                lFlags = lFlags Or DT_RTLREADING
            End If
            DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        Else
            DrawTextA lHdc, sText, -1, tRect, lFlags
        End If
        SelectObject lHdc, lFontOld
    End If
    
End Sub

Public Property Get Font() As StdFont
'/* [get] retrieve list font

    If Not (m_oFont Is Nothing) Then
        Set Font = m_oFont
    End If

End Property

Public Property Set Font(ByVal oFont As StdFont)
'*/ [set] change list font

Dim lHdc        As Long
Dim lChar       As Long
Dim uLF         As LOGFONT
Dim bteFont()   As Byte

    If Not (oFont Is Nothing) Then
        DestroyFont
        Set m_oFont = oFont
        lHdc = CreateDc("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        With uLF
            bteFont = StrConv(oFont.Name, vbFromUnicode)
            For lChar = 0 To UBound(bteFont)
                .lfFaceName(lChar) = bteFont(lChar)
            Next lChar
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
            m_lHFont = CreateFontIndirectW(uLF)
        Else
            m_lHFont = CreateFontIndirectA(uLF)
        End If
        HFont = m_lHFont
    End If

End Property

Private Function FunctionExported(ByVal sFunction As String, _
                                  ByVal sModule As String) As Boolean

Dim lhMod       As Long
Dim bLoaded     As Boolean

    lhMod = GetModuleHandleA(sModule)
    If (lhMod = 0) Then
        lhMod = LoadLibraryA(sModule)
        If lhMod Then
            bLoaded = True
        End If
    End If
    If lhMod Then
        If (GetProcAddress(lhMod, sFunction)) Then
            FunctionExported = True
        End If
    End If
    If bLoaded Then
        FreeLibrary lhMod
    End If

End Function

Private Sub DestroyCommandIml()

    If Not (m_ImlCommandHnd = 0) Then
        ImageList_Destroy m_ImlCommandHnd
        m_ImlCommandHnd = 0
    End If

End Sub

Private Sub DestroyListBoxIml()

    If Not (m_ImlListBoxHnd = 0) Then
        ImageList_Destroy m_ImlListBoxHnd
        m_ImlListBoxHnd = 0
    End If

End Sub

Private Sub DestroyChkBoxIml()

    If Not (m_ImlChkBoxHnd = 0) Then
        ImageList_Destroy m_ImlChkBoxHnd
        m_ImlChkBoxHnd = 0
    End If

End Sub

Public Property Get Focus() As Boolean
'/* [get] grid focus
    Focus = (GetFocus() = m_lCntlHwnd)
End Property

Public Property Let Focus(ByVal PropVal As Boolean)
'/* [let] grid focus
    If Not (m_lCntlHwnd = 0) Then
        If m_bIsNt Then
            If PropVal Then
                PostMessageW m_lCntlHwnd, WM_SETFOCUS, 0&, 0&
            Else
                PostMessageW m_lCntlHwnd, WM_KILLFOCUS, 0&, 0&
            End If
        Else
            If PropVal Then
                PostMessageA m_lCntlHwnd, WM_SETFOCUS, 0&, 0&
            Else
                PostMessageA m_lCntlHwnd, WM_KILLFOCUS, 0&, 0&
            End If
        End If
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Public Property Let Visible(ByVal PropVal As Boolean)

    If Not (m_lCntlHwnd = 0) Then
        If PropVal Then
            SetWindowPos m_lCntlHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
        Else
            SetWindowPos m_lCntlHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_HIDEWINDOW
        End If
    End If
    m_bVisible = PropVal
    
End Property

Public Sub ImlChkBoxAddBmp(ByVal lhBitmap As Long, _
                           Optional ByVal lMask As Long = CLR_NONE)

    If Not (m_ImlChkBoxHnd = 0) Then
        If Not lMask = CLR_NONE Then
            ImageList_AddMasked m_ImlChkBoxHnd, lhBitmap, lMask
        Else
            ImageList_Add m_ImlChkBoxHnd, lhBitmap, 0&
        End If
    End If

End Sub

Public Sub ImlCommandAddIcon(ByVal lIcon As Long)

    If (m_ImlCommandHnd = 0) Then
        InitCommandIml
    End If
    If Not (lIcon = 0) Then
        ImageList_AddIcon m_ImlCommandHnd, lIcon
    Else
        DestroyCommandIml
    End If
    
End Sub

Public Sub ImlListBoxAddIcon(ByVal lIcon As Long)

    If Not (m_ImlListBoxHnd = 0) Then
        ImageList_AddIcon m_ImlListBoxHnd, lIcon
    End If

End Sub

Public Function ImlListBoxRemoveIcon(ByVal lIndex As Long)

    If Not (m_ImlListBoxHnd = 0) Then
        ImageList_Remove m_ImlListBoxHnd, lIndex
    End If

End Function

Public Sub InitChkBoxIml(Optional ByVal lWidth As Long = 16, _
                         Optional ByVal lHeight As Long = 16)

    If Not (m_lCntlHwnd = 0) Then
        DestroyChkBoxIml
        m_ImlChkBoxHnd = ImageList_Create(lWidth, lHeight, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    End If

End Sub

Public Sub InitCommandIml(Optional ByVal lWidth As Long = 16, _
                          Optional ByVal lHeight As Long = 16)

    If Not (m_lCntlHwnd = 0) Then
        DestroyCommandIml
        m_ImlCommandHnd = ImageList_Create(lWidth, lHeight, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    End If

End Sub

Public Sub InitListBoxIml(Optional ByVal lWidth As Long = 16, _
                          Optional ByVal lHeight As Long = 16)

    If Not (m_lCntlHwnd = 0) Then
        DestroyListBoxIml
        m_ImlListBoxHnd = ImageList_Create(lWidth, lHeight, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
    End If

End Sub

Public Sub ListBoxUseExternalIml(ByVal lImlHwnd As Long)
    DestroyListBoxIml
    m_ImlListBoxHnd = lImlHwnd
End Sub

Private Sub LabelRender(ByVal lHwnd As Long, _
                        ByVal lHdc As Long)


Dim lFlags As Long
Dim tRect As RECT

    GetClientRect lHwnd, tRect
    FillRect lHdc, tRect, m_lhButtonBrush
    lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
    DrawText lHdc, lFlags, Text, tRect

End Sub

Private Sub ODRender(ByVal lHwnd As Long, _
                     ByVal lHdc As Long, _
                     ByRef tRect As RECT, _
                     Optional ByVal lState As Long, _
                     Optional ByVal lItemId As Long)
    
    Select Case m_eControlStyle
    Case ecsCommandButton
        CommandRender lHwnd, lHdc, lState
    Case ecsImageCombo
        ImageComboRender lHdc, lItemId, tRect
    Case ecsCheckBox, ecsOptionButton
        CheckBoxRender lHwnd, lHdc
    Case ecsLabel
        LabelRender lHwnd, lHdc
    Case ecsListBox, ecsListBoxMultiSelect, ecsListBoxExtended, ecsImageListBox
        ListBoxRender lHdc, lItemId, lState, tRect
    End Select
    
End Sub

Public Sub Position(ByVal lX As Long, _
                    ByVal lY As Long)
    
    If Not (m_lCntlHwnd = 0) Then
        SetWindowPos m_lCntlHwnd, HWND_TOP, lX, lY, 0&, 0&, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOOWNERZORDER
    End If
    
End Sub

Private Sub Render(ByVal lHwnd As Long, _
                   ByVal lHdc As Long, _
                   Optional ByVal lState As Long)
    
    Select Case m_eControlStyle
    Case ecsComboDropDown, ecsComboDropList, ecsImageCombo
        ComboRender lHwnd, lHdc
    Case ecsCommandButton
        CommandRender lHwnd, lHdc, lState
    End Select
    
End Sub

Private Sub SetStyle(ByVal lStyle As Long, _
                     ByVal lStyleNot As Long)

'*/ change style params

Dim lNewStyle   As Long

On Error GoTo Handler

    If Not (m_lCntlHwnd = 0) Then
        If m_bIsNt Then
            lNewStyle = GetWindowLongW(m_lCntlHwnd, GWL_STYLE)
        Else
            lNewStyle = GetWindowLongA(m_lCntlHwnd, GWL_STYLE)
        End If

        lNewStyle = lNewStyle And Not lStyleNot
        lNewStyle = lNewStyle Or lStyle
        If m_bIsNt Then
            SetWindowLongW m_lCntlHwnd, GWL_STYLE, lNewStyle
        Else
            SetWindowLongA m_lCntlHwnd, GWL_STYLE, lNewStyle
        End If
        SetWindowPos m_lCntlHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If

Handler:

End Sub

Private Sub TrackMouse(ByVal lHwnd As Long)

Dim tMe As TRACKMOUSEEVENT_STRUCT

    With tMe
        .cbSize = Len(tMe)
        .dwFlags = TME_LEAVE
        .hwndTrack = lHwnd
    End With

    If m_bTrackUser32 Then
        TrackMouseEvent tMe
    Else
        TrackMouseEventComCtl tMe
    End If

End Sub

Public Sub Size(ByVal lWidth As Long, _
                ByVal lHeight As Long)

    If Not (m_lCntlHwnd = 0) Then
        SetWindowPos m_lCntlHwnd, HWND_TOP, 0&, 0&, lWidth, lHeight, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOOWNERZORDER
    End If

End Sub

Private Sub Show()
'/* position edit box

    If Not (m_lCntlHwnd = 0) Then
        '/* position edit window
        With m_tRect
            SetWindowPos m_lCntlHwnd, HWND_TOP, 0&, 0&, 0&, 0&, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOACTIVATE
        End With
    End If

End Sub


'> Subclassing
'>>>>>>>>>>>>>>>>>>>>>>
Private Sub AttachMessages()

    With m_cCntlSubclass
        If Not (m_lCntlHwnd = 0) Then
            Select Case m_eControlStyle
            Case ecsComboSimple
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                
            Case ecsComboDropDown
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_PAINT, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                If Not (m_lEditHwnd = 0) Then
                    .Subclass m_lEditHwnd, Me
                    .AddMessage m_lEditHwnd, WM_SETTEXT, MSG_BEFORE
                End If
                If Not (m_lListHwnd = 0) Then
                    .Subclass m_lListHwnd, Me
                    .AddMessage m_lListHwnd, WM_LBUTTONUP, MSG_BEFORE
                End If
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                End If
            
            Case ecsComboDropList
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_PAINT, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                End If
            
            Case ecsImageCombo
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_PAINT, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                If Not (m_lEditHwnd = 0) Then
                    .Subclass m_lEditHwnd, Me
                    .AddMessage m_lEditHwnd, WM_SETTEXT, MSG_BEFORE
                End If
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                End If
                
            Case ecsCheckBox, ecsOptionButton
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_SETFOCUS, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KILLFOCUS, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLORBTN, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                End If
            
            Case ecsCommandButton
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_MOUSEMOVE, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_MOUSELEAVE, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                End If
       
            Case ecsLabel
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                End If
            
            Case ecsListBox, ecsListBoxMultiSelect, ecsListBoxExtended
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                End If
            
            Case ecsImageListBox
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                    .AddMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                End If
            
            Case ecsPictureBox
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                End If
            
            Case ecsTextBox
                .Subclass m_lCntlHwnd, Me
                .AddMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .AddMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                If Not (m_lParentHwnd = 0) Then
                    .Subclass m_lParentHwnd, Me
                    .AddMessage m_lParentHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                End If
            End Select
        End If
    End With
    
End Sub

Private Sub DetachMessages()
    
    With m_cCntlSubclass
        If Not (m_lCntlHwnd = 0) Then
            Select Case m_eControlStyle
            Case ecsComboSimple
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                
            Case ecsComboDropDown
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_PAINT, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lEditHwnd = 0) Then
                    .DeleteMessage m_lEditHwnd, WM_SETTEXT, MSG_BEFORE
                    .UnSubclass m_lEditHwnd
                End If
                If Not (m_lListHwnd = 0) Then
                    .DeleteMessage m_lListHwnd, WM_LBUTTONUP, MSG_BEFORE
                    .UnSubclass m_lListHwnd
                End If
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
                
            Case ecsComboDropList
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_PAINT, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
            
            Case ecsImageCombo
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_PAINT, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lListHwnd = 0) Then
                    .DeleteMessage m_lListHwnd, WM_LBUTTONUP, MSG_BEFORE
                    .UnSubclass m_lListHwnd
                End If
                If Not (m_lEditHwnd = 0) Then
                    .DeleteMessage m_lEditHwnd, WM_SETTEXT, MSG_BEFORE
                    .UnSubclass m_lEditHwnd
                End If
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
                
            Case ecsCheckBox, ecsOptionButton
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_SETFOCUS, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KILLFOCUS, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLORBTN, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
                
            Case ecsCommandButton
                .DeleteMessage m_lCntlHwnd, WM_MOUSEMOVE, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_MOUSELEAVE, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
            
            Case ecsLabel
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
            
            Case ecsListBox, ecsListBoxMultiSelect, ecsListBoxExtended
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
            
            Case ecsImageListBox
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLORLISTBOX, MSG_BEFORE
                    .DeleteMessage m_lParentHwnd, WM_DRAWITEM, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
            
            Case ecsPictureBox
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLORSTATIC, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
                
            Case ecsTextBox
                .DeleteMessage m_lCntlHwnd, WM_ERASEBKGND, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                .DeleteMessage m_lCntlHwnd, WM_KEYDOWN, MSG_BEFORE
                .UnSubclass m_lCntlHwnd
                If Not (m_lParentHwnd = 0) Then
                    .DeleteMessage m_lParentHwnd, WM_CTLCOLOREDIT, MSG_BEFORE
                    .UnSubclass m_lParentHwnd
                End If
            End Select
        End If
    End With
    
End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                ByRef bHandled As Boolean, _
                                ByRef lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                ByRef lParamUser As Long)

Dim lHdc    As Long
Dim tPaint  As PAINTSTRUCT
Dim tCBInfo As COMBOBOXINFO
Dim tOdCtl  As DRAWITEMSTRUCT

    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            If (lHwnd = m_lCntlHwnd) Then
                m_bPainting = True
                BeginPaint lHwnd, tPaint
                Render lHwnd, tPaint.hdc
                EndPaint lHwnd, tPaint
                m_bPainting = False
                bHandled = True
            Else
                lReturn = m_cCntlSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
            End If
        Else
            lReturn = m_cCntlSubclass.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If

    Case WM_DRAWITEM
        CopyMemory tOdCtl, ByVal lParam, LenB(tOdCtl)
        With tOdCtl
            If (.hwndItem = m_lCntlHwnd) Then
                ODRender .hwndItem, .hdc, .rcItem, .itemState, .itemID
            End If
        End With
        
    Case WM_KEYDOWN
        Select Case wParam
        Case VK_SPACE
            Select Case m_eControlStyle
            Case ecsCheckBox, ecsOptionButton
                Checked = Not Checked
                RaiseEvent Click
            Case ecsCommandButton
                RaiseEvent Click
            End Select
        Case VK_TAB, VK_ENTER
            If Not (m_eControlStyle = ecsTextBox) Then
                RaiseEvent LostFocus
                bHandled = True
            End If
        Case VK_ESCAPE
            RaiseEvent LostFocus
            bHandled = True
        End Select
        
    Case WM_CTLCOLORLISTBOX, WM_CTLCOLOREDIT, WM_CTLCOLORSTATIC, WM_CTLCOLORBTN
        Select Case m_eControlStyle
        Case ecsComboSimple
            ComboInfo m_lCntlHwnd, tCBInfo
            If (lParam = tCBInfo.hwndEdit) Then
                SetBkColor wParam, m_lBackColor
                SetTextColor wParam, m_lForeColor
                bHandled = True
            ElseIf (lParam = tCBInfo.hwndList) Then
                SetBkColor wParam, m_lBackColor
                SetTextColor wParam, m_lForeColor
                bHandled = True
            End If
        'Case ecsComboDropDown, ecsImageCombo
            'ComboInfo m_lCntlHwnd, tCBInfo
            'If (lParam = tCBInfo.hwndList) Then
            '    SetBkColor wParam, m_lBackColor
            '    SetTextColor wParam, m_lForeColor
            '    bHandled = True
            'ElseIf (lParam = m_lEditHwnd) Then
            '    SetBkColor wParam, m_lBackColor
            '    SetTextColor wParam, m_lForeColor
            '    bHandled = True
            'End If
        Case ecsComboDropList
            ComboInfo m_lCntlHwnd, tCBInfo
            If (lParam = tCBInfo.hwndList) Then
                SetBkColor wParam, m_lBackColor
                SetTextColor wParam, m_lForeColor
                bHandled = True
            Else
                SetBkColor wParam, m_lBackColor
                SetTextColor wParam, m_lForeColor
                bHandled = True
            End If
        Case ecsLabel
            If (lParam = m_lCntlHwnd) Then
                LabelColor lParam
                bHandled = True
            End If
        Case ecsListBox, ecsListBoxMultiSelect, ecsListBoxExtended, ecsImageListBox
            If (lParam = m_lCntlHwnd) Then
                lReturn = CreateButtonBrush
                bHandled = True
            End If
        Case ecsCheckBox, ecsOptionButton
            If (lParam = m_lCntlHwnd) Then
                ButtonColorChange lParam
                bHandled = True
            End If
        Case ecsPictureBox
            If (lParam = m_lCntlHwnd) Then
                lReturn = CreateButtonBrush
                bHandled = True
            End If
        Case ecsTextBox
            If (lParam = m_lCntlHwnd) Then
                SetBkColor wParam, m_lBackColor
                SetTextColor wParam, m_lForeColor
                lReturn = 1
                bHandled = True
            End If
        End Select
        
    Case WM_MOUSEMOVE
        If Not m_bMouseOver Then
            lHdc = GetDC(lHwnd)
            Render lHwnd, lHdc, 15
            ReleaseDC lHwnd, lHdc
            m_bMouseOver = True
        End If
        TrackMouse lHwnd
            
    Case WM_MOUSELEAVE
        lHdc = GetDC(m_lCntlHwnd)
        Render lHwnd, lHdc, 16
        ReleaseDC m_lCntlHwnd, lHdc
        m_bMouseOver = False
        
    Case WM_LBUTTONDOWN
        Select Case m_eControlStyle
        Case ecsCheckBox, ecsOptionButton
            m_bChecked = Not m_bChecked
        End Select
        Select Case m_eControlStyle
        Case ecsCheckBox, ecsCommandButton, ecsLabel, ecsOptionButton, ecsPictureBox, ecsTextBox
            RaiseEvent Click
        End Select
        
    Case WM_SETTEXT
        If (m_eControlStyle = ecsImageCombo) Then
            ComboEditBoxRender
        End If

    Case WM_LBUTTONUP
        RaiseEvent ItemChange(ListIndex)
    
    Case WM_SETFOCUS, WM_KILLFOCUS
        UpdateWindow m_lCntlHwnd
        
    Case WM_ERASEBKGND
        Select Case m_eControlStyle
        Case ecsPictureBox
            PictureBoxRefresh
        Case ecsImageCombo
            ComboEditBoxRender
        Case Else
            lReturn = 1
        End Select
    End Select

End Sub

Private Function LabelColor(ByVal lHwnd As Long) As Long

Dim lHdc    As Long
Dim lColor  As Long

    If m_bAutoBackColor Then
        lColor = GetBackColor
        If Not (lColor = -1) Then
            If Not (lColor = m_lBackColor) Then
                m_lBackColor = lColor
                CreateButtonBrush
                lHdc = GetDC(lHwnd)
                LabelRender lHwnd, lHdc
                ReleaseDC lHwnd, lHdc
            End If
        End If
    End If

End Function

Private Sub ButtonColorChange(ByVal lHwnd As Long)

Dim lHdc    As Long
Dim lColor  As Long

    If m_bAutoBackColor Then
        lColor = GetBackColor
        If Not (lColor = -1) Then
            If Not (lColor = m_lBackColor) Then
                m_lBackColor = lColor
                CreateButtonBrush
                lHdc = GetDC(lHwnd)
                CheckBoxRender lHwnd, lHdc
                ReleaseDC lHwnd, lHdc
            End If
        End If
    End If

End Sub

Public Sub Destroy()

    If Not (m_lCntlHwnd = 0) Then
        DetachMessages
        If (m_eControlStyle = ecsCommandButton) Then
            DestroyBackbuffer
        End If
        DestroyButtonBrush
        DestroyFont
        DestroyChkBoxIml
        DestroyCommandIml
        DestroyListBoxIml
        If Not m_cCntlSubclass Is Nothing Then Set m_cCntlSubclass = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cCntrlDc Is Nothing Then Set m_cCntrlDc = Nothing
        If Not m_pControlImg Is Nothing Then Set m_pControlImg = Nothing
        DestroyWindow m_lCntlHwnd
        m_lCntlHwnd = 0
    End If
    
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub
