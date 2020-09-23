Attribute VB_Name = "mSupport"
'The code in this module was modified from MSDN code of the EnumFontFamilies api.
'The point of this submission is to demonstrate how to quickly pipe font information to a
'listview control and add the appropriate icons.  BTW... sorry about the crappy 16x16 icons
'included... I made them quickly.
'... I have seen an example of this api implementation on PSC... but the
'... example parsed the font information to a listbox based on the font type...
'... then parsed the listbox, assigned an icon based on the font type... then transferred the
'... icon and font name to a list view.  Then, it re-parsed the fonts (for a different font
'... type) to the listbox, assigned an icon based on that font type... and so on.  It was
'... redundent and needed the extra listbox (invisible) control.
'In this version, I move the font information directly to the listview control with the icon needed.
'You can easily do this with a combo/listbox or any other place...
'see the comments in the EnumFontFamProc function below.
'Do what ever you'd like with this code... I didn't write the api - hope it is helpful to you.

'Font enumeration types... keep this above the Type LOGFONT type... or you'll get an error.
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

'establish some types needed by the api
Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

'declare constants needed by the api

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1
Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'  Enumerate Font Mask... use these to determine if TrueType... etc.
Public Const RASTER_FONTTYPE = &H1 ' a raster font found
Public Const DEVICE_FONTTYPE = &H2 'never found one of these... not sure what they are!
Public Const TRUETYPE_FONTTYPE = &H4 'true type font found

'declare the functions used
'get the font info
Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hDC As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, LParam As Any) As Long
'get the device context for the object (Listview1 in this example)
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'release the device context for the object (Listview1 in this example)
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, LParam As ListView) As Long
Dim FaceName As String
Dim FullName As String
    'convert the facename to unicode
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    'assign the font found to the listview at the next available location
    Set itmX = Form1.ListView1.ListItems.Add(, , left$(FaceName, InStr(FaceName, vbNullChar) - 1))
    '-- if you wanted to add this font to a list box... you'd use: list1.additem(...)
    'figure out which icon to associate with the font found
    If FontType = 4 Then
        'truetype font found
        itmX.SmallIcon = 1 'assign the TrueType icon to this entry
    End If
    If FontType = 1 Then
        'raster font found
        itmX.SmallIcon = 2 'assign the Raster icon to this entry
    End If

    EnumFontFamProc = 1 'return a true value & cycle till false
End Function

Sub FillListWithFonts(LV As ListView)
'this is the sub called to fill the listview (in this case) with the font/type information
'as you can see, it calls the EnumFontFamilies api, and cycles using the "AddressOf EnumFontFamProc" callback
Dim hDC As Long
    hDC = GetDC(LV.hWnd) ' get the device context
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProc, LV 'call the api and cycle
    ReleaseDC LV.hWnd, hDC 'release the device context
End Sub



