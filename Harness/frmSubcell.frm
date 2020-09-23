VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\prjvhGrid.vbp"
Begin VB.Form frmSubcell 
   BackColor       =   &H00F6F6F6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vhGrid - SubCells"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   11805
   StartUpPosition =   1  'CenterOwner
   Begin vhGrid.ucVHGrid ucVHGrid1 
      Height          =   5910
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   10425
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderIconColourDepth=   32
      HeaderIconCount =   6
      HeaderIcons     =   "frmSubcell.frx":0000
      CellIconSizeX   =   32
      CellIconSizeY   =   32
      CellIconColourDepth=   32
      CellIconCount   =   10
      CellIcons       =   "frmSubcell.frx":18021
      TreeIconColourDepth=   0
      AlphaBarTransparency=   70
      ForeColor       =   0
      GridLines       =   0
      HeaderDragDrop  =   0   'False
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderForeColor =   0
      HeaderForeColorFocused=   0
      HeaderForeColorPressed=   0
      HeaderHeight    =   24
      HeaderHeightSizable=   -1  'True
   End
   Begin VB.CommandButton cmdPopulate 
      Caption         =   "Populate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9540
      TabIndex        =   0
      Top             =   6300
      Width           =   2040
   End
   Begin MSComctlLib.ImageList imlHdr 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubcell.frx":40042
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubcell.frx":4182C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubcell.frx":43016
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubcell.frx":44800
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSubcell.frx":45FEA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSubcell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_lRowCount                 As Long
Private m_lSkinStyle                As Long
Private m_clblDesc                  As clsODControl
Private m_clblBck                   As clsODControl
Private WithEvents m_coptEdit0      As clsODControl
Attribute m_coptEdit0.VB_VarHelpID = -1
Private WithEvents m_coptEdit1      As clsODControl
Attribute m_coptEdit1.VB_VarHelpID = -1
Private WithEvents m_coptEdit2      As clsODControl
Attribute m_coptEdit2.VB_VarHelpID = -1
Private WithEvents m_cmbSave        As clsODControl
Attribute m_cmbSave.VB_VarHelpID = -1
Private WithEvents m_chkOptn0       As clsODControl
Attribute m_chkOptn0.VB_VarHelpID = -1
Private WithEvents m_chkOptn1       As clsODControl
Attribute m_chkOptn1.VB_VarHelpID = -1

Private Sub m_coptEdit0_Click()

    If m_coptEdit0.Checked Then
        m_coptEdit1.Checked = False
        m_coptEdit1.Refresh
        m_coptEdit2.Checked = False
        m_coptEdit2.Refresh
        ucVHGrid1.CellEdit = True
        ucVHGrid1.AdvancedEdit = False
    End If
    
End Sub

Private Sub m_coptEdit1_Click()
    If m_coptEdit1.Checked Then
        m_coptEdit0.Checked = False
        m_coptEdit0.Refresh
        m_coptEdit2.Checked = False
        m_coptEdit2.Refresh
        ucVHGrid1.CellEdit = False
        ucVHGrid1.AdvancedEdit = True
    End If
End Sub

Private Sub m_coptEdit2_Click()
    If m_coptEdit2.Checked Then
        m_coptEdit0.Checked = False
        m_coptEdit0.Refresh
        m_coptEdit1.Checked = False
        m_coptEdit1.Refresh
        ucVHGrid1.CellEdit = False
        ucVHGrid1.AdvancedEdit = False
    End If
End Sub

Private Sub ucVHGrid1_eVHEditChange(ByVal lRow As Long, ByVal lCell As Long, ByVal sText As String)
    Debug.Print "Edit: Row " & lRow & ", Cell " & lCell & ", Text " & sText
End Sub

Private Sub ucVHGrid1_eVHEditorLoaded(ByVal eEditorType As ECTEditControlType, ByVal lRow As Long, ByVal lCell As Long)

    With ucVHGrid1
        Select Case lCell
        Case 1
            .EditControlAddData m_sMedia(0), imlHdr.ListImages.Item(1).Picture, 0
            .EditControlAddData m_sMedia(1), imlHdr.ListImages.Item(2).Picture, 1
            .EditControlAddData m_sMedia(2), imlHdr.ListImages.Item(3).Picture, 2
            .EditControlAddData m_sMedia(3), imlHdr.ListImages.Item(4).Picture, 3
            .EditControlAddData m_sMedia(4), imlHdr.ListImages.Item(1).Picture, 4
            .EditControlAddData m_sMedia(5), imlHdr.ListImages.Item(2).Picture, 5
            .EditControlAddData m_sMedia(6), imlHdr.ListImages.Item(3).Picture, 6
            .EditControlAddData m_sMedia(7), imlHdr.ListImages.Item(4).Picture, 7
            .EditControlAddData m_sMedia(8), imlHdr.ListImages.Item(3).Picture, 8
            .EditControlAddData m_sMedia(9), imlHdr.ListImages.Item(4).Picture, 9
        Case 2
            .EditControlAddData m_sTitles(0), imlHdr.ListImages.Item(1).Picture, 2
            .EditControlAddData m_sTitles(1), imlHdr.ListImages.Item(2).Picture, 2
            .EditControlAddData m_sTitles(2), imlHdr.ListImages.Item(3).Picture, 2
            .EditControlAddData m_sTitles(3), imlHdr.ListImages.Item(4).Picture, 2
            .EditControlAddData m_sTitles(4), imlHdr.ListImages.Item(1).Picture, 2
            .EditControlAddData m_sTitles(5), imlHdr.ListImages.Item(2).Picture, 2
            .EditControlAddData m_sTitles(6), imlHdr.ListImages.Item(3).Picture, 2
            .EditControlAddData m_sTitles(7), imlHdr.ListImages.Item(4).Picture, 2
            .EditControlAddData m_sTitles(8), imlHdr.ListImages.Item(3).Picture, 2
            .EditControlAddData m_sTitles(9), imlHdr.ListImages.Item(4).Picture, 2
        End Select
    End With
    
End Sub

Private Sub ucVHGrid1_eVHEditRequest(ByVal lRow As Long, ByVal lCell As Long)
    
    With ucVHGrid1
        Select Case lCell
        Case 1
            .EditControlType = ectImageCombo
        Case 2
            .EditControlType = ectImageListbox
        Case Else
            .EditControlType = ectTextBox
        End Select
    End With
    
End Sub


Private Sub Form_Load()

Dim lX As Long

    With ucVHGrid1
        lX = (.Width / Screen.TwipsPerPixelX) / 8
        m_lSkinStyle = 3
        '/* auto set draw after last cell is loaded
        .FastLoad = True
        '/* add header icons
        .InitImlHeader
        '/* add row icons
        .InitImlRow 48, 48
        '/* add columns
        .ColumnAdd 0, "", (lX * 0.3), ecaColumnLeft, 0, ecsSortIcon
        .ColumnAdd 1, "Media Type", (lX * 1.5), ecaColumnLeft, 1, ecsSortDefault
        .ColumnAdd 2, "Title and Author", (lX * 2), ecaColumnLeft, 2, ecsSortDefault
        .ColumnAdd 3, "Media Sample", (lX * 2.4), ecaColumnLeft, 3, ecsSortDefault
        .ColumnAdd 4, "Statistics", (lX * 2), ecaColumnLeft, 4, ecsSortDefault
        Dim oHdr As New StdFont
        With oHdr
            .Bold = True
            .Name = "TIMES NEW ROMAN"
            .Size = 9
        End With
        Set .HeaderFont = oHdr
        '/* column tooltips
        .ColumnTipColor = &HC4B0A2
        .ColumnTipOffsetColor = &HF1DDCF
        .ColumnTipGradient = True
        .ColumnTipMultiline = True
        .ColumnTipDelayTime = 1.5
        .ColumnTipTransparency = 180
        .ColumnTipHint(1) = "Media family of found object"
        .ColumnTipHint(2) = "Search result Title and Author"
        .ColumnTipHint(3) = "Media sample data"
        .ColumnTipHint(4) = "Media statistics and state information"
        .CellEdit = True
        '/* filter menu
        .FilterBackColor = &HC4B0A2
        .FilterOffsetColor = &HF1DDCF
        .FilterControlColor = &HC4B0A2
        .FilterForeColor = &H808080
        .FilterGradient = True
        .FilterTransparency = 140
        .FilterAdd 2, "Led Zeppelin|U2|AC/DC|Alice In Chains|the Who"
        .FilterAdd 3, "sun|floodin|yesterday"
        
        '/* use xp colors: place near top
        .XPColors = True
        '/* set the row height
        .RowHeight = 50
        '/* enable cell tooltips
        .CellTips = True
        '/* custom cursors
        .CustomCursors = True
        '/* double buffer grid
        .DoubleBuffer = True
        '/* enable user sizable header height
        .HeaderHeightSizable = True
        '/* set header height
        '! Note: when using subcells, header height must => then row height..
        .HeaderHeight = 50
        '/* set alphbar transparency
        .AlphaBarTransparency = 120
        '/* enable sorting
        .CellsSorted = True
        '/* set the column focus color
        .ColumnFocusColor = &H9C541F
        '/* set grid back color
        .BackColor = &H887466
        '/* lock the third column
        .ColumnLock(3) = True
        '/* enable checkboxes
        .Checkboxes = True
        '/* use gridlines
        .GridLines = EGLBoth
        '/* set the drag effect style
        .DragEffectStyle = edsClientArrow
        '/* enable header vertical text
        .ColumnVerticalText = True
        '/* use icons with alpha channels
        .ImlUseAlphaIcons = True
        '/* apply skin
        .ThemeManager evsSilver, False, &HC4B0A2, estThemeSoft, &H333333, &H363636, &H636363, _
            &H8A8181, &HC4B0A2, &H808080, 210, True, True, &HC4B0A2, _
            &H0, True, False, False, False
        '/* apply cell decoration
        .CellDecoration erdCellChecker, &HF1DDCF, &HC4B0A2, True, 2
    End With

    '/* load data
    CreateMusicData

End Sub

Private Sub cmdPopulate_Click()

Dim lCt         As Long
Dim lRnd        As Long
Dim sTitle      As String
Dim sHeader     As String
Dim sDetails    As String
Dim oFnt        As StdFont
Dim lCust()     As Long

    m_lRowCount = 20

    sTitle = "Media Search Results:" & vbNewLine & "Time Elapsed: 1:22 Seconds.." & vbNewLine & "Found: 20 matching entries." & vbNewLine & "Media Types: Mixed"
    sHeader = "-Current Settings-" & vbNewLine & "File Edit Options: Enabled" & vbNewLine & "Media Play Types: Full" & vbNewLine & "User: Administrator" & vbNewLine & "Change Tracking: Loaded"
    sDetails = vbNewLine & "Master: Yes" & vbNewLine & "Public: Yes" & vbNewLine & "Last Access: " & CStr(Date)
    
    '/* cell font
    Set oFnt = New StdFont
    With oFnt
        .Name = "Tahoma"
        .Bold = True
        .Size = 8
    End With
    
    ReDim lCust(2)
    lCust(0) = RGB(217, 55, 0)
    lCust(1) = &H969682
    lCust(2) = &H414B25

    With ucVHGrid1
        '/* manually initialize list: rowcount, columncount
        .GridInit m_lRowCount, 5
        '/* add special cells
        '/* row 0 -spanned 2 rows
        .AddCell 0, 0, sTitle, DT_VCENTER Or DT_CENTER Or DT_WORDBREAK, , , , oFnt, 6, 2
        .RowHideCheckBox 0, True
        .RowNoFocus 0, True
        '/* cell is spanned from 1st to last column
        .CellSpanHorizontal 0, 0, 4
        
        '/* row 1 -spanned 3 rows
        .AddCell 1, 0, sHeader, DT_VCENTER Or DT_LEFT Or DT_WORDBREAK, , lCust(0), , oFnt, 6, 3
        .RowHideCheckBox 1, True
        .RowNoFocus 1, True
        '/* cell spanned from 1st to last column
        .CellSpanHorizontal 1, 0, 4
        
        '/* row 2 -spanned 3 rows
        .AddCell 2, 0, , , , lCust(0), , , , 2
        .AddCell 2, 1, m_sMedia(0), DT_LEFT Or DT_VCENTER Or DT_END_ELLIPSIS, 0, , , , 6
        .AddCell 2, 2, m_sTitles(0), DT_WORDBREAK Or DT_VCENTER
        .AddCell 2, 3, m_sLyrics(0), DT_WORDBREAK Or DT_VCENTER
        .AddCell 2, 4, m_sDesc(0) & sDetails, DT_WORDBREAK Or DT_VCENTER
        
        '/* row 3 -spanned 3 rows
        .AddCell 3, 0, , , , lCust(0), , , , 2
        .AddCell 3, 1, m_sMedia(1), DT_LEFT Or DT_VCENTER Or DT_END_ELLIPSIS, 1, , , , 6
        .AddCell 3, 2, m_sTitles(1), DT_WORDBREAK Or DT_VCENTER, , , , , 5
        .AddCell 3, 3, m_sLyrics(1), DT_WORDBREAK Or DT_VCENTER
        .AddCell 3, 4, m_sDesc(1) & sDetails, DT_WORDBREAK Or DT_VCENTER
        
        '/* row 4 -spanned 3 rows
        .AddCell 4, 0, , , , lCust(0), , , , 2
        .AddCell 4, 1, m_sMedia(2), DT_LEFT Or DT_VCENTER Or DT_END_ELLIPSIS, 2, , , , 6
        .AddCell 4, 2, m_sTitles(2), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, , , , , 5
        .AddCell 4, 3, m_sLyrics(2), DT_WORDBREAK Or DT_VCENTER
        .AddCell 4, 4, m_sDesc(2) & sDetails, DT_WORDBREAK Or DT_VCENTER

        '/* row 5 -spanned 3 rows
        .AddCell 5, 0, , , , lCust(0), , , , 2
        .AddCell 5, 1, m_sMedia(3), DT_LEFT Or DT_VCENTER Or DT_END_ELLIPSIS, 3, , , , 6
        .AddCell 5, 2, m_sTitles(3), DT_WORDBREAK Or DT_VCENTER, , , , , 5
        .AddCell 5, 3, m_sLyrics(3), DT_WORDBREAK Or DT_VCENTER
        .AddCell 5, 4, m_sDesc(3) & sDetails, DT_WORDBREAK Or DT_VCENTER

        '/* row 6 -spanned 3 rows
        .AddCell 6, 0, , , , , , , , 2
        .AddCell 6, 1, m_sMedia(4), DT_LEFT Or DT_VCENTER Or DT_END_ELLIPSIS, 4, , , , 6
        .AddCell 6, 2, m_sTitles(4), DT_WORDBREAK Or DT_VCENTER, , , , , 5
        .AddCell 6, 3, m_sLyrics(4), DT_WORDBREAK Or DT_VCENTER
        .AddCell 6, 4, m_sDesc(4) & sDetails, DT_LEFT Or DT_WORDBREAK Or DT_VCENTER
        
        '/* add the rest of the rows
        For lCt = 7 To m_lRowCount
            lRnd = RandomNum(0, 9)
            .AddCell lCt, 0, , , , lCust(0)
            .AddCell lCt, 1, m_sMedia(lRnd), DT_LEFT Or DT_END_ELLIPSIS, lRnd, , , , 6
            .AddCell lCt, 2, m_sTitles(lRnd), DT_WORDBREAK Or DT_VCENTER, , , , , 5
            .AddCell lCt, 3, m_sLyrics(lRnd), DT_WORDBREAK Or DT_VCENTER, , , , , 5
            .AddCell lCt, 4, m_sDesc(lRnd), DT_WORDBREAK Or DT_VCENTER
        Next lCt
        '/* refresh the grid
        .GridRefresh True
    End With

    '/* add subcells
    AddSubCells
    
End Sub

Private Sub AddSubCells()

    With ucVHGrid1
        Set m_clblDesc = New clsODControl
        With m_clblDesc
            .Name = "lblDesc"
            .BorderStyle ecbsNone
            .LabelTransparent = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 9, 129, 44, 13, ecsLabel
            .Text = "Local Edit Options"
            .AutoSize = True
            .AutoBackColor = True
        End With
        .SubCellAddControl 1, 3, 100, 20, m_clblDesc.hWnd, evsUserDefine, , 2, 20, True, False
        
        Set m_coptEdit0 = New clsODControl
        With m_coptEdit0
            .Name = "edit0"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_lSkinStyle
            .ThemeColor = &HC4B0A2
            .BorderStyle ecbsNone
            .AutoBackColor = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 0, 0, 100, 20, ecsOptionButton
            .Text = "Standard Edit"
            .Checked = True
        End With
        .SubCellAddControl 1, 3, 100, 20, m_coptEdit0.hWnd, evsUserDefine, , 2, 38, True, False
        
        Set m_coptEdit1 = New clsODControl
        With m_coptEdit1
            .Name = "edit0"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_lSkinStyle
            .ThemeColor = &HC4B0A2
            .BorderStyle ecbsNone
            .AutoBackColor = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 0, 0, 100, 20, ecsOptionButton
            .Text = "Advanced Editor"
        End With
        .SubCellAddControl 1, 3, 100, 20, m_coptEdit1.hWnd, evsUserDefine, , 2, 58, True, False
        
        Set m_coptEdit2 = New clsODControl
        With m_coptEdit2
            .Name = "edit0"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_lSkinStyle
            .ThemeColor = &HC4B0A2
            .BorderStyle ecbsNone
            .AutoBackColor = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 0, 0, 100, 20, ecsOptionButton
            .Text = "Editor Locked"
        End With
        .SubCellAddControl 1, 3, 100, 20, m_coptEdit2.hWnd, evsUserDefine, , 2, 78, True, False
        
        Set m_clblBck = New clsODControl
        With m_clblBck
            .Name = "lblBck"
            .BorderStyle ecbsNone
            .LabelTransparent = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 9, 129, 44, 13, ecsLabel
            .Text = "File Backup Options"
            .AutoSize = True
            .AutoBackColor = True
        End With
        .SubCellAddControl 1, 4, 100, 20, m_clblBck.hWnd, evsUserDefine, , 2, 10, True, False
        
        Set m_cmbSave = New clsODControl
        With m_cmbSave
            .Name = "cbSave"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_lSkinStyle
            .ThemeColor = &HC4B0A2
            .BorderStyle ecbsThin
            .AutoBackColor = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 2, 11, 140, 80, ecsComboDropDown
            .AddItem "No Backup"
            .AddItem "Local Backup"
            .AddItem "Remote Backup"
            .AddItem "Converging Backup"
        End With
        .SubCellAddControl 1, 4, 140, 20, m_cmbSave.hWnd, evsUserDefine, , 2, 30, True, False
        
        Set m_chkOptn0 = New clsODControl
        With m_chkOptn0
            .Name = "chkbox0"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_lSkinStyle
            .ThemeColor = &HC4B0A2
            .BorderStyle ecbsNone
            .AutoBackColor = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 0, 0, 100, 20, ecsCheckBox
            .Text = "Use Flash Editing"
        End With
        .SubCellAddControl 1, 4, 100, 20, m_chkOptn0.hWnd, evsUserDefine, , 2, 60, True, False
        
        Set m_chkOptn1 = New clsODControl
        With m_chkOptn1
            .Name = "chkbox1"
            .HiliteColor = &HCCCCCC
            .ThemeStyle = m_lSkinStyle
            .ThemeColor = &HC4B0A2
            .BorderStyle ecbsNone
            .AutoBackColor = True
            .Visible = False
            .Create ucVHGrid1.hWnd, 0, 0, 160, 20, ecsCheckBox
            .Text = "Enable Watch-File Manager"
        End With
        .SubCellAddControl 1, 4, 160, 20, m_chkOptn1.hWnd, evsUserDefine, , 2, 80, True, False
    End With
    
End Sub

Private Sub RemoveSubcells()

    If Not m_clblDesc Is Nothing Then Set m_clblDesc = Nothing
    If Not m_coptEdit0 Is Nothing Then Set m_coptEdit0 = Nothing
    If Not m_coptEdit1 Is Nothing Then Set m_coptEdit1 = Nothing
    If Not m_coptEdit2 Is Nothing Then Set m_coptEdit2 = Nothing
    If Not m_cmbSave Is Nothing Then Set m_cmbSave = Nothing
    If Not m_clblBck Is Nothing Then Set m_clblBck = Nothing
    If Not m_chkOptn0 Is Nothing Then Set m_chkOptn0 = Nothing
    If Not m_chkOptn1 Is Nothing Then Set m_chkOptn1 = Nothing
    
End Sub

Private Function RandomNum(ByVal lBase As Long, _
                           ByVal lSpan As Long) As Long
    
    RandomNum = Int(Rnd() * lSpan) + lBase

End Function

Private Sub Form_Unload(Cancel As Integer)
    RemoveSubcells
End Sub


