VERSION 5.00
Object = "*\A..\prjvhGrid.vbp"
Begin VB.Form frmUnicode 
   BackColor       =   &H00F6F6F6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vhGrid - Unicode"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11160
   StartUpPosition =   1  'CenterOwner
   Begin vhGrid.ucVHGrid ucVHGrid1 
      Height          =   6585
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   11615
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderIconColourDepth=   8
      HeaderIconCount =   5
      HeaderIcons     =   "frmUnicode.frx":0000
      CellIconSizeX   =   32
      CellIconSizeY   =   32
      CellIconColourDepth=   0
      CellIconCount   =   21
      CellIcons       =   "frmUnicode.frx":14021
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
      Left            =   9000
      TabIndex        =   0
      Top             =   7020
      Width           =   2040
   End
End
Attribute VB_Name = "frmUnicode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_aData()    As String


Private Sub cmdPopulate_Click()

    '/* add the data
    BuildList
    
End Sub

Private Sub Form_Load()

Dim lX As Long

    '/* load res data
    LoadData
    
    With ucVHGrid1
        lX = (.Width / Screen.TwipsPerPixelX) / 8
        '/* auto set draw after last cell is loaded
        .FastLoad = True
        '/* enable unicode
        .UseUnicode = True
        '/* init header icons
        .InitImlHeader
        '/* init row icons
        .InitImlRow 32, 32
        '/* add columns
        .ColumnAdd 0, "", (lX * 0.4), ecaColumnLeft, 0, ecsSortIcon
        .ColumnAdd 1, m_aData(0), (lX * 1.8), ecaColumnLeft, 1, ecsSortDefault
        .ColumnAdd 2, m_aData(1), (lX * 1.8), ecaColumnLeft, 2, ecsSortDefault
        .ColumnAdd 3, m_aData(2), (lX * 1.8), ecaColumnLeft, 3, ecsSortDefault
        .ColumnAdd 4, m_aData(3), (lX * 1.8), ecaColumnLeft, 4, ecsSortDefault
        '/* use xp colors
        .XPColors = True
        '/* grid backcolor
        .BackColor = &HC4B0A2
        '/* set the row height
        .RowHeight = 35
        '/* double buffer grid
        .DoubleBuffer = True
        '/* enable cell editing
        .CellEdit = True
        '/* lock the first column
        .LockFirstColumn = True
        '/* set alphbar transparency
        .AlphaBarTransparency = 120
        '/* enable sorting
        .CellsSorted = True
        '/* enable header drag and drop
        .HeaderDragDrop = True
        '/* set header height
        .HeaderHeight = 35
        '/* enable checkboxes
        .Checkboxes = True
        '/* use gridlines
        .GridLines = EGLBoth
        '/* set the drag effect style
        .DragEffectStyle = edsClientArrow
        '/* enable header vertical text
        .ColumnVerticalText = True
        '/* apply skin: theme, colorized, theme color, color luminence,
        '/* column font color, column font hilite color, column font preseed color
        '/* options backcolor, options offset color, options forecolor
        '/* options transparency, options use gradient, options use xp colors
        '/* options control color, options control forecolor
        '/* include: advanced edit, column tip, filter, cell tip
        .ThemeManager evsAzure, False, , , &H333333, &HDFDFDF, &HFFFFFF, _
            &H887466, &HC4B0A2, &H808080, 210, True, True, &HC4B0A2, _
            &H0, True, False, True, False
        '/* apply cell decoration
        .CellDecoration erdCellBiLinear, &HF1DDCF, &HC4B0A2, True, 2
    End With
    
End Sub

Private Sub LoadData()

    '/* init temp array
    ReDim m_aData(13)
    '/* unicode: base array
    m_aData(0) = LoadResString(105)
    m_aData(1) = LoadResString(108)
    m_aData(2) = LoadResString(109)
    m_aData(3) = LoadResString(110)
    m_aData(4) = LoadResString(101) & "|" & LoadResString(121)
    m_aData(5) = LoadResString(102) & "|" & LoadResString(122)
    m_aData(6) = LoadResString(103) & "|" & LoadResString(123)
    m_aData(7) = LoadResString(104) & "|" & LoadResString(124)
    m_aData(8) = LoadResString(105) & "|" & LoadResString(125)
    m_aData(9) = LoadResString(113) & "|" & LoadResString(123)
    m_aData(10) = LoadResString(107) & "|" & LoadResString(127)
    m_aData(11) = LoadResString(108) & "|" & LoadResString(128)
    m_aData(12) = LoadResString(109) & "|" & LoadResString(129)
    m_aData(13) = LoadResString(110) & "|" & LoadResString(130)
    
End Sub

Private Sub BuildList()

Dim lCt     As Long
Dim lRnd    As Long

    With ucVHGrid1
        '/* if a second time around
        If .GridInitialized Then
            .ClearList
        End If
        '/* if adding rows one at a a time, first init grid
        '/* with one row and number of columns
        .GridInit 1, 5
        '/* add the rest of the rows
        For lCt = 0 To 50
            lRnd = RandomNum(3, 10)
            .AddCell lCt, 0
            .AddCell lCt, 1, m_aData(lRnd), DT_LEFT Or DT_END_ELLIPSIS, lRnd
            .AddCell lCt, 2, m_aData(lRnd), DT_LEFT Or DT_END_ELLIPSIS
            .AddCell lCt, 3, m_aData(lRnd), DT_LEFT Or DT_END_ELLIPSIS
            .AddCell lCt, 4, m_aData(lRnd), DT_LEFT Or DT_END_ELLIPSIS
        Next lCt
        '/* refresh the grid
        .GridRefresh True
    End With

End Sub

Private Function RandomNum(ByVal lBase As Long, _
                           ByVal lSpan As Long) As Long

    RandomNum = Int(Rnd() * lSpan) + lBase

End Function

