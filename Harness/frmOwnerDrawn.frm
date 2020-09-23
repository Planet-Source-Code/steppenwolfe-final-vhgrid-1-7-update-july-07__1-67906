VERSION 5.00
Object = "*\A..\prjvhGrid.vbp"
Begin VB.Form frmOwnerDrawn 
   BackColor       =   &H00F6F6F6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vhGrid - Owner Drawn Cells"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin vhGrid.ucVHGrid ucVHGrid1 
      Height          =   4335
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7646
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
      HeaderIcons     =   "frmOwnerDrawn.frx":0000
      CellIconColourDepth=   0
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
   Begin VB.CommandButton cmdProgress 
      Caption         =   "Progress"
      Height          =   375
      Left            =   5850
      TabIndex        =   0
      Top             =   4770
      Width           =   1335
   End
   Begin VB.Timer tmrProgress 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmOwnerDrawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/* a simple example, but you get the idea..

Implements IOwnerDrawn


Private Sub cmdProgress_Click()

    If cmdProgress.Caption = "Progress" Then
        cmdProgress.Caption = "Stop"
        tmrProgress.Enabled = True
    Else
        cmdProgress.Caption = "Progress"
        tmrProgress.Enabled = False
    End If
   
End Sub

Private Sub Form_Load()

Dim lX As Long

    LoadColors
    With ucVHGrid1
        lX = (.Width / Screen.TwipsPerPixelX) / 8
        '/* auto set draw after last cell is loaded
        .FastLoad = True
        '/* add header icons
        .InitImlHeader
        '/* add columns
        .ColumnAdd 0, "Vert Gradient", (lX * 1.8), ecaColumnLeft, 1, ecsSortDefault
        .ColumnAdd 1, "Horz Gradient", (lX * 1.8), ecaColumnLeft, 2, ecsSortDefault
        .ColumnAdd 2, "Custom Cell", (lX * 1.8), ecaColumnLeft, 3, ecsSortDefault
        .ColumnAdd 3, "Progress", (lX * 2.2), ecaColumnLeft, 4, ecsSortDefault
        '/* use xp colors
        .XPColors = True
        '/* grid backcolor
        .BackColor = &HC4B0A2
        '/* set the row height
        .RowHeight = 35
        '/* double buffer grid
        .DoubleBuffer = True
        '/* set header height
        .HeaderHeight = 30
        '/* use gridlines
        .GridLines = EGLBoth
        '/* enable header vertical text
        .ColumnVerticalText = True
        '/* apply skin
        .ThemeManager evsClassic, True, &HC4B0A2, estThemeSoft, &HCECECE, &HDEDEDE, &HC4B0A2, _
            &H887466, &HC4B0A2, &H808080, 210, True, True, &HC4B0A2, _
            &H0, True, False, True, False
        '/* apply cell decoration
        .CellDecoration erdCellBiLinear, &HF1DDCF, &HC4B0A2, True, 2
        '/* engage ownerdrawn cells, and pass callback return param
        .OwnerDrawImpl = Me
    End With
   BuildList
   
End Sub

Private Sub BuildList()

Dim lCt     As Long

    With ucVHGrid1
        .GridInit 40, 4
        '/* add the rest of the rows
        For lCt = 0 To 40
            .AddCell lCt, 0, "", DT_LEFT Or DT_END_ELLIPSIS
            .AddCell lCt, 1, "", DT_LEFT Or DT_END_ELLIPSIS
            .AddCell lCt, 2, "", DT_LEFT Or DT_END_ELLIPSIS
            .AddCell lCt, 3, "0", DT_CENTER
        Next lCt
        '/* refresh the grid
        .GridRefresh True
    End With

End Sub

Private Sub IOwnerDrawn_Draw(GridCell As vhGrid.clsGridItem, _
                             ByVal lRow As Long, _
                             ByVal lCell As Long, _
                             ByVal lHdc As Long, _
                             ByVal eDrawStage As vhGrid.EGDDrawStage, _
                             ByVal lLeft As Long, _
                             ByVal lTop As Long, _
                             ByVal lRight As Long, _
                             ByVal lBottom As Long, _
                             bSkipDefault As Boolean)
   
   
    Select Case lCell
    Case 0
        If (eDrawStage = edgBeforeBackGround) Then
            DrawGradient lRow, lHdc, lLeft, lRight, lTop, lBottom, Fill_Vertical
            bSkipDefault = True
        End If
    Case 1
        If (eDrawStage = edgBeforeBackGround) Then
            DrawGradient lRow, lHdc, lLeft, lRight, lTop, lBottom, Fill_Horizontal
            bSkipDefault = True
        End If
    Case 2
        If (eDrawStage = edgPostDraw) Then
            ColorBox lRow, lHdc, lLeft, lRight, lTop, lBottom
        End If
    Case 3
        If (eDrawStage = edgPostDraw) Then
            DrawProgressCell GridCell, lHdc, lLeft, lRight, lTop, lBottom
        End If
    End Select
   
End Sub

Private Sub DrawProgressCell(GridCell As clsGridItem, _
                             ByVal lHdc As Long, _
                             ByVal lLeft As Long, _
                             ByVal lRight As Long, _
                             ByVal lTop As Long, _
                             ByVal lBottom As Long)

Dim lhBrush As Long
Dim tRect   As RECT

    With tRect
        .Left = lLeft + 2
        .Top = lTop + 16
        .Right = lRight - 2
        .Bottom = lBottom - 6
        .Right = .Left + ((.Right - .Left) / 100) * CLng(GridCell.Text(3))
    End With
    Gradient lHdc, tRect, RGB(10, 94, 234), RGB(10, 164, 245), Fill_Vertical
    lhBrush = CreateSolidBrush(&H303030)
    FrameRect lHdc, tRect, lhBrush
    DeleteObject lhBrush

End Sub

Private Sub ColorBox(ByVal lRow As Long, _
                     ByVal lHdc As Long, _
                     ByVal lLeft As Long, _
                     ByVal lRight As Long, _
                     ByVal lTop As Long, _
                     ByVal lBottom As Long)

Dim tRect As RECT

    With tRect
        .Left = lLeft
        .Right = lRight
        .Top = lTop
        .Bottom = lBottom
    End With
    DrawColorBox lHdc, lRow, tRect

End Sub

Private Sub tmrProgress_Timer()

Dim lCt     As Long
Dim lInt    As Long

    For lCt = 0 To (ucVHGrid1.RowCount - 1)
        If ucVHGrid1.CellText(ucVHGrid1.RowCount - 1, 3) = "100" Then
            tmrProgress.Enabled = False
            GoTo Reset
        End If
        lInt = CLng(ucVHGrid1.CellText(lCt, 3))
        If (lInt < 100) Then
            lInt = lInt + 5
            ucVHGrid1.CellText(lCt, 3) = CStr(lInt)
            ucVHGrid1.RowEnsureVisible lCt
            ucVHGrid1.GridRefresh False
            Exit For
        End If
    Next lCt

Exit Sub

Reset:
    With ucVHGrid1
        For lCt = 0 To (.RowCount - 1)
            .CellText(lCt, 3) = "0"
        Next lCt
        .GridRefresh False
    End With
    cmdProgress.Caption = "Progress"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ucVHGrid1.OwnerDrawImpl = Nothing
End Sub
