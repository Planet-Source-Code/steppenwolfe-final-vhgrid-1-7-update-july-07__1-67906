VERSION 5.00
Object = "*\A..\prjvhGrid.vbp"
Begin VB.Form frmVirtual 
   BackColor       =   &H00F6F6F6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vhGrid -Virtual Grid"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8220
   StartUpPosition =   1  'CenterOwner
   Begin vhGrid.ucVHGrid ucVHGrid1 
      Height          =   5325
      Left            =   225
      TabIndex        =   4
      Top             =   225
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   9393
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
      HeaderIcons     =   "frmVirtual.frx":0000
      CellIconSizeX   =   32
      CellIconSizeY   =   32
      CellIconColourDepth=   0
      CellIconCount   =   20
      CellIcons       =   "frmVirtual.frx":14021
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
   Begin VB.TextBox txtCount 
      Height          =   285
      Left            =   6615
      TabIndex        =   3
      Text            =   "1000000"
      Top             =   5805
      Width           =   825
   End
   Begin VB.PictureBox picBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   8160
      TabIndex        =   2
      Top             =   6240
      Width           =   8220
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Fill Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4815
      TabIndex        =   1
      Top             =   5715
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "# Rows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7470
      TabIndex        =   0
      Top             =   5850
      Width           =   570
   End
End
Attribute VB_Name = "frmVirtual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'/* not supported in virtual mode:
'-> row spanning
'-> cell spanning
'-> column filter
'-> subcells
'-> cell tips
'-> cell headers
'/* All sorting, item changes, and data management are the responsibility of
'/* the database client.

Private m_lRowCount     As Long
Private m_cTiming       As clsTiming


Private Sub Form_Load()

Dim lX As Long
    
    Set m_cTiming = New clsTiming
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
        .ColumnAdd 1, "Column 1", (lX * 1.8), ecaColumnLeft, 1, ecsSortDefault
        .ColumnAdd 2, "Column 2", (lX * 1.8), ecaColumnLeft, 2, ecsSortDefault
        .ColumnAdd 3, "Column 3", (lX * 1.8), ecaColumnLeft, 3, ecsSortDefault
        .ColumnAdd 4, "Column 4", (lX * 1.8), ecaColumnLeft, 4, ecsSortDefault
        '/* use xp colors
        .XPColors = True
        '/* grid backcolor
        .BackColor = &H8A8181
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
        .HeaderHeight = 50
        '/* enable checkboxes
        .Checkboxes = True
        '/* use gridlines
        .GridLines = EGLBoth
        '/* set the drag effect style
        .DragEffectStyle = edsClientArrow
        '/* enable header vertical text
        .ColumnVerticalText = True
        '/* enable virtual mode
        .VirtualMode = True
        '/* apply skin
        .ThemeManager evsXpGreen, False, , , &H333333, &H887466, &HC4B0A2, _
            &HF1C19B, &HC4946E, &H808080, 210, &H0, True, True, &HF1C19B, _
            True, False, True, False
        '/* apply cell decoration: layout, startcolor, offset color, xp color offset, depth
        .CellDecoration erdCellChecker, &HCDBBBC, &H8A8181, True, 1
    End With

End Sub

Private Sub cmdLoad_Click()

    If TestInput Then
        m_cTiming.Reset
        InitGrid
        picBar.Cls
        picBar.Print (m_lRowCount) & " virtual rows added to vhGrid in: " & _
            Format$(m_cTiming.Elapsed / 1000, "0.0000") & "s"
    End If

End Sub

Private Function TestInput() As Boolean
    
    If Not IsNumeric(txtCount.Text) Then
        MsgBox "Please choose a row number between 1 and 10000000", vbExclamation, "Invalid Input!"
        Exit Function
    ElseIf (txtCount.Text) > 10000000 Then
        MsgBox "Please choose a row number between 1 and 10000000", vbExclamation, "Invalid Input!"
        Exit Function
    ElseIf CLng(txtCount.Text) < 1 Then
        MsgBox "Please choose a row number between 1 and 10000000", vbExclamation, "Invalid Input!"
        Exit Function
    ElseIf ucVHGrid1.Count > 0 Then
        ucVHGrid1.ClearList
    End If
    '/* set the row count
    m_lRowCount = CLng((txtCount.Text))
    '/* success
    TestInput = True

End Function

Private Sub InitGrid()

    With ucVHGrid1
        '/* init the list
        .GridInit m_lRowCount, 5
        .Draw = True
    End With
    
End Sub

Private Sub ucVHGrid1_eVHVirtualAccess(ByVal lRow As Long, _
                                       ByVal lCell As Long, _
                                       sText As String, _
                                       lIcon As Long)
    
    '/* from database return corresponding row/field
    '/* data into the grid. Ex.
    ' .move lRow
    'If lCell = 1 Then
        'lIcon = .fetchfield(1)
    'Else
        'sText = .fetchfield(lCell)
    'End If
    Select Case lCell
    '/* item
    Case 0
        sText = ""
        lIcon = -1
    Case 1
        sText = "Row: " & Format$(lRow, "#,###,##0") & ", First Cell"
        lIcon = Left(lRow, 1)
    Case Else
        sText = "Row: " & lRow & ", Cell: " & lCell
    End Select
    
End Sub

Private Sub ucVHGrid1_DragDrop(Source As Control, x As Single, y As Single)
'item drag
End Sub

Private Sub ucVHGrid1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
'drag over item
End Sub

Private Sub ucVHGrid1_eVHAdvancedEditChange(ByVal lRow As Long, ByVal lCell As Long, ByVal lIcon As Long, ByVal lBackColor As Long, ByVal lForeColor As Long, ByVal sText As String)
'advanced edit data change
End Sub

Private Sub ucVHGrid1_eVHAdvancedEditRequest(ByVal lRow As Long, ByVal lCell As Long)
'advanced edit loaded
End Sub

Private Sub ucVHGrid1_eVHAdvancedEditRequestText(ByVal lRow As Long, ByVal lCell As Long, lIcon As Long, sText As String)
'advanced edit requesting text
End Sub

Private Sub ucVHGrid1_eVHColumnAdded(ByVal lColumn As Long, ByVal lWidth As Long, ByVal lIcon As Long, ByVal sText As String)
'column has been added
End Sub

Private Sub ucVHGrid1_eVHColumnClick(ByVal lColumn As Long)
'/* column clicked
    Debug.Print "ColumnClick " & lColumn
End Sub

Private Sub ucVHGrid1_eVHColumnDragComplete()
'column drag completed
End Sub

Private Sub ucVHGrid1_eVHColumnDragging(ByVal lColumn As Long)
'column is dragging
End Sub

Private Sub ucVHGrid1_eVHColumnHorizontalSize(ByVal lColumn As Long)
'header horizontal size change
End Sub

Private Sub ucVHGrid1_eVHColumnRemoved(ByVal lColumn As Long)
'column has been removed
End Sub

Private Sub ucVHGrid1_eVHColumnVerticalSize(ByVal lHeight As Long)
'column verical size changed
End Sub

Private Sub ucVHGrid1_eVHEditChange(ByVal lRow As Long, ByVal lCell As Long, ByVal sText As String)
'edit changed text
End Sub

Private Sub ucVHGrid1_eVHEditRequest(ByVal lRow As Long, ByVal lCell As Long)
'edit loaded
End Sub

Private Sub ucVHGrid1_eVHEditRequestText(ByVal lRow As Long, ByVal lCell As Long, sText As String)
'edit requesting text
End Sub

Private Sub ucVHGrid1_eVHErrCond(ByVal sRtn As String, ByVal lErr As Long)
'/* grid error
    Debug.Print "Error: " & sRtn & " #" & lErr
End Sub

Private Sub ucVHGrid1_eVHGridEnable(ByVal bState As Boolean)
'grid enable state change
End Sub

Private Sub ucVHGrid1_eVHGridSizeChange(ByVal lWidth As Long, ByVal lHeight As Long)
'grid size changing
End Sub

Private Sub ucVHGrid1_eVHItemCheck(ByVal lRow As Long, ByVal bState As Boolean)
'/* item check state change
    Debug.Print "check: " & lRow & " " & bState
End Sub

Private Sub ucVHGrid1_eVHItemClick(ByVal lRow As Long, ByVal lCell As Long)
'/* item clicked
    Debug.Print "Item click: Row " & lRow & ", cell " & lCell
End Sub

Private Sub ucVHGrid1_eVHItemDeleted(ByVal lRow As Long)
'item deleted
End Sub

Private Sub ucVHGrid1_eVHItemDragComplete(ByVal lSource As Long, ByVal lTarget As Long)
'item drag completed
End Sub

Private Sub ucVHGrid1_eVHItemDragging(ByVal lRow As Long)
'item dragging
End Sub

Private Sub ucVHGrid1_GotFocus()
'grid has focus
End Sub

Private Sub ucVHGrid1_LostFocus()
'grid lost focus
End Sub

Private Sub ucVHGrid1_Validate(Cancel As Boolean)
'validate drag data
End Sub
