VERSION 5.00
Object = "*\A..\prjvhGrid.vbp"
Begin VB.Form frmHyperMode 
   BackColor       =   &H00F6F6F6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "vhGrid - Hyper Mode Demonstration"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8925
   StartUpPosition =   1  'CenterOwner
   Begin vhGrid.ucVHGrid ucVHGrid1 
      Height          =   5100
      Left            =   180
      TabIndex        =   2
      Top             =   225
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   8996
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
      HeaderIcons     =   "frmHyperMode.frx":0000
      CellIconSizeX   =   32
      CellIconSizeY   =   32
      CellIconColourDepth=   0
      CellIconCount   =   5
      CellIcons       =   "frmHyperMode.frx":14021
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
      Caption         =   "Add 10,000 Rows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6660
      TabIndex        =   1
      Top             =   5490
      Width           =   2040
   End
   Begin VB.PictureBox picBar 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F6F6F6&
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
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   6075
      Width           =   8925
   End
End
Attribute VB_Name = "frmHyperMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-> this is the fast way ;o)

Private Declare Sub CopyMemBr Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                  pSrc As Any, _
                                                                  ByVal lByteLen As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long


Private m_lPointer                  As Long
Private m_lRowCount                 As Long
Private m_lSkinStyle                As Long
Private m_lDecColor                 As Long
Private m_lDecOffset                As Long
Private m_lThemeClr                 As Long
Private m_lCustomClr(15)            As Long
Private m_cTiming                   As clsTiming
Private m_cGridItems()              As clsGridItem


Private Sub cmdPopulate_Click()
    LoadAppData
    cmdPopulate.Enabled = False
End Sub

Private Sub Form_Load()

Dim lX As Long

    Set m_cTiming = New clsTiming
    LoadCustomColors
    
    With ucVHGrid1
        lX = (.Width / Screen.TwipsPerPixelX) / 8
        '/* auto set draw after last cell is loaded
        .FastLoad = True
        '/* init header iml
        .InitImlHeader
        '/* init cell iml
        .InitImlRow 32, 32
        '/* add columns
        .ColumnAdd 0, "", (lX * 0.4), ecaColumnLeft, 0, ecsSortIcon
        .ColumnAdd 1, "Context", (lX * 1.4), ecaColumnLeft, 1, ecsSortDefault
        .ColumnAdd 2, "Subject", (lX * 1.4), ecaColumnLeft, 2, ecsSortDefault
        .ColumnAdd 3, "Sample Data", (lX * 2), ecaColumnLeft, 3, ecsSortDefault
        .ColumnAdd 4, "Description", (lX * 1.4), ecaColumnLeft, 4, ecsSortDefault
        .ColumnAdd 5, "Archive", (lX * 0.4), ecaColumnCenter, , ecsSortDefault
        .ColumnAdd 6, "Public", (lX * 0.4), ecaColumnCenter, , ecsSortDefault
        .ColumnAdd 7, "Current", (lX * 0.4), ecaColumnCenter, , ecsSortDefault
        '/* column tooltips
        .ColumnTipXPColors = True
        .ColumnTipColor = &H9C541F
        .ColumnTipOffsetColor = &HFEFEFE
        .ColumnTipGradient = True
        .ColumnTipMultiline = True
        .ColumnTipDelayTime = 1
        .ColumnTipTransparency = 220
        .ColumnTipHint(1) = "Application context and document scope"
        .ColumnTipHint(2) = "Document subject header and scope"
        .ColumnTipHint(3) = "A sampling of data contained in the document"
        .ColumnTipHint(4) = "Document statistics and description"
        .ColumnTipHint(5) = "This document is Archived"
        .ColumnTipHint(6) = "This document is Public"
        .ColumnTipHint(7) = "This document is Current"
        .ForeColorAuto = True
        '/* filter menu
        .FilterGradient = True
        .FilterBackColor = &HBFC9D2
        .FilterOffsetColor = &HF7F9FB
        .FilterControlColor = &HF7F9FB
        .FilterForeColor = &HEFEFEF
        .FilterTransparency = 180
        .FilterAdd 1, "FrontPage|Word|Excel|PowerPoint|Access"
        .FilterAdd 3, "securing|notes"
        '/* set header height
        .HeaderHeight = 40
        '/* no focus on first row
        .FirstRowReserved = True
        '/* set the row height
        .RowHeight = 44
        '/* enable cell tooltips
        .CellTips = True
        '/* custom cursors
        .CustomCursors = True
        '/* double buffer grid
        .DoubleBuffer = True
        '/* enable cell editing
        .CellEdit = True
        '/* blend edit text box
        .EditBlendBackground = True
        '/* enable user sizable header height
        .HeaderHeightSizable = True
        '/* lock the first column
        .LockFirstColumn = True
        '/* set alphbar transparency
        .AlphaBarTransparency = 120
        '/* enable sorting
        .CellsSorted = True
        '/* enable column drag line
        .ColumnDragLine = True
        '/* enable header drag and drop
        .HeaderDragDrop = True
        '/* set the column focus color
        .ColumnFocusColor = &H9C541F
        '/* use xp colors
        .XPColors = True
        '/* set grid back color
        .BackColor = m_lCustomClr(8)
        
        '/* lock the third column
        .ColumnLock(3) = True
        '/* enable checkboxes
        .Checkboxes = True
        '/* use gridlines
        .GridLines = EGLBoth
        '/* enable advanced edit dialog
        .AdvancedEdit = True
        '/* set the drag effect style
        .DragEffectStyle = edsClientArrow
        '/* enable header vertical text
        .ColumnVerticalText = True
        '/* apply skin
        .ThemeManager evsMetallic, False, , , &HDDDDDD, m_lCustomClr(8), m_lCustomClr(7), _
            m_lCustomClr(8), m_lCustomClr(7), &H808080, 210, True, True, m_lCustomClr(6), _
            &H0, True, False, True, False
        '/* apply cell decoration
        .CellDecoration erdCellSplit, m_lCustomClr(7), m_lCustomClr(8), True, 1
    End With

    '/* create 10k rows worth of data
    m_lRowCount = 10000
    CreateAppData
    
End Sub

Private Sub CreateAppData()

Dim lCt     As Long
Dim lRnd    As Long
Dim lCntr   As Long
Dim lCount  As Long

    lCount = m_lRowCount - 1
    lCntr = 0
    ReDim m_cGridItems(lCount)
    Do
        Set m_cGridItems(lCntr) = New clsGridItem
        lRnd = RandomNum(0, 4)
        With m_cGridItems(lCntr)
            .Init 8
            .AddCell 0
            .AddCell 1, m_sAppName(lRnd), DT_END_ELLIPSIS, lRnd
            .AddCell 2, m_sAppDesc(lRnd), DT_END_ELLIPSIS
            .AddCell 3, m_sAppData(lRnd), DT_WORDBREAK Or DT_VCENTER
            .AddCell 4, m_sAppStats(lRnd), DT_END_ELLIPSIS
            .AddCell 5, "Yes" & lCt, DT_END_ELLIPSIS
            .AddCell 6, "Yes", DT_END_ELLIPSIS
            .AddCell 7, "Yes", DT_END_ELLIPSIS
        End With
        lCntr = lCntr + 1
    Loop Until lCntr > lCount

End Sub

Private Sub LoadAppData()
'/*  load 10,000 rows in 1/100th of a second!

    '/* reset test timer
    m_cTiming.Reset
    '/* load the pointer and show
    PutArray
    '/* get time elapsed
    picBar.Cls
    picBar.Print " " & (m_lRowCount) & " rows added to vhGrid in: " & _
        Format$(m_cTiming.Elapsed / 1000, "0.0000") & "s"
    
End Sub

Private Sub PutArray()
'/* forward struct pointer into uc

On Error GoTo Handler

    If ArrayCheck(m_cGridItems) Then
        '/* test struct
        If Not (m_lPointer = 0) Then
            DestroyItems
        End If
        '/* copy struct pointer into grid control
        CopyMemBr m_lPointer, ByVal VarPtrArray(m_cGridItems), 4&
        With ucVHGrid1
            .StructPtr = m_lPointer
            '/* load the data struct
            .LoadArray
            '/* set the item count, this will fire the callback
            '/* and populate the list
            .SetRowCount UBound(m_cGridItems) + 1
            '/* turn on draw
            .Draw = True
        End With
    End If
    
Handler:
    On Error GoTo 0

End Sub

Private Function ArrayCheck(ByRef vArray As Variant) As Boolean
'/* validity test

On Error Resume Next

    '/* an array
    If Not IsArray(vArray) Then
        GoTo Handler
    '/* dimensioned
    ElseIf IsError(UBound(vArray)) Then
        GoTo Handler
    ElseIf UBound(vArray) = -1 Then
        GoTo Handler
    End If
    ArrayCheck = True

Handler:
    On Error GoTo 0

End Function

Private Function DestroyItems() As Boolean

On Error GoTo Handler

    CopyMemBr ByVal VarPtrArray(m_cGridItems), 0&, 4&
    Erase m_cGridItems
    m_lPointer = 0
    DestroyItems = True

Handler:
    On Error GoTo 0

End Function

Private Sub LoadCustomColors()

    m_lCustomClr(0) = &HF1DDCF
    m_lCustomClr(1) = &HC4B0A2
    m_lCustomClr(2) = &H887466
    m_lCustomClr(3) = &HCFDDF1
    m_lCustomClr(4) = &HA2B0C4
    m_lCustomClr(5) = &H667488
    m_lCustomClr(6) = &HCEF9EA
    m_lCustomClr(7) = &HA1CCBD
    m_lCustomClr(8) = &H7A9B8F
    m_lCustomClr(9) = &HCDBBBC
    m_lCustomClr(10) = &HC4B0A2
    m_lCustomClr(11) = &H8A8181
    m_lCustomClr(12) = &HF1C19B
    m_lCustomClr(13) = &HC4946E
    m_lCustomClr(14) = &H976741
    m_lCustomClr(15) = &HBDCCA1
    m_lDecColor = m_lCustomClr(0)
    m_lDecOffset = m_lCustomClr(1)
    m_lThemeClr = m_lCustomClr(2)
    m_lSkinStyle = 1
    
End Sub

Private Function RandomNum(ByVal lBase As Long, _
                           ByVal lSpan As Long) As Long
    
    RandomNum = Int(Rnd() * lSpan) + lBase

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DestroyItems
    Set m_cTiming = Nothing
End Sub
