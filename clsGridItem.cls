VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGridItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_bIsInit                   As Boolean
Private m_bHasSubCells              As Boolean
Private m_bHideCheckBox             As Boolean
Private m_bRowNoFocus               As Boolean
Private m_bRowNoEdit                As Boolean
Private m_bChecked                  As Boolean
Private m_lSpanFirstCell            As Long
Private m_lSpanLastCell             As Long
Private m_lSpanRowDepth             As Long
Private m_lCellCount                As Long
Private m_lCellAlign()              As Long
Private m_lCellIconIdx()            As Long
Private m_lCellBackColor()          As Long
Private m_lCellForeColor()          As Long
Private m_lCellIndent()             As Long
Private m_hFontHnd()                As Long
Private m_lCellHeader()             As Long
Private m_lSubCellInstance()        As Long
Private m_sCellText()               As String
Private m_sCellTag                  As String


Private Sub Class_Initialize()

    m_lCellCount = -1
    m_lSpanFirstCell = -1
    m_lSpanRowDepth = -1
    ReDim m_lSubCellInstance(0)
    m_lSubCellInstance(0) = -1

End Sub

Public Property Get Text(ByVal lCellIdx As Long) As String
    Text = m_sCellText(lCellIdx)
End Property

Public Property Let Text(ByVal lCellIdx As Long, _
                         ByVal PropVal As String)
    m_sCellText(lCellIdx) = PropVal
End Property

Public Property Get Align(ByVal lCellIdx As Long) As Long
    Align = m_lCellAlign(lCellIdx)
End Property

Public Property Let Align(ByVal lCellIdx As Long, _
                          ByVal PropVal As Long)
    m_lCellAlign(lCellIdx) = PropVal
End Property

Public Property Let Checked(ByVal PropVal As Boolean)
    m_bChecked = PropVal
End Property

Public Property Get Checked() As Boolean
    Checked = m_bChecked
End Property

Public Property Get Count() As Long
    Count = m_lCellCount
End Property

Public Property Get IsInit() As Boolean
    IsInit = m_bIsInit
End Property

Public Property Get Icon(ByVal lCellIdx As Long) As Long
    Icon = m_lCellIconIdx(lCellIdx)
End Property

Public Property Let Icon(ByVal lCellIdx As Long, _
                         ByVal PropVal As Long)
    m_lCellIconIdx(lCellIdx) = PropVal
End Property

Public Property Get BackColor(ByVal lCellIdx As Long) As Long
    BackColor = m_lCellBackColor(lCellIdx)
End Property

Public Property Let BackColor(ByVal lCellIdx As Long, _
                              ByVal PropVal As Long)
    m_lCellBackColor(lCellIdx) = PropVal
End Property

Public Property Get ForeColor(ByVal lCellIdx As Long) As Long
    ForeColor = m_lCellForeColor(lCellIdx)
End Property

Public Property Let ForeColor(ByVal lCellIdx As Long, _
                              ByVal PropVal As Long)
    m_lCellForeColor(lCellIdx) = PropVal
End Property

Public Property Get FontHnd(ByVal lCellIdx As Long) As Long
    FontHnd = m_hFontHnd(lCellIdx)
End Property

Public Property Let FontHnd(ByVal lCellIdx As Long, _
                            ByVal lFontHnd As Long)
    m_hFontHnd(lCellIdx) = lFontHnd
End Property

Public Property Get Indent(ByVal lCellIdx As Long) As Long
    Indent = m_lCellIndent(lCellIdx)
End Property

Public Property Let Indent(ByVal lCellIdx As Long, _
                           ByVal PropVal As Long)
    m_lCellIndent(lCellIdx) = PropVal
End Property

Public Property Get SubHeaderCount() As Long
    SubHeaderCount = UBound(m_lCellHeader)
End Property

Public Property Get CellHeader(ByVal lCellIdx As Long) As Long
    CellHeader = m_lCellHeader(lCellIdx)
End Property

Public Property Let CellHeader(ByVal lCellIdx As Long, _
                               ByVal lInstance As Long)
    m_lCellHeader(lCellIdx) = lInstance
End Property

Public Property Get HideCheckBox() As Boolean
    HideCheckBox = m_bHideCheckBox
End Property

Public Property Let HideCheckBox(ByVal PropVal As Boolean)
    m_bHideCheckBox = PropVal
End Property

Public Property Get HasSubCells() As Boolean
    HasSubCells = m_bHasSubCells
End Property

Public Property Let HasSubCells(ByVal PropVal As Boolean)
    m_bHasSubCells = PropVal
End Property

Public Property Get SpanRowDepth() As Long
    SpanRowDepth = m_lSpanRowDepth
End Property

Public Property Let SpanRowDepth(ByVal PropVal As Long)
    m_lSpanRowDepth = PropVal
End Property

Public Property Get RowNoFocus() As Boolean
    RowNoFocus = m_bRowNoFocus
End Property

Public Property Let RowNoFocus(ByVal PropVal As Boolean)
    m_bRowNoFocus = PropVal
End Property

Public Property Get RowNoEdit() As Boolean
    RowNoEdit = m_bRowNoEdit
End Property

Public Property Let RowNoEdit(ByVal PropVal As Boolean)
    m_bRowNoEdit = PropVal
End Property

Public Property Get SpanFirstCell() As Long
    SpanFirstCell = m_lSpanFirstCell
End Property

Public Property Let SpanFirstCell(ByVal PropVal As Long)
    m_lSpanFirstCell = PropVal
End Property

Public Property Get SpanLastCell() As Long
    SpanLastCell = m_lSpanLastCell
End Property

Public Property Let SpanLastCell(ByVal PropVal As Long)
    m_lSpanLastCell = PropVal
End Property

Public Property Get SubCellCount() As Long
    SubCellCount = UBound(m_lSubCellInstance)
End Property

Public Property Let SubCellInstance(ByVal lInstance As Long)

    If (m_lSubCellInstance(0) = -1) Then
        m_lSubCellInstance(0) = lInstance
    Else
        ReDim Preserve m_lSubCellInstance(0 To SubCellCount + 1)
        m_lSubCellInstance(SubCellCount) = lInstance
    End If
    
End Property

Public Sub CellSpanHorizontal(ByVal lFirstCell As Long, _
                              ByVal lLastCell As Long)

    m_lSpanFirstCell = lFirstCell
    m_lSpanLastCell = lLastCell

End Sub

Public Sub Init(ByVal lCellCount As Long)
'/* initialize arrays to row cell count

    m_lCellCount = lCellCount
    If Not m_lCellCount = -1 Then
        ReDim m_sCellText(0 To m_lCellCount)
        ReDim m_lCellAlign(0 To m_lCellCount)
        ReDim m_lCellIconIdx(0 To m_lCellCount)
        ReDim m_lCellBackColor(0 To m_lCellCount)
        ReDim m_lCellForeColor(0 To m_lCellCount)
        ReDim m_hFontHnd(0 To m_lCellCount)
        ReDim m_lCellIndent(0 To m_lCellCount)
        ReDim m_lCellHeader(0 To m_lCellCount)
        InitCellHeader
    End If
    m_bIsInit = True

End Sub

Private Sub InitCellHeader()

Dim lCt As Long

    For lCt = 0 To UBound(m_lCellHeader)
        m_lCellHeader(lCt) = -1
    Next lCt

End Sub

Public Sub ResizeArray(ByVal lCellIndex As Long)

    If Not m_lCellCount = -1 Then
        ReDim Preserve m_sCellText(0 To lCellIndex)
        ReDim Preserve m_lCellAlign(0 To lCellIndex)
        ReDim Preserve m_lCellIconIdx(0 To lCellIndex)
        ReDim Preserve m_lCellBackColor(0 To lCellIndex)
        ReDim Preserve m_lCellForeColor(0 To lCellIndex)
        ReDim Preserve m_hFontHnd(0 To lCellIndex)
        ReDim Preserve m_lCellIndent(0 To lCellIndex)
        ReDim Preserve m_lCellHeader(0 To lCellIndex)
        m_lCellIconIdx(lCellIndex) = -1
        m_lCellBackColor(lCellIndex) = &HF8F8F8
        m_lCellForeColor(lCellIndex) = -1
        m_lCellHeader(lCellIndex) = -1
    End If
    m_lCellCount = lCellIndex

End Sub

Public Sub AddCell(ByVal lCellIndex As Long, _
                   Optional ByVal sText As String, _
                   Optional ByVal lAlign As Long, _
                   Optional ByVal lIconIdx As Long = -1, _
                   Optional ByVal lBackColor As Long = &HF8F8F8, _
                   Optional ByVal lForeColor As Long = -1, _
                   Optional ByVal lFontHandle As Long = -1, _
                   Optional ByVal lIndent As Long = 0)

On Error GoTo Handler

    If (lCellIndex > m_lCellCount) Then
        ResizeArray lCellIndex
    End If
    
    m_sCellText(lCellIndex) = sText
    m_lCellAlign(lCellIndex) = lAlign
    m_lCellIconIdx(lCellIndex) = lIconIdx
    m_lCellBackColor(lCellIndex) = lBackColor
    m_lCellForeColor(lCellIndex) = lForeColor
    m_hFontHnd(lCellIndex) = lFontHandle
    m_lCellIndent(lCellIndex) = lIndent
    m_lCellHeader(lCellIndex) = -1
    
Handler:

End Sub

Public Property Get RowTag() As String
    RowTag = m_sCellTag
End Property

Public Property Let RowTag(ByVal sText As String)
    m_sCellTag = sText
End Property

Public Sub RemoveCell(ByVal lCell As Long)

    DeleteCell lCell, m_sCellText
    DeleteCell lCell, m_lCellAlign
    DeleteCell lCell, m_lCellIconIdx
    DeleteCell lCell, m_lCellBackColor
    DeleteCell lCell, m_lCellForeColor
    DeleteCell lCell, m_lCellIndent
    DeleteCell lCell, m_lCellHeader
    If (m_lSpanLastCell = lCell) Then
        If (lCell > 1) Then
            m_lSpanLastCell = lCell - 1
        Else
            m_lSpanFirstCell = -1
            m_lSpanLastCell = -1
        End If
    ElseIf (m_lSpanFirstCell = lCell) Then
        m_lSpanFirstCell = -1
        m_lSpanLastCell = -1
    End If
    m_lCellCount = (m_lCellCount - 1)

End Sub

Private Sub DeleteCell(ByVal lCell As Long, _
                       ByRef cArray)
    
Dim lLb As Long
Dim lUb As Long
Dim lCt As Long

On Error GoTo Handler

    lLb = LBound(cArray)
    lUb = UBound(cArray)

    If (lUb = -1) Or (lUb - lLb = 0) Then
        Erase cArray
        Exit Sub
    End If

    '/* if invalid Pos
    If (lCell > lUb) Or (lCell = -1) Then
        lCell = lUb
    ElseIf lCell < lLb Then
        lCell = lLb
    ElseIf lCell = lUb Then
        ReDim Preserve cArray(lUb - 1)
        Exit Sub
    End If

    For lCt = lCell + 1 To lUb
        cArray(lCt - 1) = cArray(lCt)
    Next lCt
    ReDim Preserve cArray(lUb - 1)
    
Handler:

End Sub
