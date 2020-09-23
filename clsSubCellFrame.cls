VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Enum ESCFramePosition
    escUserDefine = 0&
    escTopLeft = 1&
    escTopCenter = 2&
    escTopRight = 3&
    escCenterLeft = 4&
    escCenterCell = 5&
    escCenterRight = 6&
    escBottomLeft = 7&
    escBottomCenter = 8&
    escBottomRight = 9&
End Enum

Public Enum ESCFrameConnector
    escTopCap = 1&
    escJoined = 2&
    escBottomCap = 3&
End Enum

Private m_bSpanUseVirtualRow                    As Boolean
Private m_bSpanUseVirtualCell                   As Boolean
Private m_lRowIndex                             As Long
Private m_lCellIndex                            As Long
Private m_lLeft                                 As Long
Private m_lTop                                  As Long
Private m_lWidth                                As Long
Private m_lHeight                               As Long
Private m_lCntlHwnd                             As Long
Private m_eFramePosition                        As ESCFramePosition
Private m_eFrameConnector                       As ESCFrameConnector


Public Property Get RowIndex() As Long
    RowIndex = m_lRowIndex
End Property

Public Property Let RowIndex(ByVal PropVal As Long)
    m_lRowIndex = PropVal
End Property

Public Property Get CellIndex() As Long
    CellIndex = m_lCellIndex
End Property

Public Property Let CellIndex(ByVal PropVal As Long)
    m_lCellIndex = PropVal
End Property

Public Property Get Left() As Long
    Left = m_lLeft
End Property

Public Property Let Left(ByVal PropVal As Long)
    m_lLeft = PropVal
End Property

Public Property Get Top() As Long
    Top = m_lTop
End Property

Public Property Let Top(ByVal PropVal As Long)
    m_lTop = PropVal
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

Public Property Get FramePosition() As ESCFramePosition
    FramePosition = m_eFramePosition
End Property

Public Property Let FramePosition(ByVal PropVal As ESCFramePosition)
    m_eFramePosition = PropVal
End Property

Public Property Get FrameConnector() As ESCFrameConnector
    FrameConnector = m_eFrameConnector
End Property

Public Property Let FrameConnector(ByVal PropVal As ESCFrameConnector)
    m_eFrameConnector = PropVal
End Property

Public Property Get CntlHwnd() As Long
    CntlHwnd = m_lCntlHwnd
End Property

Public Property Let CntlHwnd(ByVal PropVal As Long)
    m_lCntlHwnd = PropVal
End Property

Public Property Get SpanUseVirtualRow() As Boolean
    SpanUseVirtualRow = m_bSpanUseVirtualRow
End Property

Public Property Let SpanUseVirtualRow(ByVal PropVal As Boolean)
    m_bSpanUseVirtualRow = PropVal
End Property

Public Property Get SpanUseVirtualCell() As Boolean
    SpanUseVirtualCell = m_bSpanUseVirtualCell
End Property

Public Property Let SpanUseVirtualCell(ByVal PropVal As Boolean)
    m_bSpanUseVirtualCell = PropVal
End Property
