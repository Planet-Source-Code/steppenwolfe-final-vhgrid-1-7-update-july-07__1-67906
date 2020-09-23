VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCellHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Enum ECCFrameConnector
    escTopCap = 1&
    escJoined = 2&
    escBottomCap = 3&
End Enum


Private m_bUseSpannedCell                       As Boolean
Private m_lRowIndex                             As Long
Private m_lCellIndex                            As Long
Private m_lAlignFlag                            As Long
Private m_lFontHandle                           As Long
Private m_lForeColor                            As Long
Private m_lFocusForeColor                       As Long
Private m_lIndent                               As Long
Private m_sText                                 As String
Private m_eFrameConnector                       As ECCFrameConnector


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

Public Property Get Indent() As Long
    Indent = m_lIndent
End Property

Public Property Let Indent(ByVal PropVal As Long)
    m_lIndent = PropVal
End Property

Public Property Get AlignFlag() As Long
    AlignFlag = m_lAlignFlag
End Property

Public Property Let AlignFlag(ByVal PropVal As Long)
    m_lAlignFlag = PropVal
End Property

Public Property Get FontHandle() As Long
    FontHandle = m_lFontHandle
End Property

Public Property Let FontHandle(ByVal PropVal As Long)
    m_lFontHandle = PropVal
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_lForeColor
End Property

Public Property Let ForeColor(ByVal PropVal As Long)
    m_lForeColor = PropVal
End Property

Public Property Get FocusForeColor() As Long
    FocusForeColor = m_lFocusForeColor
End Property

Public Property Let FocusForeColor(ByVal PropVal As Long)
    m_lFocusForeColor = PropVal
End Property

Public Property Get UseSpannedCell() As Boolean
    UseSpannedCell = m_bUseSpannedCell
End Property

Public Property Let UseSpannedCell(ByVal PropVal As Boolean)
    m_bUseSpannedCell = PropVal
End Property

Public Property Get Text() As String
    Text = m_sText
End Property

Public Property Let Text(ByVal PropVal As String)
    m_sText = PropVal
End Property

Public Property Get FrameConnector() As ECCFrameConnector
    FrameConnector = m_eFrameConnector
End Property

Public Property Let FrameConnector(ByVal PropVal As ECCFrameConnector)
    m_eFrameConnector = PropVal
End Property
