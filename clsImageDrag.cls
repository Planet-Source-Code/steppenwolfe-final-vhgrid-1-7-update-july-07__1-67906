VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImageDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Type RECT
    Left                                        As Long
    Top                                         As Long
    Right                                       As Long
    bottom                                      As Long
End Type


Private Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, _
                                                                 ByVal iTrack As Long, _
                                                                 ByVal dxHotspot As Long, _
                                                                 ByVal dyHotspot As Long) As Long

Private Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()

Private Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, _
                                                                 ByVal x As Long, _
                                                                 ByVal y As Long) As Long

Private Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As Long

Private Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Long) As Long

Private Declare Function ImageList_SetDragCursorImage Lib "comctl32.dll" (ByVal himlDrag As Long, _
                                                                          ByVal iDrag As Long, _
                                                                          ByVal dxHotspot As Long, _
                                                                          ByVal dyHotspot As Long) As Long

Private Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal x As Long, _
                                                                ByVal y As Long) As Long

Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal hImageList As Long, _
                                                           ByVal hIcon As Long) As Long

Private Declare Function ImageList_Add Lib "COMCTL32" (ByVal hImageList As Long, _
                                                       ByVal hBitmap As Long, _
                                                       ByVal hBitmapMask As Long) As Long

Private Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long


Private m_bInDrag                           As Boolean
Private m_bStartDrag                        As Boolean
Private m_lImlHwnd                          As Long
Private m_lParentHnd                        As Long
Private m_lLastHwnd                         As Long


Public Property Let hImageList(ByVal himl As Long)
    m_lImlHwnd = himl
End Property

Public Property Let Parent(ByVal PropVal As Long)
    m_lParentHnd = PropVal
End Property


Public Sub StartDrag(ByVal lImageIdx As Long, _
                     Optional ByVal lX As Long = 0, _
                     Optional ByVal lY As Long = 0)
Dim lResult As Long

    CompleteDrag
    lResult = ImageList_BeginDrag(m_lImlHwnd, lImageIdx, lX, lY)
    If Not (lResult = 0) Then
        m_bInDrag = True
        m_bStartDrag = True
    End If

End Sub

Public Sub DragDrop()

Dim lX          As Long
Dim lY          As Long
Dim lParHnd     As Long

    If m_bInDrag Then
        Convert lParHnd, lX, lY
        If m_bStartDrag Then
            ImageList_DragEnter lParHnd, lX, lY
            m_lLastHwnd = lParHnd
            m_bStartDrag = False
        End If
        ImageList_DragMove lX, lY
    End If
   
End Sub

Public Sub HideDragImage(ByVal bState As Boolean)

    If m_bInDrag Then
        If bState Then
            ImageList_DragLeave m_lLastHwnd
            m_bStartDrag = True
        Else
            DragDrop
        End If
    End If

End Sub

Public Sub CompleteDrag()

    If m_bInDrag Then
        ImageList_EndDrag
        ImageList_DragLeave m_lLastHwnd
        m_lLastHwnd = 0
        m_bInDrag = False
    End If
   
End Sub

Private Sub Convert(lParHwnd As Long, _
                    lX As Long, _
                    lY As Long)

Dim tPnt  As POINTAPI
Dim tRect  As RECT
   
    GetCursorPos tPnt
    '/* convert x & y to screen coordinates
    With tPnt
        If (m_lParentHnd = 0) Then
            '/* relative to the screen
            lX = .x
            lY = .y
        Else
            lParHwnd = m_lParentHnd
            GetWindowRect lParHwnd, tRect
            lX = (.x - tRect.Left)
            lY = (.y - tRect.Top)
        End If
    End With
    
End Sub

Private Sub Class_Terminate()
    CompleteDrag
End Sub




'Private Sub CreateCursor()
'doesn't work, figures..

    'If Not (m_lImlHwnd = 0) Then
    '    DestroyCursor
    '    m_lCursorHwnd = ImageList_Create(16&, 16&, ILC_MASK Or ILC_COLOR32, 0&, 0&)
    '    Set pDragCursor = LoadResPicture("CURSOR-DRAG", vbResCursor)
    '    lhDragCursor = pDragCursor.Handle
    '    ImageList_Add m_lCursorHwnd, lhDragCursor, 0&
    '    ImageList_SetDragCursorImage m_lCursorHwnd, 1&, 0&, 0&
    'End If

'End Sub

'Private Sub DestroyCursor()

    'If Not (m_lCursorHwnd = 0) Then
    '    If ImageList_Destroy(m_lCursorHwnd) Then
    '        m_lCursorHwnd = 0
    '    End If
    'End If
    
'End Sub
