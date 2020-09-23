Attribute VB_Name = "mUtility"
Option Explicit

Private Const DT_LEFT                           As Long = &H0&
Private Const DT_VCENTER                        As Long = &H4&

Private Const CLRGREEN = "&H00FF00 &H00F000 &H00E100 &H00D200 &H00C300 &H00B400 &H00A500 &H009600 &H008700 &H007800 &H006900" & _
                        " &H9CFF9C &H8DF08D &H7EE17E &H6FD26F &H60C360 &H51B451 &H42A542 &H339633 &H248724 &H157815 &H066906" & _
                        " &HD2FFD2 &HC3F0C3 &HB4E1B4 &HA5D2A5 &H96C396 &H87B487 &H78A578 &H699669 &H5A875A &H4B784B &H3C693C" & _
                        " &HEBFFEB &HDCF0DC &HCDE1CD &HBED2BE &HAFC3AF &HA0B4A0 &H91A591 &H829682 &H738773 &H647864 &H556955"

Public Enum GRADIENT_DIRECTION
    Fill_None = -1
    Fill_Horizontal = 0
    Fill_Vertical = 1
End Enum


Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private Type POINTAPI
    x           As Long
    y           As Long
End Type

Private Type TRIVERTEX
    x           As Long
    y           As Long
    Red         As Integer
    Green       As Integer
    Blue        As Integer
    alpha       As Integer
End Type

Private Declare Function GradientFill Lib "Msimg32.dll" (ByVal hdc As Long, _
                                                         pVertex As TRIVERTEX, _
                                                         ByVal dwNumVertex As Long, _
                                                         pMesh As GRADIENT_RECT, _
                                                         ByVal dwNumMesh As Long, _
                                                         ByVal dwMode As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function FrameRect Lib "USER32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function CopyRect Lib "USER32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function InflateRect Lib "USER32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function FillRect Lib "USER32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function DrawTextA Lib "USER32" (ByVal hdc As Long, _
                                                 ByVal lpStr As String, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private m_sBoxColor()       As String


Public Sub Gradient(ByVal lHdc As Long, _
                    ByRef tRect As RECT, _
                    ByVal lStartColor As Long, _
                    ByVal lEndColor As Long, _
                    ByVal eDirection As GRADIENT_DIRECTION, _
                    Optional ByVal bJuxtapose As Boolean)

Dim btClrs(3)           As Byte
Dim btVert(7)           As Byte
Dim tGradRect           As GRADIENT_RECT
Dim tVert(1)            As TRIVERTEX
  
    '/* Check If the Fill is From Left to Right
    With tRect
        If bJuxtapose Then
            '/* Init vertices : Set Position : Define Size
            tVert(0).x = .Left
            tVert(1).x = .Left + .Right
            tVert(0).y = .Top
            tVert(1).y = .Top + .Bottom
        Else
            '/* Init vertices : Set Position : Define Size
            tVert(0).x = .Right
            tVert(1).x = .Left
            tVert(0).y = .Bottom
            tVert(1).y = .Top
        End If
    End With
    '/* Init vertices :colors, initial
    CopyMemory btClrs(0), lEndColor, &H4
    '/* Red
    btVert(1) = btClrs(0)
    '/* Green
    btVert(3) = btClrs(1)
    '/* Blue
    btVert(5) = btClrs(2)
    CopyMemory tVert(0).Red, btVert(0), &H8
    '/* Init vertices :colors, final
    CopyMemory btClrs(0), lStartColor, &H4
    '/* Red
    btVert(1) = btClrs(0)
    '/* Green
    btVert(3) = btClrs(1)
    '/* Blue
    btVert(5) = btClrs(2)
    CopyMemory tVert(1).Red, btVert(0), &H8
    '/* Init gradient rect
    With tGradRect
        .UpperLeft = 0
        .LowerRight = 1
    End With
    '/* Fill the DC
    GradientFill lHdc, tVert(0), 2, tGradRect, 1, eDirection

End Sub

Public Sub LoadColors()

Dim sColor  As String

    '/* split color const
    sColor = CLRGREEN
    m_sBoxColor = Split(sColor, Chr$(32))

End Sub

Public Sub DrawGradient(ByVal lRow As Long, _
                         ByVal lHdc As Long, _
                         ByVal lLeft As Long, _
                         ByVal lRight As Long, _
                         ByVal lTop As Long, _
                         ByVal lBottom As Long, _
                         ByVal eGradDir As GRADIENT_DIRECTION)

Dim tRect   As RECT

    With tRect
        .Left = lLeft
        .Right = lRight
        .Top = lTop
        .Bottom = lBottom
    End With
    
    Gradient lHdc, tRect, &HFFFFFF, &H887466, eGradDir
    SetBkMode lHdc, 1&
    InflateRect tRect, -5, -10
    DrawTextA lHdc, ("Row: " & lRow & Chr(0)), -1, tRect, DT_LEFT Or DT_VCENTER

End Sub

Public Sub DrawColorBox(ByVal lHdc As Long, _
                        ByVal lIColorIdx As Long, _
                        ByRef tRect As RECT)

Dim lhPen       As Long
Dim lhPenOld    As Long
Dim lhBrush     As Long
Dim tPnt        As POINTAPI
Dim tRcpy       As RECT

    CopyRect tRcpy, tRect
    With tRcpy
        .Left = .Left + 2
        .Right = .Left + 14
        .Top = .Top + ((.Bottom - .Top) - 14) / 2
        .Bottom = .Top + 14
    End With
    With tRcpy
        MoveToEx lHdc, .Left, .Top, tPnt
        lhPen = CreatePen(0&, 1, &H0)
        lhPenOld = SelectObject(lHdc, lhPen)
        LineTo lHdc, .Right - 1, .Top
        LineTo lHdc, .Right - 1, .Bottom - 1
        LineTo lHdc, .Left, .Bottom - 1
        LineTo lHdc, .Left, .Top
    End With
    
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen
    lhBrush = CreateSolidBrush(CLng(m_sBoxColor(lIColorIdx)))
    
    InflateRect tRcpy, -1, -1
    FillRect lHdc, tRcpy, lhBrush
    DeleteObject lhBrush
    With tRcpy
        .Left = .Left + 16
        .Right = tRect.Right
    End With
    SetTextColor lHdc, vbBlack
    SetBkMode lHdc, 1&
    DrawTextA lHdc, ("&" & m_sBoxColor(lIColorIdx) & Chr(0)), -1, tRcpy, DT_LEFT Or DT_VCENTER

End Sub


