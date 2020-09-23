VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GXMSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/* Paul Catons excellent winsubhook3 subclasser with a few minor alterations
'/* Unicode support added Oct. 18 2006

Private Const MSG_ENTRIES                       As Long = 32
Private Const WNDPROC_OFF                       As Long = &H38
Private Const GWL_WNDPROC                       As Long = -4
Private Const CODE_LEN                          As Long = 260
Private Const MEM_LEN                           As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))
Private Const PAGE_RWX                          As Long = &H40&
Private Const MEM_COMMIT                        As Long = &H1000&
Private Const MEM_RELEASE                       As Long = &H8000&
Private Const IDX_SHUTDOWN                      As Long = 1
Private Const IDX_HWND                          As Long = 2
Private Const IDX_EBMODE                        As Long = 3
Private Const IDX_CWP                           As Long = 4
Private Const IDX_SWL                           As Long = 5
Private Const IDX_FREE                          As Long = 6
Private Const IDX_BADPTR                        As Long = 7
Private Const IDX_OWNER                         As Long = 8
Private Const IDX_WNDPROC                       As Long = 9
Private Const IDX_CALLBACK                      As Long = 10
Private Const IDX_BTABLE                        As Long = 11
Private Const IDX_ATABLE                        As Long = 12
Private Const IDX_PARM_USER                     As Long = 13
Private Const IDX_EBX                           As Long = 16
Private Const SUB_NAME                          As String = "GXMSubclass"


Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hwnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long

Private Declare Function CallWindowProcW Lib "user32" (ByVal lpPrevWndFunc As Long, _
                                                       ByVal hwnd As Long, _
                                                       ByVal Msg As Long, _
                                                       ByVal wParam As Long, _
                                                       ByVal lParam As Long) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, _
                                                                lpdwProcessId As Long) As Long

Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long

Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, _
                                                      ByVal dwSize As Long, _
                                                      ByVal flAllocationType As Long, _
                                                      ByVal flProtect As Long) As Long

Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, _
                                                     ByVal dwSize As Long, _
                                                     ByVal dwFreeType As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, _
                                                  ByVal Source As Long, _
                                                  ByVal Length As Long)

Private Declare Function GetVersion Lib "kernel32" () As Long

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long


Private m_bIsNt                                 As Boolean
Private m_lResMem                               As Long
Private m_lTData(64)                            As Long
Private m_cWndHnd                               As Collection


Private Sub Class_Initialize()
    Version
End Sub

Private Function Version()

Dim lVer As Long

    lVer = GetVersion()
    m_bIsNt = ((lVer And &H80000000) = 0)
    
End Function

Public Property Get p_ParamUser(ByVal lHwnd As Long) As Long
'Get the subclasser p_ParamUser callback parameter

    If (IsBadCodePtr(MapHandle(lHwnd)) = 0) Then
        p_ParamUser = p_CAddress(IDX_PARM_USER)
    End If

End Property

Public Property Let p_ParamUser(ByVal lHwnd As Long, _
                                ByVal lValue As Long)
'Let the subclasser p_ParamUser callback parameter

    If (IsBadCodePtr(MapHandle(lHwnd)) = 0) Then
        p_CAddress(IDX_PARM_USER) = lValue
    End If
  
End Property

Private Property Get p_CAddress(ByVal lIndex As Long) As Long
    RtlMoveMemory VarPtr(p_CAddress), m_lResMem + (lIndex * 4), 4&
End Property

Private Property Let p_CAddress(ByVal lIndex As Long, ByVal lValue As Long)
    RtlMoveMemory m_lResMem + (lIndex * 4), VarPtr(lValue), 4&
End Property

Public Function Subclass(ByVal lHwnd As Long, _
                         ByVal oCallback As GXISubclass, _
                         Optional ByVal lParamUser As Long = 0, _
                         Optional ByVal bIdeSafety As Boolean = True) As Boolean

Dim lAddr         As Long
Dim lID           As Long
Dim lProc         As Long
  
    'Ensure the window handle is valid
    If (IsWindow(lHwnd) = 0) Then
        ErrorCond SUB_NAME, "Invalid window handle"
        Exit Function
    End If

    'Get this process's ID
    lProc = GetCurrentProcessId
    'Get the process ID associated with the window handle
    GetWindowThreadProcessId lHwnd, lID
    'Ensure that the window handle doesn't belong to another process
    If Not (lID = lProc) Then
        ErrorCond SUB_NAME, "Window handle belongs to another process"
        Exit Function
    End If
  
  'If this is the first time through, do the one-time initialization
    If (m_cWndHnd Is Nothing) Then
        'Create the hWnd/thunk-address collection
        Set m_cWndHnd = New Collection
        BuildTable
     End If

    'Get the address of the owner's vTable
    RtlMoveMemory VarPtr(lAddr), ObjPtr(oCallback), 4&
    'Get the address of the implemented interface
    RtlMoveMemory VarPtr(lAddr), lAddr + &H1C, 4&
    'Allocate executable memory
    m_lResMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)

    'Ensure the allocation succeeded
    If Not (m_lResMem = 0) Then
        'Catch double subclassing
        On Error GoTo CatchDoubleSub
        'Add the hWnd/thunk-address to the collection
        m_cWndHnd.Add m_lResMem, "h" & lHwnd
        On Error GoTo 0
        'If the user wants IDE protection
        'Store the EbMode function address in the thunk data
        If bIdeSafety Then
            m_lTData(IDX_EBMODE) = ReturnAddr("vba6", "EbMode")
        End If
        'Patch the thunk data address
        m_lTData(IDX_EBX) = m_lResMem
        'Store the window handle in the thunk data
        m_lTData(IDX_HWND) = lHwnd
        'Store the address of the before table in the thunk data
        m_lTData(IDX_BTABLE) = m_lResMem + CODE_LEN
        'Store the address of the after table in the thunk data
        m_lTData(IDX_ATABLE) = m_lResMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)
        'Store owner's object address in the thunk data
        m_lTData(IDX_OWNER) = ObjPtr(oCallback)
        'Store the callback address in the thunk data
        m_lTData(IDX_CALLBACK) = lAddr
        'Store the lParamUser callback parameter in the thunk data
        m_lTData(IDX_PARM_USER) = lParamUser
        'Set the new WndProc, return the address of the original WndProc
        If m_bIsNt Then
            lAddr = SetWindowLongW(lHwnd, GWL_WNDPROC, m_lResMem + WNDPROC_OFF)
        Else
            lAddr = SetWindowLongA(lHwnd, GWL_WNDPROC, m_lResMem + WNDPROC_OFF)
        End If
        'Ensure the new WndProc was set correctly
        If (lAddr = 0) Then
            ErrorCond SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
            GoTo ReleaseMemory
        End If
        'Store the original WndProc address in the thunk data
        m_lTData(IDX_WNDPROC) = lAddr
        'Copy the thunk code/data to the allocated memory
        RtlMoveMemory m_lResMem, VarPtr(m_lTData(0)), CODE_LEN
        Subclass = True
    Else
        ErrorCond SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
    End If
  
Exit Function

CatchDoubleSub:
    ErrorCond SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
    VirtualFree m_lResMem, 0, MEM_RELEASE

End Function

Private Sub BuildTable()

    m_lTData(14) = &HD231C031
    m_lTData(15) = &HBBE58960
    m_lTData(17) = &H4339F631
    m_lTData(18) = &H4A21750C
    m_lTData(19) = &HE82C7B8B
    m_lTData(20) = &H74&
    m_lTData(21) = &H75147539
    m_lTData(22) = &H21E80F
    m_lTData(23) = &HD2310000
    m_lTData(24) = &HE8307B8B
    m_lTData(25) = &H60&
    m_lTData(26) = &H10C261
    m_lTData(27) = &H830C53FF
    m_lTData(28) = &HD77401F8
    m_lTData(29) = &H2874C085
    m_lTData(30) = &H2E8&
    m_lTData(31) = &HFFE9EB00
    m_lTData(32) = &H75FF3075
    m_lTData(33) = &H2875FF2C
    m_lTData(34) = &HFF2475FF
    m_lTData(35) = &H3FF2473
    m_lTData(36) = &H891053FF
    m_lTData(37) = &HBFF1C45
    m_lTData(38) = &H73396775
    m_lTData(39) = &H58627404
    m_lTData(40) = &H6A2473FF
    m_lTData(41) = &H873FFFC
    m_lTData(42) = &H891453FF
    m_lTData(43) = &H7589285D
    m_lTData(44) = &H3045C72C
    m_lTData(45) = &H8000&
    m_lTData(46) = &H8920458B
    m_lTData(47) = &H4589145D
    m_lTData(48) = &HC4836124
    m_lTData(49) = &H1862FF04
    m_lTData(50) = &H35E30F8B
    m_lTData(51) = &HA78C985
    m_lTData(52) = &H8B04C783
    m_lTData(53) = &HAFF22845
    m_lTData(54) = &H73FF2775
    m_lTData(55) = &H1C53FF28
    m_lTData(56) = &H438D1F75
    m_lTData(57) = &H144D8D34
    m_lTData(58) = &H1C458D50
    m_lTData(59) = &HFF3075FF
    m_lTData(60) = &H75FF2C75
    m_lTData(61) = &H873FF28
    m_lTData(62) = &HFF525150
    m_lTData(63) = &H53FF2073
    m_lTData(64) = &HC328&
    
    If m_bIsNt Then
        'Store CallWindowProcW function address in the thunk data
        m_lTData(IDX_CWP) = ReturnAddr("user32", "CallWindowProcW")
        'Store the SetWindowLongw function address in the thunk data
        m_lTData(IDX_SWL) = ReturnAddr("user32", "SetWindowLongW")
    Else
        'Store CallWindowProcA function address in the thunk data
        m_lTData(IDX_CWP) = ReturnAddr("user32", "CallWindowProcA")
        'Store the SetWindowLongA function address in the thunk data
        m_lTData(IDX_SWL) = ReturnAddr("user32", "SetWindowLongA")
    End If
    
    'Store the VirtualFree function address in the thunk data
    m_lTData(IDX_FREE) = ReturnAddr("kernel32", "VirtualFree")
    'Store the IsBadCodePtr function address in the thunk data
    m_lTData(IDX_BADPTR) = ReturnAddr("kernel32", "IsBadCodePtr")

End Sub

Public Sub AddMessage(ByVal lHwnd As Long, _
                      ByVal uMsg As eMsg, _
                      Optional ByVal eWhen As eMsgWhen)
'Add the message value to the window handle's specified callback table

    'Ensure that the thunk hasn't already released its memory
    If (IsBadCodePtr(MapHandle(lHwnd)) = 0) Then
        'Add the message to the before table
        If (eWhen And MSG_BEFORE) Then
            AddMsg uMsg, IDX_BTABLE
        End If
        'Add the message to the after table
        If (eWhen And MSG_AFTER) Then
            AddMsg uMsg, IDX_ATABLE
        End If
    End If

End Sub

Public Sub DeleteMessage(ByVal lHwnd As Long, _
                         ByVal uMsg As eMsg, _
                         Optional ByVal eWhen As eMsgWhen)
'Delete the message value from the window handle's specified callback table

    'Ensure that the thunk hasn't already released its memory
    If (IsBadCodePtr(MapHandle(lHwnd)) = 0) Then
        If (eWhen And MSG_BEFORE) Then
            DelMsg uMsg, IDX_BTABLE
        End If
        If (eWhen And MSG_AFTER) Then
            DelMsg uMsg, IDX_ATABLE
        End If
    End If
  
End Sub

Public Function CallOldWndProc(ByVal lHwnd As Long, _
                               ByVal uMsg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long
'Call the original WndProc

    'Ensure that the thunk hasn't already released its memory
    If (IsBadCodePtr(MapHandle(lHwnd)) = 0) Then
        'Call the original WndProc of the passed window handle parameter
        If m_bIsNt Then
            CallOldWndProc = CallWindowProcW(p_CAddress(IDX_WNDPROC), lHwnd, uMsg, wParam, lParam)
        Else
            CallOldWndProc = CallWindowProcA(p_CAddress(IDX_WNDPROC), lHwnd, uMsg, wParam, lParam)
        End If
    End If

End Function

Private Sub AddMsg(ByVal uMsg As Long, _
                   ByVal nTable As Long)
'Add the message to the specified table of the window handle

Dim lCount As Long
Dim lBase  As Long
Dim I      As Long

    'Remember m_lResMem so that we can restore its value on exit
    lBase = m_lResMem
    'Map p_CAddress() to the specified table
    m_lResMem = p_CAddress(nTable)

    'If ALL_MESSAGES are being added to the table
    If (uMsg = ALL_MESSAGES) Then
        'Set the table entry count to ALL_MESSAGES
        lCount = ALL_MESSAGES
    Else
        'Get the current table entry count
        lCount = p_CAddress(0)
        If lCount >= MSG_ENTRIES Then
            'Check for message table overflow
            ErrorCond "AddMsg", "Message table overflow. Max message entries exceeded."
            GoTo Handler
        End If
        'Loop through the table entries
        For I = 1 To lCount
            'If the element is free...
            If p_CAddress(I) = 0 Then
                p_CAddress(I) = uMsg
                GoTo Handler
            'If the message is already in the table
            ElseIf p_CAddress(I) = uMsg Then
                GoTo Handler
            End If
        Next I
        'On drop through: i = lCount + 1, the new table entry count
        lCount = I
        p_CAddress(lCount) = uMsg
    End If
    'Store the message in the appended table entry
    p_CAddress(0) = lCount
  
Handler:
    m_lResMem = lBase

End Sub

Private Sub DelMsg(ByVal uMsg As Long, _
                   ByVal nTable As Long)
'Delete the message from the specified table of the window handle

Dim lCount As Long
Dim lBase  As Long
Dim I      As Long

    'Remember m_lResMem so that we can restore its value on exit
    lBase = m_lResMem
    'Map p_CAddress() to the specified table
    m_lResMem = p_CAddress(nTable)

    'If ALL_MESSAGES are being deleted from the table
    If (uMsg = ALL_MESSAGES) Then
        'Zero the table entry count
        p_CAddress(0) = 0
    Else
        'Get the table entry count
        lCount = p_CAddress(0)
        'Loop through the table entries
        For I = 1 To lCount
            If p_CAddress(I) = uMsg Then
                p_CAddress(I) = 0
                GoTo Handler
            End If
        Next I
        ErrorCond "DelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
    End If
  
Handler:
    m_lResMem = lBase

End Sub

Private Sub ErrorCond(ByVal sRoutine As String, _
                      ByVal sMsg As String)
'Error handler

  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine

End Sub

Private Function ReturnAddr(ByVal sDLL As String, _
                            ByVal sProc As String) As Long
'Return the address of the specified DLL/procedure

    'Get the specified procedure address
    If m_bIsNt Then
        ReturnAddr = GetProcAddress(GetModuleHandleW(StrPtr(sDLL)), sProc)
    Else
        ReturnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    End If
    'In the IDE, validate that the procedure address was located
    Debug.Assert ReturnAddr
  
End Function

Private Function MapHandle(ByVal lHwnd As Long) As Long
'Map p_CAddress() to the thunk address for the specified window handle

    'Ensure that subclassing has been started
    If (m_cWndHnd Is Nothing) Then
        ErrorCond "MapHandle", "Subclassing hasn't been started"
    Else
        On Error GoTo Handler
        'Catch unsubclassed window handles
        m_lResMem = m_cWndHnd("h" & lHwnd)
        'Get the thunk address
        MapHandle = m_lResMem
    End If
  
Exit Function

Handler:
    ErrorCond "MapHandle", "Window handle isn't subclassed"

End Function

Public Sub UnSubclass(ByVal lHwnd As Long)
'UnSubclass the specified window handle

    'Ensure that subclassing has been started
    If (m_cWndHnd Is Nothing) Then
        ErrorCond "UnSubclass", "Window handle isn't subclassed"
    Else
        'Ensure that the thunk hasn't already released its memory
        If (IsBadCodePtr(MapHandle(lHwnd)) = 0) Then
            'Set the shutdown indicator
            p_CAddress(IDX_SHUTDOWN) = -1
            'Delete all before messages
            DelMsg ALL_MESSAGES, IDX_BTABLE
            'Delete all after messages
            DelMsg ALL_MESSAGES, IDX_ATABLE
        End If
        'Remove the specified window handle from the collection
        m_cWndHnd.Remove "h" & lHwnd
    End If
    '/* a fix?
    SendMessageLongA lHwnd, 0&, 0&, 0&

End Sub

Public Sub Terminate()
'Terminate all subclassing

Dim I As Long

    If Not (m_cWndHnd Is Nothing) Then
        'Ensure that subclassing has been started
        With m_cWndHnd
            'Loop through the collection of window handles in reverse order
            For I = .Count To 1 Step -1
                'Get the thunk address
                m_lResMem = .Item(I)
                'Ensure that the thunk hasn't already released its memory
                If (IsBadCodePtr(m_lResMem) = 0) Then
                    UnSubclass p_CAddress(IDX_HWND)
                End If
            Next I
        End With
        'Destroy the hWnd/thunk-address collection
        Set m_cWndHnd = Nothing
    End If
  
End Sub

Private Sub Class_Terminate()
    Terminate
End Sub
