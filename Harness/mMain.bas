Attribute VB_Name = "mMain"
Option Explicit

Private Const ICC_LISTVIEW_CLASSES              As Long = &H1

Public Enum ColorFlag
  CC_RGBINIT = &H1
  CC_FULLOPEN = &H2
  CC_PREVENTFULLOPEN = &H4
  CC_SHOWHELP = &H8
  CC_ENABLEHOOK = &H10
  CC_ENABLETEMPLATE = &H20
  CC_ENABLETEMPLATEHANDLE = &H40
  CC_SOLIDCOLOR = &H80
  CC_ANYCOLOR = &H100
End Enum

Private Type TCOLORDLG
  lStructSize     As Long
  hwndOwner       As Long
  hInstance       As Long
  rgbResult       As Long
  lpCustColors    As Long
  Flags           As Long
  lCustData       As Long
  lpfnHook        As Long
  lpTemplateName  As String
End Type

Private Type tagINITCOMMONCONTROLSEX
    dwSize As Long
    dwICC As Long
End Type


Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As TCOLORDLG) As Long

Public Sub Main()

    InitComctl32
    Load frmGridTest
    frmGridTest.Show
    'frmUnicode.Show
    'frmVirtual.Show
    'frmOwnerDrawn.Show
    'frmSubcell.Show
    
End Sub

Private Function InitComctl32() As Boolean

Dim icc As tagINITCOMMONCONTROLSEX

On Error GoTo Handler
  
    icc.dwSize = Len(icc)
    icc.dwICC = ICC_LISTVIEW_CLASSES
    InitComctl32 = InitCommonControlsEx(icc)

On Error GoTo 0
Exit Function

Handler:
  InitCommonControls

End Function

Public Function ShowColor(ByVal lOwnerHwnd As Long, _
                          ByVal lDfltClr As Long, _
                          ByRef lCustomClr() As Long, _
                          Optional ByVal ShowMode As Integer = 0) As Long

Dim tTCD As TCOLORDLG

    With tTCD
        .lStructSize = Len(tTCD)
        .hwndOwner = lOwnerHwnd
        .hInstance = App.hInstance
        .Flags = CC_ANYCOLOR
        
        Select Case ShowMode
        Case 1
            .Flags = .Flags Or CC_FULLOPEN
        Case 2
            .Flags = .Flags Or CC_PREVENTFULLOPEN
        End Select

        .Flags = .Flags Or CC_RGBINIT
        .rgbResult = lDfltClr
        .lpCustColors = VarPtr(lCustomClr(0))
        
        If ChooseColor(tTCD) = 1 Then
            ShowColor = .rgbResult
        Else
            ShowColor = -1
        End If
    End With
  
End Function

