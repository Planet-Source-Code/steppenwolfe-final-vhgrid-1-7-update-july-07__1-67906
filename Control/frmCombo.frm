VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCombo 
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditBox 
      Caption         =   "Edit Box"
      Height          =   510
      Left            =   6480
      TabIndex        =   0
      Top             =   4905
      Width           =   1005
   End
   Begin MSComctlLib.ImageList imlHdr 
      Left            =   7335
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCombo.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCombo.frx":01DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCombo.frx":03B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCombo.frx":058E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_cComboSimple          As clsODControl
Private m_cComboDropDown        As clsODControl
Private m_cComboDropList        As clsODControl
Private m_cImageCombo           As clsODControl
Private m_cList                 As clsODControl
Private m_cListMultiSelect      As clsODControl
Private m_cListExtended         As clsODControl
Private m_cListIcons            As clsODControl
Private m_cButton               As clsODControl
Private m_cOptBtn               As clsODControl
Private m_cCheckBox             As clsODControl
Private m_cPictureBox           As clsODControl
Private m_cLabel                As clsODControl
Private m_cTextBox              As clsODControl

Private m_cEditText             As clsAdvancedEdit

Private Sub cmdEditBox_Click()
    m_cEditText.CreateEditBox Me.hWnd, 200, 200, 3, &HC4B0A2, &HCFDDF1, "Test string", imlHdr.hImageList, 1, eiaIcon
    ' m_lHGHwnd, .Left, .Top, m_eEditorThemeStyle, m_lAdvancedEditThemeColor, m_lAdvancedEditOffsetColor, sText, m_lImlRowHndl, lIcon, eiaIcon
End Sub

Private Sub Form_Load()

    Set m_cComboSimple = New clsODControl
    Set m_cComboDropDown = New clsODControl
    Set m_cComboDropList = New clsODControl
    Set m_cImageCombo = New clsODControl
    Set m_cList = New clsODControl
    Set m_cListMultiSelect = New clsODControl
    Set m_cListExtended = New clsODControl
    Set m_cListIcons = New clsODControl
    Set m_cButton = New clsODControl
    Set m_cOptBtn = New clsODControl
    Set m_cCheckBox = New clsODControl
    Set m_cPictureBox = New clsODControl
    Set m_cLabel = New clsODControl
    Set m_cTextBox = New clsODControl
    LoadControls
    Set m_cEditText = New clsAdvancedEdit
    
End Sub

Private Sub LoadControls()

    '/* combo
    With m_cComboSimple
        .Name = "Combo1"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 20, 20, 140, 100, ecsComboSimple
        .AddItem .Name
    End With
    
    '/* combo dropdown
    With m_cComboDropDown
        .Name = "Combo2"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 20, 130, 140, 80, ecsComboDropDown
        .AddItem .Name
    End With
        
    '/* combo droplist
    With m_cComboDropList
        .Name = "Combo3"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 20, 160, 140, 80, ecsComboDropList
        .AddItem .Name
    End With

    '/* image combo
    With m_cImageCombo
        .Name = "Combo4"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 20, 190, 140, 80, ecsImageCombo
        .ImlListBoxAddIcon imlHdr.ListImages.Item(1).Picture
        .ImlListBoxAddIcon imlHdr.ListImages.Item(2).Picture
        .ImlListBoxAddIcon imlHdr.ListImages.Item(3).Picture
        .ImlListBoxAddIcon imlHdr.ListImages.Item(4).Picture
        .AddItem "test1", 0
        .AddItem "test2", 1
        .AddItem "test3", 2
        .AddItem "test4", 3
    End With

    '/* listbox
    With m_cList
        .Name = "list1"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 180, 20, 100, 80, ecsListBox
        .AddItem .Name
    End With
    
    '/* listbox multi select
    With m_cListMultiSelect
        .Name = "list2"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 180, 110, 100, 80, ecsListBoxMultiSelect
        .AddItem .Name
    End With
    
    '/* listbox extended
    With m_cListExtended
        .Name = "list3"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 180, 200, 100, 80, ecsListBoxExtended
        .AddItem .Name
    End With

    '/* listbox icons
    With m_cListIcons
        .Name = "list4"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 180, 290, 100, 80, ecsImageListBox
        .ImlListBoxAddIcon imlHdr.ListImages.Item(1).Picture
        .ImlListBoxAddIcon imlHdr.ListImages.Item(2).Picture
        .ImlListBoxAddIcon imlHdr.ListImages.Item(3).Picture
        .ImlListBoxAddIcon imlHdr.ListImages.Item(4).Picture
        .AddItem "test1", 0
        .AddItem "test2", 1
        .AddItem "test3", 2
        .AddItem "test4", 3
    End With
    
    '/* button
    With m_cButton
        .Name = "button1"
        .HiliteColor = &HCCCCCC
        .Create Me.hWnd, 300, 20, 60, 25, ecsCommandButton
        .ImlCommandAddIcon imlHdr.ListImages.Item(1).Picture
        .Text = .Name
    End With
    
    '/* option button
    With m_cOptBtn
        .Name = "option1"
        .AutoBackColor = True
        .Create Me.hWnd, 300, 55, 80, 15, ecsOptionButton
        .Text = .Name
    End With
    
    '/* checkbox
    With m_cCheckBox
        .Name = "checkbox1"
        .AutoBackColor = True
        .Create Me.hWnd, 300, 80, 80, 15, ecsCheckBox
        .Text = .Name
    End With
    
    '/* picturebox
    With m_cPictureBox
        .Name = "checkbox1"
        .AutoSize = True
        .BorderStyle ecbsThin
        .BackColor = vbWhite
        .Create Me.hWnd, 300, 130, 100, 80, ecsPictureBox
        .PictureBoxLoadImage imlHdr.ListImages.Item(4).Picture, eiaIcon, 16, 16
    End With
    
    '/* label
    With m_cLabel
        .Name = "label1"
        .AutoSize = True
        .AutoBackColor = True
        .Create Me.hWnd, 300, 155, 100, 20, ecsLabel
        .Text = .Name & ": test 123"
    End With
    
    '/* textbox
    With m_cTextBox
        .Name = "textbox1"
        .BorderStyle ecbsThin
        .Create Me.hWnd, 300, 185, 100, 80, ecsTextBox
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set m_cComboSimple = Nothing
    Set m_cComboDropDown = Nothing
    Set m_cComboDropList = Nothing
    Set m_cImageCombo = New clsODControl
    Set m_cList = Nothing
    Set m_cListMultiSelect = Nothing
    Set m_cListExtended = Nothing
    Set m_cListIcons = Nothing
    Set m_cButton = Nothing
    Set m_cOptBtn = Nothing
    Set m_cCheckBox = Nothing
    Set m_cTextBox = New clsODControl
    Set m_cPictureBox = New clsODControl
    Set m_cLabel = New clsODControl
    Set m_cEditText = Nothing

End Sub
