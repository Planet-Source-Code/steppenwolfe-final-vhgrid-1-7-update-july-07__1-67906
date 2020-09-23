VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5325
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Column Number"
      Height          =   645
      Left            =   3465
      TabIndex        =   9
      Top             =   630
      Width           =   1680
      Begin VB.TextBox txtColumn 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   945
         TabIndex        =   11
         Text            =   "0"
         Top             =   270
         Width           =   285
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         ScaleHeight     =   330
         ScaleWidth      =   1545
         TabIndex        =   10
         Top             =   270
         Width           =   1545
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Column #"
            Height          =   210
            Left            =   180
            TabIndex        =   12
            Top             =   45
            Width           =   660
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Direction"
      Height          =   645
      Left            =   1665
      TabIndex        =   5
      Top             =   630
      Width           =   1680
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         ScaleHeight     =   330
         ScaleWidth      =   1545
         TabIndex        =   6
         Top             =   270
         Width           =   1545
         Begin VB.OptionButton optDirection 
            Caption         =   "Down"
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   8
            Top             =   90
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton optDirection 
            Caption         =   "Up"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   7
            Top             =   90
            Width           =   600
         End
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   4185
      TabIndex        =   4
      Top             =   90
      Width           =   1050
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Exact Match"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1035
      Width           =   1275
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Match Case"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   2
      Top             =   765
      Width           =   1275
   End
   Begin VB.TextBox txtFind 
      Height          =   330
      Left            =   1035
      TabIndex        =   1
      Top             =   135
      Width           =   2850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find What:"
      Height          =   210
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   765
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bNext     As Boolean


Private Sub cmdFind_Click()

Dim lCol As Long

    If Not IsNumeric(txtColumn.Text) Then
        MsgBox "Please enter a valid column number in the Find dialog Column textbox!", vbExclamation, "Invalid Column!"
        Exit Sub
    ElseIf CLng(txtColumn.Text) > (frmGridTest.ucVHGrid1.ColumnCount - 1) Then
        MsgBox "Please enter a valid column number in the Find dialog Column textbox!", vbExclamation, "Invalid Column!"
        Exit Sub
    ElseIf Len(txtFind.Text) = 0 Then
        Exit Sub
    Else
        lCol = CLng(txtColumn.Text)
    End If

    frmGridTest.ucVHGrid1.Find txtFind.Text, lCol, CBool(chkMatch(0).Value), _
        CBool(chkMatch(1).Value), optDirection(0).Value, m_bNext, True
    m_bNext = True

End Sub

Private Sub optDirection_Click(Index As Integer)
    m_bNext = False
End Sub
