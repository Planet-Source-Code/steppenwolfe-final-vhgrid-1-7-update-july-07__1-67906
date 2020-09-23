VERSION 5.00
Object = "*\A..\prjvhGrid.vbp"
Begin VB.Form frmGridTest 
   BackColor       =   &H00F6F6F6&
   Caption         =   "vhGrid - Features Demonstration"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15840
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   15840
   StartUpPosition =   2  'CenterScreen
   Begin vhGrid.ucVHGrid ucVHGrid1 
      Height          =   7035
      Left            =   225
      TabIndex        =   116
      Top             =   270
      Width           =   9780
      _extentx        =   17251
      _extenty        =   12409
      font            =   "frmGridTest.frx":0000
      headerfont      =   "frmGridTest.frx":002C
      headericoncount =   5
      headericons     =   "frmGridTest.frx":0058
      celliconsizex   =   32
      celliconsizey   =   32
      celliconcolourdepth=   0
      celliconcount   =   20
      cellicons       =   "frmGridTest.frx":14079
      treeiconcolourdepth=   8
      treeiconcount   =   3
      treeicons       =   "frmGridTest.frx":6409A
      alphabartransparency=   70
      forecolor       =   0
      gridlines       =   0
      headerdragdrop  =   0   'False
      headerforecolor =   0
      headerforecolorfocused=   0
      headerforecolorpressed=   0
      headerheight    =   24
      headerheightsizable=   -1  'True
      oledropmode     =   1
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   10305
      ScaleHeight     =   7050
      ScaleWidth      =   5340
      TabIndex        =   41
      Top             =   225
      Width           =   5370
      Begin VB.Frame frmProp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Focus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Index           =   0
         Left            =   2745
         TabIndex        =   98
         Top             =   135
         Width           =   2400
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   90
            ScaleHeight     =   915
            ScaleWidth      =   2220
            TabIndex        =   99
            Top             =   180
            Width           =   2220
            Begin VB.OptionButton optFocus 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Normal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   45
               TabIndex        =   105
               Top             =   90
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton optFocus 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Alpha Blend"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   45
               TabIndex        =   104
               Top             =   360
               Width           =   1185
            End
            Begin VB.OptionButton optFocus 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Alpha Bar"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   45
               TabIndex        =   103
               Top             =   630
               Width           =   1050
            End
            Begin VB.PictureBox picFocus 
               Appearance      =   0  'Flat
               BackColor       =   &H00C56A31&
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   1755
               ScaleHeight     =   210
               ScaleWidth      =   390
               TabIndex        =   102
               Top             =   585
               Width           =   420
            End
            Begin VB.CheckBox chkTextOnly 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Text Only"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1170
               TabIndex        =   101
               Top             =   90
               Width           =   1005
            End
            Begin VB.PictureBox picSelect 
               Appearance      =   0  'Flat
               BackColor       =   &H00D8E9EC&
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   1260
               ScaleHeight     =   210
               ScaleWidth      =   390
               TabIndex        =   100
               Top             =   585
               Width           =   420
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Focus"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   1755
               TabIndex        =   107
               Top             =   405
               Width           =   405
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Select"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   1260
               TabIndex        =   106
               Top             =   405
               Width           =   390
            End
         End
      End
      Begin VB.Frame frmProp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grid Lines"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Index           =   1
         Left            =   2745
         TabIndex        =   90
         Top             =   1305
         Width           =   2400
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1140
            Left            =   45
            ScaleHeight     =   1140
            ScaleWidth      =   2310
            TabIndex        =   91
            Top             =   180
            Width           =   2310
            Begin VB.OptionButton optLines 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Vertical"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   96
               Top             =   90
               Width           =   915
            End
            Begin VB.OptionButton optLines 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Horizontal"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   95
               Top             =   360
               Width           =   1050
            End
            Begin VB.OptionButton optLines 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Both"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   94
               Top             =   630
               Value           =   -1  'True
               Width           =   735
            End
            Begin VB.PictureBox picGridline 
               Appearance      =   0  'Flat
               BackColor       =   &H0099A8AC&
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   1800
               ScaleHeight     =   210
               ScaleWidth      =   390
               TabIndex        =   93
               Top             =   855
               Width           =   420
            End
            Begin VB.OptionButton optLines 
               BackColor       =   &H00FFFFFF&
               Caption         =   "None"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   92
               Top             =   900
               Width           =   1050
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   1800
               TabIndex        =   97
               Top             =   675
               Width           =   360
            End
         End
      End
      Begin VB.Frame frmProp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Border Style"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   2
         Left            =   2745
         TabIndex        =   85
         Top             =   2700
         Width           =   2400
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   870
            Left            =   45
            ScaleHeight     =   870
            ScaleWidth      =   2265
            TabIndex        =   86
            Top             =   180
            Width           =   2265
            Begin VB.OptionButton optBorder 
               BackColor       =   &H00FFFFFF&
               Caption         =   "3D"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   89
               Top             =   90
               Width           =   600
            End
            Begin VB.OptionButton optBorder 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Thin"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   88
               Top             =   360
               Value           =   -1  'True
               Width           =   690
            End
            Begin VB.OptionButton optBorder 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Line"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   135
               TabIndex        =   87
               Top             =   630
               Width           =   735
            End
         End
      End
      Begin VB.Frame frmProp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Drag Effect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   3
         Left            =   2745
         TabIndex        =   80
         Top             =   3825
         Width           =   2400
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   870
            Left            =   45
            ScaleHeight     =   870
            ScaleWidth      =   2265
            TabIndex        =   81
            Top             =   180
            Width           =   2265
            Begin VB.OptionButton optDrag 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Arrow"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   84
               Top             =   90
               Value           =   -1  'True
               Width           =   825
            End
            Begin VB.OptionButton optDrag 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Thin Line"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   83
               Top             =   360
               Width           =   1005
            End
            Begin VB.OptionButton optDrag 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Thick Line"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   82
               Top             =   630
               Width           =   1050
            End
         End
      End
      Begin VB.Frame frmProp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sorting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   4
         Left            =   2745
         TabIndex        =   75
         Top             =   4950
         Width           =   2400
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   870
            Left            =   45
            ScaleHeight     =   870
            ScaleWidth      =   2265
            TabIndex        =   76
            Top             =   180
            Width           =   2265
            Begin VB.OptionButton optSorting 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Case Sensitive"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   79
               Top             =   90
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton optSorting 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Case Insensitive"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   78
               Top             =   360
               Width           =   1590
            End
            Begin VB.OptionButton optSorting 
               BackColor       =   &H00FFFFFF&
               Caption         =   "None"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   77
               Top             =   630
               Width           =   780
            End
         End
      End
      Begin VB.Frame frmProp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cell Decoration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Index           =   5
         Left            =   2745
         TabIndex        =   71
         Top             =   6075
         Width           =   2400
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   45
            ScaleHeight     =   555
            ScaleWidth      =   2265
            TabIndex        =   72
            Top             =   225
            Width           =   2265
            Begin VB.OptionButton optDecoration 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Structured"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   74
               Top             =   315
               Value           =   -1  'True
               Width           =   1230
            End
            Begin VB.OptionButton optDecoration 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Per Cell"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   73
               Top             =   45
               Width           =   915
            End
         End
      End
      Begin VB.Frame frmStyles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Base Styles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Index           =   0
         Left            =   180
         TabIndex        =   63
         Top             =   135
         Width           =   2400
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1725
            Left            =   90
            ScaleHeight     =   1725
            ScaleWidth      =   2220
            TabIndex        =   64
            Top             =   180
            Width           =   2220
            Begin VB.CheckBox chkBase 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Check Boxes"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   45
               TabIndex        =   70
               Top             =   90
               Value           =   1  'Checked
               Width           =   1275
            End
            Begin VB.CheckBox chkBase 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Double Buffer"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   45
               TabIndex        =   69
               Top             =   360
               Value           =   1  'Checked
               Width           =   1410
            End
            Begin VB.CheckBox chkBase 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Enabled"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   45
               TabIndex        =   68
               Top             =   630
               Value           =   1  'Checked
               Width           =   960
            End
            Begin VB.CheckBox chkBase 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Full Row Select"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   45
               TabIndex        =   67
               Top             =   900
               Width           =   1545
            End
            Begin VB.CheckBox chkBase 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Unicode (NT/2K/XP)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   45
               TabIndex        =   66
               Top             =   1170
               Width           =   1860
            End
            Begin VB.CheckBox chkBase 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Visible"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   45
               TabIndex        =   65
               Top             =   1440
               Value           =   1  'Checked
               Width           =   870
            End
         End
      End
      Begin VB.Frame frmStyles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Advanced Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   1
         Left            =   180
         TabIndex        =   56
         Top             =   2115
         Width           =   2400
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1410
            Left            =   45
            ScaleHeight     =   1410
            ScaleWidth      =   2265
            TabIndex        =   57
            Top             =   225
            Width           =   2265
            Begin VB.CheckBox chkOptions 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cell Edit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   90
               TabIndex        =   62
               Top             =   45
               Value           =   1  'Checked
               Width           =   1005
            End
            Begin VB.CheckBox chkOptions 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Advanced Edit (Ctrl+A)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   90
               TabIndex        =   61
               Top             =   315
               Value           =   1  'Checked
               Width           =   2040
            End
            Begin VB.CheckBox chkOptions 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Custom Cursors"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   90
               TabIndex        =   60
               Top             =   585
               Value           =   1  'Checked
               Width           =   1590
            End
            Begin VB.CheckBox chkOptions 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Lock First Column"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   90
               TabIndex        =   59
               Top             =   855
               Value           =   1  'Checked
               Width           =   1590
            End
            Begin VB.CheckBox chkOptions 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Row Drag and Drop"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   90
               TabIndex        =   58
               Top             =   1125
               Value           =   1  'Checked
               Width           =   1770
            End
         End
      End
      Begin VB.Frame frmStyles 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Extended Styles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Index           =   2
         Left            =   180
         TabIndex        =   42
         Top             =   3825
         Width           =   2400
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2760
            Left            =   45
            ScaleHeight     =   2760
            ScaleWidth      =   2265
            TabIndex        =   43
            Top             =   225
            Width           =   2265
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cell Hot Track"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   45
               TabIndex        =   54
               Top             =   45
               Width           =   1365
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Cell Tool Tip"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   45
               TabIndex        =   53
               Top             =   315
               Value           =   1  'Checked
               Width           =   1365
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Header Height Sizable"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   45
               TabIndex        =   52
               Top             =   585
               Value           =   1  'Checked
               Width           =   2040
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column Width Sizable"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   45
               TabIndex        =   51
               Top             =   855
               Value           =   1  'Checked
               Width           =   1905
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column Lock"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   45
               TabIndex        =   50
               Top             =   1125
               Value           =   1  'Checked
               Width           =   1320
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column Filters"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   45
               TabIndex        =   49
               Top             =   1395
               Value           =   1  'Checked
               Width           =   1320
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column Hints"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   6
               Left            =   45
               TabIndex        =   48
               Top             =   1665
               Value           =   1  'Checked
               Width           =   1320
            End
            Begin VB.TextBox txtColumn 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1755
               TabIndex        =   47
               Text            =   "1"
               Top             =   1125
               Width           =   195
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column Focus"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   7
               Left            =   45
               TabIndex        =   46
               Top             =   1935
               Width           =   1320
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column Drag Line"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   8
               Left            =   45
               TabIndex        =   45
               Top             =   2205
               Value           =   1  'Checked
               Width           =   1635
            End
            Begin VB.CheckBox chkExtended 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column Vertical Text"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   9
               Left            =   45
               TabIndex        =   44
               Top             =   2475
               Value           =   1  'Checked
               Width           =   1950
            End
            Begin VB.Label lblStyles 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Col #"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   3
               Left            =   1395
               TabIndex        =   55
               Top             =   1170
               Width           =   345
            End
         End
      End
   End
   Begin VB.Frame frmProp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Method Demos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Index           =   6
      Left            =   6030
      TabIndex        =   22
      Top             =   7515
      Width           =   2580
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   135
         ScaleHeight     =   1590
         ScaleWidth      =   2355
         TabIndex        =   23
         Top             =   180
         Width           =   2355
         Begin VB.CommandButton cmdDemo 
            Caption         =   "OwnerDrawn Cells"
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   28
            Top             =   675
            Width           =   2085
         End
         Begin VB.CommandButton cmdDemo 
            Caption         =   "Unicode (NT/2K/XP)"
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   27
            Top             =   405
            Width           =   2085
         End
         Begin VB.CommandButton cmdDemo 
            Caption         =   "Virtual Mode"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   26
            Top             =   135
            Width           =   2085
         End
         Begin VB.CommandButton cmdDemo 
            Caption         =   "SubCell Controls"
            Height          =   240
            Index           =   3
            Left            =   135
            TabIndex        =   25
            Top             =   945
            Width           =   2085
         End
         Begin VB.CommandButton cmdDemo 
            Caption         =   "Hyper Mode"
            Height          =   240
            Index           =   4
            Left            =   135
            TabIndex        =   24
            Top             =   1215
            Width           =   2085
         End
      End
   End
   Begin VB.Frame fmExtended 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Extended Styles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   225
      TabIndex        =   2
      Top             =   7515
      Width           =   5640
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1545
         Left            =   90
         ScaleHeight     =   1545
         ScaleWidth      =   5505
         TabIndex        =   3
         Top             =   270
         Width           =   5505
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "QuickSilver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   4140
            TabIndex        =   117
            Top             =   225
            Width           =   1185
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Vista"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   3375
            TabIndex        =   115
            Top             =   225
            Width           =   735
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "XP-Green"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   4410
            TabIndex        =   114
            Top             =   585
            Width           =   1005
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "XP-Blue"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   3510
            TabIndex        =   113
            Top             =   585
            Width           =   915
         End
         Begin VB.PictureBox picThemeClr 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1DDCF&
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1125
            ScaleHeight     =   210
            ScaleWidth      =   390
            TabIndex        =   19
            Top             =   585
            Width           =   420
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Azure"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   225
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Classic"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   945
            TabIndex        =   15
            Top             =   225
            Width           =   825
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Gloss"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   2
            Left            =   1845
            TabIndex        =   14
            Top             =   225
            Width           =   735
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Metal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   3
            Left            =   2655
            TabIndex        =   13
            Top             =   225
            Width           =   735
         End
         Begin VB.OptionButton optStyles 
            BackColor       =   &H00FFFFFF&
            Caption         =   "XP-Silver"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   4
            Left            =   2520
            TabIndex        =   12
            Top             =   585
            Width           =   960
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Left            =   0
            ScaleHeight     =   690
            ScaleWidth      =   6045
            TabIndex        =   5
            Top             =   810
            Width           =   6045
            Begin VB.OptionButton optRowDec 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Checker"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   2295
               TabIndex        =   21
               Top             =   405
               Width           =   960
            End
            Begin VB.OptionButton optRowDec 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Column"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   1440
               TabIndex        =   18
               Top             =   405
               Width           =   825
            End
            Begin VB.OptionButton optRowDec 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Row"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   90
               TabIndex        =   9
               Top             =   405
               Value           =   -1  'True
               Width           =   645
            End
            Begin VB.OptionButton optRowDec 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Split"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   765
               TabIndex        =   8
               Top             =   405
               Width           =   645
            End
            Begin VB.TextBox txtDepth 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   3780
               TabIndex        =   7
               Text            =   "2"
               Top             =   360
               Width           =   240
            End
            Begin VB.CommandButton cmdReset 
               Caption         =   "Reset"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4275
               TabIndex        =   6
               Top             =   315
               Width           =   1095
            End
            Begin VB.Label lblStyles 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Row Decoration"
               Height          =   210
               Index           =   1
               Left            =   90
               TabIndex        =   11
               Top             =   135
               Width           =   1290
            End
            Begin VB.Label lblStyles 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Depth:"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   2
               Left            =   3330
               TabIndex        =   10
               Top             =   405
               Width           =   405
            End
         End
         Begin VB.CheckBox chkSkin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Colorize"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   4
            Top             =   585
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1620
            TabIndex        =   20
            Top             =   585
            Width           =   795
         End
         Begin VB.Label lblStyles 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skin Styles"
            Height          =   210
            Index           =   0
            Left            =   90
            TabIndex        =   17
            Top             =   -45
            Width           =   915
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Base Functions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   8775
      TabIndex        =   1
      Top             =   7515
      Width           =   6900
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   135
         ScaleHeight     =   1590
         ScaleWidth      =   6675
         TabIndex        =   29
         Top             =   180
         Width           =   6675
         Begin VB.OptionButton OPTpOS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Left"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   3690
            TabIndex        =   112
            Top             =   855
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton OPTpOS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Right"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   3690
            TabIndex        =   111
            Top             =   1125
            Width           =   780
         End
         Begin VB.OptionButton OPTpOS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bottom"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3690
            TabIndex        =   110
            Top             =   585
            Width           =   915
         End
         Begin VB.OptionButton OPTpOS 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Top"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   3690
            TabIndex        =   109
            Top             =   315
            Width           =   645
         End
         Begin VB.CommandButton cmdPopulate 
            Caption         =   "Populate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   4905
            TabIndex        =   39
            Top             =   1080
            Width           =   1635
         End
         Begin VB.TextBox txtCount 
            Height          =   285
            Left            =   4905
            TabIndex        =   38
            Text            =   "1000"
            Top             =   765
            Width           =   555
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Add Row"
            Height          =   285
            Index           =   0
            Left            =   135
            TabIndex        =   37
            Top             =   135
            Width           =   1545
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Rem Row"
            Height          =   285
            Index           =   1
            Left            =   135
            TabIndex        =   36
            Top             =   495
            Width           =   1545
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Add Column"
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   35
            Top             =   855
            Width           =   1545
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Rem Column"
            Height          =   285
            Index           =   3
            Left            =   135
            TabIndex        =   34
            Top             =   1215
            Width           =   1545
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Check All"
            Height          =   285
            Index           =   4
            Left            =   1890
            TabIndex        =   33
            Top             =   135
            Width           =   1545
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "UnCheck All"
            Height          =   285
            Index           =   5
            Left            =   1890
            TabIndex        =   32
            Top             =   495
            Width           =   1545
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Find"
            Height          =   285
            Index           =   6
            Left            =   1890
            TabIndex        =   31
            Top             =   855
            Width           =   1545
         End
         Begin VB.CommandButton cmdMethods 
            Caption         =   "Clear List"
            Height          =   285
            Index           =   7
            Left            =   1890
            TabIndex        =   30
            Top             =   1215
            Width           =   1545
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TreeView Orientation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3690
            TabIndex        =   108
            Top             =   45
            Width           =   1770
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "# Rows"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5490
            TabIndex        =   40
            Top             =   810
            Width           =   570
         End
      End
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
      ScaleWidth      =   15780
      TabIndex        =   0
      Top             =   9525
      Width           =   15840
   End
End
Attribute VB_Name = "frmGridTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_lRowCount                 As Long
Private m_lSkinStyle                As Long
Private m_lDecoration               As Long
Private m_lDecColor                 As Long
Private m_lDecOffset                As Long
Private m_lThemeClr                 As Long
Private m_lCustomClr(15)            As Long
Private m_cTiming                   As clsTiming


Private Sub ucVHGrid1_eTVNodeClick(ByVal hNode As Long)
    Debug.Print "Treeview Click: " & ucVHGrid1.TreeViewNodeText(hNode)
End Sub

Private Sub Form_Load()

Dim lX      As Long

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
       ' .ForeColorAuto = True
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
        .ForeColorFocused = m_lDecColor
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
        .BackColor = m_lThemeClr
        .DisabledBackColor = &HCCCCCC
        .DisabledForeColor = &H999999
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
        .ImlUseAlphaIcons = True
        '/* apply skin
        .ThemeManager evsXpSilver, False, , , &H333333, m_lCustomClr(11), m_lCustomClr(10), _
            m_lCustomClr(11), m_lCustomClr(10), &H808080, 210, True, True, m_lCustomClr(9), _
            &H0, True, False, True, False
        '/* apply cell decoration
        .CellDecoration erdCellChecker, m_lCustomClr(9), m_lCustomClr(10), True, 1
        .TreeViewAlignment = etvLeftAlign
        .TreeViewInit 150, 150, , , , , 8, True, True, True, True, True, , , , True
        .TreeViewOLEDragMode = tvdManual
        .TreeViewUseUnicode = True
    End With

    '/* load data
    CreateAppData
    AddNodes
    
End Sub

Private Sub AddNodes()

Dim lKey        As Long
Dim lBranch     As Long
Dim lFolder     As Long
Dim lSubFolder  As Long

    With ucVHGrid1
        .TreeViewClear
        .TreeViewInitIml
        For lBranch = 1 To 2
            lKey = lKey + 1
            lFolder = .TreeViewAddNode(, , lKey, "Folder #" & lBranch, 0, 1)
            .TreeViewDraw False
            For lSubFolder = 1 To 20
                lKey = lKey + 1
                .TreeViewAddNode lFolder, , lKey, "SubFolder" & lBranch & "." & lSubFolder, 0, 1
            Next lSubFolder
        Next lBranch
        .TreeViewDraw True
        .TreeViewExpand 0, True
    End With

End Sub

Private Sub cmdPopulate_Click()

Dim lCt         As Long
Dim lRnd        As Long
Dim lCust()     As Long
Dim sTitle      As String
Dim sHeader     As String
Dim sDetails    As String
Dim oFnt        As StdFont

    '/* test user params
    If Not TestInput Then
        Exit Sub
    End If

    sTitle = "Summary of Drive Search; ACX1, Index 3" & vbNewLine & "Access Date; Sept 17, 2006" & vbNewLine & "Storage Facility; GreenBorough Mass."
    sHeader = "Server: SXX3" & vbNewLine & "Media: Tape" & vbNewLine & "Status: Archived" & vbNewLine & "Last Administrative Access: Dec 3, 2006" & vbNewLine & "Clearence: Public"
    sDetails = vbNewLine & "Master: Yes" & vbNewLine & "Public: Yes" & vbNewLine & "Last Access: " & CStr(Date)
    '/* per-cell colors
    ReDim lCust(8)
    lCust(0) = &HFFFFEB
    lCust(1) = &HF0F0DC
    lCust(2) = &HE1E1CD
    lCust(3) = &HD2D2BE
    lCust(4) = &HC3C3AF
    lCust(5) = &HB4B4A0
    lCust(6) = &HA5A591
    lCust(7) = &H969682
    lCust(8) = &H414B25
    
    '/* cell font
    Set oFnt = New StdFont
    With oFnt
        .Name = "Times New Roman"
        .Bold = True
        .Size = 8
    End With
    
    '/* reset test timer
    m_cTiming.Reset

    With ucVHGrid1
        '/* manually initialize list: rowcount, columncount
        .ClearList
        .FastLoad = True '<- put this before gridinit if re-adding multiple rows!!! loads 1000x faster!
        .GridInit m_lRowCount, 8
        m_cTiming.Reset
        '/* add special cells
        '/* row 0 -spanned 2 rows
        .AddCell 0, 0, sTitle, DT_VCENTER Or DT_CENTER Or DT_WORDBREAK, , lCust(1), , oFnt, 6, 3
        '/* hide checkbox in this row
        .RowHideCheckBox 0, True
        '/* row does not accept focus
        .RowNoFocus 0, True
        '/* row is not editable
        .RowNoEdit 0, True
        '/* cell is spanned from 1st to last column
        .CellSpanHorizontal 0, 0, 7
        '/* row 1 -spanned 3 rows
        .AddCell 1, 0, sHeader, DT_VCENTER Or DT_LEFT Or DT_WORDBREAK, , lCust(2), , oFnt, 6, 2
        .RowHideCheckBox 1, True
        .RowNoFocus 1, True
        .RowNoEdit 1, True
        '/* cell hz spanned from first to last column
        .CellSpanHorizontal 1, 0, 7
        
        '/* row 2 -spanned 3 rows
        .AddCell 2, 0, , , , lCust(0), , , , 2
        .AddCell 2, 1, m_sAppName(0), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 0, lCust(1), , , 6
        .AddCell 2, 2, m_sAppDesc(0), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 1, lCust(2), , oFnt
        .AddCell 2, 3, m_sAppData(0), DT_WORDBREAK Or DT_TOP, 2, lCust(3)
        .AddCell 2, 4, m_sAppStats(0) & sDetails, DT_LEFT Or DT_WORDBREAK, 3, lCust(4)
        .AddCell 2, 5, "Yes", DT_LEFT Or DT_END_ELLIPSIS, , lCust(5)
        .AddCell 2, 6, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(6)
        .AddCell 2, 7, "Yes", DT_LEFT Or DT_END_ELLIPSIS, , lCust(7)
        .AddCellHeader 2, 1, "Archive Type", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 2, 2, "Document Title", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 2, 3, "MASTER DOCUMENT", vbRed, &HAAAAFF, oFnt, DT_CENTER
        .AddCellHeader 2, 4, "Document Details", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        
        '/* row 3 -spanned 2 rows
        .AddCell 3, 0, , , , lCust(0), , , , 2
        .AddCell 3, 1, m_sAppName(1), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 4, lCust(1), , , 6
        .AddCell 3, 2, m_sAppDesc(1), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 5, lCust(2), , oFnt, 5
        .AddCell 3, 3, m_sAppData(1), DT_WORDBREAK Or DT_TOP, 6, lCust(3)
        .AddCell 3, 4, m_sAppStats(1) & sDetails, DT_LEFT Or DT_WORDBREAK, 7, lCust(4)
        .AddCell 3, 5, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(5)
        .AddCell 3, 6, "Yes", DT_LEFT Or DT_END_ELLIPSIS, , lCust(6)
        .AddCell 3, 7, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(7)
        .AddCellHeader 3, 1, "Archive Type", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 3, 2, "Document Title", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 3, 3, "MASTER DOCUMENT", vbRed, &HAAAAFF, oFnt, DT_CENTER
        .AddCellHeader 3, 4, "Document Details", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        
        '/* row 4 -spanned 2 rows
        .AddCell 4, 0, , , , lCust(0), , , , 2
        .AddCell 4, 1, m_sAppName(2), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 8, lCust(1), , , 6
        .AddCell 4, 2, m_sAppDesc(2), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 9, lCust(2), , oFnt, 5
        .AddCell 4, 3, m_sAppData(2), DT_WORDBREAK Or DT_TOP, 10, lCust(3)
        .AddCell 4, 4, m_sAppStats(2) & sDetails, DT_LEFT Or DT_WORDBREAK, 11, lCust(4)
        .AddCell 4, 5, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(5)
        .AddCell 4, 6, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(6)
        .AddCell 4, 7, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(7)
        .AddCellHeader 4, 1, "Archive Type", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 4, 2, "Document Title", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 4, 3, "MASTER DOCUMENT", vbRed, &HAAAAFF, oFnt, DT_CENTER
        .AddCellHeader 4, 4, "Document Details", lCust(8), &HCEF9EA, oFnt, DT_LEFT

        '/* row 5 -spanned 2 rows
        .AddCell 5, 0, , , , lCust(0), , , , 2
        .AddCell 5, 1, m_sAppName(3), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 12, lCust(1), , , 6
        .AddCell 5, 2, m_sAppDesc(3), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 13, lCust(2), , oFnt, 5
        .AddCell 5, 3, m_sAppData(3), DT_WORDBREAK Or DT_TOP, 14, lCust(3)
        .AddCell 5, 4, m_sAppStats(3) & sDetails, DT_LEFT Or DT_WORDBREAK, 15, lCust(4)
        .AddCell 5, 5, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(5)
        .AddCell 5, 6, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(6)
        .AddCell 5, 7, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(7)
        .AddCellHeader 5, 1, "Archive Type", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 5, 2, "Document Title", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 5, 3, "MASTER DOCUMENT", vbRed, &HAAAAFF, oFnt, DT_CENTER
        .AddCellHeader 5, 4, "Document Details", lCust(8), &HCEF9EA, oFnt, DT_LEFT

        '/* row 6 -spanned 2 rows
        .AddCell 6, 0, , , , lCust(0), , , , 2
        .AddCell 6, 1, m_sAppName(4), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 16, lCust(1), , , 6
        .AddCell 6, 2, m_sAppDesc(4), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, 17, lCust(2), , oFnt, 5
        .AddCell 6, 3, m_sAppData(4), DT_WORDBREAK Or DT_TOP, 18, lCust(3)
        .AddCell 6, 4, m_sAppStats(4) & sDetails, DT_LEFT Or DT_WORDBREAK, 19, lCust(4)
        .AddCell 6, 5, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(5)
        .AddCell 6, 6, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(6)
        .AddCell 6, 7, "No", DT_LEFT Or DT_END_ELLIPSIS, , lCust(7)
        .AddCellHeader 6, 1, "Archive Type", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 6, 2, "Document Title", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        .AddCellHeader 6, 3, "MASTER DOCUMENT", vbRed, &HAAAAFF, oFnt, DT_CENTER
        .AddCellHeader 6, 4, "Document Details", lCust(8), &HCEF9EA, oFnt, DT_LEFT
        
        '/* add the rest of the rows
        For lCt = 7 To (m_lRowCount - 1)
            lRnd = RandomNum(0, 4)
            .AddCell lCt, 0, , , , lCust(0)
            .AddCell lCt, 1, m_sAppName(lRnd), DT_LEFT Or DT_END_ELLIPSIS, lRnd * 5, lCust(1), , , 6
            .AddCell lCt, 2, m_sAppDesc(lRnd), DT_LEFT Or DT_END_ELLIPSIS Or DT_VCENTER, , lCust(2), , , 5
            .AddCell lCt, 3, m_sAppData(lRnd), DT_LEFT Or DT_WORDBREAK, , lCust(3), , , 5
            .AddCell lCt, 4, m_sAppStats(lRnd), DT_LEFT Or DT_END_ELLIPSIS, , lCust(4)
            .AddCell lCt, 5, "Yes" & lCt, DT_LEFT Or DT_END_ELLIPSIS, , lCust(5)
            .AddCell lCt, 6, "Yes", DT_LEFT Or DT_END_ELLIPSIS, , lCust(6)
            .AddCell lCt, 7, "Yes", DT_LEFT Or DT_END_ELLIPSIS, , lCust(7)
        Next lCt
        '/* refresh the grid
        .Draw = True
        .GridRefresh True
    End With

    '/* get time elapsed
    picBar.Cls
    picBar.Print " " & (m_lRowCount) & " rows added to vhGrid in: " & _
        Format$(m_cTiming.Elapsed / 1000, "0.0000") & "s"
    
End Sub

Private Sub chkBase_Click(Index As Integer)
'/* base options

Dim bVal As Boolean

    bVal = (chkBase(Index).Value = 1)
    
    With ucVHGrid1
        Select Case Index
        '/* checkboxes
        Case 0
            .Checkboxes = bVal
        '/* double buffer
        Case 1
            .DoubleBuffer = bVal
        '/* enabled
        Case 2
            .Enabled = bVal
        '/* full row select
        Case 3
            .FullRowSelect = bVal
            .GridRefresh False
        '/* unicode
        Case 4
            .UseUnicode = bVal
        '/* visible
        Case 5
            .Visible = bVal
        End Select
    End With

End Sub

Private Sub chkExtended_Click(Index As Integer)
'/* extended styles

Dim bVal As Boolean

    bVal = (chkExtended(Index).Value = 1)
    
    With ucVHGrid1
        Select Case Index
        '/* cell hot track
        Case 0
            .CellHotTrack = bVal
            .GridRefresh False
        '/* cell tool tips
        Case 1
            .CellTips = bVal
        '/* header vert sizing
        Case 2
            .HeaderHeightSizable = bVal
        '/* column horz sizing
        Case 3
            .HeaderFixedWidth = (Not bVal)
        '/* column lock
        Case 4
            If IsNumeric(txtColumn.Text) Then
                If CLng(txtColumn.Text) < (.ColumnCount - 1) Then
                    If CLng(txtColumn.Text) > 0 Then
                        .ColumnLock(CLng(txtColumn.Text)) = bVal
                    Else
                        MsgBox "Please choose a column number between 1 and " & (.ColumnCount - 1), vbExclamation, "Invalid Input!"
                        Exit Sub
                    End If
                Else
                    MsgBox "Please choose a column number between 1 and " & (.ColumnCount - 1), vbExclamation, "Invalid Input!"
                    Exit Sub
                End If
            Else
                MsgBox "Please choose a column number between 1 and " & (.ColumnCount - 1), vbExclamation, "Invalid Input!"
                Exit Sub
            End If
        '/* column filter
        Case 5
            .ColumnFilters = bVal
        '/* column hints
        Case 6
            .ColumnToolTips = bVal
        '/* column focus
        Case 7
            .ColumnFocus = bVal
            .GridRefresh False
        '/* column drag line
        Case 8
            .ColumnDragLine = bVal
        '/* column verical text
        Case 9
            .ColumnVerticalText = bVal
            .Resize
        End Select
    End With

End Sub

Private Sub chkOptions_Click(Index As Integer)
'/* advanced options

Dim bVal As Boolean

    bVal = (chkOptions(Index).Value = 1)
    
    With ucVHGrid1
        Select Case Index
        '/* cell edit
        Case 0
            .CellEdit = bVal
        '/* advanced edit
        Case 1
            .AdvancedEdit = bVal
        '/* custom cursors
        Case 2
            .CustomCursors = bVal
        '/* lock first column
        Case 3
            .LockFirstColumn = bVal
        '/* drag and drop rows
        Case 4
            If bVal Then
                .OLEDropMode = vbOLEDropManual
            Else
                .OLEDropMode = vbOLEDropNone
            End If
        End Select
    End With

End Sub

Private Sub chkTextOnly_Click()
    ucVHGrid1.FocusTextOnly = CBool(chkTextOnly.Value = 1)
End Sub

Private Sub cmdDemo_Click(Index As Integer)

    Select Case Index
    Case 0
        Load frmVirtual
        frmVirtual.Show
    Case 1
        Load frmUnicode
        frmUnicode.Show
    Case 2
        Load frmOwnerDrawn
        frmOwnerDrawn.Show
    Case 3
        Load frmSubcell
        frmSubcell.Show
    Case 4
        Load frmHyperMode
        frmHyperMode.Show
    End Select
    
End Sub

Private Sub cmdMethods_Click(Index As Integer)
    
    With ucVHGrid1
        Select Case Index
        Case 0
            AddGridCell
        Case 1
            RemoveGridCell
        Case 2
            AddColumn
        Case 3
            RemoveColumn
        Case 4
            .CheckAll
        Case 5
            .UnCheckAll
        Case 6
            frmFind.Show vbModeless, Me
        Case 7
            .ClearList
        End Select
    End With
    
End Sub

Private Sub AddGridCell()

Dim lCount  As Long
Dim lRnd    As Long

    With ucVHGrid1
        '/* test for init before loading the item
        If Not .GridInitialized Then
            '/* init with 1 row, 8 columns
            .GridInit 1, 8
            '/* if adding one row at a time, turn the draw flag on first
            .Draw = True
            lCount = 0
        Else
            '/* the [Count] property returns all rows including span virtual members
            '/* the [RowCount] property returns the number of visible rows
            lCount = .RowCount
        End If
        lRnd = RandomNum(0, 4)
        '/* add the cell data
        .AddCell lCount, 0
        .AddCell lCount, 1, m_sAppName(lRnd), DT_LEFT Or DT_END_ELLIPSIS, lRnd, , , , 6
        .AddCell lCount, 2, m_sAppDesc(lRnd), DT_LEFT Or DT_END_ELLIPSIS
        .AddCell lCount, 3, m_sAppData(lRnd), DT_WORDBREAK Or DT_LEFT Or DT_VCENTER
        .AddCell lCount, 4, m_sAppStats(lRnd), DT_LEFT Or DT_END_ELLIPSIS
        .AddCell lCount, 5, "Yes", DT_LEFT Or DT_END_ELLIPSIS
        .AddCell lCount, 6, "Yes", DT_LEFT Or DT_END_ELLIPSIS
        .AddCell lCount, 7, CStr(lCount), DT_LEFT Or DT_END_ELLIPSIS
        .RowEnsureVisible lCount
    End With
    
End Sub

Private Sub RemoveGridCell()

Dim lCount  As Long
    
    With ucVHGrid1
        '/* removing the top row
        lCount = (.RowCount - 1)
        '/* delete the row
        .RowRemove lCount
        .RowEnsureVisible (lCount - 1)
    End With

End Sub

Private Sub AddColumn()

Dim lCt     As Long
Dim lColumn As Long

    With ucVHGrid1
        lColumn = .ColumnCount
        .ColumnAdd lColumn, "Column " & lColumn, 40, ecaColumnLeft, 1
        For lCt = 0 To (.RowCount - 1)
            .AddCell lCt, lColumn, "Row " & lCt, DT_LEFT Or DT_END_ELLIPSIS
        Next lCt
    End With

End Sub

Private Sub RemoveColumn()

Dim lColumn As Long

    With ucVHGrid1
        lColumn = (.ColumnCount - 1)
        .ColumnRemove lColumn
    End With

End Sub

Private Sub cmdReset_Click()
'/* reset skin style

Dim lFntClr     As Long
Dim lFntHilite  As Long
Dim lFntPress   As Long

    If Not IsNumeric(txtDepth.Text) Then
        MsgBox "Please choose a depth number between 0 and 4", vbExclamation, "Invalid Input!"
        Exit Sub
    End If
    Select Case m_lSkinStyle
    '/* azure
    Case 0
        m_lDecColor = m_lCustomClr(0)
        m_lDecOffset = m_lCustomClr(1)
        m_lThemeClr = m_lCustomClr(2)
        lFntClr = &H333333
        lFntHilite = m_lDecOffset
        lFntPress = m_lDecColor
    '/* classic
    Case 1
        m_lDecColor = m_lCustomClr(3)
        m_lDecOffset = m_lCustomClr(4)
        m_lThemeClr = m_lCustomClr(5)
        lFntClr = &H444444
        lFntHilite = m_lDecOffset
        lFntPress = m_lThemeClr
    '/* gloss
    Case 2
        m_lDecColor = m_lCustomClr(6)
        m_lDecOffset = m_lCustomClr(7)
        m_lThemeClr = m_lCustomClr(8)
        lFntClr = &H232323
        lFntHilite = m_lDecColor
        lFntPress = m_lDecOffset
    '/* metal
    Case 3
        m_lDecColor = m_lCustomClr(9)
        m_lDecOffset = m_lCustomClr(10)
        m_lThemeClr = m_lCustomClr(11)
        lFntClr = &H333333
        lFntHilite = m_lThemeClr
        lFntPress = m_lDecOffset
    '/* xp
    Case 4
        m_lDecColor = m_lCustomClr(12)
        m_lDecOffset = m_lCustomClr(13)
        m_lThemeClr = m_lCustomClr(14)
        lFntClr = &H333333
        lFntHilite = m_lDecColor
        lFntPress = m_lDecOffset
    End Select

    With ucVHGrid1
        .BackColor = m_lThemeClr
        .CellDecoration m_lDecoration, m_lDecColor, m_lDecOffset, True, CLng(txtDepth.Text)
        .ThemeManager m_lSkinStyle, CBool(chkSkin.Value), picThemeClr.BackColor, estThemeSoft, _
            lFntClr, lFntHilite, lFntPress, m_lThemeClr, m_lDecOffset, &H808080, 210, True, True, _
            m_lDecColor, lFntClr, False, True, True, False, True
        .Resize
        .GridRefresh False
    End With

End Sub

Private Sub optBorder_Click(Index As Integer)
'/* border style

    With ucVHGrid1
        Select Case Index
        '/* 3d
        Case 0
            .BorderStyle = ebsThick
        '/* thin
        Case 1
            .BorderStyle = ebsThin
        '/* none
        Case 2
            .BorderStyle = ebsNone
        End Select
    End With

End Sub

Private Sub optDecoration_Click(Index As Integer)
'/* cell decoration

    With ucVHGrid1
        Select Case Index
        '/* per cell
        Case 0
            .CellUseDecoration = False
        '/* structured
        Case 1
            .CellUseDecoration = True
        End Select
        .GridRefresh False
    End With

End Sub

Private Sub optDrag_Click(Index As Integer)
'/* drag and drop effect

    With ucVHGrid1
        Select Case Index
        '/* arrow
        Case 0
            .DragEffectStyle = edsClientArrow
        '/* thick line
        Case 1
            .DragEffectStyle = edsThinLine
        '/* thin line
        Case 2
            .DragEffectStyle = edsThickLine
        End Select
    End With

End Sub

Private Sub optFocus_Click(Index As Integer)
'/* focus effect

    With ucVHGrid1
        Select Case Index
        '/* normal
        Case 0
            .FocusAlphaBlend = False
            .AlphaBarActive = False
        '/* alpha blend
        Case 1
            .AlphaBarActive = False
            .FocusAlphaBlend = True
        '/* alpha bar
        Case 2
            .FocusAlphaBlend = False
            .AlphaBarActive = True
        End Select
    End With

End Sub

Private Sub optLines_Click(Index As Integer)
'/* grid lines

    With ucVHGrid1
        Select Case Index
        '/* verical
        Case 0
            .GridLines = EGLVertical
        '/* horizontal
        Case 1
            .GridLines = EGLHorizontal
        '/* both
        Case 2
            .GridLines = EGLBoth
        '/* none
        Case 3
            .GridLines = EGLNone
        End Select
        .GridRefresh False
    End With

End Sub

Private Sub OPTpOS_Click(Index As Integer)
    Select Case Index
    Case 0
        ucVHGrid1.TreeViewHeight = 150
        ucVHGrid1.TreeViewAlignment = etvTopAlign
    Case 1
        ucVHGrid1.TreeViewHeight = 150
        ucVHGrid1.TreeViewAlignment = etvBottomAlign
    Case 2
        ucVHGrid1.TreeViewWidth = 150
        ucVHGrid1.TreeViewAlignment = etvRightAlign
    Case 3
        ucVHGrid1.TreeViewWidth = 150
        ucVHGrid1.TreeViewAlignment = etvLeftAlign
    End Select
End Sub

Private Sub optRowDec_Click(Index As Integer)
    m_lDecoration = Index
End Sub

Private Sub optSorting_Click(Index As Integer)
'/* advanced options

    With ucVHGrid1
        Select Case Index
        '/* case sensitive
        Case 0
            .SortType = estCaseSensitive
        '/* case insensitive
        Case 1
            .SortType = estCaseInsensitive
        '/* none
        Case 2
            .SortType = estNone
        End Select
        .Resize
    End With

End Sub

Private Sub optStyles_Click(Index As Integer)
    m_lSkinStyle = Index
End Sub

Private Sub picFocus_Click()

Dim lRet        As Long

    lRet = ShowColor(Me.hWnd, &HFFFFFF, m_lCustomClr, 1)
    If Not (lRet = -1) Then
        picFocus.BackColor = lRet
        ucVHGrid1.CellFocusedColor = lRet
    End If

End Sub

Private Sub picGridline_Click()

Dim lRet As Long

    lRet = ShowColor(Me.hWnd, &HFFFFFF, m_lCustomClr, 1)
    If Not (lRet = -1) Then
        picGridline.BackColor = lRet
        ucVHGrid1.GridLineColor = lRet
        ucVHGrid1.GridRefresh False
    End If

End Sub

Private Sub picSelect_Click()

Dim lRet As Long

    lRet = ShowColor(Me.hWnd, &HFFFFFF, m_lCustomClr, 1)
    If Not (lRet = -1) Then
        picSelect.BackColor = lRet
        ucVHGrid1.CellSelectedColor = lRet
    End If

End Sub

Private Sub picThemeClr_Click()

Dim lRet As Long
    
    lRet = ShowColor(Me.hWnd, &HFFFFFF, m_lCustomClr, 1)
    If Not (lRet = -1) Then
        picThemeClr.BackColor = lRet
    End If
    
End Sub

Private Function TestInput() As Boolean
    
    If Not IsNumeric(txtCount.Text) Then
        MsgBox "Please choose a row number between 5 and 10000", vbExclamation, "Invalid Input!"
        Exit Function
    ElseIf (txtCount.Text) > 10000 Then
        MsgBox "Please choose a row number between 5 and 10000", vbExclamation, "Invalid Input!"
        Exit Function
    ElseIf CLng(txtCount.Text) < 5 Then
        MsgBox "Please choose a row number between 5 and 10000", vbExclamation, "Invalid Input!"
        Exit Function
    ElseIf ucVHGrid1.Count > 0 Then
        ucVHGrid1.ClearList
    End If
    '/* set the row count
    m_lRowCount = CLng((txtCount.Text))
    '/* success
    TestInput = True

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

Private Sub Form_Resize()

On Error Resume Next

    Picture5.Left = Me.Width - (Picture5.Width + 300)
    ucVHGrid1.Width = Picture5.Left - 500
    
On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not m_cTiming Is Nothing Then Set m_cTiming = Nothing
End Sub
