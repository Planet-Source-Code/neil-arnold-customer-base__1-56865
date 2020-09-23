VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCalMnth 
   BackColor       =   &H00CBD3D6&
   Caption         =   "By Month"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalMnth.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   9900
      TabIndex        =   144
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   4575
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9825
      Top             =   3900
   End
   Begin VB.PictureBox picTooltip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   7425
      ScaleHeight     =   1785
      ScaleWidth      =   3210
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   1725
      Visible         =   0   'False
      Width           =   3240
      Begin VB.TextBox txtTooltip 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   1800
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   0
         Width           =   3240
      End
   End
   Begin VB.PictureBox picSat 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   5
      Left            =   8175
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1290
      Begin VB.TextBox txtSat 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSat 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "7"
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   94
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSat 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   4
      Left            =   8175
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      Begin VB.TextBox txtSat 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSat 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "7"
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   92
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSat 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   3
      Left            =   8175
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      Begin VB.TextBox txtSat 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSat 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "7"
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   90
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSat 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   8175
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      Begin VB.TextBox txtSat 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSat 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "7"
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   88
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSat 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   8175
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      Begin VB.TextBox txtSat 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSat 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "7"
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   86
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSat 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   8175
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      Begin VB.TextBox txtSat 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSat 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "7"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   84
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picFri 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   5
      Left            =   6825
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1290
      Begin VB.TextBox txtFri 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblFri 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "6"
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   82
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picFri 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   4
      Left            =   6825
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      Begin VB.TextBox txtFri 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblFri 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "6"
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   80
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picFri 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   3
      Left            =   6825
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      Begin VB.TextBox txtFri 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblFri 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "6"
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   78
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picFri 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   6825
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      Begin VB.TextBox txtFri 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblFri 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "6"
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   76
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picFri 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   6825
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      Begin VB.TextBox txtFri 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblFri 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "6"
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   74
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picFri 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   6825
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      Begin VB.TextBox txtFri 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblFri 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "6"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   72
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picThu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   5
      Left            =   5475
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1290
      Begin VB.TextBox txtThu 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblThu 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "5"
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   70
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picThu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   4
      Left            =   5475
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      Begin VB.TextBox txtThu 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblThu 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "5"
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   68
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picThu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   3
      Left            =   5475
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      Begin VB.TextBox txtThu 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblThu 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "5"
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   66
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picThu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   5475
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      Begin VB.TextBox txtThu 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblThu 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "5"
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   64
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picThu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   5475
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      Begin VB.TextBox txtThu 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblThu 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "5"
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   62
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picThu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   5475
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      Begin VB.TextBox txtThu 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblThu 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "5"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   60
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picWed 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   5
      Left            =   4125
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1290
      Begin VB.TextBox txtWed 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblWed 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "4"
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   58
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picWed 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   4
      Left            =   4125
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      Begin VB.TextBox txtWed 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblWed 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "4"
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   56
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picWed 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   3
      Left            =   4125
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      Begin VB.TextBox txtWed 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblWed 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "4"
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   54
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picWed 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   4125
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      Begin VB.TextBox txtWed 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblWed 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "4"
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   52
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picWed 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   4125
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      Begin VB.TextBox txtWed 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblWed 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "4"
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   50
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picWed 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   4125
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      Begin VB.TextBox txtWed 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblWed 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "4"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   48
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picTue 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   5
      Left            =   2775
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1290
      Begin VB.TextBox txtTue 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblTue 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "3"
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   46
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picTue 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   4
      Left            =   2775
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      Begin VB.TextBox txtTue 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblTue 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "3"
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   44
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picTue 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   3
      Left            =   2775
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      Begin VB.TextBox txtTue 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblTue 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "3"
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   42
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picTue 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   2775
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      Begin VB.TextBox txtTue 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblTue 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "3"
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   40
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picTue 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   2775
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      Begin VB.TextBox txtTue 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblTue 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "3"
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   38
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picTue 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   2775
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      Begin VB.TextBox txtTue 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblTue 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "3"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   36
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picMon 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   5
      Left            =   1425
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1290
      Begin VB.TextBox txtMon 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "2"
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   34
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picMon 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   4
      Left            =   1425
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      Begin VB.TextBox txtMon 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "2"
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   32
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picMon 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   3
      Left            =   1425
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      Begin VB.TextBox txtMon 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "2"
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   30
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picMon 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   1425
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      Begin VB.TextBox txtMon 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "2"
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   28
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picMon 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   1425
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      Begin VB.TextBox txtMon 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "2"
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   26
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picMon 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   1425
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      Begin VB.TextBox txtMon 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblMon 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "2"
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   24
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSun 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   5
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6375
      Width           =   1290
      Begin VB.TextBox txtSun 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   5
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSun 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   22
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSun 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   4
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5325
      Width           =   1290
      Begin VB.TextBox txtSun 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSun 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   20
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSun 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   3
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
      Begin VB.TextBox txtSun 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSun 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   18
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSun 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   2
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1290
      Begin VB.TextBox txtSun 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   2
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSun 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   16
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSun 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   1
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2175
      Width           =   1290
      Begin VB.TextBox txtSun 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   1
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSun 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   14
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picSun 
      Appearance      =   0  'Flat
      BackColor       =   &H00E3E9EB&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Index           =   0
      Left            =   75
      ScaleHeight     =   990
      ScaleWidth      =   1290
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1290
      Begin VB.TextBox txtSun 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   690
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label lblSun 
         Alignment       =   2  'Center
         BackColor       =   &H00E3E9EB&
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   12
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H004A4A4A&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11355
      Begin VB.ComboBox cboYear 
         BackColor       =   &H004A4A4A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5100
         Style           =   2  'Dropdown List
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   50
         Width           =   1065
      End
      Begin VB.ComboBox cboMonth 
         BackColor       =   &H004A4A4A&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2475
         Style           =   2  'Dropdown List
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   50
         Width           =   2040
      End
      Begin MSComctlLib.TabStrip tbsMain 
         Height          =   315
         Left            =   8850
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   150
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Day"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Week"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Month"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "List"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   4575
         TabIndex        =   141
         Top             =   75
         Width           =   465
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " View month:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   1275
         TabIndex        =   139
         Top             =   75
         Width           =   1140
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmCalMnth.frx":0442
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Calendar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   450
         TabIndex        =   2
         Top             =   75
         Width           =   840
      End
   End
   Begin VB.Label lblDate 
      BackColor       =   &H004A4A4A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1275
      TabIndex        =   96
      Top             =   7500
      Width           =   4365
   End
   Begin VB.Label Label1 
      BackColor       =   &H004A4A4A&
      Caption         =   "Current Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   75
      TabIndex        =   95
      Top             =   7500
      Width           =   1215
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00B1CBD4&
      Height          =   90
      Left            =   9525
      Top             =   5925
      Width           =   240
   End
   Begin VB.Label lblWkDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   8175
      TabIndex        =   10
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label lblWkDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fri"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   6825
      TabIndex        =   9
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label lblWkDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Thu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   5475
      TabIndex        =   8
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label lblWkDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Wed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4125
      TabIndex        =   7
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label lblWkDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tue"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2775
      TabIndex        =   6
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label lblWkDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1425
      TabIndex        =   5
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label lblWkDay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sun"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   1290
   End
   Begin VB.Label lblCalHdr 
      Alignment       =   2  'Center
      BackColor       =   &H00B1CBD4&
      Caption         =   "month / year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   600
      Width           =   11265
   End
End
Attribute VB_Name = "frmCalMnth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_vDate As Variant
Dim m_strMonth As String
Dim m_intYear As Integer

Private Sub cboMonth_Click()
   Const sMOD_NAME As String = "frmCalMnth.cboMonth_Click"
   On Error GoTo Error_Handler
   
   m_strMonth = cboMonth.Text
   If (cboYear.Text = "") Then
      m_intYear = Year(Now())
   Else
      m_intYear = cboYear.Text
   End If
   
   'load the calendar
   Call EmptyDateValues
   Call CreateCalendar
   Call HighlightCalendarDates
   Call LoadAppointments
   
   lblCalHdr.Caption = m_strMonth & ", " & m_intYear
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while changing Calendar Dates!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub cboYear_Click()
   Const sMOD_NAME As String = "frmCalMnth.cboYear_Click"
   On Error GoTo Error_Handler
   
   m_intYear = cboYear.Text
   If (cboMonth.Text <> "") Then
      m_strMonth = cboMonth.Text
   End If
   
   'load the calendar
   Call EmptyDateValues
   Call CreateCalendar
   Call HighlightCalendarDates
   Call LoadAppointments
   
   lblCalHdr.Caption = m_strMonth & ", " & m_intYear
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while changing Calendar Dates!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   
   lblCalHdr.Caption = Format(m_vDate, "mmmm, yyyy")
   lblDate.Caption = Format(m_vDate, "dddd - mmmm dd, yyyy")
   cboMonth.Text = m_strMonth
   cboYear = m_intYear
   
   tbsMain.Tabs(3).Selected = True
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmCalMnth.Form_Load"
   On Error GoTo Error_Handler
   
   Dim intYrs As Integer
   Dim iYrCtr As Integer
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Calendar Month View", True
   frmMain.picStatus.BackColor = &H4A4A4A
   
   Me.Hide
   
   'load month combo
   cboMonth.AddItem "January"
   cboMonth.AddItem "February"
   cboMonth.AddItem "March"
   cboMonth.AddItem "April"
   cboMonth.AddItem "May"
   cboMonth.AddItem "June"
   cboMonth.AddItem "July"
   cboMonth.AddItem "August"
   cboMonth.AddItem "September"
   cboMonth.AddItem "October"
   cboMonth.AddItem "November"
   cboMonth.AddItem "December"
   
   'load year combo
   intYrs = Year(Now())
   For iYrCtr = intYrs To 3000
      cboYear.AddItem iYrCtr
   Next iYrCtr
   
   m_vDate = Date
   
   'set current month & year
   m_strMonth = Format(m_vDate, "mmmm")
   m_intYear = Format(m_vDate, "yyyy")
   
   'set global form identifier
   g_strFormFlag = "CMnth"
   
   Me.Show
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   ShowError
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
   
   On Error Resume Next
   
   LockWindowUpdate frmCalMnth.hWnd
   
   'adjust calendar header
   lblCalHdr.Move 75, picBanner.Height + 150, Me.ScaleWidth - 150
   'adjust weekday labels
   lblWkDay(0).Move lblCalHdr.Left, lblCalHdr.Top + 240, lblCalHdr.Width \ 7
   lblWkDay(1).Move lblWkDay(0).Left + lblWkDay(0).Width, lblWkDay(0).Top, lblCalHdr.Width \ 7
   lblWkDay(2).Move lblWkDay(1).Left + lblWkDay(1).Width, lblWkDay(1).Top, lblCalHdr.Width \ 7
   lblWkDay(3).Move lblWkDay(2).Left + lblWkDay(2).Width, lblWkDay(2).Top, lblCalHdr.Width \ 7
   lblWkDay(4).Move lblWkDay(3).Left + lblWkDay(3).Width, lblWkDay(3).Top, lblCalHdr.Width \ 7
   lblWkDay(5).Move lblWkDay(4).Left + lblWkDay(4).Width, lblWkDay(4).Top, lblCalHdr.Width \ 7
   lblWkDay(6).Move lblWkDay(5).Left + lblWkDay(5).Width, lblWkDay(5).Top, lblCalHdr.Width \ 7
   'adjust sunday grid cells
   picSun(0).Move lblWkDay(0).Left, lblWkDay(0).Top + 254, lblWkDay(0).Width - 7
   picSun(1).Move picSun(0).Left, picSun(0).Top + 1004, picSun(0).Width
   picSun(2).Move picSun(1).Left, picSun(1).Top + 1004, picSun(1).Width
   picSun(3).Move picSun(2).Left, picSun(2).Top + 1004, picSun(2).Width
   picSun(4).Move picSun(3).Left, picSun(3).Top + 1004, picSun(3).Width
   picSun(5).Move picSun(4).Left, picSun(4).Top + 1004, picSun(4).Width
   'adjust monday grid cells
   picMon(0).Move lblWkDay(1).Left + 7, lblWkDay(1).Top + 254, lblWkDay(1).Width - 14
   picMon(1).Move picMon(0).Left, picSun(1).Top, picMon(0).Width
   picMon(2).Move picMon(1).Left, picSun(2).Top, picMon(1).Width
   picMon(3).Move picMon(2).Left, picSun(3).Top, picMon(2).Width
   picMon(4).Move picMon(3).Left, picSun(4).Top, picMon(3).Width
   picMon(5).Move picMon(4).Left, picSun(5).Top, picMon(4).Width
   'adjust tuesday grid cells
   picTue(0).Move lblWkDay(2).Left + 7, lblWkDay(2).Top + 254, lblWkDay(2).Width - 14
   picTue(1).Move picTue(0).Left, picMon(1).Top, picTue(0).Width
   picTue(2).Move picTue(1).Left, picMon(2).Top, picTue(1).Width
   picTue(3).Move picTue(2).Left, picMon(3).Top, picTue(2).Width
   picTue(4).Move picTue(3).Left, picMon(4).Top, picTue(3).Width
   picTue(5).Move picTue(4).Left, picMon(5).Top, picTue(4).Width
   'adjust wednesday grid cells
   picWed(0).Move lblWkDay(3).Left + 7, lblWkDay(3).Top + 254, lblWkDay(3).Width - 14
   picWed(1).Move picWed(0).Left, picTue(1).Top, picWed(0).Width
   picWed(2).Move picWed(1).Left, picTue(2).Top, picWed(1).Width
   picWed(3).Move picWed(2).Left, picTue(3).Top, picWed(2).Width
   picWed(4).Move picWed(3).Left, picTue(4).Top, picWed(3).Width
   picWed(5).Move picWed(4).Left, picTue(5).Top, picWed(4).Width
   'adjust thursday grid cells
   picThu(0).Move lblWkDay(4).Left + 7, lblWkDay(4).Top + 254, lblWkDay(4).Width - 14
   picThu(1).Move picThu(0).Left, picWed(1).Top, picThu(0).Width
   picThu(2).Move picThu(1).Left, picWed(2).Top, picThu(1).Width
   picThu(3).Move picThu(2).Left, picWed(3).Top, picThu(2).Width
   picThu(4).Move picThu(3).Left, picWed(4).Top, picThu(3).Width
   picThu(5).Move picThu(4).Left, picWed(5).Top, picThu(4).Width
   'adjust friday grid cells
   picFri(0).Move lblWkDay(5).Left + 7, lblWkDay(5).Top + 254, lblWkDay(5).Width - 14
   picFri(1).Move picFri(0).Left, picThu(1).Top, picFri(0).Width
   picFri(2).Move picFri(1).Left, picThu(2).Top, picFri(1).Width
   picFri(3).Move picFri(2).Left, picThu(3).Top, picFri(2).Width
   picFri(4).Move picFri(3).Left, picThu(4).Top, picFri(3).Width
   picFri(5).Move picFri(4).Left, picThu(5).Top, picFri(4).Width
   'adjust saturday grid cells
   picSat(0).Move lblWkDay(6).Left + 7, lblWkDay(6).Top + 254, lblWkDay(6).Width - 14
   picSat(1).Move picSat(0).Left, picFri(1).Top, picSat(0).Width
   picSat(2).Move picSat(1).Left, picFri(2).Top, picSat(1).Width
   picSat(3).Move picSat(2).Left, picFri(3).Top, picSat(2).Width
   picSat(4).Move picSat(3).Left, picFri(4).Top, picSat(3).Width
   picSat(5).Move picSat(4).Left, picFri(5).Top, picSat(4).Width
   'adjust border
   shpBorder.Move lblCalHdr.Left - 15, lblCalHdr.Top - 15, lblCalHdr.Width + 30, 6540
   'adjust date monitor
   Label1.Move lblCalHdr.Left, shpBorder.Top + shpBorder.Height + 150
   lblDate.Move Label1.Left + Label1.Width, Label1.Top
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   Set frmCalMnth = Nothing
End Sub

Private Sub picBanner_Resize()
   tbsMain.Move picBanner.ScaleWidth - tbsMain.Width
End Sub

'////Original Code by Gaurav Arora////
Public Sub CreateCalendar()
   'set up the calendar grid
   Const sMOD_NAME As String = "frmCalMnth.CreateCalendar"
   On Error GoTo Error_Handler
   
   Dim dtDate As Date
   Dim strWkDay As String
   Dim intDaysMonth As Integer
   Dim intCC As Integer
   Dim intCCC As Integer
   Dim intM As Integer
    
   dtDate = CDate(m_strMonth & " 01," & m_intYear)
   
   intDaysMonth = GetDaysPerMonth(m_strMonth, m_intYear)
   
   strWkDay = WeekdayName(Weekday(dtDate))
   
   Select Case strWkDay
      Case "Monday"
         intCC = 1
         lblMon(intM).Caption = "1"
         For intM = 1 To 5 'calculate monday days
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblMon(intM).Caption = intCC
         Next intM
         'calculate tuesday
         intM = 0
         intCC = 2
         lblTue(intM).Caption = "2"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblTue(intM).Caption = intCC
         Next
         'calculate wednesday
         intM = 0
         intCC = 3
         lblWed(intM).Caption = "3"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblWed(intM).Caption = intCC
         Next intM
         'calculate thursday
         intM = 0
         intCC = 4
         lblThu(intM).Caption = "4"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblThu(intM).Caption = intCC
         Next intM
         'calculate friday
         intM = 0
         intCC = 5
         lblFri(intM).Caption = "5"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblFri(intM).Caption = intCC
         Next intM
         'calculate saturday
         intM = 0
         intCC = 6
         lblSat(intM).Caption = "6"
         For intM = 1 To 5
            intCC = intCC + 7
            intCCC = intCC
            If intCC > intDaysMonth Then Exit For
            lblSat(intM).Caption = intCC
         Next intM
         'calculate sunday
         intM = 1
         intCC = 7
         lblSun(intM).Caption = "7"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSun(intM).Caption = intCC
         Next intM
         
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblSun(0).Caption = CStr(intCCC)
         End If
      Case "Tuesday"
         intCC = 1
         lblTue(intM).Caption = "1"
         For intM = 1 To 5 'calculate tuesday days
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblTue(intM).Caption = intCC
         Next intM
         'calculate wednesday
         intM = 0
         intCC = 2
         lblWed(intM).Caption = "2"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblWed(intM).Caption = intCC
         Next intM
         'calculate thursday
         intM = 0
         intCC = 3
         lblThu(intM).Caption = "3"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblThu(intM).Caption = intCC
         Next intM
         'calculate friday
         intM = 0
         intCC = 4
         lblFri(intM).Caption = "4"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblFri(intM).Caption = intCC
         Next intM
         'calculate saturday
         intM = 0
         intCC = 5
         lblSat(intM).Caption = "5"
         For intM = 1 To 5
            intCC = intCC + 7
            intCCC = intCC
            If intCC > intDaysMonth Then Exit For
            lblSat(intM).Caption = intCC
         Next intM
         'calculate sunday
         intM = 1
         intCC = 6
         lblSun(intM).Caption = "6"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSun(intM).Caption = intCC
         Next intM
         'calculate monday
         intM = 1
         intCC = 7
         lblMon(intM).Caption = "7"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblMon(intM).Caption = intCC
         Next intM
         
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblSun(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblMon(0).Caption = CStr(intCCC)
         End If
      Case "Wednesday"
         intCC = 1
         lblWed(intM).Caption = "1"
         For intM = 1 To 5 'calculate wednesday days
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblWed(intM).Caption = intCC
         Next intM
         'calculate thursday
         intM = 0
         intCC = 2
         lblThu(intM).Caption = "2"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblThu(intM).Caption = intCC
         Next intM
         'calculate friday
         intM = 0
         intCC = 3
         lblFri(intM).Caption = "3"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblFri(intM).Caption = intCC
         Next intM
         'calculate saturday
         intM = 0
         intCC = 4
         lblSat(intM).Caption = "4"
         For intM = 1 To 5
            intCC = intCC + 7
            intCCC = intCC
            If intCC > intDaysMonth Then Exit For
            lblSat(intM).Caption = intCC
         Next intM
         'calculate sunday
         intM = 1
         intCC = 5
         lblSun(intM).Caption = "5"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSun(intM).Caption = intCC
         Next intM
         'calculate monday
         intM = 1
         intCC = 6
         lblMon(intM).Caption = "6"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblMon(intM).Caption = intCC
         Next intM
         'calculate tuesday
         intM = 1
         intCC = 7
         lblTue(intM).Caption = "7"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblTue(intM).Caption = intCC
         Next intM
         
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblSun(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblMon(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblTue(0).Caption = CStr(intCCC)
         End If
      Case "Thursday"
         intCC = 1
         lblThu(intM).Caption = "1"
         For intM = 1 To 5 'calculate thursday days
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblThu(intM).Caption = intCC
         Next intM
         'calculate friday
         intM = 0
         intCC = 2
         lblFri(intM).Caption = "2"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblFri(intM).Caption = intCC
         Next intM
         'calculate saturday
         intM = 0
         intCC = 3
         lblSat(intM).Caption = "3"
         For intM = 1 To 5
            intCC = intCC + 7
            intCCC = intCC
            If intCC > intDaysMonth Then Exit For
            lblSat(intM).Caption = intCC
         Next intM
         'calculate sunday
         intM = 1
         intCC = 4
         lblSun(intM).Caption = "4"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSun(intM).Caption = intCC
         Next intM
         'calculate monday
         intM = 1
         intCC = 5
         lblMon(intM).Caption = "5"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblMon(intM).Caption = intCC
         Next intM
         'calculate tuesday
         intM = 1
         intCC = 6
         lblTue(intM).Caption = "6"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblTue(intM).Caption = intCC
         Next intM
         'calculate wednesday
         intM = 1
         intCC = 7
         lblWed(intM).Caption = "7"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblWed(intM).Caption = intCC
         Next intM
         
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblSun(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblMon(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblTue(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblWed(0).Caption = CStr(intCCC)
         End If
      Case "Friday"
         intCC = 1
         lblFri(intM).Caption = "1"
         For intM = 1 To 5 'calculate friday days
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblFri(intM).Caption = intCC
         Next intM
         'calculate saturday
         intM = 0
         intCC = 2
         lblSat(intM).Caption = "2"
         For intM = 1 To 5
            intCC = intCC + 7
            intCCC = intCC
            If intCC > intDaysMonth Then Exit For
            lblSat(intM).Caption = intCC
         Next intM
         'calculate sunday
         intM = 1
         intCC = 3
         lblSun(intM).Caption = "3"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSun(intM).Caption = intCC
         Next intM
         'calculate monday
         intM = 1
         intCC = 4
         lblMon(intM).Caption = "4"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblMon(intM).Caption = intCC
         Next intM
         'calculate tuesday
         intM = 1
         intCC = 5
         lblTue(intM).Caption = "5"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblTue(intM).Caption = intCC
         Next intM
         'calculate wednesday
         intM = 1
         intCC = 6
         lblWed(intM).Caption = "6"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblWed(intM).Caption = intCC
         Next intM
         'calculate thursday
         intM = 1
         intCC = 7
         lblThu(intM).Caption = "7"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblThu(intM).Caption = intCC
         Next intM
         
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblSun(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblMon(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblTue(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblWed(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblThu(0).Caption = CStr(intCCC)
         End If
      Case "Saturday"
         intM = 0 'calculate saturday days
         intCC = 1
         lblSat(intM).Caption = "1"
         For intM = 1 To 5
            intCC = intCC + 7
            intCCC = intCC
            If intCC > intDaysMonth Then Exit For
            lblSat(intM).Caption = intCC
         Next intM
         'calculate sunday
         intM = 1
         intCC = 2
         lblSun(intM).Caption = "2"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSun(intM).Caption = intCC
         Next intM
         'calculate monday
         intM = 1
         intCC = 3
         lblMon(intM).Caption = "3"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblMon(intM).Caption = intCC
         Next intM
         'calculate tuesday
         intM = 1
         intCC = 4
         lblTue(intM).Caption = "4"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblTue(intM).Caption = intCC
         Next intM
         'calculate wednesday
         intM = 1
         intCC = 5
         lblWed(intM).Caption = "5"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblWed(intM).Caption = intCC
         Next intM
         'calculate thursday
         intM = 1
         intCC = 6
         lblThu(intM).Caption = "6"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblThu(intM).Caption = intCC
         Next intM
         'calculate friday
         intM = 1
         intCC = 7
         lblFri(intM).Caption = "7"
         For intM = 2 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblFri(intM).Caption = intCC
         Next intM
         
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblSun(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblMon(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblTue(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblWed(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblThu(0).Caption = CStr(intCCC)
         End If
         If intCCC < intDaysMonth Then
            intCCC = intCCC + 1
            lblFri(0).Caption = CStr(intCCC)
         End If
      Case "Sunday"
         intCC = 1 'calculate sunday days
         lblSun(intM).Caption = "1"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSun(intM).Caption = intCC
         Next intM
         'calculate monday
         intM = 0
         intCC = 2
         lblMon(intM).Caption = "2"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblMon(intM).Caption = intCC
         Next intM
         'calculate tuesday
         intM = 0
         intCC = 3
         lblTue(intM).Caption = "3"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblTue(intM).Caption = intCC
         Next intM
         'calculate wednesday
         intM = 0
         intCC = 4
         lblWed(intM).Caption = "4"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblWed(intM).Caption = intCC
         Next intM
         'calculate thursday
         intM = 0
         intCC = 5
         lblThu(intM).Caption = "5"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblThu(intM).Caption = intCC
         Next intM
         'calculate friday
         intM = 0
         intCC = 6
         lblFri(intM).Caption = "6"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblFri(intM).Caption = intCC
         Next intM
         'calculate saturday
         intM = 0
         intCC = 7
         lblSat(intM).Caption = "7"
         For intM = 1 To 5
            intCC = intCC + 7
            If intCC > intDaysMonth Then Exit For
            lblSat(intM).Caption = intCC
         Next intM
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Function GetDaysPerMonth(sMonth As String, iYear As Integer) As Integer
   'get the number of days in the month passed
   Const sMOD_NAME As String = "frmCalMnth.GetDaysPerMonth"
   On Error GoTo Error_Handler
   
   Select Case sMonth
      Case "January"
         GetDaysPerMonth = 31
      Case "February"
         If (iYear Mod 4) = 0 Or (iYear Mod 400) = 0 Then 'leap year
            GetDaysPerMonth = 29
         Else
            GetDaysPerMonth = 28
         End If
      Case "March"
         GetDaysPerMonth = 31
      Case "April"
         GetDaysPerMonth = 30
      Case "May"
         GetDaysPerMonth = 31
      Case "June"
         GetDaysPerMonth = 30
      Case "July"
         GetDaysPerMonth = 31
      Case "August"
         GetDaysPerMonth = 31
      Case "September"
         GetDaysPerMonth = 30
      Case "October"
         GetDaysPerMonth = 31
      Case "November"
         GetDaysPerMonth = 30
      Case "December"
         GetDaysPerMonth = 31
   End Select
   
   Exit Function
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

Public Sub EmptyDateValues()
   Dim iCtr As Integer
   
   For iCtr = 0 To 5
      lblSun(iCtr).Caption = ""
      txtSun(iCtr).Text = ""
      lblMon(iCtr).Caption = ""
      txtMon(iCtr).Text = ""
      lblTue(iCtr).Caption = ""
      txtTue(iCtr).Text = ""
      lblWed(iCtr).Caption = ""
      txtWed(iCtr).Text = ""
      lblThu(iCtr).Caption = ""
      txtThu(iCtr).Text = ""
      lblFri(iCtr).Caption = ""
      txtFri(iCtr).Text = ""
      lblSat(iCtr).Caption = ""
      txtSat(iCtr).Text = ""
   Next iCtr
End Sub

Public Sub HighlightCalendarDates()
   Const sMOD_NAME As String = "frmCalMnth.HighlightCalendarDates"
   On Error GoTo Error_Handler
   
   Dim iCtr As Integer
   
   'make sunday days bold & all other regular
   For iCtr = 0 To 5
      If lblSun(iCtr).Caption <> CStr(Day(Date)) Then
         lblSun(iCtr).FontBold = True
         lblSun(iCtr).ForeColor = vbRed
         lblSun(iCtr).BackColor = &HE3E9EB
      End If
      If lblMon(iCtr).Caption <> CStr(Day(Date)) Then
         lblMon(iCtr).ForeColor = vbBlack
         lblMon(iCtr).FontBold = False
         lblMon(iCtr).BackColor = &HE3E9EB
      End If
      If lblTue(iCtr).Caption <> CStr(Day(Date)) Then
         lblTue(iCtr).ForeColor = vbBlack
         lblTue(iCtr).FontBold = False
         lblTue(iCtr).BackColor = &HE3E9EB
      End If
      If lblWed(iCtr).Caption <> CStr(Day(Date)) Then
         lblWed(iCtr).ForeColor = vbBlack
         lblWed(iCtr).FontBold = False
         lblWed(iCtr).BackColor = &HE3E9EB
      End If
      If lblThu(iCtr).Caption <> CStr(Day(Date)) Then
         lblThu(iCtr).ForeColor = vbBlack
         lblThu(iCtr).FontBold = False
         lblThu(iCtr).BackColor = &HE3E9EB
      End If
      If lblFri(iCtr).Caption <> CStr(Day(Date)) Then
         lblFri(iCtr).ForeColor = vbBlack
         lblFri(iCtr).FontBold = False
         lblFri(iCtr).BackColor = &HE3E9EB
      End If
      If lblSat(iCtr).Caption <> CStr(Day(Date)) Then
         lblSat(iCtr).ForeColor = vbBlack
         lblSat(iCtr).FontBold = False
         lblSat(iCtr).BackColor = &HE3E9EB
      End If
   Next iCtr
   
   'show current day
   For iCtr = 0 To 5
      If lblSun(iCtr).Caption = CStr(Day(Date)) Then
         lblSun(iCtr).FontBold = True
         lblSun(iCtr).ForeColor = vbRed
         lblSun(iCtr).BackColor = &H4A4A4A
      End If
      If lblMon(iCtr).Caption = CStr(Day(Date)) Then
         lblMon(iCtr).FontBold = True
         lblMon(iCtr).ForeColor = vbWhite
         lblMon(iCtr).BackColor = &H4A4A4A
      End If
      If lblTue(iCtr).Caption = CStr(Day(Date)) Then
         lblTue(iCtr).FontBold = True
         lblTue(iCtr).ForeColor = vbWhite
         lblTue(iCtr).BackColor = &H4A4A4A
      End If
      If lblWed(iCtr).Caption = CStr(Day(Date)) Then
         lblWed(iCtr).FontBold = True
         lblWed(iCtr).ForeColor = vbWhite
         lblWed(iCtr).BackColor = &H4A4A4A
      End If
      If lblThu(iCtr).Caption = CStr(Day(Date)) Then
         lblThu(iCtr).FontBold = True
         lblThu(iCtr).ForeColor = vbWhite
         lblThu(iCtr).BackColor = &H4A4A4A
      End If
      If lblFri(iCtr).Caption = CStr(Day(Date)) Then
         lblFri(iCtr).FontBold = True
         lblFri(iCtr).ForeColor = vbWhite
         lblFri(iCtr).BackColor = &H4A4A4A
      End If
      If lblSat(iCtr).Caption = CStr(Day(Date)) Then
         lblSat(iCtr).FontBold = True
         lblSat(iCtr).ForeColor = vbWhite
         lblSat(iCtr).BackColor = &H4A4A4A
      End If
   Next iCtr
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub picFri_Resize(Index As Integer)
   txtFri(Index).Left = 0
   txtFri(Index).Width = picFri(Index).ScaleWidth
End Sub

Private Sub picMon_Resize(Index As Integer)
   txtMon(Index).Left = 0
   txtMon(Index).Width = picMon(Index).ScaleWidth
End Sub

Private Sub picSat_Resize(Index As Integer)
   txtSat(Index).Left = 0
   txtSat(Index).Width = picSat(Index).ScaleWidth
End Sub

Private Sub picSun_Resize(Index As Integer)
   txtSun(Index).Left = 0
   txtSun(Index).Width = picSun(Index).ScaleWidth
End Sub

Private Sub picThu_Resize(Index As Integer)
   txtThu(Index).Left = 0
   txtThu(Index).Width = picThu(Index).ScaleWidth
End Sub

Private Sub picTue_Resize(Index As Integer)
   txtTue(Index).Left = 0
   txtTue(Index).Width = picTue(Index).ScaleWidth
End Sub

Private Sub picWed_Resize(Index As Integer)
   txtWed(Index).Left = 0
   txtWed(Index).Width = picWed(Index).ScaleWidth
End Sub

Public Sub LoadAppointments()
   Const sMOD_NAME As String = "frmCalMnth.LoadAppointments"
   On Error GoTo Error_Handler
   
   'load all appointments for the month into the calendar grid
   Dim intMonth As Integer
   Dim intEndDay As Integer
   Dim vMonthStart As Variant
   Dim vStartDate As Variant
   Dim vEndDate As Variant
   Dim strMnthMod As String
   Dim SQL As String
   Dim strTime As String
   Dim strSubj As String
   Dim strAppt As String
   
   'current month: mmmm/dd/yyyy
   vMonthStart = m_strMonth & "/01/" & m_intYear
   
   'get number of current month
   intMonth = Month(vMonthStart)
   
   'convert month number to string, if less than 10 add leading zero
   If (intMonth <= 9) Then
      strMnthMod = "0" & CStr(intMonth)
   Else
      strMnthMod = CStr(intMonth)
   End If
   
   'build start date string
   vStartDate = "#" & strMnthMod & "/" & "01" & "/" & m_intYear & "#"
   
   'get ending date
   intEndDay = GetDaysPerMonth(m_strMonth, m_intYear)
   vEndDate = "#" & strMnthMod & "/" & intEndDay & "/" & m_intYear & "#"
   
   SQL = "SELECT DateFrom, TimeFrom, Subject FROM Appts "
   SQL = SQL & "WHERE DateFrom BETWEEN " & vStartDate
   SQL = SQL & " AND " & vEndDate
   SQL = SQL & " ORDER BY TimeFrom Desc"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!TimeFrom)) Then
               strTime = Format(!TimeFrom, "h:nn AMPM") & " "
            Else
               strTime = "No Time"
            End If
            If (Not IsNull(!Subject)) Then
               strSubj = !Subject
            End If
            strAppt = strTime & vbCrLf & strSubj & vbCrLf
            'txtMon(3).Text = strAppt
            If (Not IsNull(!DateFrom)) Then
               Call SetTextInCell(!DateFrom, strAppt)
            End If
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the Appointments!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub SetTextInCell(dtDay As Date, strText As String)
   Const sMOD_NAME As String = "frmCalMnth.SetTextInCell"
   On Error GoTo Error_Handler
   
   Dim intDayNum As Integer
   Dim strDayNum As String
   Dim iCtr As Integer
   
   intDayNum = DatePart("d", dtDay)
   strDayNum = CStr(intDayNum)
   
   For iCtr = 0 To 5
      If (lblSun(iCtr).Caption = strDayNum) Then
         txtSun(iCtr).SelText = strText
         txtSun(iCtr).SelStart = 0
      End If
      If (lblMon(iCtr).Caption = strDayNum) Then
         txtMon(iCtr).SelText = strText
         txtMon(iCtr).SelStart = 0
      End If
      If (lblTue(iCtr).Caption = strDayNum) Then
         txtTue(iCtr).SelText = strText
         txtTue(iCtr).SelStart = 0
      End If
      If (lblWed(iCtr).Caption = strDayNum) Then
         txtWed(iCtr).SelText = strText
         txtWed(iCtr).SelStart = 0
      End If
      If (lblThu(iCtr).Caption = strDayNum) Then
         txtThu(iCtr).SelText = strText
         txtThu(iCtr).SelStart = 0
      End If
      If (lblFri(iCtr).Caption = strDayNum) Then
         txtFri(iCtr).SelText = strText
         txtFri(iCtr).SelStart = 0
      End If
      If (lblSat(iCtr).Caption = strDayNum) Then
         txtSat(iCtr).SelText = strText
         txtSat(iCtr).SelStart = 0
      End If
   Next iCtr
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while setting the Appointment Text!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub tbsMain_Click()
   Select Case tbsMain.SelectedItem.Index
      Case 1 'Day
         UnloadAllForms
         frmCalDay.m_blnIsSystem = True
         Load frmCalDay
      Case 2 'Week
         UnloadAllForms
         Load frmCalWeek
      Case 3 'Month
         'take no action
      Case 4 'List
         UnloadAllForms
         Load frmCalList
   End Select
End Sub

Private Sub Timer1_Timer()
   Text1.Text = Text1.Text + 1
   If Text1.Text = 7 Then
      picTooltip.Visible = False
      Timer1.Enabled = False
      Text1.Text = "1"
   End If
End Sub

Private Sub txtFri_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalMnth.txtFri_Click"
   On Error GoTo Error_Handler
   
   If (txtFri(Index).Text = "") Then
      picFri(Index).SetFocus
      Exit Sub
   End If
   If (Index = 4 Or Index = 5) Then
      picTooltip.Move picFri(Index).Left, picFri(Index).Top - picTooltip.Height - 15
   Else
      picTooltip.Move picFri(Index).Left, picFri(Index).Top + picFri(Index).Height + 15
   End If
   picTooltip.Visible = True
   txtTooltip.Text = txtFri(Index).Text
   Text1.Text = "1"
   Timer1.Enabled = True
   picTooltip.SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtMon_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalMnth.txtMon_Click"
   On Error GoTo Error_Handler
   
   If (txtMon(Index).Text = "") Then
      picMon(Index).SetFocus
      Exit Sub
   End If
   If (Index = 4 Or Index = 5) Then
      picTooltip.Move picMon(Index).Left, picMon(Index).Top - picTooltip.Height - 15
   Else
      picTooltip.Move picMon(Index).Left, picMon(Index).Top + picMon(Index).Height + 15
   End If
   picTooltip.Visible = True
   txtTooltip.Text = txtMon(Index).Text
   Text1.Text = "1"
   Timer1.Enabled = True
   picTooltip.SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtSat_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalMnth.txtSat_Click"
   On Error GoTo Error_Handler
   
   If (txtSat(Index).Text = "") Then
      picSat(Index).SetFocus
      Exit Sub
   End If
   If (Index = 4 Or Index = 5) Then
      picTooltip.Move picSat(Index).Left - (picTooltip.Width - picSat(Index).Width), picSat(Index).Top - picTooltip.Height - 15
   Else
      picTooltip.Move picSat(Index).Left - (picTooltip.Width - picSat(Index).Width), picSat(Index).Top + picSat(Index).Height + 15
   End If
   picTooltip.Visible = True
   txtTooltip.Text = txtSat(Index).Text
   Text1.Text = "1"
   Timer1.Enabled = True
   picTooltip.SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtSun_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalMnth.txtSun_Click"
   On Error GoTo Error_Handler
   
   If (txtSun(Index).Text = "") Then
      picSun(Index).SetFocus
      Exit Sub
   End If
   If (Index = 4 Or Index = 5) Then
      picTooltip.Move picSun(Index).Left, picSun(Index).Top - picTooltip.Height - 15
   Else
      picTooltip.Move picSun(Index).Left, picSun(Index).Top + picSun(Index).Height + 15
   End If
   picTooltip.Visible = True
   txtTooltip.Text = txtSun(Index).Text
   Text1.Text = "1"
   Timer1.Enabled = True
   picTooltip.SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtThu_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalMnth.txtThu_Click"
   On Error GoTo Error_Handler
   
   If (txtThu(Index).Text = "") Then
      picThu(Index).SetFocus
      Exit Sub
   End If
   If (Index = 4 Or Index = 5) Then
      picTooltip.Move picThu(Index).Left, picThu(Index).Top - picTooltip.Height - 15
   Else
      picTooltip.Move picThu(Index).Left, picThu(Index).Top + picThu(Index).Height + 15
   End If
   picTooltip.Visible = True
   txtTooltip.Text = txtThu(Index).Text
   Text1.Text = "1"
   Timer1.Enabled = True
   picTooltip.SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtTue_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalMnth.txtTue_Click"
   On Error GoTo Error_Handler
   
   If (txtTue(Index).Text = "") Then
      picTue(Index).SetFocus
      Exit Sub
   End If
   If (Index = 4 Or Index = 5) Then
      picTooltip.Move picTue(Index).Left, picTue(Index).Top - picTooltip.Height - 15
   Else
      picTooltip.Move picTue(Index).Left, picTue(Index).Top + picTue(Index).Height + 15
   End If
   picTooltip.Visible = True
   txtTooltip.Text = txtTue(Index).Text
   Text1.Text = "1"
   Timer1.Enabled = True
   picTooltip.SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtWed_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalMnth.txtWed_Click"
   On Error GoTo Error_Handler
   
   If (txtWed(Index).Text = "") Then
      picWed(Index).SetFocus
      Exit Sub
   End If
   If (Index = 4 Or Index = 5) Then
      picTooltip.Move picWed(Index).Left, picWed(Index).Top - picTooltip.Height - 15
   Else
      picTooltip.Move picWed(Index).Left, picWed(Index).Top + picWed(Index).Height + 15
   End If
   picTooltip.Visible = True
   txtTooltip.Text = txtWed(Index).Text
   Text1.Text = "1"
   Timer1.Enabled = True
   picTooltip.SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
