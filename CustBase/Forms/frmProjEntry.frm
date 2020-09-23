VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProjEntry 
   Caption         =   "Projects"
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
   Icon            =   "frmProjEntry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPane3 
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   7575
      ScaleHeight     =   2715
      ScaleWidth      =   2115
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   4725
      Visible         =   0   'False
      Width           =   2115
      Begin VB.ComboBox cboHistFilter 
         BackColor       =   &H00B0C0D6&
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   150
         Width           =   1890
      End
      Begin VB.CommandButton cmdAddFile 
         BackColor       =   &H00B0C0D6&
         Caption         =   "Add A File ..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10125
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   150
         Visible         =   0   'False
         Width           =   1065
      End
      Begin MSComctlLib.ImageList imlLView 
         Left            =   300
         Top             =   525
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProjEntry.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProjEntry.frx":095C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProjEntry.frx":0E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProjEntry.frx":1208
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProjEntry.frx":155A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProjEntry.frx":187C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProjEntry.frx":1BCE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvHistPane 
         Height          =   6240
         Left            =   150
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   900
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   11007
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlLView"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Subject"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblHistBanner 
         BackColor       =   &H00B0C0D6&
         Caption         =   "History -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   83
         Top             =   150
         Width           =   11115
      End
      Begin VB.Label lblHistView 
         BackStyle       =   0  'Transparent
         Caption         =   "View :"
         Height          =   315
         Left            =   975
         TabIndex        =   82
         Top             =   150
         Width           =   465
      End
      Begin VB.Label lblHistHdr1 
         BackColor       =   &H00DEE6F0&
         Caption         =   " Date:"
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
         Left            =   150
         TabIndex        =   81
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label lblHistHdr2 
         BackColor       =   &H00DEE6F0&
         Caption         =   " Type:"
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
         Left            =   2100
         TabIndex        =   80
         Top             =   600
         Width           =   2340
      End
      Begin VB.Label lblHistHdr3 
         BackColor       =   &H00DEE6F0&
         Caption         =   " Subject / File:"
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
         Left            =   4425
         TabIndex        =   79
         Top             =   600
         Width           =   6840
      End
      Begin VB.Shape shpHistPane 
         BorderColor     =   &H00CCCCB4&
         Height          =   1365
         Left            =   75
         Top             =   2250
         Width           =   465
      End
   End
   Begin VB.PictureBox picPane2 
      BorderStyle     =   0  'None
      Height          =   2115
      Left            =   7800
      ScaleHeight     =   2115
      ScaleWidth      =   1890
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   1890
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         Height          =   6165
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   600
         Width           =   7890
      End
      Begin VB.PictureBox picTimeSaver 
         BorderStyle     =   0  'None
         Height          =   3540
         Left            =   8250
         ScaleHeight     =   3540
         ScaleWidth      =   2265
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2265
         Begin VB.Image Image2 
            Height          =   240
            Left            =   375
            Picture         =   "frmProjEntry.frx":1F60
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Time - Saver"
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
            Left            =   675
            TabIndex        =   68
            Top             =   0
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "1. Highlight"
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
            Left            =   0
            TabIndex        =   67
            Top             =   375
            Width           =   990
         End
         Begin VB.Label Label2 
            Caption         =   "part of the"
            Height          =   240
            Index           =   2
            Left            =   975
            TabIndex        =   66
            Top             =   375
            Width           =   915
         End
         Begin VB.Label Label2 
            Caption         =   " text with your mouse."
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   65
            Top             =   600
            Width           =   1740
         End
         Begin VB.Label Label2 
            Caption         =   "2. Click"
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
            Left            =   0
            TabIndex        =   64
            Top             =   825
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "on an action below to"
            Height          =   240
            Index           =   5
            Left            =   600
            TabIndex        =   63
            Top             =   825
            Width           =   1590
         End
         Begin VB.Label Label2 
            Caption         =   " transform the selection."
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   62
            Top             =   1050
            Width           =   2040
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   0
            Left            =   150
            Picture         =   "frmProjEntry.frx":2272
            Stretch         =   -1  'True
            Top             =   1425
            Width           =   240
         End
         Begin VB.Label lblHyper 
            Caption         =   "Save it as a To Do"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   0
            Left            =   450
            MouseIcon       =   "frmProjEntry.frx":25B4
            MousePointer    =   99  'Custom
            TabIndex        =   61
            Top             =   1425
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   1
            Left            =   150
            Picture         =   "frmProjEntry.frx":28BE
            Stretch         =   -1  'True
            Top             =   1725
            Width           =   240
         End
         Begin VB.Label lblHyper 
            Caption         =   "Save it as an E-mail"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   450
            MouseIcon       =   "frmProjEntry.frx":2BD0
            MousePointer    =   99  'Custom
            TabIndex        =   60
            Top             =   1725
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   2
            Left            =   150
            Picture         =   "frmProjEntry.frx":2EDA
            Stretch         =   -1  'True
            Top             =   2025
            Width           =   240
         End
         Begin VB.Label lblHyper 
            Caption         =   "Save it as a Letter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   2
            Left            =   450
            MouseIcon       =   "frmProjEntry.frx":325C
            MousePointer    =   99  'Custom
            TabIndex        =   59
            Top             =   2025
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   3
            Left            =   150
            Picture         =   "frmProjEntry.frx":3566
            Stretch         =   -1  'True
            Top             =   2325
            Width           =   240
         End
         Begin VB.Label lblHyper 
            Caption         =   "Save it as an Appt"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   3
            Left            =   450
            MouseIcon       =   "frmProjEntry.frx":38A8
            MousePointer    =   99  'Custom
            TabIndex        =   58
            Top             =   2325
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   4
            Left            =   150
            Picture         =   "frmProjEntry.frx":3BB2
            Stretch         =   -1  'True
            Top             =   2625
            Width           =   240
         End
         Begin VB.Label lblHyper 
            Caption         =   "Save it as a Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   4
            Left            =   450
            MouseIcon       =   "frmProjEntry.frx":3EF4
            MousePointer    =   99  'Custom
            TabIndex        =   57
            Top             =   2625
            Width           =   1440
         End
         Begin VB.Label Label2 
            Caption         =   "3. Click"
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
            Index           =   7
            Left            =   0
            TabIndex        =   56
            Top             =   3000
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   " OK to file the"
            Height          =   240
            Index           =   8
            Left            =   600
            TabIndex        =   55
            Top             =   3000
            Width           =   1590
         End
         Begin VB.Label Label2 
            Caption         =   " highlighted text in History."
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   54
            Top             =   3225
            Width           =   2040
         End
      End
      Begin VB.Label lblCommentBanner 
         BackColor       =   &H00B0C0D6&
         Caption         =   "Free-Form Project Comments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   74
         Top             =   75
         Width           =   11115
      End
      Begin VB.Label lblMsg1 
         Caption         =   "Type"
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
         Left            =   8850
         TabIndex        =   73
         Top             =   600
         Width           =   465
      End
      Begin VB.Label lblMsg2 
         Caption         =   " a comment of any length"
         Height          =   240
         Left            =   9300
         TabIndex        =   72
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lblMsg3 
         Caption         =   "in the box. It will be saved"
         Height          =   240
         Left            =   8850
         TabIndex        =   71
         Top             =   825
         Width           =   2340
      End
      Begin VB.Label lblMsg4 
         Caption         =   "automatically."
         Height          =   240
         Left            =   8850
         TabIndex        =   70
         Top             =   1050
         Width           =   2340
      End
   End
   Begin MSComCtl2.MonthView mnvEnd 
      Height          =   2310
      Left            =   3675
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1865
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      MonthBackColor  =   15593715
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   53608449
      TitleBackColor  =   11583702
      CurrentDate     =   38258
   End
   Begin MSComCtl2.MonthView mnvStart 
      Height          =   2310
      Left            =   3675
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   1605
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   15593715
      BorderStyle     =   1
      Appearance      =   0
      MonthBackColor  =   15593715
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   53608449
      TitleBackColor  =   11583702
      CurrentDate     =   38258
   End
   Begin VB.CommandButton cmdRelCon 
      BackColor       =   &H00B0C0D6&
      Caption         =   "Add Related Contact ..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4200
      Width           =   1665
   End
   Begin VB.CommandButton cmdUserFld 
      BackColor       =   &H00B0C0D6&
      Caption         =   "Add User Field ..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3975
      Width           =   1290
   End
   Begin VB.CommandButton cmdAppts 
      BackColor       =   &H00B0C0D6&
      Caption         =   "New Appt ..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2775
      Width           =   1065
   End
   Begin VB.OptionButton optNType 
      BackColor       =   &H00B0C0D6&
      Caption         =   "Note"
      Height          =   240
      Index           =   0
      Left            =   8175
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   600
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.OptionButton optNType 
      BackColor       =   &H00B0C0D6&
      Caption         =   "Call"
      Height          =   240
      Index           =   1
      Left            =   9150
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   600
      Width           =   915
   End
   Begin VB.TextBox txtNewNote 
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   7650
      MultiLine       =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   825
      Width           =   3690
   End
   Begin VB.CommandButton cmdHist 
      BackColor       =   &H00B0C0D6&
      Caption         =   "Add To History"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10125
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   600
      Width           =   1140
   End
   Begin MSComctlLib.ListView lvRelCon 
      Height          =   1590
      Left            =   75
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4650
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2805
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Phone"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00DEE6F0&
      Height          =   390
      Left            =   10950
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   7350
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picProfile 
      BackColor       =   &H00B0C0D6&
      BorderStyle     =   0  'None
      Height          =   3420
      Left            =   75
      ScaleHeight     =   3420
      ScaleWidth      =   3600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   3605
      Begin VB.ListBox lstPrjType 
         Appearance      =   0  'Flat
         ForeColor       =   &H00696969&
         Height          =   1785
         Left            =   1350
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   975
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ListBox lstStatus 
         Appearance      =   0  'Flat
         ForeColor       =   &H00696969&
         Height          =   1980
         Left            =   1350
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   750
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   4
         Left            =   3340
         Picture         =   "frmProjEntry.frx":41FE
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Modify the setting attribute of this record"
         Top             =   765
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   3
         Left            =   3340
         Picture         =   "frmProjEntry.frx":4470
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Modify the End Date of this Project"
         Top             =   1530
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   2
         Left            =   3340
         Picture         =   "frmProjEntry.frx":46E2
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Modify the Start Date of this Project"
         Top             =   1275
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   1
         Left            =   3340
         Picture         =   "frmProjEntry.frx":4954
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Modify the Type setting for this Project"
         Top             =   1020
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   0
         Left            =   3340
         Picture         =   "frmProjEntry.frx":4BC6
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Select a Status setting for this Project"
         Top             =   510
         Width           =   220
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   0
         Top             =   255
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   510
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1020
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1275
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   5
         Top             =   1530
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   1390
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1785
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   2
         Top             =   765
         Width           =   2190
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EDF0F3&
         Caption         =   "Setting"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   6
         Left            =   15
         TabIndex        =   84
         Top             =   765
         Width           =   1365
      End
      Begin VB.Shape shpFiller 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   1370
         Left            =   15
         Top             =   2040
         Width           =   3570
      End
      Begin VB.Label lblProfileHdr 
         BackStyle       =   0  'Transparent
         Caption         =   " Project Profile"
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
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EDF0F3&
         Caption         =   "Name"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   22
         Top             =   255
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EDF0F3&
         Caption         =   "Status"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   21
         Top             =   510
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EDF0F3&
         Caption         =   "Project Type"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   20
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EDF0F3&
         Caption         =   "Start Date"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   19
         Top             =   1275
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EDF0F3&
         Caption         =   "End Date"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   18
         Top             =   1530
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EDF0F3&
         Caption         =   "Budget"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   5
         Left            =   15
         TabIndex        =   17
         Top             =   1785
         Width           =   1365
      End
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H002A59A0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11355
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   11355
      Begin VB.ComboBox cboProjList 
         BackColor       =   &H002A59A0&
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
         Left            =   2100
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   50
         Width           =   4740
      End
      Begin MSComctlLib.TabStrip tbsMain 
         Height          =   315
         Left            =   8850
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   150
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Info"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Comments"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "History"
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
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmProjEntry.frx":4E38
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   " Project Record :"
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
         TabIndex        =   11
         Top             =   75
         Width           =   1590
      End
   End
   Begin MSComctlLib.ListView lvHist 
      Height          =   1815
      Left            =   3825
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   825
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlHist"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvAppts 
      Height          =   840
      Left            =   7650
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   1482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvToDo 
      Height          =   645
      Left            =   3825
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3225
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   1138
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Due"
         Object.Width           =   2011
      EndProperty
   End
   Begin MSComctlLib.ListView lvUserDef 
      Height          =   1215
      Left            =   3825
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   4425
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Field Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Field Value"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlHist 
      Left            =   75
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjEntry.frx":B0C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjEntry.frx":B5DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblUserDef 
      BackColor       =   &H00B0C0D6&
      Caption         =   " User Defined Project Fields"
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
      Left            =   3825
      TabIndex        =   46
      Top             =   3975
      Width           =   7515
   End
   Begin VB.Label lblUserDefHdr1 
      BackColor       =   &H00DEE6F0&
      Caption         =   " Field Name:"
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
      Left            =   3825
      TabIndex        =   45
      Top             =   4200
      Width           =   3765
   End
   Begin VB.Label lblUserDefHdr2 
      BackColor       =   &H00DEE6F0&
      Caption         =   " Field Value:"
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
      Left            =   7575
      TabIndex        =   44
      Top             =   4200
      Width           =   3765
   End
   Begin VB.Shape shpUserDef 
      BorderColor     =   &H00B0C0D6&
      Height          =   690
      Left            =   4800
      Top             =   5100
      Width           =   465
   End
   Begin VB.Label lblToDo 
      BackColor       =   &H00B0C0D6&
      Caption         =   " To Do"
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
      Left            =   3825
      TabIndex        =   41
      Top             =   2775
      Width           =   3690
   End
   Begin VB.Label lblToDoHdr 
      BackColor       =   &H00DEE6F0&
      Caption         =   " Subject:"
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
      Left            =   3825
      TabIndex        =   40
      Top             =   3000
      Width           =   2250
   End
   Begin VB.Label lblToDoHdrDue 
      BackColor       =   &H00DEE6F0&
      Caption         =   "   Due:"
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
      Left            =   6075
      TabIndex        =   39
      Top             =   3000
      Width           =   1425
   End
   Begin VB.Shape shpToDo 
      BorderColor     =   &H00B0C0D6&
      Height          =   690
      Left            =   3750
      Top             =   3075
      Width           =   465
   End
   Begin VB.Label lblAppts 
      BackColor       =   &H00B0C0D6&
      Caption         =   " Appointments"
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
      Left            =   7650
      TabIndex        =   38
      Top             =   2775
      Width           =   3690
   End
   Begin VB.Shape shpAppts 
      BorderColor     =   &H00B0C0D6&
      Height          =   990
      Left            =   7575
      Top             =   2850
      Width           =   465
   End
   Begin VB.Label lblHist 
      BackColor       =   &H00B0C0D6&
      Caption         =   " Project History"
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
      Left            =   3825
      TabIndex        =   34
      Top             =   600
      Width           =   3690
   End
   Begin VB.Shape shpHist 
      BorderColor     =   &H00B0C0D6&
      Height          =   1140
      Left            =   3750
      Top             =   1575
      Width           =   615
   End
   Begin VB.Label lblNewHist 
      BackColor       =   &H00B0C0D6&
      Caption         =   "New:"
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
      Left            =   7650
      TabIndex        =   33
      Top             =   600
      Width           =   3690
   End
   Begin VB.Shape shpNewHist 
      BorderColor     =   &H00B0C0D6&
      Height          =   1440
      Left            =   7575
      Top             =   1275
      Width           =   765
   End
   Begin VB.Shape shpRelCon 
      BorderColor     =   &H00B0C0D6&
      Height          =   765
      Left            =   1275
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label lblRelConPhoneHdr 
      BackColor       =   &H00DEE6F0&
      Caption         =   " Phone"
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
      Left            =   2100
      TabIndex        =   26
      Top             =   4425
      Width           =   1590
   End
   Begin VB.Label lblRelConNameHdr 
      BackColor       =   &H00DEE6F0&
      Caption         =   " Name"
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
      TabIndex        =   25
      Top             =   4425
      Width           =   2040
   End
   Begin VB.Label lblRelCon 
      BackColor       =   &H00B0C0D6&
      Caption         =   " Related Contacts"
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
      TabIndex        =   24
      Top             =   4200
      Width           =   3615
   End
End
Attribute VB_Name = "frmProjEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsProject As Recordset 'main recordset
Dim rsNote As Recordset 'for Attach(notes) table
Dim rsList As Recordset 'all other data work
Dim rsComment As Recordset 'for contact comments entry

Dim m_strOnEnter As String 'for text on entry into textbox
Dim m_strOnLeave As String 'for text when leaving textbox
Dim m_strNoteType As String 'for note type "C" = Call, "N" = Note

Dim m_lngToDo As Long 'for selected to do item
Dim m_lngNotes As Long 'for selected notes item
Dim m_lngAppts As Long 'for selected appts item

'***for comments
Dim m_blnIsNewComment As Boolean
Dim m_blnChanged As Boolean
Dim m_lngCommID As Long

'***for full history grid
Dim m_blnFullHist As Boolean

Private Sub cboHistFilter_Click()
   'load lvHistPane contents according to cbo item selected
   Const sMOD_NAME As String = "frmContEntry.cboHistFilter_Click"
   On Error GoTo Error_Handler
   
   lvHistPane.ListItems.Clear
   
   Select Case cboHistFilter.Text
      Case "All Items"
         m_blnFullHist = True
         Call LoadNotesCallsHistory
         Call LoadToDoHistory
         Call LoadApptsHistory
      Case "Notes & Calls"
         m_blnFullHist = False
         Call LoadNotesCallsHistory
      Case "Documents"
         m_blnFullHist = False
      Case "To Do's"
         m_blnFullHist = False
         Call LoadToDoHistory
      Case "E-mails"
         m_blnFullHist = False
      Case "Appointments"
         m_blnFullHist = False
         Call LoadApptsHistory
      Case "Invoices"
         m_blnFullHist = False
      Case "Bills"
         m_blnFullHist = False
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the list!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub cboProjList_Click()
   Const sMOD_NAME As String = "frmProjEntry.cboProjList_Click"
   On Error GoTo Error_Handler
   
   Dim lngProjID As Long
   
   lngProjID = cboProjList.ItemData(cboProjList.ListIndex)
   
   frmSwtProject.m_lngProjectID = lngProjID
   Load frmSwtProject
   frmSwtProject.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the Project information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub cmdAddFile_Click()
   MsgBox "Sorry, this feature is not available yet.", , APP_MSG_NAME
End Sub

Private Sub cmdAppts_Click()
   'Add new appointment
   Const sMOD_NAME As String = "frmProjEntry.cmdAppts_Click"
   On Error GoTo Error_Handler
   
   icurState = NOW_ADDING
   Load frmAppt
   frmAppt.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the Appointments dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub cmdHist_Click()
   Const sMOD_NAME As String = "frmProjEntry.cmdHist_Click"
   On Error GoTo Error_Handler
   
   Call PostNewHistoryItem
   Call LoadProjectHistory
   
   txtNewNote.Text = ""
   optNType(0).Value = True
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub cmdRelCon_Click()
   Load frmSetRelContact
   frmSetRelContact.Show vbModeless, frmMain
End Sub

Private Sub cmdUserFld_Click()
   Const sMOD_NAME As String = "frmProjEntry.cmdUserFld_Click"
   On Error GoTo Error_Handler
   
   Load frmUserPrjFields
   frmUserPrjFields.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the User Defined Field dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Activate()
   Const sMOD_NAME As String = "frmProjEntry.Form_Activate"
   On Error GoTo Error_Handler
   
   If (Text1(6).Text = "Default") Then
      cboProjList.Text = Text1(0).Text
   End If
   cboHistFilter.Text = "All Items"
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub Form_Load()
   'add code
   Const sMOD_NAME As String = "frmProjEntry.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Project Information Screen", True
   frmMain.picStatus.BackColor = &H2A59A0
   
   Set rsProject = dbContact.OpenRecordset("Projects", dbOpenTable)
   Set rsNote = dbContact.OpenRecordset("Attach", dbOpenTable)
   Set rsComment = dbContact.OpenRecordset("PComments", dbOpenTable)
   
   'set note type
   m_strNoteType = "N"
   
   'set both month views
   mnvStart.Value = Date
   mnvEnd.Value = Date
   
   'load history pane combo box
   With cboHistFilter
      .AddItem "All Items"
      .AddItem "Notes & Calls"
      .AddItem "Documents"
      .AddItem "To Do's"
      .AddItem "E-mails"
      .AddItem "Appointments"
   End With
   
   'Load all needed data
   Call LoadProjectCombo
   Call LoadMainProjectInfo
   Call LoadProjectHistory
   Call LoadToDoInfo
   Call LoadApptsInfo
   Call LoadRelContactInfo
   Call LoadUserDefInfo
   Call LoadStatusItems
   Call LoadPrjTypeItems
   Call LoadComments
   
   'set screen flag
   g_strFormFlag = "PEnt"
   
   'set gridline preference
   lvAppts.GridLines = g_blnShowLines
   lvHist.GridLines = g_blnShowLines
   lvHistPane.GridLines = g_blnShowLines
   lvRelCon.GridLines = g_blnShowLines
   lvToDo.GridLines = g_blnShowLines
   lvUserDef.GridLines = g_blnShowLines
   
   'enable frmMain menu delete & print options
   frmMain.mnuEditDelete.Enabled = True
   frmMain.mnuFilePrint.Enabled = True
   frmMain.tbrMain.Buttons(7).Enabled = True
   
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
   
   LockWindowUpdate frmProjEntry.hWnd
   
   'adjust related contact items
   lblRelCon.Move picProfile.Left, picProfile.Top + picProfile.Height + 225, picProfile.Width
   lblRelConNameHdr.Move lblRelCon.Left, lblRelCon.Top + 240
   lblRelConPhoneHdr.Move lblRelConNameHdr.Left + lblRelConNameHdr.Width, lblRelConNameHdr.Top, lblRelCon.Width - lblRelConNameHdr.Width '- 15
   lvRelCon.Move lblRelCon.Left, lblRelConNameHdr.Top + 240, lblRelCon.Width, Me.ScaleHeight - picProfile.Height - 1415
   shpRelCon.Move lvRelCon.Left - 15, lvRelCon.Top - 15, lvRelCon.Width + 30, lvRelCon.Height + 30
   cmdRelCon.Move lblRelCon.Left + lblRelCon.Width - 1740, lblRelCon.Top
   'adjust history items
   lblHist.Move picProfile.Left + picProfile.Width + 150, picProfile.Top, (Me.ScaleWidth - picProfile.Width - 450) * 0.5
   lvHist.Move lblHist.Left, lblHist.Top + 240, lblHist.Width, Me.ScaleHeight * 0.2
   '***adjust lvHist header widths
   lvHist.ColumnHeaders(1).Width = 1000
   lvHist.ColumnHeaders(2).Width = lvHist.Width - 1285
   shpHist.Move lvHist.Left - 15, lvHist.Top - 15, lvHist.Width + 30, lvHist.Height + 30
   'adjust new hist items
   lblNewHist.Move lblHist.Left + lblHist.Width + 150, lblHist.Top, (Me.ScaleWidth - picProfile.Width - 450) * 0.5
   cmdHist.Move lblNewHist.Left + lblNewHist.Width - 1215, lblNewHist.Top
   optNType(0).Move lblNewHist.Left + 525, lblNewHist.Top
   optNType(1).Move lblNewHist.Left + 1500, lblNewHist.Top
   txtNewNote.Move lblNewHist.Left, lblNewHist.Top + 240, lblNewHist.Width, lvHist.Height
   shpNewHist.Move txtNewNote.Left - 15, txtNewNote.Top - 15, txtNewNote.Width + 30, txtNewNote.Height + 30
   'adjust to do items
   lblToDo.Move lblHist.Left, lvHist.Top + lvHist.Height + 150, lblHist.Width
   lblToDoHdrDue.Move lblToDo.Left + lblToDo.Width - 1440, lblToDo.Top + 240, 1435 '1425
   lblToDoHdr.Move lblToDo.Left, lblToDo.Top + 240, lblToDo.Width - 1425
   lvToDo.Move lblToDo.Left, lblToDoHdr.Top + 240, lblToDo.Width
   '***adjust lvToDo header widths
   lvToDo.ColumnHeaders(1).Width = lvToDo.Width - 1425
   shpToDo.Move lvToDo.Left - 15, lvToDo.Top - 15, lvToDo.Width + 30, lvToDo.Height + 30
   'adjust appts items
   lblAppts.Move lblNewHist.Left, lblToDo.Top, lblNewHist.Width
   lvAppts.Move lblAppts.Left, lblAppts.Top + 240, lblAppts.Width, lvToDo.Height + 240
   '***adjust lvAppts header widths
   lvAppts.ColumnHeaders(2).Width = lvAppts.Width - 1225
   shpAppts.Move lvAppts.Left - 15, lvAppts.Top - 15, lvAppts.Width + 30, lvAppts.Height + 30
   cmdAppts.Move lblAppts.Left + lblAppts.Width - 1140, lblAppts.Top
   'adjust user defined items
   lblUserDef.Move lblToDo.Left, lvAppts.Top + lvAppts.Height + 150, Me.ScaleWidth - picProfile.Width - 300
   cmdUserFld.Move lblUserDef.Left + lblUserDef.Width - 1365, lblUserDef.Top
   lblUserDefHdr1.Move lblUserDef.Left, lblUserDef.Top + 240, lblUserDef.Width * 0.5
   lblUserDefHdr2.Move lblUserDef.Left + lblUserDefHdr1.Width, lblUserDefHdr1.Top, lblUserDef.Width * 0.5
   lvUserDef.Move lblUserDef.Left, lblUserDefHdr1.Top + 240, lblUserDef.Width, Me.ScaleHeight * 0.33
   '***adjust lvUserDef header widths
   lvUserDef.ColumnHeaders(1).Width = lvUserDef.Width * 0.5
   lvUserDef.ColumnHeaders(2).Width = lvUserDef.Width * 0.5 - 265
   shpUserDef.Move lvUserDef.Left - 15, lvUserDef.Top - 15, lvUserDef.Width + 30, lvUserDef.Height + 30
   'adjust lstStatus item
   lstStatus.Move Text1(1).Left, Text1(1).Top + 255, Text1(1).Width
   'adjust lstPrjType item
   lstPrjType.Move Text1(2).Left, Text1(2).Top + 255, Text1(2).Width
   'adjust start date month view
   mnvStart.Left = picProfile.Left + picProfile.Width
   'adjust end date month view
   mnvEnd.Left = picProfile.Left + picProfile.Width
   'adjust all pane 2 items*************************************************
   picPane2.Move 0, 465, Me.ScaleWidth, Me.ScaleHeight - 465
   lblCommentBanner.Move 225, picPane2.ScaleTop + 225, picPane2.ScaleWidth - 450
   txtComments.Move 225, lblCommentBanner.Top + lblCommentBanner.Height + 225, picPane2.ScaleWidth * 0.75, picPane2.ScaleHeight - 1350
   lblMsg1.Move txtComments.Left + txtComments.Width + 225, txtComments.Top
   lblMsg2.Move lblMsg1.Left + lblMsg1.Width, lblMsg1.Top
   lblMsg3.Move lblMsg1.Left, lblMsg1.Top + 240
   lblMsg4.Move lblMsg1.Left, lblMsg3.Top + 240
   picTimeSaver.Move lblMsg1.Left
   'adjust all pane 3 items*************************************************
   picPane3.Move 0, 465, Me.ScaleWidth, Me.ScaleHeight - 465
   lblHistBanner.Move 225, picPane3.ScaleTop + 225, picPane3.ScaleWidth - 450
   lblHistView.Move lblHistBanner.Left + 825, lblHistBanner.Top
   cboHistFilter.Move lblHistBanner.Left + 1350, lblHistBanner.Top
   cmdAddFile.Move lblHistBanner.Left + lblHistBanner.Width - 1140, lblHistBanner.Top + 37
   lblHistHdr1.Move lblHistBanner.Left, lblHistBanner.Top + 540, lblHistBanner.Width * 0.12
   lblHistHdr2.Move lblHistHdr1.Left + lblHistHdr1.Width, lblHistHdr1.Top, lblHistHdr1.Width
   lblHistHdr3.Move lblHistHdr2.Left + lblHistHdr2.Width, lblHistHdr2.Top, (lblHistBanner.Width * 0.76) + 15
   lvHistPane.Move lblHistBanner.Left, lblHistHdr1.Top + 240, lblHistBanner.Width, picPane3.ScaleHeight - 1155
   '***adjust lvHistPane column widths
   lvHistPane.ColumnHeaders(1).Width = lvHistPane.Width * 0.12
   lvHistPane.ColumnHeaders(2).Width = lvHistPane.Width * 0.12
   lvHistPane.ColumnHeaders(3).Width = (lvHistPane.Width * 0.76) - 265
   shpHistPane.Move lvHistPane.Left - 15, lvHistPane.Top - 15, lvHistPane.Width + 30, lvHistPane.Height + 30
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'save any entered notes not already saved
   If (txtNewNote.Text <> "") Then
      Call PostNewHistoryItem
   End If
   'remove data & form reference
   rsProject.Close
   Set rsProject = Nothing
   rsNote.Close
   Set rsNote = Nothing
   rsComment.Close
   Set rsComment = Nothing
   
   'disable frmMain menu delete & print options
   frmMain.mnuEditDelete.Enabled = False
   frmMain.mnuFilePrint.Enabled = False
   frmMain.tbrMain.Buttons(7).Enabled = False
   
   Set frmProjEntry = Nothing
End Sub

Private Sub lblHyper_Click(Index As Integer)
   Const sMOD_NAME As String = "frmProjEntry.lblHyper_Click"
   On Error GoTo Error_Handler
   
   Dim strSelect As String
   
   strSelect = txtComments.SelText
   
   If (strSelect = "") Then
      MsgBox "There is no selected text, from Project Comments.", , APP_MSG_NAME
      Exit Sub
   End If
   
   Select Case Index
      Case 0 'to do
         icurState = NOW_ADDING
         Load frmToDo
         frmToDo.Show vbModeless, frmMain
         frmToDo.Text1(1).Text = strSelect
      Case 1 'e-mail
         MsgBox "Sorry, this feature is not available yet.", , APP_MSG_NAME
      Case 2 'letter
         MsgBox "Sorry, this feature is not available yet.", , APP_MSG_NAME
      Case 3 'appt
         icurState = NOW_ADDING
         Load frmAppt
         frmAppt.Show vbModeless, frmMain
         frmAppt.Text1(1).Text = strSelect
      Case 4 'note
         icurState = NOW_ADDING
         Load frmNotes
         frmNotes.Show vbModeless, frmMain
         frmNotes.Text1.Text = strSelect
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub lstPrjType_Click()
   Const sMOD_NAME As String = "frmProjEntry.lstPrjType_Click"
   On Error GoTo Error_Handler
   
   If (lstPrjType.Text = "<Add Field...>") Then
      lstPrjType.Visible = False
      Text1(2).SetFocus
      Load frmSetProjType
      frmSetProjType.Show vbModeless, frmMain
      Exit Sub
   Else
      Text1(2).Text = lstPrjType.Text
      lstPrjType.Visible = False
      Text1(2).SetFocus
      
      Call EditProjectType
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   lstPrjType.Visible = False
End Sub

Private Sub lstStatus_Click()
   Const sMOD_NAME As String = "frmProjEntry.lstStatus_Click"
   On Error GoTo Error_Handler
   
   If (lstStatus.Text = "<Add Field...>") Then
      lstStatus.Visible = False
      Text1(1).SetFocus
      Load frmSetProjStatus
      frmSetProjStatus.Show vbModeless, frmMain
      Exit Sub
   Else
      Text1(1).Text = lstStatus.Text
      lstStatus.Visible = False
      Text1(1).SetFocus
      
      Call EditProjectStatus
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   lstStatus.Visible = False
End Sub

Private Sub lvAppts_Click()
   Const sMOD_NAME As String = "frmProjEntry.lvAppts_Click"
   On Error GoTo Error_Handler
   
   m_lngAppts = CLng(Mid$(lvAppts.SelectedItem.Key, 3, Len(lvAppts.SelectedItem.Key)))
   
   'code to open Appts entry screen
   icurState = NOW_EDITING
   frmAppt.m_lngApptID = m_lngAppts
   Load frmAppt
   frmAppt.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while loading the Appointments dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvHist_Click()
   Const sMOD_NAME As String = "frmProjEntry.lvHist_Click"
   On Error GoTo Error_Handler
   
   m_lngNotes = CLng(Mid$(lvHist.SelectedItem.Key, 3, Len(lvHist.SelectedItem.Key)))
   
   'code to open Notes entry screen
   icurState = NOW_EDITING
   frmNotes.m_lngNoteID = m_lngNotes
   Load frmNotes
   frmNotes.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred loading the Notes/Calls dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvRelCon_Click()
   Const sMOD_NAME As String = "frmProjEntry.lvRelCon_Click"
   On Error GoTo Error_Handler
   
   Dim lngContID As Long
   
   lngContID = CLng(Mid$(lvRelCon.SelectedItem.Key, 3, Len(lvRelCon.SelectedItem.Key)))
   
   g_lngContID = lngContID
   UnloadAllForms
   Load frmContEntry
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while loading the Related Contact information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvToDo_Click()
   Const sMOD_NAME As String = "frmProjEntry.lvToDo_Click"
   On Error GoTo Error_Handler
   
   m_lngToDo = CLng(Mid$(lvToDo.SelectedItem.Key, 3, Len(lvToDo.SelectedItem.Key)))
   
   'code to open ToDo entry screen
   icurState = NOW_EDITING
   frmToDo.m_lngToDoID = m_lngToDo
   Load frmToDo
   frmToDo.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while loading the To Do dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvUserDef_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Const sMOD_NAME As String = "frmProjEntry.lvUserDef_MouseUp"
   On Error GoTo Error_Handler
   Dim lngUserDefID As Long
   
   If Button = vbRightButton Then
      If lvUserDef.ListItems.Count = 0 Then Exit Sub
      lngUserDefID = CLng(Mid$(lvUserDef.SelectedItem.Key, 3, Len(lvUserDef.SelectedItem.Key)))
      
      'code to Delete user defined field
      frmMain.m_lngProjUserFld = lngUserDefID
      frmMain.m_strProjUserFld = lvUserDef.SelectedItem
      PopupMenu frmMain.mnuDelProjFld
   End If
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub mnvEnd_DateClick(ByVal DateClicked As Date)
   Const sMOD_NAME As String = "frmProjEntry.mnvEnd_DateClick"
   On Error GoTo Error_Handler
   
   Text1(4).Text = mnvEnd.Value
   mnvEnd.Visible = False
   Text1(4).SetFocus
   
   Call EditEndDate
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   mnvEnd.Visible = False
End Sub

Private Sub mnvStart_DateClick(ByVal DateClicked As Date)
   Const sMOD_NAME As String = "frmProjEntry.mnvStart_DateClick"
   On Error GoTo Error_Handler
   
   Text1(3).Text = mnvStart.Value
   mnvStart.Visible = False
   Text1(3).SetFocus
   
   Call EditStartDate
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   mnvStart.Visible = False
End Sub

Private Sub optNType_Click(Index As Integer)
   Select Case Index
      Case 0 'note
         m_strNoteType = "N"
      Case 1 'call
         m_strNoteType = "C"
   End Select
End Sub

Private Sub picBanner_Resize()
   tbsMain.Left = picBanner.ScaleWidth - tbsMain.Width
End Sub

Public Sub LoadProjectCombo()
   'load all contact ShownNames for selection into cboProjList
   Const sMOD_NAME As String = "frmProjEntry.LoadProjectCombo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strSetting As String
   
   strSetting = "Default"
   
   SQL = "SELECT ProjID, PName, Setting FROM Projects "
   SQL = SQL & "WHERE Setting = '" & strSetting & "' "
   SQL = SQL & "ORDER BY PName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   cboProjList.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ProjID)) Then cboProjList.AddItem !PName
            cboProjList.ItemData(cboProjList.NewIndex) = !ProjID
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadMainProjectInfo()
   'load the main contact fields from Contacts table
   Const sMOD_NAME As String = "frmProjEntry.LoadMainProjectInfo"
   On Error GoTo Error_Handler
   
   With rsProject
      If (.RecordCount > 0) Then
         .MoveFirst
         .Index = "PrimaryKey"
         .Seek "=", g_lngProjID
      End If
      
      Dim Indx As Integer
      For Indx = 0 To 6
         Text1(Indx).Text = ""
      Next
      
      If (Not IsNull(!PName)) Then Text1(0) = !PName
      If (Not IsNull(!Status)) Then Text1(1) = !Status
      If (Not IsNull(!Setting)) Then Text1(6) = !Setting
      If (Not IsNull(!PrjType)) Then Text1(2) = !PrjType
      If (Not IsNull(!StartDate)) Then Text1(3) = Format(!StartDate, "mm/dd/yyyy")
      If (Not IsNull(!EndDate)) Then Text1(4) = Format(!EndDate, "mm/dd/yyyy")
      If (Not IsNull(!Budget)) Then Text1(5) = !Budget
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadProjectHistory()
   'load all history items for this contact
   Const sMOD_NAME As String = "frmProjEntry.LoadProjectHistory"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, fkProjID, NType, TextBody, DateStamp "
   SQL = SQL & "FROM Attach "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " ORDER BY DateStamp DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvHist.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               If !NType = "C" Then
                  If (Not IsNull(!DateStamp)) Then
                     Set Item = lvHist.ListItems.Add(, "ID" & !RefNum, Format(!DateStamp, "mm/dd"), , 2)
                     Item.SubItems(1) = Replace(!TextBody, vbCrLf, "") 'so we don't the Cr & Lf representaion
                  End If
               ElseIf !NType = "N" Then
                  If (Not IsNull(!DateStamp)) Then
                     Set Item = lvHist.ListItems.Add(, "ID" & !RefNum, Format(!DateStamp, "mm/dd"), , 1)
                     Item.SubItems(1) = Replace(!TextBody, vbCrLf, "") 'so we don't the Cr & Lf representaion
                  End If
               End If
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvHist, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadToDoInfo()
   'load all associated to do information
   Const sMOD_NAME As String = "frmProjEntry.LoadToDoInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, Subject, fkProjID, DueDate, Completed "
   SQL = SQL & "FROM ToDo "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " AND Completed = False "
   SQL = SQL & " ORDER BY DueDate Desc"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvToDo.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               If (Not IsNull(!Subject)) Then
                  Set Item = lvToDo.ListItems.Add(, "ID" & !RefNum, !Subject)
                  If (Not IsNull(!DueDate)) Then Item.SubItems(1) = !DueDate
               End If
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvToDo, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadApptsInfo()
   'load all associated appointment information
   Const sMOD_NAME As String = "frmProjEntry.LoadApptsInfo"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim vDate As Variant
   Dim sDate As String
   
   vDate = Format(Date, "mm/dd/yyyy")
   vDate = "#" & vDate & "#"
   
   SQL = "SELECT RefNum, fkProjID, Subject, DateFrom FROM Appts "
   SQL = SQL & "WHERE DateFrom >= " & vDate
   SQL = SQL & " AND fkProjID = " & g_lngProjID
   SQL = SQL & " ORDER BY DateFrom DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvAppts.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               Set Item = lvAppts.ListItems.Add(, "ID" & !RefNum, Format(!DateFrom, "ddd mm/dd"))
               Item.SubItems(1) = !Subject
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvAppts, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub PostNewHistoryItem()
   'save any new note / call item
   Const sMOD_NAME As String = "frmProjEntry.PostNewHistoryItem"
   On Error GoTo Error_Handler
   
   If (txtNewNote.Text = "") Then Exit Sub
   
   rsNote.AddNew
   
   With rsNote
      If (Len(txtNewNote)) Then !TextBody = txtNewNote
      !fkProjID = g_lngProjID
      !NType = m_strNoteType
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while posting this History item!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Public Sub LoadUserDefInfo()
   'load all user defined fields for this contact
   Const sMOD_NAME As String = "frmProjEntry.LoadUserDefInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, fkProjID, fkUserFld, Value FROM PUFldValues "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " ORDER BY fkUserFld"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvUserDef.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkUserFld)) Then
               Set Item = lvUserDef.ListItems.Add(, "ID" & !RefNum, !fkUserFld)
               If (Not IsNull(!Value)) Then Item.SubItems(1) = !Value
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvUserDef, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadStatusItems()
   'load all project status items into lstStatus from Lookup table
   Const sMOD_NAME As String = "frmProjEntry.LoadStatusItems"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "PSTAT"
   
   SQL = "SELECT ItemID, Description FROM Lookup "
   SQL = SQL & "WHERE ItemID = '" & strType & "' "
   SQL = SQL & "ORDER BY Description"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lstStatus.Clear
   lstStatus.AddItem "<Add Field...>"
   lstStatus.AddItem " "
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then lstStatus.AddItem !Description
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadPrjTypeItems()
   'load all project types into lstPrjType from Lookup table
   Const sMOD_NAME As String = "frmProjEntry.LoadPrjTypeItems"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "PTYPE"
   
   SQL = "SELECT ItemID, Description FROM Lookup "
   SQL = SQL & "WHERE ItemID = '" & strType & "' "
   SQL = SQL & "ORDER BY Description"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lstPrjType.Clear
   lstPrjType.AddItem "<Add Field...>"
   lstPrjType.AddItem " "
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then lstPrjType.AddItem !Description
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub picPop_Click(Index As Integer)
   Select Case Index
      Case 0 'status
         Text1(1).SetFocus
         
         If (lstStatus.Visible = True) Then
            lstStatus.Visible = False
         Else
            lstStatus.Visible = True
         End If
      Case 1 'project type
         Text1(2).SetFocus
         
         If (lstPrjType.Visible = True) Then
            lstPrjType.Visible = False
         Else
            lstPrjType.Visible = True
         End If
      Case 2 'start date
         Text1(3).SetFocus
         
         If (mnvStart.Visible = True) Then
            mnvStart.Visible = False
         Else
            mnvStart.Visible = True
         End If
      Case 3 'end date
         Text1(4).SetFocus
         
         If (mnvEnd.Visible = True) Then
            mnvEnd.Visible = False
         Else
            mnvEnd.Visible = True
         End If
      Case 4 'setting
         Text1(6).SetFocus
         Load frmSetting
         frmSetting.Show vbModeless, frmMain
   End Select
End Sub

Private Sub picPop_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      picPop(Index).BackColor = &H808080
   End If
End Sub

Private Sub picPop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      picPop(Index).BackColor = vbWhite
   End If
End Sub

Private Sub tbsMain_Click()
   Const sMOD_NAME As String = "frmProjEntry.tbsMain_Click"
   On Error GoTo Error_Handler
   
   picPane2.Visible = False
   picPane3.Visible = False
   
   Select Case tbsMain.SelectedItem.Index
      Case 1 'info
         picPane2.Visible = False
         picPane3.Visible = False
         Text1(0).SetFocus
      Case 2 'comments
         picPane2.Visible = True
         picPane3.Visible = False
         txtComments.SetFocus
      Case 3 'history
         picPane2.Visible = False
         picPane3.Visible = True
   End Select
   
   'save any added comments if needed
   If (m_blnChanged = True) Then
      If (m_blnIsNewComment = True) Then
         icurState = NOW_ADDING
         Call PostCommentEntry
      Else
         icurState = NOW_EDITING
         Call PostCommentEntry
      End If
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   highLight
   m_strOnEnter = Text1(Index).Text
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Const sMOD_NAME As String = "frmProjEntry.Text1_LostFocus"
   On Error GoTo Error_Handler
   
   Dim intStrLen As Integer
   
   m_strOnLeave = Text1(Index).Text
   intStrLen = Len(m_strOnEnter)
   
   If m_strOnEnter <> m_strOnLeave Then
      Select Case Index
         Case 1 'Status
            If (m_strOnLeave = "") Then Exit Sub
            Call EditProjectStatus
         Case 2 'Project Type
            If (m_strOnLeave = "") Then Exit Sub
            Call EditProjectType
         Case 3 'Start Date
            If (m_strOnLeave = "") Then Exit Sub
            Call EditStartDate
         Case 4 'End Date
            If (m_strOnLeave = "") Then Exit Sub
            Call EditEndDate
         Case 5 'Budget
            If (m_strOnLeave = "") Then Exit Sub
            Call EditBudget
         Case 6 'Setting
            If (m_strOnLeave = "") Then Exit Sub
            Call EditSetting
      End Select
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub EditProjectStatus()
   'Edit the Project Status
   Const sMOD_NAME As String = "frmProjEntry.EditProjectStatus"
   On Error GoTo Error_Handler
   
   rsProject.Edit
   
   With rsProject
      !Status = Text1(1).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the Project Status record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditProjectType()
   'Edit the Project Type
   Const sMOD_NAME As String = "frmProjEntry.EditProjectType"
   On Error GoTo Error_Handler
   
   rsProject.Edit
   
   With rsProject
      !PrjType = Text1(2).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the Project Type record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditStartDate()
   'Edit the Start Date
   Const sMOD_NAME As String = "frmProjEntry.EditStartDate"
   On Error GoTo Error_Handler
   
   rsProject.Edit
   
   With rsProject
      !StartDate = Text1(3).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the Project Start Date record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditEndDate()
   'Edit the End Date
   Const sMOD_NAME As String = "frmProjEntry.EditEndDate"
   On Error GoTo Error_Handler
   
   rsProject.Edit
   
   With rsProject
      !EndDate = Text1(4).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the Project End Date record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditBudget()
   'Edit the Budget
   Const sMOD_NAME As String = "frmProjEntry.EditBudget"
   On Error GoTo Error_Handler
   
   rsProject.Edit
   
   With rsProject
      !Budget = Text1(5).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the Project Budget record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub LoadComments()
   'load any listed comments for this contact
   Const sMOD_NAME As String = "frmProjEntry.LoadComments"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT RefNum, fkProjID, Comments FROM PComments "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         m_blnIsNewComment = False
         .MoveFirst
         m_lngCommID = !RefNum
         If (Not IsNull(!Comments)) Then txtComments = !Comments
         m_blnChanged = False
      Else
         m_blnIsNewComment = True
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub txtComments_Change()
   m_blnChanged = True
End Sub

Private Sub PostCommentEntry()
   'post any comment entered into the database
   Const sMOD_NAME As String = "frmProjEntry.PostCommentEntry"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting Contact Comments", True
   
   If (icurState = NOW_ADDING) Then
      rsComment.AddNew
   Else
      With rsComment
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngCommID
            If Not .NoMatch Then
               rsComment.Edit
            Else
               MsgBox "Error Posting Record", , "Posting Error"
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
      End With
   End If
   
   With rsComment
      !fkProjID = g_lngProjID
      If (Len(txtComments)) Then !Comments = txtComments
      
      .Update
   End With
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   MsgBox "An un-known error occurred while Posting Project Comments!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub LoadNotesCallsHistory()
   'load all history items for this contact
   Const sMOD_NAME As String = "frmProjEntry.LoadNotesCallsHistory"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, fkProjID, NType, TextBody, DateStamp "
   SQL = SQL & "FROM Attach "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " ORDER BY DateStamp DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   If (m_blnFullHist <> True) Then
      lvHistPane.ListItems.Clear
   End If
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               If !NType = "C" Then
                  If (Not IsNull(!DateStamp)) Then
                     Set Item = lvHistPane.ListItems.Add(, "CA" & !RefNum, Format(!DateStamp, "mm/dd/yyyy"), , 2)
                     'Set Item = lvHistPane.ListItems.Add(, , Format(!DateStamp, "mm/dd/yyyy"), , 2)
                     Item.SubItems(1) = "Call"
                     Item.SubItems(2) = Replace(!TextBody, vbCrLf, "") 'so we don't the Cr & Lf representaion
                  End If
               ElseIf !NType = "N" Then
                  If (Not IsNull(!DateStamp)) Then
                     Set Item = lvHistPane.ListItems.Add(, "NT" & !RefNum, Format(!DateStamp, "mm/dd/yyyy"), , 1)
                     Item.SubItems(1) = "Note"
                     Item.SubItems(2) = Replace(!TextBody, vbCrLf, "") 'so we don't the Cr & Lf representaion
                  End If
               End If
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvHistPane, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub LoadToDoHistory()
   'load all associated to do information
   Const sMOD_NAME As String = "frmProjEntry.LoadToDoHistory"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   Dim strSubject As String
   Dim strTextBody As String
   
   SQL = "SELECT RefNum, Subject, fkProjID, DueDate, TextBody "
   SQL = SQL & "FROM ToDo "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " ORDER BY DueDate Desc"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   If (m_blnFullHist <> True) Then
      lvHistPane.ListItems.Clear
   End If
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               If (Not IsNull(!DueDate)) Then
                  Set Item = lvHistPane.ListItems.Add(, "TD" & !RefNum, Format(!DueDate, "mm/dd/yyyy"), , 4)
                  'Set Item = lvHistPane.ListItems.Add(, , Format(!DueDate, "mm/dd/yyyy"), , 4)
                  Item.SubItems(1) = "To Do"
                  If (Not IsNull(!Subject)) Then
                     strSubject = !Subject
                  End If
                  If (Not IsNull(!TextBody)) Then
                     strTextBody = Replace(!TextBody, vbCrLf, "") 'so we don't the Cr & Lf representaion
                  End If
                  Item.SubItems(2) = strSubject & " - " & strTextBody
               End If
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvHistPane, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub LoadApptsHistory()
   'load all associated appointment information
   Const sMOD_NAME As String = "frmProjEntry.LoadApptsHistory"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim strSubject As String
   Dim strTextBody As String
   
   SQL = "SELECT RefNum, fkProjID, Subject, DateFrom, TextBody FROM Appts "
   'SQL = SQL & "WHERE DateFrom >= " & vDate
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " ORDER BY DateFrom DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   If (m_blnFullHist <> True) Then
      lvHistPane.ListItems.Clear
   End If
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               Set Item = lvHistPane.ListItems.Add(, "AP" & !RefNum, Format(!DateFrom, "mm/dd/yyyy"), , 6)
               Item.SubItems(1) = "Appointment"
               If (Not IsNull(!Subject)) Then
                  strSubject = !Subject
               End If
               If (Not IsNull(!TextBody)) Then
                  strTextBody = Replace(!TextBody, vbCrLf, "") 'so we don't the Cr & Lf representaion
               End If
               Item.SubItems(2) = strSubject & " - " & strTextBody
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvHistPane, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub EditSetting()
   'edit the project show setting
   Const sMOD_NAME As String = "frmProjEntry.EditSetting"
   On Error GoTo Error_Handler
   
   rsProject.Edit
   
   With rsProject
      !Setting = Text1(6).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the Project Setting record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Public Sub LoadRelContactInfo()
   'load all contacts related to this project
   Const sMOD_NAME As String = "frmProjEntry.LoadRelContactInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strPhone As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, fkProjID, fkContID, ContShowName "
   SQL = SQL & "FROM RelateProject "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " ORDER BY ContShowName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvRelCon.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkContID)) Then
               If (Not IsNull(!ContShowName)) Then
                  Set Item = lvRelCon.ListItems.Add(, "ID" & !fkContID, !ContShowName)
                  strPhone = GetPhoneNum(!fkContID)
                  If (strPhone <> "") Then
                     Item.SubItems(1) = strPhone
                  End If
               End If
            End If
            .MoveNext
         Wend
      Else
         rsList.Close
         Set rsList = Nothing
         Exit Sub
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvRelCon, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub
