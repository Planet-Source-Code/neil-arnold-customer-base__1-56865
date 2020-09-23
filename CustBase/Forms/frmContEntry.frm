VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContEntry 
   Caption         =   "Contact Information"
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
   Icon            =   "frmContEntry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPane3 
      BorderStyle     =   0  'None
      Height          =   2340
      Left            =   6600
      ScaleHeight     =   2340
      ScaleWidth      =   2715
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   2715
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
               Picture         =   "frmContEntry.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContEntry.frx":095C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContEntry.frx":0E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContEntry.frx":1208
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContEntry.frx":155A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContEntry.frx":187C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmContEntry.frx":1BCE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvHistPane 
         Height          =   6240
         Left            =   150
         TabIndex        =   109
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
      Begin VB.CommandButton cmdAddFile 
         BackColor       =   &H00CCCCB4&
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
         TabIndex        =   105
         Top             =   150
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.ComboBox cboHistFilter 
         BackColor       =   &H00CCCCB4&
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   150
         Width           =   1890
      End
      Begin VB.Shape shpHistPane 
         BorderColor     =   &H00CCCCB4&
         Height          =   1365
         Left            =   75
         Top             =   2250
         Width           =   465
      End
      Begin VB.Label lblHistHdr3 
         BackColor       =   &H00ECECE1&
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
         TabIndex        =   108
         Top             =   600
         Width           =   6840
      End
      Begin VB.Label lblHistHdr2 
         BackColor       =   &H00ECECE1&
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
         TabIndex        =   107
         Top             =   600
         Width           =   2340
      End
      Begin VB.Label lblHistHdr1 
         BackColor       =   &H00ECECE1&
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
         TabIndex        =   106
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label lblHistView 
         BackStyle       =   0  'Transparent
         Caption         =   "View :"
         Height          =   315
         Left            =   975
         TabIndex        =   103
         Top             =   150
         Width           =   465
      End
      Begin VB.Label lblHistBanner 
         BackColor       =   &H00CCCCB4&
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
         TabIndex        =   102
         Top             =   150
         Width           =   11115
      End
   End
   Begin VB.PictureBox picPane2 
      BorderStyle     =   0  'None
      Height          =   2190
      Left            =   6600
      ScaleHeight     =   2190
      ScaleWidth      =   2865
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   2865
      Begin VB.PictureBox picTimeSaver 
         BorderStyle     =   0  'None
         Height          =   3540
         Left            =   8250
         ScaleHeight     =   3540
         ScaleWidth      =   2265
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2265
         Begin VB.Label Label2 
            Caption         =   " highlighted text in History."
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   100
            Top             =   3225
            Width           =   2040
         End
         Begin VB.Label Label2 
            Caption         =   " OK to file the"
            Height          =   240
            Index           =   8
            Left            =   600
            TabIndex        =   99
            Top             =   3000
            Width           =   1590
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
            TabIndex        =   98
            Top             =   3000
            Width           =   615
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
            MouseIcon       =   "frmContEntry.frx":1F60
            MousePointer    =   99  'Custom
            TabIndex        =   97
            Top             =   2625
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   4
            Left            =   150
            Picture         =   "frmContEntry.frx":226A
            Stretch         =   -1  'True
            Top             =   2625
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
            MouseIcon       =   "frmContEntry.frx":25AC
            MousePointer    =   99  'Custom
            TabIndex        =   96
            Top             =   2325
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   3
            Left            =   150
            Picture         =   "frmContEntry.frx":28B6
            Stretch         =   -1  'True
            Top             =   2325
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
            MouseIcon       =   "frmContEntry.frx":2BF8
            MousePointer    =   99  'Custom
            TabIndex        =   95
            Top             =   2025
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   2
            Left            =   150
            Picture         =   "frmContEntry.frx":2F02
            Stretch         =   -1  'True
            Top             =   2025
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
            MouseIcon       =   "frmContEntry.frx":3284
            MousePointer    =   99  'Custom
            TabIndex        =   94
            Top             =   1725
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   1
            Left            =   150
            Picture         =   "frmContEntry.frx":358E
            Stretch         =   -1  'True
            Top             =   1725
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
            MouseIcon       =   "frmContEntry.frx":38A0
            MousePointer    =   99  'Custom
            TabIndex        =   93
            Top             =   1425
            Width           =   1440
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   0
            Left            =   150
            Picture         =   "frmContEntry.frx":3BAA
            Stretch         =   -1  'True
            Top             =   1425
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   " transform the selection."
            Height          =   240
            Index           =   6
            Left            =   150
            TabIndex        =   92
            Top             =   1050
            Width           =   2040
         End
         Begin VB.Label Label2 
            Caption         =   "on an action below to"
            Height          =   240
            Index           =   5
            Left            =   600
            TabIndex        =   91
            Top             =   825
            Width           =   1590
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
            TabIndex        =   90
            Top             =   825
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   " text with your mouse."
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   89
            Top             =   600
            Width           =   1740
         End
         Begin VB.Label Label2 
            Caption         =   "part of the"
            Height          =   240
            Index           =   2
            Left            =   975
            TabIndex        =   88
            Top             =   375
            Width           =   915
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
            TabIndex        =   87
            Top             =   375
            Width           =   990
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
            TabIndex        =   86
            Top             =   0
            Width           =   1140
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   375
            Picture         =   "frmContEntry.frx":3EEC
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         Height          =   6165
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   600
         Width           =   7890
      End
      Begin VB.Label lblMsg4 
         Caption         =   "automatically."
         Height          =   240
         Left            =   8850
         TabIndex        =   84
         Top             =   1050
         Width           =   2340
      End
      Begin VB.Label lblMsg3 
         Caption         =   "in the box. It will be saved"
         Height          =   240
         Left            =   8850
         TabIndex        =   83
         Top             =   825
         Width           =   2340
      End
      Begin VB.Label lblMsg2 
         Caption         =   " a comment of any length"
         Height          =   240
         Left            =   9300
         TabIndex        =   82
         Top             =   600
         Width           =   1890
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
         TabIndex        =   81
         Top             =   600
         Width           =   465
      End
      Begin VB.Label lblCommentBanner 
         BackColor       =   &H00CCCCB4&
         Caption         =   "Free-Form Contact Comments"
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
         TabIndex        =   79
         Top             =   75
         Width           =   11115
      End
   End
   Begin VB.CommandButton cmdUserFld 
      BackColor       =   &H00CCCCB4&
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
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   5925
      Width           =   1290
   End
   Begin VB.CommandButton cmdRelCon 
      BackColor       =   &H00CCCCB4&
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   4050
      Width           =   1665
   End
   Begin VB.CommandButton cmdHist 
      BackColor       =   &H00CCCCB4&
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
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   600
      Width           =   1140
   End
   Begin MSComctlLib.ListView lvUserDef 
      Height          =   1215
      Left            =   3825
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6375
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
   Begin MSComctlLib.ListView lvRelCon 
      Height          =   1290
      Left            =   3825
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   4500
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   2275
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Memo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "E-Mail"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdAppts 
      BackColor       =   &H00CCCCB4&
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
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   2775
      Width           =   1065
   End
   Begin MSComctlLib.ListView lvAppts 
      Height          =   840
      Left            =   7650
      TabIndex        =   62
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
      TabIndex        =   60
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
   Begin VB.TextBox txtNewNote 
      BackColor       =   &H00EAFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   7650
      MultiLine       =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   825
      Width           =   3690
   End
   Begin VB.OptionButton optNType 
      BackColor       =   &H00CCCCB4&
      Caption         =   "Call"
      Height          =   240
      Index           =   1
      Left            =   9150
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   600
      Width           =   915
   End
   Begin VB.OptionButton optNType 
      BackColor       =   &H00CCCCB4&
      Caption         =   "Note"
      Height          =   240
      Index           =   0
      Left            =   8175
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   600
      Value           =   -1  'True
      Width           =   915
   End
   Begin MSComctlLib.ListView lvHist 
      Height          =   1815
      Left            =   3825
      TabIndex        =   52
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
   Begin MSComctlLib.ImageList imlHist 
      Left            =   4500
      Top             =   7200
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
            Picture         =   "frmContEntry.frx":41FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContEntry.frx":4718
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvRelPrj 
      Height          =   990
      Left            =   75
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   6750
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1746
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
         Text            =   "Project"
         Object.Width           =   3732
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2098
      EndProperty
   End
   Begin VB.PictureBox picProfile 
      BackColor       =   &H00CCCCB4&
      BorderStyle     =   0  'None
      Height          =   5620
      Left            =   75
      ScaleHeight     =   5625
      ScaleWidth      =   3600
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   3605
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   15
         Left            =   3340
         Picture         =   "frmContEntry.frx":4C32
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   121
         TabStop         =   0   'False
         ToolTipText     =   "Delete this E-Mail Record"
         Top             =   5370
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   14
         Left            =   3340
         Picture         =   "frmContEntry.frx":4EA4
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   120
         TabStop         =   0   'False
         ToolTipText     =   "Delete this E-Mail Record"
         Top             =   5115
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   13
         Left            =   3340
         Picture         =   "frmContEntry.frx":5116
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   119
         TabStop         =   0   'False
         ToolTipText     =   "Delete this E-Mail Record"
         Top             =   4860
         Width           =   220
      End
      Begin VB.ListBox lstJobTitle 
         Appearance      =   0  'Flat
         ForeColor       =   &H00696969&
         Height          =   2565
         Left            =   1425
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1800
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   12
         Left            =   3340
         Picture         =   "frmContEntry.frx":5388
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   118
         TabStop         =   0   'False
         ToolTipText     =   "Delete this Phone Number Record"
         Top             =   4605
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   11
         Left            =   3340
         Picture         =   "frmContEntry.frx":55FA
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   117
         TabStop         =   0   'False
         ToolTipText     =   "Delete this Phone Number Record"
         Top             =   4350
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   10
         Left            =   3340
         Picture         =   "frmContEntry.frx":586C
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   116
         TabStop         =   0   'False
         ToolTipText     =   "Delete this Phone Number Record"
         Top             =   4095
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   9
         Left            =   3340
         Picture         =   "frmContEntry.frx":5ADE
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   115
         TabStop         =   0   'False
         ToolTipText     =   "Delete this Phone Number Record"
         Top             =   3840
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   8
         Left            =   3340
         Picture         =   "frmContEntry.frx":5D50
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   114
         TabStop         =   0   'False
         ToolTipText     =   "Delete this Phone Number Record"
         Top             =   3585
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   7
         Left            =   3340
         Picture         =   "frmContEntry.frx":5FC2
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   113
         TabStop         =   0   'False
         ToolTipText     =   "Modify the setting attribute of this record"
         Top             =   1020
         Width           =   220
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   17
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1020
         Width           =   2190
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   6
         Left            =   3340
         Picture         =   "frmContEntry.frx":6234
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Modify this Address Record"
         Top             =   3135
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   5
         Left            =   3340
         Picture         =   "frmContEntry.frx":64A6
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Modify this Address Record"
         Top             =   2685
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   4
         Left            =   3340
         Picture         =   "frmContEntry.frx":6718
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Modify this Address Record"
         Top             =   2235
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   3
         Left            =   3340
         Picture         =   "frmContEntry.frx":698A
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Modify this Address Record"
         Top             =   1785
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   2
         Left            =   3340
         Picture         =   "frmContEntry.frx":6BFC
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Add a Job Title to this record"
         Top             =   1530
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   1
         Left            =   3340
         Picture         =   "frmContEntry.frx":6E6E
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Add or remove Group settings for this Contact"
         Top             =   765
         Width           =   220
      End
      Begin VB.PictureBox picPop 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   220
         Index           =   0
         Left            =   3340
         Picture         =   "frmContEntry.frx":70E0
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Modify this Name entry"
         Top             =   255
         Width           =   220
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   16
         Left            =   1390
         MaxLength       =   50
         TabIndex        =   17
         Top             =   5370
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   15
         Left            =   1390
         MaxLength       =   50
         TabIndex        =   16
         Top             =   5115
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   14
         Left            =   1390
         MaxLength       =   50
         TabIndex        =   15
         Top             =   4860
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   13
         Left            =   1390
         MaxLength       =   25
         TabIndex        =   14
         Top             =   4605
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   12
         Left            =   1390
         MaxLength       =   25
         TabIndex        =   13
         Top             =   4350
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   11
         Left            =   1390
         MaxLength       =   25
         TabIndex        =   12
         Top             =   4095
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   10
         Left            =   1390
         MaxLength       =   25
         TabIndex        =   11
         Top             =   3840
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   9
         Left            =   1390
         MaxLength       =   25
         TabIndex        =   10
         Top             =   3585
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   440
         Index           =   8
         Left            =   1390
         MaxLength       =   170
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3135
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   440
         Index           =   7
         Left            =   1390
         MaxLength       =   170
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2685
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   440
         Index           =   6
         Left            =   1390
         MaxLength       =   170
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2235
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   440
         Index           =   5
         Left            =   1390
         MaxLength       =   170
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1785
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   1390
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1530
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   1390
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1275
         Width           =   2190
      End
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   2
         Top             =   765
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
         Index           =   0
         Left            =   1390
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   0
         Top             =   255
         Width           =   2190
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Setting"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   17
         Left            =   15
         TabIndex        =   112
         Top             =   1020
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "E-Mail : Other"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   16
         Left            =   15
         TabIndex        =   39
         Top             =   5370
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "E-Mail : Work"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   15
         Left            =   15
         TabIndex        =   38
         Top             =   5115
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "E-Mail : Personal"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   14
         Left            =   15
         TabIndex        =   37
         Top             =   4860
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Phone : Fax"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   13
         Left            =   15
         TabIndex        =   36
         Top             =   4605
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Phone : Other"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   12
         Left            =   15
         TabIndex        =   35
         Top             =   4350
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Phone : Mobile"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   11
         Left            =   15
         TabIndex        =   34
         Top             =   4095
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Phone : Work"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   10
         Left            =   15
         TabIndex        =   33
         Top             =   3840
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Phone : Home"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   9
         Left            =   15
         TabIndex        =   32
         Top             =   3585
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Address : Ship To"
         ForeColor       =   &H00696969&
         Height          =   435
         Index           =   8
         Left            =   15
         TabIndex        =   31
         Top             =   3135
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Address : Bill To"
         ForeColor       =   &H00696969&
         Height          =   435
         Index           =   7
         Left            =   15
         TabIndex        =   30
         Top             =   2685
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Address : Work"
         ForeColor       =   &H00696969&
         Height          =   435
         Index           =   6
         Left            =   15
         TabIndex        =   29
         Top             =   2235
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Address : Home"
         ForeColor       =   &H00696969&
         Height          =   435
         Index           =   5
         Left            =   15
         TabIndex        =   28
         Top             =   1785
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Job Title"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   4
         Left            =   15
         TabIndex        =   27
         Top             =   1530
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Company"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   3
         Left            =   15
         TabIndex        =   26
         Top             =   1275
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Groups"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   2
         Left            =   15
         TabIndex        =   25
         Top             =   765
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Shown As"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   1
         Left            =   15
         TabIndex        =   24
         Top             =   510
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00ECECE1&
         Caption         =   "Name"
         ForeColor       =   &H00696969&
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   23
         Top             =   255
         Width           =   1365
      End
      Begin VB.Label lblProfileHdr 
         BackStyle       =   0  'Transparent
         Caption         =   " Name Profile"
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
         TabIndex        =   22
         Top             =   0
         Width           =   2115
      End
   End
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00F3F3ED&
      Height          =   390
      Left            =   10950
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7350
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H00666633&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11355
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   11355
      Begin MSComctlLib.TabStrip tbsMain 
         Height          =   315
         Left            =   8850
         TabIndex        =   77
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
      Begin VB.ComboBox cboContList 
         BackColor       =   &H00666633&
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
         Left            =   2025
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   50
         Width           =   4740
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   " Name Record :"
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
         TabIndex        =   19
         Top             =   75
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmContEntry.frx":7352
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdRelProj 
      BackColor       =   &H00CCCCB4&
      Caption         =   "Add Related Project ..."
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
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   6300
      Width           =   1665
   End
   Begin VB.Shape shpUserDef 
      BorderColor     =   &H00CCCCB4&
      Height          =   690
      Left            =   4800
      Top             =   7050
      Width           =   465
   End
   Begin VB.Label lblUserDefHdr2 
      BackColor       =   &H00ECECE1&
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
      TabIndex        =   72
      Top             =   6150
      Width           =   3765
   End
   Begin VB.Label lblUserDefHdr1 
      BackColor       =   &H00ECECE1&
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
      TabIndex        =   71
      Top             =   6150
      Width           =   3765
   End
   Begin VB.Label lblUserDef 
      BackColor       =   &H00CCCCB4&
      Caption         =   " User Defined Contact Fields"
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
      TabIndex        =   70
      Top             =   5925
      Width           =   7515
   End
   Begin VB.Shape shpRelCon 
      BorderColor     =   &H00CCCCB4&
      Height          =   1065
      Left            =   3750
      Top             =   4350
      Width           =   540
   End
   Begin VB.Label lblRelConHdr4 
      BackColor       =   &H00ECECE1&
      Caption         =   " E-mail:"
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
      Left            =   9750
      TabIndex        =   68
      Top             =   4275
      Width           =   1590
   End
   Begin VB.Label lblRelConHdr3 
      BackColor       =   &H00ECECE1&
      Caption         =   " Phone:"
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
      Left            =   8175
      TabIndex        =   67
      Top             =   4275
      Width           =   1590
   End
   Begin VB.Label lblRelConHdr2 
      BackColor       =   &H00ECECE1&
      Caption         =   " Memo:"
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
      Left            =   6150
      TabIndex        =   66
      Top             =   4275
      Width           =   2040
   End
   Begin VB.Label lblRelConHdr1 
      BackColor       =   &H00ECECE1&
      Caption         =   " Name:"
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
      TabIndex        =   65
      Top             =   4275
      Width           =   2340
   End
   Begin VB.Label lblRelCon 
      BackColor       =   &H00CCCCB4&
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
      Left            =   3825
      TabIndex        =   64
      Top             =   4050
      Width           =   7515
   End
   Begin VB.Shape shpAppts 
      BorderColor     =   &H00CCCCB4&
      Height          =   990
      Left            =   7575
      Top             =   2850
      Width           =   465
   End
   Begin VB.Label lblAppts 
      BackColor       =   &H00CCCCB4&
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
      TabIndex        =   61
      Top             =   2775
      Width           =   3690
   End
   Begin VB.Shape shpToDo 
      BorderColor     =   &H00CCCCB4&
      Height          =   690
      Left            =   3750
      Top             =   3075
      Width           =   465
   End
   Begin VB.Label lblToDoHdrDue 
      BackColor       =   &H00ECECE1&
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
      TabIndex        =   59
      Top             =   3000
      Width           =   1425
   End
   Begin VB.Label lblToDoHdr 
      BackColor       =   &H00ECECE1&
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
      TabIndex        =   58
      Top             =   3000
      Width           =   2250
   End
   Begin VB.Label lblToDo 
      BackColor       =   &H00CCCCB4&
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
      TabIndex        =   57
      Top             =   2775
      Width           =   3690
   End
   Begin VB.Shape shpNewHist 
      BorderColor     =   &H00CCCCB4&
      Height          =   1440
      Left            =   7575
      Top             =   1275
      Width           =   765
   End
   Begin VB.Label lblNewHist 
      BackColor       =   &H00CCCCB4&
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
      TabIndex        =   53
      Top             =   600
      Width           =   3690
   End
   Begin VB.Shape shpHist 
      BorderColor     =   &H00CCCCB4&
      Height          =   1140
      Left            =   3750
      Top             =   1575
      Width           =   615
   End
   Begin VB.Label lblHist 
      BackColor       =   &H00CCCCB4&
      Caption         =   " Contact History"
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
      TabIndex        =   51
      Top             =   600
      Width           =   3690
   End
   Begin VB.Shape shpRelPrj 
      BorderColor     =   &H00CCCCB4&
      Height          =   690
      Left            =   3225
      Top             =   6525
      Width           =   540
   End
   Begin VB.Label lblRelPrjHdr 
      BackColor       =   &H00ECECE1&
      Caption         =   " Project:                                      Status:"
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
      TabIndex        =   49
      Top             =   6525
      Width           =   3615
   End
   Begin VB.Label lblRelPrj 
      BackColor       =   &H00CCCCB4&
      Caption         =   " Related Projects"
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
      TabIndex        =   48
      Top             =   6300
      Width           =   3615
   End
End
Attribute VB_Name = "frmContEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsContact As Recordset 'main recordset
Dim rsNote As Recordset 'for Attach(notes) table
Dim rsList As Recordset 'all other data work
Dim rsAddress As Recordset 'for address entry
Dim rsPhone As Recordset 'for phone number entry
Dim rsEmail As Recordset 'for e-mail entry
Dim rsComment As Recordset 'for contact comments entry

Dim m_strOnEnter As String 'for text on entry into textbox
Dim m_strOnLeave As String 'for text when leaving textbox
Dim m_strNoteType As String 'for note type "C" = Call, "N" = Note
Dim m_strCType As String 'for contact type
'***for address parsing routine
Dim m_strStreetAddr As String
Dim m_strCity As String
Dim m_strState As String
Dim m_strZipCode As String
Dim m_strCountry As String '**Update
'******************************
Dim m_lngToDo As Long 'for selected to do item
Dim m_lngNotes As Long 'for selected notes item
Dim m_lngAppts As Long 'for selected appts item
'***for comments
Dim m_blnIsNewComment As Boolean
Dim m_blnChanged As Boolean
Dim m_lngCommID As Long
'***for full history grid
Dim m_blnFullHist As Boolean

Private Sub cboContList_Click()
   Const sMOD_NAME As String = "frmContEntry.cboContList_Click"
   On Error GoTo Error_Handler
   
   Dim lngContID As Long
   
   '***new code added 09/27/04
   lngContID = cboContList.ItemData(cboContList.ListIndex)
   
   frmSwtContact.m_lngContactID = lngContID
   Load frmSwtContact
   frmSwtContact.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the Contact List!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

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
   MsgBox "An un-known error occurred while gathering the Contact History!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub cmdAddFile_Click()
   MsgBox "Sorry, this feature is not available yet.", , APP_MSG_NAME
End Sub

Private Sub cmdAppts_Click()
   'Add new appointment
   Const sMOD_NAME As String = "frmContEntry.cmdAppts_Click"
   On Error GoTo Error_Handler
   
   icurState = NOW_ADDING
   Load frmAppt
   frmAppt.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while setting the Appointment screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub cmdHist_Click()
   Const sMOD_NAME As String = "frmContEntry.cmdHist_Click"
   On Error GoTo Error_Handler
   
   Call PostNewHistoryItem
   Call LoadContactHistory
   
   txtNewNote.Text = ""
   optNType(0).Value = True
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub cmdRelCon_Click()
   Const sMOD_NAME As String = "frmContEntry.cmdRelCon_Click"
   On Error GoTo Error_Handler
   
   Load frmNewLinkCont
   frmNewLinkCont.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while setting the Related Contact screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub cmdRelProj_Click()
   'show the dialog to add a related project
   Load frmSetRelProject
   frmSetRelProject.Show vbModeless, frmMain
End Sub

Private Sub cmdUserFld_Click()
   Const sMOD_NAME As String = "frmContEntry.cmdUserFld_Click"
   On Error GoTo Error_Handler
   
   Load frmUserFields
   frmUserFields.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while setting the User Defined Fields screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Activate()
   Const sMOD_NAME As String = "frmContEntry.Form_Activate"
   On Error GoTo Error_Handler
   
   If (Text1(17).Text = "Default") Then
      cboContList.Text = Text1(1).Text
   End If
   cboHistFilter.Text = "All Items"
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmContEntry.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Contact Information Screen", True
   frmMain.picStatus.BackColor = &H666633
   
   Set rsContact = dbContact.OpenRecordset("Contacts", dbOpenTable)
   Set rsNote = dbContact.OpenRecordset("Attach", dbOpenTable)
   Set rsAddress = dbContact.OpenRecordset("CAddress", dbOpenTable)
   Set rsPhone = dbContact.OpenRecordset("CPhone", dbOpenTable)
   Set rsEmail = dbContact.OpenRecordset("CEMail", dbOpenTable)
   Set rsComment = dbContact.OpenRecordset("CComments", dbOpenTable)
   
   'set note type
   m_strNoteType = "N"
   
   'load history pane combo box
   With cboHistFilter
      .AddItem "All Items"
      .AddItem "Notes & Calls"
      .AddItem "Documents"
      .AddItem "To Do's"
      .AddItem "E-mails"
      .AddItem "Appointments"
      .AddItem "Invoices"
      .AddItem "Bills"
   End With
   
   'Load all needed data
   Call LoadMainContactInfo
   Call LoadContactCombo
   Call LoadAddressInfo
   Call LoadPhoneInfo
   Call LoadEmailInfo
   Call LoadRelatedProjects
   Call LoadContactHistory
   Call LoadToDoInfo
   Call LoadApptsInfo
   Call LoadRelContactInfo
   Call LoadUserDefInfo
   Call LoadJobTitles
   Call LoadComments
   
   'set screen flag
   g_strFormFlag = "CEnt"
   
   'set gridline preference
   lvAppts.GridLines = g_blnShowLines
   lvHist.GridLines = g_blnShowLines
   lvHistPane.GridLines = g_blnShowLines
   lvRelCon.GridLines = g_blnShowLines
   lvRelPrj.GridLines = g_blnShowLines
   lvToDo.GridLines = g_blnShowLines
   lvUserDef.GridLines = g_blnShowLines
   
   'enable frmMain menu delete, & Print options
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
   
   LockWindowUpdate frmContEntry.hWnd
   
   'adjust rel proj items
   lblRelPrj.Move picProfile.Left, picProfile.Top + picProfile.Height + 225
   cmdRelProj.Move lblRelPrj.Left + lblRelPrj.Width - 1740, lblRelPrj.Top
   lblRelPrjHdr.Move lblRelPrj.Left, lblRelPrj.Top + lblRelPrj.Height, lblRelPrj.Width
   lvRelPrj.Move lblRelPrj.Left, lblRelPrjHdr.Top + 240, lblRelPrj.Width, Me.ScaleHeight - picProfile.Height - 1470
   shpRelPrj.Move lvRelPrj.Left - 15, lvRelPrj.Top - 15, lvRelPrj.Width + 30, lvRelPrj.Height + 30
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
   lblToDoHdrDue.Move lblToDo.Left + lblToDo.Width - 1440, lblToDo.Top + 240, 1425
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
   'adjust related contacts items
   lblRelCon.Move lblToDo.Left, lvAppts.Top + lvAppts.Height + 150, Me.ScaleWidth - picProfile.Width - 300
   cmdRelCon.Move lblRelCon.Left + lblRelCon.Width - 1740, lblRelCon.Top
   lblRelConHdr1.Move lblRelCon.Left, lblRelCon.Top + 240, lblRelCon.Width * 0.3
   lblRelConHdr2.Move lblRelConHdr1.Left + lblRelConHdr1.Width, lblRelConHdr1.Top, lblRelCon.Width * 0.3
   lblRelConHdr3.Move lblRelConHdr2.Left + lblRelConHdr2.Width, lblRelConHdr1.Top, lblRelCon.Width * 0.2
   lblRelConHdr4.Move lblRelConHdr3.Left + lblRelConHdr3.Width, lblRelConHdr3.Top, lblRelCon.Width * 0.2
   lvRelCon.Move lblRelCon.Left, lblRelConHdr1.Top + 240, lblRelCon.Width, Me.ScaleHeight * 0.15
   '***adjust lvRelCon header widths
   lvRelCon.ColumnHeaders(1).Width = lvRelCon.Width * 0.3
   lvRelCon.ColumnHeaders(2).Width = lvRelCon.Width * 0.3
   lvRelCon.ColumnHeaders(3).Width = lvRelCon.Width * 0.2
   lvRelCon.ColumnHeaders(4).Width = lvRelCon.Width * 0.2 - 265
   shpRelCon.Move lvRelCon.Left - 15, lvRelCon.Top - 15, lvRelCon.Width + 30, lvRelCon.Height + 30
   'adjust user defined items
   lblUserDef.Move lblRelCon.Left, lvRelCon.Top + lvRelCon.Height + 150, lblRelCon.Width
   cmdUserFld.Move lblUserDef.Left + lblUserDef.Width - 1365, lblUserDef.Top
   lblUserDefHdr1.Move lblUserDef.Left, lblUserDef.Top + 240, lblUserDef.Width * 0.5
   lblUserDefHdr2.Move lblUserDef.Left + lblUserDefHdr1.Width, lblUserDefHdr1.Top, lblUserDef.Width * 0.5
   lvUserDef.Move lblUserDef.Left, lblUserDefHdr1.Top + 240, lblUserDef.Width, Me.ScaleHeight * 0.23
   '***adjust lvUserDef header widths
   lvUserDef.ColumnHeaders(1).Width = lvUserDef.Width * 0.5
   lvUserDef.ColumnHeaders(2).Width = lvUserDef.Width * 0.5 - 265
   shpUserDef.Move lvUserDef.Left - 15, lvUserDef.Top - 15, lvUserDef.Width + 30, lvUserDef.Height + 30
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
   'adjust Job Title listbox
   lstJobTitle.Move Text1(4).Left, Text1(4).Top + Text1(4).Height + 15, Text1(4).Width

   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'save any entered notes not already saved
   If (txtNewNote.Text <> "") Then
      Call PostNewHistoryItem
   End If
   'remove data & form reference
   rsContact.Close
   Set rsContact = Nothing
   rsNote.Close
   Set rsNote = Nothing
   rsAddress.Close
   Set rsAddress = Nothing
   rsPhone.Close
   Set rsPhone = Nothing
   rsEmail.Close
   Set rsEmail = Nothing
   rsComment.Close
   Set rsComment = Nothing
   
   'disable frmMain menu delete & print options
   frmMain.mnuEditDelete.Enabled = False
   frmMain.mnuFilePrint.Enabled = False
   frmMain.tbrMain.Buttons(7).Enabled = False
   
   Set frmContEntry = Nothing
End Sub

Public Sub LoadMainContactInfo()
   'load the main contact fields from Contacts table
   Const sMOD_NAME As String = "frmContEntry.LoadMainContactInfo"
   On Error GoTo Error_Handler
   
   With rsContact
      If (.RecordCount > 0) Then
         .MoveFirst
         .Index = "PrimaryKey"
         .Seek "=", g_lngContID
      End If
      
      Dim Indx As Integer
      For Indx = 0 To 4
         Text1(Indx).Text = ""
      Next
      
      If (Not IsNull(!FullName)) Then Text1(0) = !FullName
      If (Not IsNull(!ShownName)) Then Text1(1) = !ShownName
      If (Not IsNull(!Group)) Then Text1(2) = !Group
      If (Not IsNull(!CompName)) Then Text1(3) = !CompName
      If (Not IsNull(!JobTitle)) Then Text1(4) = !JobTitle
      If (Not IsNull(!CTYPE)) Then
         m_strCType = !CTYPE
         If !CTYPE = "I" Then
            Text1(3).Locked = False
         ElseIf !CTYPE = "C" Then
            Text1(0).Locked = False
         End If
      End If
      If (Not IsNull(!Setting)) Then Text1(17) = !Setting
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadContactCombo()
   'load all contact ShownNames for selection into cboContList
   Const sMOD_NAME As String = "frmContEntry.LoadContactCombo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strSetting As String
   
   strSetting = "Default"
   
   SQL = "SELECT ContID, Setting, ShownName FROM Contacts "
   SQL = SQL & "WHERE Setting = '" & strSetting & "' "
   SQL = SQL & "ORDER BY ShownName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   cboContList.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ContID)) Then cboContList.AddItem !ShownName
            cboContList.ItemData(cboContList.NewIndex) = !ContID
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

Private Sub lblHyper_Click(Index As Integer)
   Const sMOD_NAME As String = "frmContEntry.lblHyper_Click"
   On Error GoTo Error_Handler
   
   Dim strSelect As String
   
   strSelect = txtComments.SelText
   
   If (strSelect = "") Then
      MsgBox "There is no selected text, from Contact Comments.", , APP_MSG_NAME
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

Private Sub lstJobTitle_Click()
   Const sMOD_NAME As String = "frmContEntry.lstJobTitle_Click"
   On Error GoTo Error_Handler
   
   If (lstJobTitle.Text = "<Add Field...>") Then
      lstJobTitle.Visible = False
      Load frmSetJobTitle
      frmSetJobTitle.Show vbModeless, frmMain
      Exit Sub
   Else
      Text1(4).Text = lstJobTitle.Text
      lstJobTitle.Visible = False
      Text1(4).SetFocus
      
      Call EditJobTitle
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   lstJobTitle.Visible = False
End Sub

Private Sub lvAppts_Click()
   Const sMOD_NAME As String = "frmContEntry.lvAppts_Click"
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
      MsgBox "An un-known error occurred while opening the Appointments dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvHist_Click()
   Const sMOD_NAME As String = "frmContEntry.lvHist_Click"
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
      MsgBox "An un-known error occurred while opening the Calls/Notes dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvHistPane_Click()
   Const sMOD_NAME As String = "frmContEntry.lvHistPane_Click"
   On Error GoTo Error_Handler
   
   Dim strKeyType As String
   Dim lngAttach As Long
   Dim lngToDo As Long
   Dim lngAppt As Long
   
   strKeyType = Left(lvHistPane.SelectedItem.Key, 2)
   
   Select Case strKeyType
      Case "CA", "NT" 'call or note
         lngAttach = CLng(Mid$(lvHistPane.SelectedItem.Key, 3, Len(lvHistPane.SelectedItem.Key)))
         'code to open Notes entry screen
         icurState = NOW_EDITING
         frmNotes.m_lngNoteID = lngAttach
         Load frmNotes
         frmNotes.Show vbModeless, frmMain
      Case "TD" 'to do item
         lngToDo = CLng(Mid$(lvHistPane.SelectedItem.Key, 3, Len(lvHistPane.SelectedItem.Key)))
         'code to open ToDo entry screen
         icurState = NOW_EDITING
         frmToDo.m_lngToDoID = lngToDo
         Load frmToDo
         frmToDo.Show vbModeless, frmMain
      Case "AP" 'appointment
         lngAppt = CLng(Mid$(lvHistPane.SelectedItem.Key, 3, Len(lvHistPane.SelectedItem.Key)))
         'code to open Appts entry screen
         icurState = NOW_EDITING
         frmAppt.m_lngApptID = lngAppt
         Load frmAppt
         frmAppt.Show vbModeless, frmMain
   End Select
   
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

Private Sub lvRelCon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Const sMOD_NAME As String = "frmContEntry.lvRelCon_MouseUp"
   On Error GoTo Error_Handler
   Dim lngContID As Long
   
   If Button = vbLeftButton Then
      lngContID = CLng(Mid$(lvRelCon.SelectedItem.Key, 3, Len(lvRelCon.SelectedItem.Key)))
   
      frmSwtContact.m_lngContactID = lngContID
      Load frmSwtContact
      frmSwtContact.Show , frmMain
   ElseIf Button = vbRightButton Then
      If (lvRelCon.ListItems.Count > 0) Then
         lngContID = CLng(Mid$(lvRelCon.SelectedItem.Key, 3, Len(lvRelCon.SelectedItem.Key)))
   
         frmMain.m_lngRelContID = lngContID
         frmMain.m_strOldRCMemo = lvRelCon.SelectedItem.SubItems(1)
         PopupMenu frmMain.mnuRCMemoPop
      End If
   End If
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
   End If
End Sub

Private Sub lvRelPrj_Click()
   Const sMOD_NAME As String = "frmContEntry.lvRelPrj_Click"
   On Error GoTo Error_Handler
   
   Dim lngProjID As Long
   
   lngProjID = CLng(Mid$(lvRelPrj.SelectedItem.Key, 3, Len(lvRelPrj.SelectedItem.Key)))
   
   g_lngProjID = lngProjID
   UnloadAllForms
   Load frmProjEntry
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while opening the Related Project Information screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvToDo_Click()
   Const sMOD_NAME As String = "frmContEntry.lvToDo_Click"
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
      MsgBox "An un-known error occurred while opening the To Do entry screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvUserDef_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Const sMOD_NAME As String = "frmContEntry.lvUserDef_MouseUp"
   On Error GoTo Error_Handler
   Dim lngUserDefID As Long
   
   If Button = vbRightButton Then
      If lvUserDef.ListItems.Count = 0 Then Exit Sub
      lngUserDefID = CLng(Mid$(lvUserDef.SelectedItem.Key, 3, Len(lvUserDef.SelectedItem.Key)))
   
      'code to Delete user defined field
      frmMain.m_lngContUserFld = lngUserDefID
      frmMain.m_strContUserFld = lvUserDef.SelectedItem
      PopupMenu frmMain.mnuDelContFld
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

Private Sub picPop_Click(Index As Integer)
   Const sMOD_NAME As String = "frmContEntry.picPop_Click"
   On Error GoTo Error_Handler
   
   Dim iMsg As VbMsgBoxResult
   
   Select Case Index
      Case 0 'name
         Text1(0).SetFocus
         Load frmNameDetail
         frmNameDetail.Show vbModeless, frmMain
      Case 1 'Groups
         Text1(2).SetFocus
         Load frmSelectGrp
         frmSelectGrp.Show vbModeless, frmMain
      Case 2 'Job Title
         Text1(4).SetFocus
         
         If (lstJobTitle.Visible = True) Then
            lstJobTitle.Visible = False
         Else
            lstJobTitle.Visible = True
         End If
      Case 3 'addr-home
         If (Text1(5).Text = "") Then Exit Sub
         Text1(5).SetFocus
         frmAddrDetail.m_strAddrType = "Home"
         Load frmAddrDetail
         frmAddrDetail.Show vbModeless, frmMain
      Case 4 'addr-work
         If (Text1(6).Text = "") Then Exit Sub
         Text1(6).SetFocus
         frmAddrDetail.m_strAddrType = "Work"
         Load frmAddrDetail
         frmAddrDetail.Show vbModeless, frmMain
      Case 5 'addr-bill to
         If (Text1(7).Text = "") Then Exit Sub
         Text1(7).SetFocus
         frmAddrDetail.m_strAddrType = "Bill To"
         Load frmAddrDetail
         frmAddrDetail.Show vbModeless, frmMain
      Case 6 'addr-ship to
         If (Text1(8).Text = "") Then Exit Sub
         Text1(8).SetFocus
         frmAddrDetail.m_strAddrType = "Ship To"
         Load frmAddrDetail
         frmAddrDetail.Show vbModeless, frmMain
      Case 7 'contact Setting
         If (Text1(17).Text = "") Then Exit Sub
         Text1(17).SetFocus
         Load frmSetting
         frmSetting.Show vbModeless, frmMain
      Case 8 'delete home phone
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeletePhoneNums("Home")
      Case 9 'delete work phone
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeletePhoneNums("Work")
      Case 10 'delete mobile phone
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeletePhoneNums("Mobile")
      Case 11 'delete other phone
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeletePhoneNums("Other")
      Case 12 'delete fax phone
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeletePhoneNums("Fax")
      Case 13 'delete personal e-mail
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeleteEMails("Personal")
      Case 14 'delete work e-mail
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeleteEMails("Work")
      Case 15 'delete other e-mail
         iMsg = MsgBox("Are you sure?", vbQuestion + vbYesNo, "Confirm Delete")
         If (iMsg <> vbYes) Then Exit Sub
         
         Call DeleteEMails("Other")
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub picPop_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      picPop(Index).BackColor = &H808080
   End If
End Sub

Private Sub picPop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      picPop(Index).BackColor = vbWhite
   End If
End Sub

Private Sub tbsMain_Click()
   Const sMOD_NAME As String = "frmContEntry.tbsMain_Click"
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
   Const sMOD_NAME As String = "frmContEntry.Text1_LostFocus"
   On Error GoTo Error_Handler
   
   Dim intStrLen As Integer
   Dim strAddress As String
   
   m_strOnLeave = Text1(Index).Text
   intStrLen = Len(m_strOnEnter)
   
   If m_strOnEnter <> m_strOnLeave Then
      Select Case Index
         Case 0 'Contacts.FullName
            If (m_strOnLeave = "") Then Exit Sub
            Call EditContactName
         Case 3 'Contacts.CompName
            If (m_strOnLeave = "") Then Exit Sub
            Call EditCompName
         Case 4 'Contacts.JobTitle
            If (m_strOnLeave = "") Then Exit Sub
            Call EditJobTitle
         Case 5 'CAddress.Home
            If (intStrLen <= 0) Then
               strAddress = Text1(5).Text
               ParseStrAddr (strAddress)
               icurState = NOW_ADDING
               Call EditHomeAddress
            ElseIf (intStrLen > 0) Then
               strAddress = Text1(5).Text
               ParseStrAddr (strAddress)
               icurState = NOW_EDITING
               Call EditHomeAddress
            End If
         Case 6 'CAddress.Work
            If (intStrLen <= 0) Then
               strAddress = Text1(6).Text
               ParseStrAddr (strAddress)
               icurState = NOW_ADDING
               Call EditWorkAddress
            ElseIf (intStrLen > 0) Then
               strAddress = Text1(6).Text
               ParseStrAddr (strAddress)
               icurState = NOW_EDITING
               Call EditWorkAddress
            End If
         Case 7 'CAddress.BillTo
            If (intStrLen <= 0) Then
               strAddress = Text1(7).Text
               ParseStrAddr (strAddress)
               icurState = NOW_ADDING
               Call EditBillingAddress
            ElseIf (intStrLen > 0) Then
               strAddress = Text1(7).Text
               ParseStrAddr (strAddress)
               icurState = NOW_EDITING
               Call EditBillingAddress
            End If
         Case 8 'CAddress.Ship To
            If (intStrLen <= 0) Then
               strAddress = Text1(8).Text
               ParseStrAddr (strAddress)
               icurState = NOW_ADDING
               Call EditShippingAddress
            ElseIf (intStrLen > 0) Then
               strAddress = Text1(8).Text
               ParseStrAddr (strAddress)
               icurState = NOW_EDITING
               Call EditShippingAddress
            End If
         Case 9 'CPhone.Home
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditHomePhone
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditHomePhone
            End If
         Case 10 'CPhone.Work
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditWorkPhone
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditWorkPhone
            End If
         Case 11 'CPhone.Mobile
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditMobilePhone
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditMobilePhone
            End If
         Case 12 'CPhone.Other
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditOtherPhone
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditOtherPhone
            End If
         Case 13 'CPhone.Fax
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditFaxPhone
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditFaxPhone
            End If
         Case 14 'CEMail.Personal
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditPersonalEmail
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditPersonalEmail
            End If
         Case 15 'CEMail.Work
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditWorkEmail
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditWorkEmail
            End If
         Case 16 'CEMail.Other
            If (intStrLen <= 0) Then
               icurState = NOW_ADDING
               Call EditOtherEmail
            ElseIf (intStrLen > 0) Then
               icurState = NOW_EDITING
               Call EditOtherEmail
            End If
         Case 17 'Contacts.Setting
            If (m_strOnLeave = "") Then Exit Sub
            Call EditSetting
      End Select
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim strTip As String
   
   strTip = Text1(Index).Text
   strTip = Replace(strTip, vbCrLf, "  ")
   
   Text1(Index).ToolTipText = strTip
End Sub

Public Sub LoadRelatedProjects()
   'load all projects related to this contact
   Const sMOD_NAME As String = "frmContEntry.LoadRelatedProjects"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT fkProjID, fkContID FROM RelateProject "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvRelPrj.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkProjID)) Then
               Call GetProjectInfo(!fkProjID)
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
End Sub

Private Sub GetProjectInfo(lngPrjID As Long)
   'get the project info from the proj link in contact table
   Const sMOD_NAME As String = "frmContEntry.GetProjectInfo"
   On Error GoTo Error_Handler
   
   Dim cSQL As String
   Dim rsProj As Recordset
   Dim Item As ListItem
   
   cSQL = "SELECT ProjID, PName, Status FROM Projects "
   cSQL = cSQL & "WHERE ProjID = " & lngPrjID
   
   Set rsProj = dbContact.OpenRecordset(cSQL)
   
   With rsProj
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!PName)) Then Set Item = lvRelPrj.ListItems.Add(, "ID" & !ProjID, !PName)
         Item.SubItems(1) = "* " & !Status
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvRelPrj, picGrdClr
   End If
   
   rsProj.Close
   Set rsProj = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadAddressInfo()
   'load any stored address info for this contact
   Const sMOD_NAME As String = "frmContEntry.LoadAddressInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip, Country "
   SQL = SQL & "FROM CAddress "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkLookup)) Then
               Select Case !fkLookup
                  Case "Home"
                     Text1(5) = !Street & vbCrLf & !City & ", " & !State & " " & !Zip & vbCrLf & !Country
                     Text1(5).Locked = True 'added 10.28.04
                  Case "Work"
                     Text1(6) = !Street & vbCrLf & !City & ", " & !State & " " & !Zip & vbCrLf & !Country
                     Text1(6).Locked = True 'added 10.28.04
                  Case "Bill To"
                     Text1(7) = !Street & vbCrLf & !City & ", " & !State & " " & !Zip & vbCrLf & !Country
                     Text1(7).Locked = True 'added 10.28.04
                  Case "Ship To"
                     Text1(8) = !Street & vbCrLf & !City & ", " & !State & " " & !Zip & vbCrLf & !Country
                     Text1(8).Locked = True 'added 10.28.04
               End Select
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
End Sub

Public Sub LoadContactHistory()
   'load all history items for this contact
   Const sMOD_NAME As String = "frmContEntry.LoadContactHistory"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, fkContID, NType, TextBody, DateStamp "
   SQL = SQL & "FROM Attach "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
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

Public Sub LoadPhoneInfo()
   'load any stored phone number info for this contact
   Const sMOD_NAME As String = "frmContEntry.LoadPhoneInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT fkContID, fkLookup, PhoneNum "
   SQL = SQL & "FROM CPhone "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkLookup)) Then
               Select Case !fkLookup
                  Case "Home"
                     Text1(9) = !PhoneNum
                  Case "Work"
                     Text1(10) = !PhoneNum
                  Case "Mobile"
                     Text1(11) = !PhoneNum
                  Case "Other"
                     Text1(12) = !PhoneNum
                  Case "Fax"
                     Text1(13) = !PhoneNum
               End Select
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
End Sub

Public Sub LoadEmailInfo()
   'load any stored phone number info for this contact
   Const sMOD_NAME As String = "frmContEntry.LoadEmailInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT fkContID, fkLookup, Email "
   SQL = SQL & "FROM CEMail "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkLookup)) Then
               Select Case !fkLookup
                  Case "Personal"
                     Text1(14) = !Email
                  Case "Work"
                     Text1(15) = !Email
                  Case "Other"
                     Text1(16) = !Email
               End Select
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
End Sub

Public Sub LoadToDoInfo()
   'load all associated to do information
   Const sMOD_NAME As String = "frmContEntry.LoadToDoInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, Subject, fkContID, DueDate, Completed "
   SQL = SQL & "FROM ToDo "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
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
   Const sMOD_NAME As String = "frmContEntry.LoadApptsInfo"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim vDate As Variant
   Dim sDate As String
   
   vDate = Format(Date, "mm/dd/yyyy")
   vDate = "#" & vDate & "#"
   
   SQL = "SELECT RefNum, fkContID, Subject, DateFrom FROM Appts "
   SQL = SQL & "WHERE DateFrom >= " & vDate
   SQL = SQL & " AND fkContID = " & g_lngContID
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

Public Sub LoadRelContactInfo()
   'load all contacts related to the currently viewed one
   Const sMOD_NAME As String = "frmContEntry.LoadRelContactInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strPhone As String
   Dim strEmail As String
   Dim Item As ListItem
   
   SQL = "SELECT MasterContID, LinkMemo, SubContID, SubContShowName "
   SQL = SQL & "FROM RelateCont "
   SQL = SQL & "WHERE MasterContID = " & g_lngContID
   SQL = SQL & " ORDER BY SubContShowName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvRelCon.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!SubContID)) Then
               If (Not IsNull(!SubContShowName)) Then
                  Set Item = lvRelCon.ListItems.Add(, "ID" & !SubContID, !SubContShowName)
                  If (Not IsNull(!LinkMemo)) Then Item.SubItems(1) = !LinkMemo
                  'code for phone & email
                  'strPhone = GetRelPhoneNum(!ContID)
                  strPhone = GetPhoneNum(!SubContID)
                  If (Not IsNull(strPhone)) Then Item.SubItems(2) = strPhone
                  'strEmail = GetRelEMail(!ContID)
                  strEmail = GetEMail(!SubContID)
                  If (Not IsNull(strEmail)) Then Item.SubItems(3) = strEmail
               End If
            End If
            .MoveNext
         Wend
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

Public Sub LoadUserDefInfo()
   'load all user defined fields for this contact
   Const sMOD_NAME As String = "frmContEntry.LoadUserDefInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, fkContID, fkUserFld, Value FROM CUFldValues "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
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

Private Sub PostNewHistoryItem()
   'save any new note / call item
   Const sMOD_NAME As String = "frmContEntry.PostNewHistoryItem"
   On Error GoTo Error_Handler
   
   If (txtNewNote.Text = "") Then Exit Sub
   
   rsNote.AddNew
   
   With rsNote
      If (Len(txtNewNote)) Then !TextBody = txtNewNote
      !fkContID = g_lngContID
      !NType = m_strNoteType
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while posting the History item!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditContactName()
   'add a contact name (only if listing is a company)
   Const sMOD_NAME As String = "frmContEntry.EditContactName"
   On Error GoTo Error_Handler
   
   rsContact.Edit
   
   With rsContact
      !FullName = Text1(0).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred editing the Contact Name!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Public Sub LoadJobTitles()
   'load all job titles into lstJobTitles from Lookup table
   Const sMOD_NAME As String = "frmContEntry.LoadJobTitles"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "TITLE"
   
   SQL = "SELECT ItemID, Description FROM Lookup "
   SQL = SQL & "WHERE ItemID = '" & strType & "' "
   SQL = SQL & "ORDER BY Description"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lstJobTitle.Clear
   lstJobTitle.AddItem "<Add Field...>"
   lstJobTitle.AddItem " "
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then lstJobTitle.AddItem !Description
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

Private Sub EditJobTitle()
   'add a job title
   Const sMOD_NAME As String = "frmContEntry.EditJobTitle"
   On Error GoTo Error_Handler
   
   rsContact.Edit
   
   With rsContact
      !JobTitle = Text1(4).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Job Title!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub ParseStrAddr(strAddrToParse As String)
   'parse the street, city, state, zip from address string
   Const sMOD_NAME As String = "frmContEntry.ParseStrAddr"
   On Error GoTo Error_Handler
   
   Dim strFullAddr() As String
   Dim strCity() As String
   Dim strCtyStZipPart As String
   Dim strRemoveCity As String
   Dim strRemoveStZip As String
   Dim strStZip() As String
   
   'find the carriage return / line feed after street addr & Cty St Zip
   'if a country was entered
   strAddrToParse = Replace(strAddrToParse, Chr(13) + Chr(10), "_")
   
   'split address into street _ cty st zip _ country
   strFullAddr = Split(strAddrToParse, "_")
   
   'set street address & country
   m_strStreetAddr = strFullAddr(0)
   If UBound(strFullAddr) > 1 Then
      m_strCountry = strFullAddr(2)
   End If
   'setup city, state, zip part
   strCtyStZipPart = strFullAddr(1)
   
   'remove any comma/space (between "City, State")
   strRemoveCity = Replace(strCtyStZipPart, ", ", "_")
   'peel off city
   strCity = Split(strRemoveCity, "_")
   'assign city
   m_strCity = strCity(0)
   
   'remove any comma/space (between "City, State Zip"), or blank spaces
   strRemoveStZip = Replace(strCtyStZipPart, ", ", "_")
   strRemoveStZip = Replace(strRemoveStZip, " ", "_")
   
   'assign st, zip variables
   'split State Zip part
   strStZip = Split(strRemoveStZip, "_")
   If UBound(strStZip) = 3 Then 'if space between city names (Forest Grove)
      m_strState = strStZip(2)
      m_strZipCode = strStZip(3)
   Else 'no space, single name (Portland)
      m_strState = strStZip(1)
      m_strZipCode = strStZip(2)
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub EditHomeAddress()
   'add or edit the home address listing
   Const sMOD_NAME As String = "frmContEntry.EditHomeAddress"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Home"
   
   Select Case icurState
      Case NOW_ADDING
         rsAddress.AddNew
         With rsAddress
            !fkContID = g_lngContID
            !fkLookup = strType
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip "
         SQL = SQL & "FROM CAddress "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   m_strStreetAddr = ""
   m_strCity = ""
   m_strState = ""
   m_strZipCode = ""
   m_strCountry = ""
   
   Call LoadAddressInfo 'added 10.28.04
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Home address!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditCompName()
   'add the company name
   Const sMOD_NAME As String = "frmContEntry.EditCompName"
   On Error GoTo Error_Handler
   
   rsContact.Edit
   
   With rsContact
      !CompName = Text1(3).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Company Name!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditSetting()
   'edit the contact setting value
   Const sMOD_NAME As String = "frmContEntry.EditSetting"
   On Error GoTo Error_Handler
   
   rsContact.Edit
   
   With rsContact
      !Setting = Text1(17).Text
      
      .Update
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Contact Setting!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditWorkAddress()
   'add or edit the work address listing
   Const sMOD_NAME As String = "frmContEntry.EditWorkAddress"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Work"
   
   Select Case icurState
      Case NOW_ADDING
         rsAddress.AddNew
         With rsAddress
            !fkContID = g_lngContID
            !fkLookup = strType
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip "
         SQL = SQL & "FROM CAddress "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   m_strStreetAddr = ""
   m_strCity = ""
   m_strState = ""
   m_strZipCode = ""
   m_strCountry = ""
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Work address!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditBillingAddress()
   'add or edit the billing address listing
   Const sMOD_NAME As String = "frmContEntry.EditBillingAddress"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Bill To"
   
   Select Case icurState
      Case NOW_ADDING
         rsAddress.AddNew
         With rsAddress
            !fkContID = g_lngContID
            !fkLookup = strType
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip "
         SQL = SQL & "FROM CAddress "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   m_strStreetAddr = ""
   m_strCity = ""
   m_strState = ""
   m_strZipCode = ""
   m_strCountry = ""
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Billing address!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditShippingAddress()
   'add or edit the Shipping address listing
   Const sMOD_NAME As String = "frmContEntry.EditShippingAddress"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Ship To"
   
   Select Case icurState
      Case NOW_ADDING
         rsAddress.AddNew
         With rsAddress
            !fkContID = g_lngContID
            !fkLookup = strType
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip "
         SQL = SQL & "FROM CAddress "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (m_strStreetAddr <> "") Then !Street = m_strStreetAddr
            If (m_strCity <> "") Then !City = m_strCity
            If (m_strState <> "") Then !State = m_strState
            If (m_strZipCode <> "") Then !Zip = m_strZipCode
            If (m_strCountry <> "") Then !Country = m_strCountry
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   m_strStreetAddr = ""
   m_strCity = ""
   m_strState = ""
   m_strZipCode = ""
   m_strCountry = ""
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Shipping address!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditHomePhone()
   'add or edit the home phone listing
   Const sMOD_NAME As String = "frmContEntry.EditHomePhone"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Home"
   
   Select Case icurState
      Case NOW_ADDING
         rsPhone.AddNew
         With rsPhone
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(9))) Then !PhoneNum = Text1(9)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CPhone "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(9))) Then !PhoneNum = Text1(9)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Home Phone!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditWorkPhone()
   'add or edit the work phone listing
   Const sMOD_NAME As String = "frmContEntry.EditWorkPhone"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Work"
   
   Select Case icurState
      Case NOW_ADDING
         rsPhone.AddNew
         With rsPhone
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(10))) Then !PhoneNum = Text1(10)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CPhone "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(10))) Then !PhoneNum = Text1(10)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Work Phone!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditMobilePhone()
   'add or edit the mobile phone listing
   Const sMOD_NAME As String = "frmContEntry.EditMobilePhone"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Mobile"
   
   Select Case icurState
      Case NOW_ADDING
         rsPhone.AddNew
         With rsPhone
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(11))) Then !PhoneNum = Text1(11)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CPhone "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(11))) Then !PhoneNum = Text1(11)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Mobile Phone!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditOtherPhone()
   'add or edit the other phone listing
   Const sMOD_NAME As String = "frmContEntry.EditOtherPhone"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Other"
   
   Select Case icurState
      Case NOW_ADDING
         rsPhone.AddNew
         With rsPhone
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(12))) Then !PhoneNum = Text1(12)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CPhone "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(12))) Then !PhoneNum = Text1(12)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Other Phone!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditFaxPhone()
   'add or edit the fax phone listing
   Const sMOD_NAME As String = "frmContEntry.EditFaxPhone"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Fax"
   
   Select Case icurState
      Case NOW_ADDING
         rsPhone.AddNew
         With rsPhone
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(13))) Then !PhoneNum = Text1(13)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CPhone "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(13))) Then !PhoneNum = Text1(13)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Fax Phone!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditPersonalEmail()
   'add or edit the personal email listing
   Const sMOD_NAME As String = "frmContEntry.EditPersonalEmail"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Personal"
   
   Select Case icurState
      Case NOW_ADDING
         rsEmail.AddNew
         With rsEmail
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(14))) Then !Email = Text1(14)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CEMail "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(14))) Then !Email = Text1(14)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Personal E-Mail!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditWorkEmail()
   'add or edit the work email listing
   Const sMOD_NAME As String = "frmContEntry.EditWorkEmail"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Work"
   
   Select Case icurState
      Case NOW_ADDING
         rsEmail.AddNew
         With rsEmail
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(15))) Then !Email = Text1(15)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CEMail "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(15))) Then !Email = Text1(15)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Work E-Mail!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EditOtherEmail()
   'add or edit the other email listing
   Const sMOD_NAME As String = "frmContEntry.EditOtherEmail"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strType As String
   
   strType = "Other"
   
   Select Case icurState
      Case NOW_ADDING
         rsEmail.AddNew
         With rsEmail
            !fkContID = g_lngContID
            !fkLookup = strType
            If (Len(Text1(16))) Then !Email = Text1(16)
            
            .Update
         End With
      Case NOW_EDITING
         SQL = "SELECT fkContID, fkLookup, PhoneNum "
         SQL = SQL & "FROM CEMail "
         SQL = SQL & "WHERE fkContID = " & g_lngContID
         SQL = SQL & " AND fkLookup = '" & strType & "' "
         
         Set rsList = dbContact.OpenRecordset(SQL)
         
         rsList.Edit
         With rsList
            If (Len(Text1(16))) Then !Email = Text1(16)
            
            .Update
         End With
         
         rsList.Close
         Set rsList = Nothing
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Other E-Mail!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub LoadComments()
   'load any listed comments for this contact
   Const sMOD_NAME As String = "frmContEntry.LoadComments"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT RefNum, fkContID, Comments FROM CComments "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   
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

Private Sub PostCommentEntry()
   'post any comment entered into the database
   Const sMOD_NAME As String = "frmContEntry.PostCommentEntry"
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
      !fkContID = g_lngContID
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
   
   MsgBox "An un-known error occurred while Posting Contact Comments!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtComments_Change()
   m_blnChanged = True
End Sub

Private Sub LoadNotesCallsHistory()
   'load all history items for this contact
   Const sMOD_NAME As String = "frmContEntry.LoadNotesCallsHistory"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, fkContID, NType, TextBody, DateStamp "
   SQL = SQL & "FROM Attach "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
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
   Const sMOD_NAME As String = "frmContEntry.LoadToDoHistory"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   Dim strSubject As String
   Dim strTextBody As String
   
   SQL = "SELECT RefNum, Subject, fkContID, DueDate, TextBody "
   SQL = SQL & "FROM ToDo "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
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
   Const sMOD_NAME As String = "frmContEntry.LoadApptsHistory"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim strSubject As String
   Dim strTextBody As String
   
   SQL = "SELECT RefNum, fkContID, Subject, DateFrom, TextBody FROM Appts "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
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

Private Sub DeletePhoneNums(strType As String)
   'delete desired contact phone numbers
   Const sMOD_NAME As String = "frmContEntry.DeletePhoneNums"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim iCtr As Integer
   
   SQL = "DELETE * FROM CPhone "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   SQL = SQL & " AND fkLookup = '" & strType & "' "
   
   dbContact.Execute (SQL)
   
   For iCtr = 9 To 13
      Text1(iCtr).Text = ""
   Next
   Call LoadPhoneInfo
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub DeleteEMails(strType As String)
   'delete desired contact e-mail addresses
   Const sMOD_NAME As String = "frmContEntry.DeleteEMails"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim iCtr As Integer
   
   SQL = "DELETE * FROM CEMail "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   SQL = SQL & " AND fkLookup = '" & strType & "' "
   
   dbContact.Execute (SQL)
   
   For iCtr = 14 To 16
      Text1(iCtr).Text = ""
   Next
   Call LoadEmailInfo
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub
