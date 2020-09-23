VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCalWeek 
   BackColor       =   &H00CAD9DB&
   Caption         =   "By Week"
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
   Icon            =   "frmCalWeek.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picQckAdd 
      BackColor       =   &H00CAD9DB&
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   8700
      ScaleHeight     =   2715
      ScaleWidth      =   2565
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5025
      Width           =   2565
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00CAD9DB&
         Caption         =   "&Add It"
         Default         =   -1  'True
         Height          =   345
         Left            =   750
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2250
         Width           =   990
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   1800
         Width           =   1740
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1425
         Width           =   1740
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         ItemData        =   "frmCalWeek.frx":0442
         Left            =   1725
         List            =   "frmCalWeek.frx":04D6
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1050
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmCalWeek.frx":0696
         Left            =   750
         List            =   "frmCalWeek.frx":072A
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   1050
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   750
         MaxLength       =   255
         TabIndex        =   43
         Top             =   300
         Width           =   1740
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   0
         Left            =   750
         TabIndex        =   45
         Top             =   675
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   13294043
         CalendarTitleBackColor=   11652052
         Format          =   53608449
         CurrentDate     =   38252
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   315
         Index           =   1
         Left            =   1725
         TabIndex        =   47
         Top             =   675
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   13294043
         CalendarTitleBackColor=   11652052
         Format          =   53608449
         CurrentDate     =   38252
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Project:"
         Height          =   240
         Index           =   6
         Left            =   75
         TabIndex        =   54
         Top             =   1837
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   52
         Top             =   1462
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   240
         Index           =   4
         Left            =   1425
         TabIndex        =   50
         Top             =   1087
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   48
         Top             =   1087
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   240
         Index           =   2
         Left            =   1425
         TabIndex        =   46
         Top             =   712
         Width           =   240
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   44
         Top             =   712
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Appt:"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   42
         Top             =   337
         Width           =   615
      End
      Begin VB.Label lblQckAdd 
         Alignment       =   2  'Center
         BackColor       =   &H004A4A4A&
         Caption         =   "Quick Add - New Appointment"
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
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   2565
      End
   End
   Begin VB.PictureBox picCalSelect 
      BackColor       =   &H00CAD9DB&
      BorderStyle     =   0  'None
      Height          =   4440
      Left            =   8700
      ScaleHeight     =   4440
      ScaleWidth      =   2565
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   525
      Width           =   2565
      Begin MSComCtl2.MonthView mvwCal 
         Height          =   4455
         Left            =   0
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   7858
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthRows       =   2
         MonthBackColor  =   13294043
         StartOfWeek     =   53608449
         TitleBackColor  =   11652052
         TrailingForeColor=   11119017
         CurrentDate     =   38261
      End
   End
   Begin VB.PictureBox picOuter 
      BorderStyle     =   0  'None
      Height          =   6540
      Left            =   75
      ScaleHeight     =   6540
      ScaleWidth      =   8265
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1050
      Width           =   8270
      Begin VB.PictureBox picInner 
         BackColor       =   &H00E3E9EB&
         BorderStyle     =   0  'None
         Height          =   6540
         Left            =   0
         ScaleHeight     =   6540
         ScaleWidth      =   8265
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   8270
         Begin VB.TextBox txtAppt 
            Appearance      =   0  'Flat
            Height          =   1365
            Index           =   0
            Left            =   750
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   750
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   23
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   22
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   21
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   20
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   19
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   18
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   17
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   16
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   15
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   14
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   13
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   12
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   11
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   10
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   9
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   8
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   7
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   6
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   5
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   4
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   3
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   2
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   1
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay5 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   0
            Left            =   6750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   23
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   22
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   21
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   20
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   19
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   18
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   17
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   16
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   15
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   14
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   13
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   12
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   11
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   10
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   9
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   8
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   7
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   6
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   5
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   4
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   3
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   2
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   1
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay4 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   0
            Left            =   5250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   23
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   22
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   21
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   20
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   19
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   18
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   17
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   16
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   15
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   14
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   13
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   12
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   11
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   10
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   9
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   8
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   7
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   6
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   5
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   4
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   3
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   2
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   1
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay3 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   0
            Left            =   3750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   23
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   22
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   21
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   20
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   19
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   18
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   17
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   16
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   15
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   14
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   13
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   12
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   11
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   10
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   9
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   8
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   7
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   6
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   5
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   4
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   3
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   2
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   1
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay2 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   0
            Left            =   2250
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   23
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   22
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   21
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   20
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   19
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   18
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   17
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   16
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   15
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   14
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   13
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   12
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   11
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   10
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   9
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   8
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   7
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   6
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   5
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   4
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   3
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   2
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   1
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Shape shpDay1 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   0
            Left            =   750
            Top             =   0
            Width           =   1515
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "11:00p"
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
            Index           =   23
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "10:00p"
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
            Index           =   22
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "9:00p"
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
            Index           =   21
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "8:00p"
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
            Index           =   20
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "7:00p"
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
            Index           =   19
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "6:00p"
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
            Index           =   18
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "5:00p"
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
            Index           =   17
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "4:00p"
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
            Index           =   16
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "3:00p"
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
            Index           =   15
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "2:00p"
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
            Index           =   14
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "1:00p"
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
            Index           =   13
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "12:00p"
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
            Index           =   12
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "11:00a"
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
            Index           =   11
            Left            =   0
            TabIndex        =   23
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "10:00a"
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
            Index           =   10
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "9:00a"
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
            Index           =   9
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "8:00a"
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
            Index           =   8
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "7:00a"
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
            TabIndex        =   19
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "6:00a"
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
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "5:00a"
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
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "4:00a"
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
            TabIndex        =   16
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "3:00a"
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
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "2:00a"
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
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "1:00a"
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
            TabIndex        =   13
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lblHour 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E3E9EB&
            Caption         =   "12:00a"
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
            TabIndex        =   12
            Top             =   75
            Width           =   615
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   23
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   22
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   21
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   20
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   19
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   18
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   17
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   16
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   15
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   14
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   13
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   12
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   11
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   10
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   9
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   8
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   7
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   6
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   5
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   4
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   3
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   2
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   1
            Left            =   0
            Top             =   0
            Width           =   765
         End
         Begin VB.Shape shpHour 
            BackColor       =   &H00E3E9EB&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000F&
            Height          =   690
            Index           =   0
            Left            =   0
            Top             =   0
            Width           =   765
         End
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
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmCalWeek.frx":08EA
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
   Begin MSComCtl2.FlatScrollBar vBarScroll 
      Height          =   6765
      Left            =   8330
      TabIndex        =   36
      Top             =   825
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   11933
      _Version        =   393216
      Orientation     =   1245184
   End
   Begin VB.Image imgUp 
      Height          =   105
      Left            =   8400
      Picture         =   "frmCalWeek.frx":0BF4
      ToolTipText     =   "Go forward one week"
      Top             =   675
      Width           =   105
   End
   Begin VB.Image imgDown 
      Height          =   105
      Left            =   150
      Picture         =   "frmCalWeek.frx":0CDE
      ToolTipText     =   "Go back one week"
      Top             =   675
      Width           =   105
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00CBD3D6&
      X1              =   6825
      X2              =   6825
      Y1              =   825
      Y2              =   1050
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00CBD3D6&
      X1              =   5325
      X2              =   5325
      Y1              =   825
      Y2              =   1050
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00CBD3D6&
      X1              =   3825
      X2              =   3825
      Y1              =   825
      Y2              =   1050
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00CBD3D6&
      X1              =   2325
      X2              =   2325
      Y1              =   825
      Y2              =   1050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00CBD3D6&
      X1              =   825
      X2              =   825
      Y1              =   825
      Y2              =   1050
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B1CBD4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   6825
      TabIndex        =   9
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B1CBD4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   5325
      TabIndex        =   8
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B1CBD4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   3825
      TabIndex        =   7
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B1CBD4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   2325
      TabIndex        =   6
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00B1CBD4&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   825
      TabIndex        =   5
      Top             =   825
      Width           =   1515
   End
   Begin VB.Label lblFill 
      BackColor       =   &H00B1CBD4&
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   825
      Width           =   765
   End
   Begin VB.Label lblHdr 
      Alignment       =   2  'Center
      BackColor       =   &H00C8D0D4&
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
      Width           =   8535
   End
End
Attribute VB_Name = "frmCalWeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsAppt As Recordset 'main recordset
Dim rsList As Recordset

Dim m_vDate As Variant
Dim m_vWeekStart As Variant
Dim m_vWeekEnd As Variant
Dim m_vSaveStartDate As Variant
Dim m_vSaveEndDate As Variant
Dim m_lngContID As Long
Dim m_lngProjID As Long

Private Sub cmdAdd_Click()
   If (Not ValidateEntry()) Then Exit Sub
   
   Call PostEntry
End Sub

Private Sub Combo1_Click(Index As Integer)
   Const sMOD_NAME As String = "frmCalWeek.Combo1_Click"
   On Error GoTo Error_Handler
   
   Dim intCboIndx As Integer
   
   If Index = 0 Then
      intCboIndx = Combo1(0).ListIndex
      Combo1(1).ListIndex = intCboIndx + 2
   End If
   If Index = 2 Then
      m_lngContID = Combo1(2).ItemData(Combo1(2).ListIndex)
   End If
   If Index = 3 Then
      m_lngProjID = Combo1(3).ItemData(Combo1(3).ListIndex)
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub dtpDate_CloseUp(Index As Integer)
   Select Case Index
      Case 0 'from date
         m_vSaveStartDate = dtpDate(0).Value
      Case 1 'to date
         m_vSaveEndDate = dtpDate(1).Value
   End Select
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   
   'set current tab
   tbsMain.Tabs(2).Selected = True
   'set 8:00am position
   vBarScroll.Value = -5448
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmCalWeek.Form_Load"
   On Error GoTo Error_Handler
   
   Dim iCtr As Integer
   
   'set main recordset
   Set rsAppt = dbContact.OpenRecordset("Appts", dbOpenTable)
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Calendar Weekly View", True
   frmMain.picStatus.BackColor = &H4A4A4A
   
   
   'set the current date
   m_vDate = Date
   m_vSaveStartDate = m_vDate
   m_vSaveEndDate = m_vDate
   
   'set both date pickers to current date
   For iCtr = 0 To 1
      dtpDate(iCtr).Value = m_vDate
   Next
   'set calendar to current date
   mvwCal.Value = m_vDate
   
   'flatten all needed ctrls
   FlatBorder Text1.hWnd
   FlatBorder dtpDate(0).hWnd
   FlatBorder dtpDate(1).hWnd
   
   For iCtr = 0 To 3
      FlatBorder Combo1(iCtr).hWnd
   Next iCtr
   
   'setup the date/time grid
   Call AdjustControls
   
   'set the scroll bar
   picInner.Height = 16250
   Call SetScrollBars
   
   'setup screen
   'set a blank space in contact & project combos
   Combo1(2).AddItem " "
   Combo1(3).AddItem " "
   'set date system
   Call CalculateDates
   'add appointments
   Call LoadAppointments
   'setup contact & project combos
   Call LoadContactNames(Combo1(2))
   Call LoadProjectNames(Combo1(3))
   Call SetShapeBorder
   
   'set global from identifier
   g_strFormFlag = "CWk"
   
   Screen.MousePointer = vbDefault
   'MsgBar vbNullString, False
   
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
   
   'adjust calendar ctrl
   picCalSelect.Move vBarScroll.Left + vBarScroll.Width + 75, picBanner.Height + 75, Me.ScaleWidth - picOuter.Width - vBarScroll.Width - 225
   'adjust appt entry items
   picQckAdd.Move picCalSelect.Left, picCalSelect.Top + picCalSelect.Height, picCalSelect.Width, Me.ScaleHeight - picBanner.Height - picCalSelect.Height - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsAppt.Close
   Set rsAppt = Nothing
   
   Set frmCalWeek = Nothing
End Sub

Private Sub AdjustControls()
   Const sMOD_NAME As String = "frmCalWeek.AdjustControls"
   On Error GoTo Error_Handler
   
   Dim iCtr As Integer
   
   For iCtr = 1 To 23
      shpHour(iCtr).Top = shpHour(iCtr - 1).Top + shpHour(iCtr - 1).Height - 15
      shpDay1(iCtr).Top = shpDay1(iCtr - 1).Top + shpDay1(iCtr - 1).Height - 15
      shpDay2(iCtr).Top = shpDay2(iCtr - 1).Top + shpDay2(iCtr - 1).Height - 15
      shpDay3(iCtr).Top = shpDay3(iCtr - 1).Top + shpDay3(iCtr - 1).Height - 15
      shpDay4(iCtr).Top = shpDay4(iCtr - 1).Top + shpDay4(iCtr - 1).Height - 15
      shpDay5(iCtr).Top = shpDay5(iCtr - 1).Top + shpDay5(iCtr - 1).Height - 15
      lblHour(iCtr).Move shpHour(iCtr).Left + 75, shpHour(iCtr).Top + 75
   Next
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while setting up the Calendar Grid!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Public Sub SetScrollBars()
   On Error Resume Next
   
   vBarScroll.Min = 0
   vBarScroll.Max = picOuter.Height - picInner.Height
   vBarScroll.LargeChange = picOuter.ScaleHeight / 7
   vBarScroll.SmallChange = picOuter.ScaleHeight / 9.6
End Sub

Private Sub imgDown_Click()
   m_vDate = DateValue(m_vDate - 7)
   Call CalculateDates
   Call DestroyTextBoxArray
   Call LoadAppointments
End Sub

Private Sub imgUp_Click()
   m_vDate = DateValue(m_vDate + 7)
   Call CalculateDates
   Call DestroyTextBoxArray
   Call LoadAppointments
End Sub

Private Sub mvwCal_DateClick(ByVal DateClicked As Date)
   frmCalDay.m_blnIsSystem = False
   frmCalDay.m_vSrchDate = mvwCal.Value
   UnloadAllForms
   Load frmCalDay
End Sub

Private Sub picBanner_Resize()
   On Error Resume Next
   
   tbsMain.Move picBanner.ScaleWidth - tbsMain.Width
End Sub

Private Sub picCalSelect_Resize()
   On Error Resume Next
   
   mvwCal.Left = (picCalSelect.ScaleWidth - mvwCal.Width) / 2
   mvwCal.Top = (picCalSelect.ScaleHeight - mvwCal.Height) / 2
End Sub

Private Sub picQckAdd_Resize()
   On Error Resume Next
   
   lblQckAdd.Width = picQckAdd.ScaleWidth
   Text1.Width = picQckAdd.ScaleWidth - 840
   dtpDate(0).Width = (picQckAdd.ScaleWidth - 1230) / 2
   Label1(2).Left = dtpDate(0).Left + dtpDate(0).Width + 75
   dtpDate(1).Left = Label1(2).Left + Label1(2).Width + 75
   dtpDate(1).Width = (picQckAdd.ScaleWidth - 1230) / 2
   Combo1(0).Width = dtpDate(0).Width
   Label1(4).Left = Label1(2).Left
   Combo1(1).Left = dtpDate(1).Left
   Combo1(1).Width = dtpDate(1).Width
   Combo1(2).Width = Text1.Width
   Combo1(3).Width = Text1.Width
   cmdAdd.Left = (picQckAdd.ScaleWidth - cmdAdd.Width) / 2
End Sub

Private Sub tbsMain_Click()
   Select Case tbsMain.SelectedItem.Index
      Case 1 'Day
         UnloadAllForms
         frmCalDay.m_blnIsSystem = True
         Load frmCalDay
      Case 2 'Week
         'take no action
      Case 3 'Month
         UnloadAllForms
         Load frmCalMnth
      Case 4 'List
         UnloadAllForms
         Load frmCalList
   End Select
End Sub

Private Sub Text1_Change()
   Combo1(0).Text = "8:00 AM"
End Sub

Private Sub txtAppt_Click(Index As Integer)
   Dim lngApptID As Long
   
   lngApptID = CLng(txtAppt(Index).Tag)
   
   frmAppt.m_lngApptID = lngApptID
   icurState = NOW_EDITING
   Load frmAppt
   frmAppt.Show vbModeless, frmMain
End Sub

Private Sub vBarScroll_Change()
   picInner.Top = vBarScroll.Value
End Sub

Private Sub vBarScroll_Scroll()
   picInner.Top = vBarScroll.Value
End Sub

Sub CalculateDates()
   Const sMOD_NAME As String = "frmCalWeek.CalculateDates"
   On Error GoTo Error_Handler
   
   Dim strWkDay As String
   Dim iCtr As Integer
   
   '*clear previous captions
   lblHdr.Caption = ""
   For iCtr = 0 To 4
      lblDay(iCtr).Caption = ""
   Next
   
   '*get the weekday name of the current date
   strWkDay = WeekdayName(Weekday(m_vDate, vbMonday), True, vbMonday)
   
   '*calculate the start of the 5 day week, and the end of the 5 day week
   Call SetupHeaders(strWkDay)
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while changing dates!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub SetupHeaders(sWkDay As String)
   Const sMOD_NAME As String = "frmCalWeek.SetupHeaders"
   On Error GoTo Error_Handler
   
   Dim iDayCtr As Integer
   Dim Indx As Integer
   Dim iStartMnth As Integer
   Dim iEndMnth As Integer
   Dim dtChkDate As Date
   
   Select Case sWkDay
      Case "Mon"
         '*show monday date
         m_vWeekStart = m_vDate
         lblDay(0) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(0).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show middle 4 days
         For iDayCtr = 1 To 4
            m_vWeekEnd = DateValue(m_vWeekStart + iDayCtr)
            lblDay(iDayCtr) = Format(m_vWeekEnd, "ddd m/dd")
            lblDay(iDayCtr).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         Next iDayCtr
         '*show date in header
         'find out if both dates are in the same month
         iStartMnth = Month(m_vWeekStart)
         iEndMnth = Month(m_vWeekEnd)
         'if months overlap
         If (iStartMnth <> iEndMnth) Then
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "mmm d, yyyy")
         Else
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "d, yyyy")
         End If
      Case "Tue"
         '*show tuesday date
         m_vWeekStart = m_vDate
         lblDay(1) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(1).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show monday date
         m_vWeekStart = DateValue(m_vWeekStart - 1)
         lblDay(0) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(0).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show rest of week
         For iDayCtr = 2 To 4
            m_vWeekEnd = DateValue(m_vWeekStart + iDayCtr)
            lblDay(iDayCtr) = Format(m_vWeekEnd, "ddd m/dd")
            lblDay(iDayCtr).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         Next iDayCtr
         'find out if both dates are in the same month
         iStartMnth = Month(m_vWeekStart)
         iEndMnth = Month(m_vWeekEnd)
         'if months overlap
         If (iStartMnth <> iEndMnth) Then
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "mmm d, yyyy")
         Else
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "d, yyyy")
         End If
      Case "Wed"
         '*show wednesday date
         m_vWeekStart = m_vDate
         lblDay(2) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(2).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show tuesday date
         m_vWeekStart = DateValue(m_vWeekStart - 1)
         lblDay(1) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(1).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show monday date
         m_vWeekStart = DateValue(m_vWeekStart - 1)
         lblDay(0) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(0).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show thursday date
         m_vWeekEnd = DateValue(m_vWeekStart + 3)
         lblDay(3) = Format(m_vWeekEnd, "ddd m/dd")
         lblDay(3).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         '*show friday date
         m_vWeekEnd = DateValue(m_vWeekStart + 4)
         lblDay(4) = Format(m_vWeekEnd, "ddd m/dd")
         lblDay(4).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         'find out if both dates are in the same month
         iStartMnth = Month(m_vWeekStart)
         iEndMnth = Month(m_vWeekEnd)
         'if months overlap
         If (iStartMnth <> iEndMnth) Then
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "mmm d, yyyy")
         Else
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "d, yyyy")
         End If
      Case "Thu"
         '*show thursday date
         m_vWeekStart = m_vDate
         lblDay(3) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(3).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show friday date
         m_vWeekEnd = DateValue(m_vWeekStart + 1)
         lblDay(4) = Format(m_vWeekEnd, "ddd m/dd")
         lblDay(4).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         '*show monday date
         m_vWeekStart = DateValue(m_vWeekEnd - 4)
         lblDay(0) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(0).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show tuesday date
         m_vWeekStart = DateValue(m_vWeekStart + 1)
         lblDay(1) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(1).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         '*show wednesday date
         m_vWeekStart = DateValue(m_vWeekStart + 1)
         lblDay(2) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(2).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         'set start of week date
         m_vWeekStart = DateValue(m_vWeekStart - 2)
         'find out if both dates are in the same month
         iStartMnth = Month(m_vWeekStart)
         iEndMnth = Month(m_vWeekEnd)
         'if months overlap
         If (iStartMnth <> iEndMnth) Then
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "mmm d, yyyy")
         Else
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "d, yyyy")
         End If
      Case "Fri"
         'show monday date
         m_vWeekStart = DateValue(m_vDate - 4)
         lblDay(0) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(0).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         'show tue to fri dates
         For iDayCtr = 1 To 4
            m_vWeekEnd = DateValue(m_vWeekStart + iDayCtr)
            lblDay(iDayCtr) = Format(m_vWeekEnd, "ddd m/dd")
            lblDay(iDayCtr).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         Next iDayCtr
         'find out if both dates are in the same month
         iStartMnth = Month(m_vWeekStart)
         iEndMnth = Month(m_vWeekEnd)
         'if months overlap
         If (iStartMnth <> iEndMnth) Then
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "mmm d, yyyy")
         Else
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "d, yyyy")
         End If
      Case "Sat"
         'show monday date
         m_vWeekStart = DateValue(m_vDate - 5)
         lblDay(0) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(0).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         'show tue to fri dates
         For iDayCtr = 1 To 4
            m_vWeekEnd = DateValue(m_vWeekStart + iDayCtr)
            lblDay(iDayCtr) = Format(m_vWeekEnd, "ddd m/dd")
            lblDay(iDayCtr).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         Next iDayCtr
         'find out if both dates are in the same month
         iStartMnth = Month(m_vWeekStart)
         iEndMnth = Month(m_vWeekEnd)
         'if months overlap
         If (iStartMnth <> iEndMnth) Then
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "mmm d, yyyy")
         Else
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "d, yyyy")
         End If
      Case "Sun"
         'show monday date
         m_vWeekStart = DateValue(m_vDate + 1)
         lblDay(0) = Format(m_vWeekStart, "ddd m/dd")
         lblDay(0).Tag = CStr(Format(m_vWeekStart, "mm/dd/yyyy"))
         'show tue to fri dates
         For iDayCtr = 1 To 4
            m_vWeekEnd = DateValue(m_vWeekStart + iDayCtr)
            lblDay(iDayCtr) = Format(m_vWeekEnd, "ddd m/dd")
            lblDay(iDayCtr).Tag = CStr(Format(m_vWeekEnd, "mm/dd/yyyy"))
         Next iDayCtr
         'find out if both dates are in the same month
         iStartMnth = Month(m_vWeekStart)
         iEndMnth = Month(m_vWeekEnd)
         'if months overlap
         If (iStartMnth <> iEndMnth) Then
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "mmm d, yyyy")
         Else
            lblHdr = Format(m_vWeekStart, "mmm d - ") & Format(m_vWeekEnd, "d, yyyy")
         End If
   End Select
   
   'if the current date is shown in lblDay, make forecolor red
   For Indx = 0 To 4
      dtChkDate = CVDate(lblDay(Indx).Tag)
      If (dtChkDate = m_vDate) Then
         lblDay(Indx).ForeColor = vbRed
      Else
         lblDay(Indx).ForeColor = vbBlack
      End If
   Next Indx
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while setting up the Calendar!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Public Sub LoadAppointments()
   'lookup the appts between m_vWeekStart and m_vWeekEnd
   Const sMOD_NAME As String = "frmCalWeek.LoadAppointments"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim vStart As Variant
   Dim vEnd As Variant
   Dim vApptDate As Variant
   Dim strRefID As String
   Dim strTimeFrom As String
   Dim strTimeTo As String
   Dim strShowTime As String
   Dim strSubject As String
   Dim strContact As String
   Dim strProject As String
   Dim strText As String
   
   vStart = "#" & m_vWeekStart & "#"
   vEnd = "#" & m_vWeekEnd & "#"
   
   SQL = "SELECT RefNum, fkContID, fkProjID, Subject, DateFrom, "
   SQL = SQL & "DateTo, TimeFrom, TimeTo FROM Appts "
   SQL = SQL & "WHERE DateFrom BETWEEN " & vStart
   SQL = SQL & " AND " & vEnd
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then strRefID = CStr(!RefNum)
            If (Not IsNull(!DateFrom)) Then vApptDate = !DateFrom
            If (Not IsNull(!TimeFrom)) Then
               strTimeFrom = CStr(Format(!TimeFrom, "h:nna/p"))
            End If
            If (Not IsNull(!TimeTo)) Then
               strTimeTo = CStr(Format(!TimeTo, "h:nna/p"))
            End If
            If (Not IsNull(!Subject)) Then strSubject = !Subject
            If (!fkContID > 0) Then
               strContact = ConvertContactName(!fkContID)
            End If
            If (!fkProjID > 0) Then
               strProject = ConvertProjectName(!fkProjID)
            End If
            'setup text for appt text
            strText = strTimeFrom & " - " & strTimeTo & vbCrLf
            strText = strText & strSubject
            If (strContact <> "") Then
               strText = strText & vbCrLf & strContact
            End If
            If (strProject <> "") Then
               strText = strText & vbCrLf & strProject
            End If
            'add send code
            Call SetAppointmentText(vApptDate, !TimeFrom, !TimeTo, strRefID, strText)
            'clear variables for next use
            strRefID = ""
            strTimeFrom = ""
            strTimeTo = ""
            strSubject = ""
            strContact = ""
            strProject = ""
            strText = ""
            
            MsgBar "There are " & .RecordCount & " appointments listed for this date range.", False
            .MoveNext
         Wend
      Else
         MsgBar vbNullString, False
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the Appoinments!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub SetAppointmentText(vSetDate As Variant, vStartTime As Variant, _
                                 vEndTime As Variant, sApptID As String, _
                                 sApptText As String)
   'load the appointment in the calendar grid
   Const sMOD_NAME As String = "frmCalWeek.SetAppointmentText"
   On Error GoTo Error_Handler
   
   Dim intWkday As Integer
   Dim intRowStart As Integer
   Dim strStartTime As String
   Dim intRowEnd As Integer
   Dim strEndTime As String
   Dim intNewIndex As Integer
   
   'check to see if date is mon-fri, if not exit sub
   intWkday = Weekday(vSetDate - 1)
   
   'get row start position, (lblHour.Index)
   strStartTime = CStr(Format(vStartTime, "h:00a/p"))
   intRowStart = GetRowStart(strStartTime)
   
   'get row end position, (lblHour.Index)
   strEndTime = CStr(Format(vEndTime, "h:00a/p"))
   intRowEnd = GetRowEnd(strEndTime)
   
   'set proper column
   Select Case intWkday
      Case 1 'mon, col 1
            intNewIndex = txtAppt.UBound + 1
            Load txtAppt(intNewIndex)
            'txtAppt(intNewIndex).Move shpDay1(intRowStart).Left, shpDay1(intRowStart).Top, shpDay1(intRowStart).Width, shpDay1(intRowEnd + 1).Top - shpDay1(intRowStart).Top
            txtAppt(intNewIndex).Left = shpDay1(intRowStart).Left
            If Minute(vStartTime) = 30 Then
               txtAppt(intNewIndex).Top = shpDay1(intRowStart).Top + (shpDay1(intRowStart).Height / 2)
            Else
               txtAppt(intNewIndex).Top = shpDay1(intRowStart).Top
            End If
            txtAppt(intNewIndex).Width = shpDay1(intRowStart).Width
            If Minute(vEndTime) = 30 Then
               txtAppt(intNewIndex).Height = (shpDay1(intRowEnd + 1).Top - txtAppt(intNewIndex).Top) - 345
            Else
               txtAppt(intNewIndex).Height = shpDay1(intRowEnd).Top - txtAppt(intNewIndex).Top
            End If
            txtAppt(intNewIndex).Visible = True
            txtAppt(intNewIndex).Tag = sApptID
            txtAppt(intNewIndex).Text = sApptText
      Case 2 'tue, col 2
            intNewIndex = txtAppt.UBound + 1
            Load txtAppt(intNewIndex)
            'txtAppt(intNewIndex).Move shpDay2(intRowStart).Left, shpDay2(intRowStart).Top, shpDay2(intRowStart).Width, shpDay2(intRowEnd + 1).Top - shpDay2(intRowStart).Top
            txtAppt(intNewIndex).Left = shpDay2(intRowStart).Left
            If Minute(vStartTime) = 30 Then
               txtAppt(intNewIndex).Top = shpDay2(intRowStart).Top + (shpDay2(intRowStart).Height / 2)
            Else
               txtAppt(intNewIndex).Top = shpDay2(intRowStart).Top
            End If
            txtAppt(intNewIndex).Width = shpDay2(intRowStart).Width
            If Minute(vEndTime) = 30 Then
               txtAppt(intNewIndex).Height = (shpDay2(intRowEnd + 1).Top - txtAppt(intNewIndex).Top) - 345
            Else
               txtAppt(intNewIndex).Height = shpDay2(intRowEnd).Top - txtAppt(intNewIndex).Top
            End If
            txtAppt(intNewIndex).Visible = True
            txtAppt(intNewIndex).Tag = sApptID
            txtAppt(intNewIndex).Text = sApptText
      Case 3 'wed, col 3
            intNewIndex = txtAppt.UBound + 1
            Load txtAppt(intNewIndex)
            'txtAppt(intNewIndex).Move shpDay3(intRowStart).Left, shpDay3(intRowStart).Top, shpDay3(intRowStart).Width, shpDay3(intRowEnd + 1).Top - shpDay3(intRowStart).Top
            txtAppt(intNewIndex).Left = shpDay3(intRowStart).Left
            If Minute(vStartTime) = 30 Then
               txtAppt(intNewIndex).Top = shpDay3(intRowStart).Top + (shpDay3(intRowStart).Height / 2)
            Else
               txtAppt(intNewIndex).Top = shpDay3(intRowStart).Top
            End If
            txtAppt(intNewIndex).Width = shpDay3(intRowStart).Width
            If Minute(vEndTime) = 30 Then
               txtAppt(intNewIndex).Height = (shpDay3(intRowEnd + 1).Top - txtAppt(intNewIndex).Top) - 345
            Else
               txtAppt(intNewIndex).Height = shpDay3(intRowEnd).Top - txtAppt(intNewIndex).Top
            End If
            txtAppt(intNewIndex).Visible = True
            txtAppt(intNewIndex).Tag = sApptID
            txtAppt(intNewIndex).Text = sApptText
      Case 4 'thu, col 4
            intNewIndex = txtAppt.UBound + 1
            Load txtAppt(intNewIndex)
            'txtAppt(intNewIndex).Move shpDay4(intRowStart).Left, shpDay4(intRowStart).Top, shpDay4(intRowStart).Width, shpDay4(intRowEnd + 1).Top - shpDay4(intRowStart).Top
            txtAppt(intNewIndex).Left = shpDay4(intRowStart).Left
            If Minute(vStartTime) = 30 Then
               txtAppt(intNewIndex).Top = shpDay4(intRowStart).Top + (shpDay4(intRowStart).Height / 2)
            Else
               txtAppt(intNewIndex).Top = shpDay4(intRowStart).Top
            End If
            txtAppt(intNewIndex).Width = shpDay4(intRowStart).Width
            If Minute(vEndTime) = 30 Then
               txtAppt(intNewIndex).Height = (shpDay4(intRowEnd + 1).Top - txtAppt(intNewIndex).Top) - 345
            Else
               txtAppt(intNewIndex).Height = shpDay4(intRowEnd).Top - txtAppt(intNewIndex).Top
            End If
            txtAppt(intNewIndex).Visible = True
            txtAppt(intNewIndex).Tag = sApptID
            txtAppt(intNewIndex).Text = sApptText
      Case 5 'fri, col 5
            intNewIndex = txtAppt.UBound + 1
            Load txtAppt(intNewIndex)
            'txtAppt(intNewIndex).Move shpDay5(intRowStart).Left, shpDay5(intRowStart).Top, shpDay5(intRowStart).Width, shpDay5(intRowEnd + 1).Top - shpDay5(intRowStart).Top
            txtAppt(intNewIndex).Left = shpDay5(intRowStart).Left
            If Minute(vStartTime) = 30 Then
               txtAppt(intNewIndex).Top = shpDay5(intRowStart).Top + (shpDay5(intRowStart).Height / 2)
            Else
               txtAppt(intNewIndex).Top = shpDay5(intRowStart).Top
            End If
            txtAppt(intNewIndex).Width = shpDay5(intRowStart).Width
            If Minute(vEndTime) = 30 Then
               txtAppt(intNewIndex).Height = (shpDay5(intRowEnd + 1).Top - txtAppt(intNewIndex).Top) - 345
            Else
               txtAppt(intNewIndex).Height = shpDay5(intRowEnd).Top - txtAppt(intNewIndex).Top
            End If
            txtAppt(intNewIndex).Visible = True
            txtAppt(intNewIndex).Tag = sApptID
            txtAppt(intNewIndex).Text = sApptText
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Function GetRowStart(sStTime As String) As Integer
   'find the appropriate starting row
   Dim Indx As Integer
   
   For Indx = 0 To 23
      If lblHour(Indx).Caption = sStTime Then
         GetRowStart = Indx
         Exit For
      End If
   Next Indx
End Function

Private Function GetRowEnd(sEndTime As String) As Integer
   'find the appropriate ending row
   Dim Indx As Integer
   
   For Indx = 0 To 23
      If lblHour(Indx).Caption = sEndTime Then
         GetRowEnd = Indx
         Exit For
      End If
   Next Indx
End Function

Public Sub DestroyTextBoxArray()
   'remove all textboxes
   Dim intAryIndex As Integer
   
   intAryIndex = txtAppt.UBound
   
   Do While intAryIndex > 0
      Unload txtAppt(intAryIndex)
      intAryIndex = intAryIndex - 1
   Loop
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(Text1) < 1) Then
      Indx = MsgBox("You Must Enter An Appointment Subject", _
         vbInformation + vbOKOnly, "Validate : Appointment Subject")
      Text1.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmCalWeek.PostEntry"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting To Do Entry", True
   
   rsAppt.AddNew
   
   With rsAppt
      If (Len(Text1)) Then !Subject = Text1
      
      !DateFrom = m_vSaveStartDate
      !DateTo = m_vSaveEndDate
      
      If (m_lngContID > 0) Then !fkContID = m_lngContID
      If (m_lngProjID > 0) Then !fkProjID = m_lngProjID
      
      If (Len(Combo1(0).Text)) Then
            !TimeFrom = Format(Combo1(0).Text, "hh:nn AMPM")
      End If
      If (Len(Combo1(1).Text)) Then
            !TimeTo = Format(Combo1(1).Text, "hh:nn AMPM")
      End If
      
      .Update
   End With
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Call ClearControls
   'add appointments
   Call LoadAppointments
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   MsgBox "An un-known error occurred while Posting this entry!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub ClearControls()
   Text1.Text = ""
   dtpDate(0).Value = m_vDate
   dtpDate(1).Value = m_vDate
   Combo1(0).Text = "8:00 AM"
   Combo1(1).Text = "9:00 AM"
   Combo1(2).Text = " "
   Combo1(3).Text = " "
   
   Text1.SetFocus
End Sub

Private Sub SetShapeBorder()
   'set all shape borders to gray (to show up in XP)
   Dim iCtr As Integer
   
   For iCtr = 0 To 23
      shpHour(iCtr).BorderColor = &HCBD3D6
      shpDay1(iCtr).BorderColor = &HCBD3D6
      shpDay2(iCtr).BorderColor = &HCBD3D6
      shpDay3(iCtr).BorderColor = &HCBD3D6
      shpDay4(iCtr).BorderColor = &HCBD3D6
      shpDay5(iCtr).BorderColor = &HCBD3D6
   Next
End Sub
