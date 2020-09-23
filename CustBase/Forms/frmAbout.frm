VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00F3F3ED&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4065
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6540
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4875
      TabIndex        =   2
      Top             =   3525
      Width           =   1440
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   240
      Left            =   150
      TabIndex        =   3
      Top             =   3675
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CBD3D6&
      Caption         =   $"frmAbout.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   990
      Left            =   150
      TabIndex        =   1
      Top             =   2400
      Width           =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   150
      X2              =   6375
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   2
      X1              =   150
      X2              =   6375
      Y1              =   2325
      Y2              =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00CBD3D6&
      Caption         =   "Customer Base - Contact Manager"
      BeginProperty Font 
         Name            =   "Bad"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00696969&
      Height          =   540
      Left            =   150
      TabIndex        =   0
      Top             =   1725
      Width           =   6165
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   825
      Picture         =   "frmAbout.frx":0195
      Top             =   85
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00CBD3D6&
      BackStyle       =   1  'Opaque
      Height          =   3915
      Left            =   75
      Top             =   75
      Width           =   6390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
   lblVersion = "Version / Build " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmAbout = Nothing
End Sub
