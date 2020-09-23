VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPassLogon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Logon"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPassLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   225
      Top             =   1575
   End
   Begin MSComctlLib.ProgressBar prbTimer 
      Height          =   240
      Left            =   75
      TabIndex        =   10
      Top             =   2475
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   60
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   4725
      TabIndex        =   8
      Top             =   1425
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Height          =   390
      Index           =   0
      Left            =   4725
      TabIndex        =   7
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1425
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1275
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1425
      MaxLength       =   50
      TabIndex        =   4
      Top             =   900
      Width           =   2940
   End
   Begin VB.Label lblTimeTicker 
      Height          =   240
      Left            =   75
      TabIndex        =   9
      Top             =   2175
      Width           =   5940
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "[Note: This password Verification Is Case-Sensitive]"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   1
      Left            =   1425
      TabIndex        =   6
      Top             =   1575
      Width           =   2940
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Password:"
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   3
      Top             =   1275
      Width           =   1290
   End
   Begin VB.Label Label3 
      Caption         =   "Enter User Name:"
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   900
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "[Note: You have 60 seconds to login to the system with the correct password, if un-successful the program will terminate]"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   375
      Width           =   5940
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Customer Base - User Login"
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
      TabIndex        =   0
      Top             =   75
      Width           =   5940
   End
End
Attribute VB_Name = "frmPassLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_strUserName As String 'entered user name
Dim m_strExistPwd As String 'existing password
Dim m_strEntPwd As String 'entered password
Dim m_blnPassed As Boolean 'True = correct password was entered
Dim m_intSeconds As Integer 'for timer

Private Sub cmdOpts_Click(Index As Integer)
   Select Case Index
      Case 0 'OK
         'check for nulls, user name not found would open with empty string
         If ((m_strExistPwd = "") Or (m_strEntPwd = "")) Then Exit Sub
         
         If (StrComp(m_strExistPwd, m_strEntPwd) <> 0) Then
            MsgBox "In-Valid Password, Please Re-Enter!", vbCritical + vbOKOnly, _
               "Password Did Not Match"
            Text1(1).SetFocus
            Exit Sub
         Else
            frmMain.m_blnEmerShutdown = False
            m_blnPassed = True
            Unload Me
         End If
      Case 1 'cancel
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmPassLogon.Form_Load"
   On Error GoTo Error_Handler
   
   'flatten all necessary items
   FlatBorder Text1(0).hWnd
   FlatBorder Text1(1).hWnd
   
   'start timer
   Timer1.Enabled = True
   
   'set passed toggle
   m_blnPassed = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (m_blnPassed = False) Then
      frmMain.m_blnEmerShutdown = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmPassLogon = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   highLight
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Const sMOD_NAME As String = "frmPassLogon.Text1_LostFocus"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'user name
         If (Text1(0).Text = "") Then Exit Sub
         
         m_strUserName = Text1(0).Text
         Call GetUserPassword
      Case 1 'password
         If (Text1(1).Text = "") Then Exit Sub
         
         m_strEntPwd = Text1(1).Text
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub Timer1_Timer()
   Dim intTimeRemain As Integer
   
   m_intSeconds = m_intSeconds + 1
   
   intTimeRemain = 60 - m_intSeconds
   
   prbTimer.Value = m_intSeconds
   lblTimeTicker.Caption = "Time remaining is " & intTimeRemain & " seconds."
   
   If (m_intSeconds = 60) Then
      m_blnPassed = False
      Unload Me
   End If
End Sub

Private Sub GetUserPassword()
   'get the password for the current user
   Const sMOD_NAME As String = "frmPassLogon.GetUserPassword"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT UserName, Password FROM Security "
   SQL = SQL & "WHERE UserName = '" & m_strUserName & "' "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!Password)) Then
            m_strExistPwd = Base64Decode(!Password)
         End If
      Else
         MsgBox "In-Correct User Name, Please Re-Enter.", , "User Name Not On File"
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub
