VERSION 5.00
Begin VB.Form frmReminder 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5970
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4200
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   2265
      Index           =   5
      Left            =   1050
      TabIndex        =   15
      Top             =   2775
      Width           =   3015
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   4
      Left            =   1050
      TabIndex        =   14
      Top             =   2475
      Width           =   3015
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   1050
      TabIndex        =   13
      Top             =   2175
      Width           =   3015
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   1050
      TabIndex        =   12
      Top             =   1875
      Width           =   3015
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   1050
      TabIndex        =   11
      Top             =   1575
      Width           =   3015
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   1050
      TabIndex        =   10
      Top             =   1275
      Width           =   3015
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   240
      Index           =   5
      Left            =   150
      TabIndex        =   9
      Top             =   2775
      Width           =   840
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   240
      Index           =   4
      Left            =   150
      TabIndex        =   8
      Top             =   2475
      Width           =   840
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Due:"
      Height          =   240
      Index           =   3
      Left            =   150
      TabIndex        =   7
      Top             =   2175
      Width           =   840
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   240
      Index           =   2
      Left            =   150
      TabIndex        =   6
      Top             =   1875
      Width           =   840
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Project:"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   1575
      Width           =   840
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact:"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   1275
      Width           =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DBFDFF&
      Caption         =   " Details:"
      ForeColor       =   &H00896A4B&
      Height          =   240
      Left            =   225
      TabIndex        =   3
      Top             =   1020
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00896A4B&
      Height          =   3990
      Left            =   75
      Top             =   1125
      Width           =   4055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      TabIndex        =   2
      Top             =   675
      Width           =   4065
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2700
      TabIndex        =   1
      Top             =   5455
      Width           =   1290
   End
   Begin VB.Shape shpClose 
      Height          =   315
      Left            =   2625
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   1440
   End
   Begin VB.Line Line1 
      X1              =   75
      X2              =   4125
      Y1              =   5325
      Y2              =   5325
   End
   Begin VB.Label lblBanner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Base - Contact Manager Reminder"
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
      Left            =   150
      TabIndex        =   0
      Top             =   50
      Width           =   3915
   End
   Begin VB.Image imgMain 
      Height          =   5985
      Left            =   -10
      Picture         =   "frmReminder.frx":0000
      Top             =   0
      Width           =   4230
   End
End
Attribute VB_Name = "frmReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Public m_lngRemindID As Long
Public m_lngToDoID As Long
Public m_lngApptID As Long
Public m_strType As String

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmReminder.Form_Load"
   On Error GoTo Error_Handler
   
   'set this form on top of all others
   SetWindowPos frmReminder.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
   
   Call InitializeScreen
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Const sMOD_NAME As String = "frmReminder.Form_Unload"
   On Error GoTo Error_Handler
   
   'mark this reminder as complete
   Call MarkAsComplete
   'remove tray icon
   Call RemoveFromTray
   SetWindowPos frmReminder.hWnd, -2, 0, 0, 0, 0, &H1 Or &H2
   Me.Hide
   
   frmMain.Show
   
   If (m_strType = "TD") Then
      icurState = NOW_EDITING
      frmToDo.m_lngToDoID = m_lngToDoID
      Load frmToDo
      frmToDo.Show vbModeless, frmMain
   ElseIf (m_strType = "AP") Then
      icurState = NOW_EDITING
      frmAppt.m_lngApptID = m_lngApptID
      Load frmAppt
      frmAppt.Show vbModeless, frmMain
   End If
   
   Set frmReminder = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub lblClose_Click()
   Unload Me
End Sub

Public Sub InitializeScreen()
   'setup the opening screen
   Const sMOD_NAME As String = "frmReminder.InitializeScreen"
   On Error GoTo Error_Handler
   
   Dim lngSound As Long
   
   Select Case m_strType
      Case "AP" 'appointment
         Call SetupApptScreen
      Case "TD" 'to do
         lblCaption(4).Visible = False
         Call SetupToDoScreen
   End Select
   
   lngSound = sndPlaySound(App.Path & "\alert.wav", 1)
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub SetupToDoScreen()
   'setup the screen for a To Do reminder
   Dim SQL As String
   Dim strContact As String
   Dim strProject As String
   Dim strDueDate As String
   Dim strDueTime As String
   
   SQL = "SELECT RefNum, Subject, fkContID, fkProjID, DueDate, "
   SQL = SQL & "DueTime, TextBody FROM ToDo "
   SQL = SQL & "WHERE RefNum = " & m_lngToDoID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!fkContID)) Then
            strContact = ConvertContactName(!fkContID)
            lblText(0) = strContact
         End If
         If (Not IsNull(!fkProjID)) Then
            strProject = ConvertProjectName(!fkProjID)
            lblText(1) = strProject
         End If
         If (Not IsNull(!Subject)) Then lblText(2) = !Subject
         If (Not IsNull(!DueDate)) Then strDueDate = Format(!DueDate, "m/d/yyyy")
         If (Not IsNull(!DueTime)) Then strDueTime = Format(!DueTime, "h:nn AMPM")
         lblText(3) = strDueDate & " at " & strDueTime
         If (Not IsNull(!TextBody)) Then lblText(5) = !TextBody
      End If
   End With
   
   lblTitle.Caption = "Reminder for scheduled To Do item"
   
   rsList.Close
   Set rsList = Nothing
End Sub

Private Sub MarkAsComplete()
   'mark the current reminder as completed
   Dim SQL As String
   
   SQL = "SELECT RefNum, Completed FROM Remind "
   SQL = SQL & "WHERE RefNum = " & m_lngRemindID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   'rsList.Edit
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         .Edit
         !Completed = True
         
         .Update
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
End Sub

Private Sub SetupApptScreen()
   'setup the screen for an Appt reminder
   Dim SQL As String
   Dim strContact As String
   Dim strProject As String
   Dim strDueDate As String
   Dim strDueTime As String
   Dim strToDate As String
   Dim strToTime As String
   
   SQL = "SELECT RefNum, Subject, fkContID, fkProjID, DateFrom, "
   SQL = SQL & "DateTo, TimeFrom, TimeTo, TextBody FROM Appts "
   SQL = SQL & "WHERE RefNum = " & m_lngApptID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!fkContID)) Then
            strContact = ConvertContactName(!fkContID)
            lblText(0) = strContact
         End If
         If (Not IsNull(!fkProjID)) Then
            strProject = ConvertProjectName(!fkProjID)
            lblText(1) = strProject
         End If
         If (Not IsNull(!Subject)) Then lblText(2) = !Subject
         If (Not IsNull(!DateFrom)) Then strDueDate = Format(!DateFrom, "m/d/yyyy")
         If (Not IsNull(!TimeFrom)) Then strDueTime = Format(!TimeFrom, "h:nn AMPM")
         lblText(3) = strDueDate & " at " & strDueTime
         If (Not IsNull(!DateTo)) Then strToDate = Format(!DateTo, "m/d/yyyy")
         If (Not IsNull(!TimeTo)) Then strToTime = Format(!TimeTo, "h:nn AMPM")
         lblText(4) = strToDate & " at " & strToTime
         If (Not IsNull(!TextBody)) Then lblText(5) = !TextBody
      End If
   End With
   
   lblTitle.Caption = "Reminder for scheduled Appointment"
   
   rsList.Close
   Set rsList = Nothing
End Sub
