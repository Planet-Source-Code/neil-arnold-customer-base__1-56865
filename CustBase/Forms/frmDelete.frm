VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDelete 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   990
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4665
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
   ScaleHeight     =   990
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   75
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   525
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   4200
      Top             =   525
   End
   Begin MSComctlLib.ProgressBar prbProg 
      Height          =   240
      Left            =   150
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   " Preparing To Delete Files ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   4365
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_intType As Integer '1 = Contact, 2 = Project

Private Sub Form_Activate()
   Const sMOD_NAME As String = "frmDelete.Form_Activate"
   On Error GoTo Error_Handler
   
   Timer1.Enabled = True
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmDelete = Nothing
End Sub

Private Sub DeleteContact()
   'delete the select contact from the database
   Const sMOD_NAME As String = "frmDelete.DeleteContact"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Deleting Contact Information", True
   prbProg.Visible = True
   
   Dim SQL As String
   
   'delete contact
   lblProgress.Caption = "Deleting Contact file ..."
   prbProg.Value = 9
   SQL = "DELETE * FROM Contacts WHERE ContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact address
   lblProgress.Caption = "Deleting Contact Address files ..."
   prbProg.Value = 18
   SQL = "DELETE * FROM CAddress WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact phone numbers
   lblProgress.Caption = "Deleting Phone Number files ..."
   prbProg.Value = 27
   SQL = "DELETE * FROM CPhone WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact comments
   lblProgress.Caption = "Deleting Contact Comment files ..."
   prbProg.Value = 36
   SQL = "DELETE * FROM CComments WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact user fields
   lblProgress.Caption = "Deleting Contact User Field files ..."
   prbProg.Value = 45
   SQL = "DELETE * FROM CUFldValues WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact notes & calls
   lblProgress.Caption = "Deleting Contact Notes/Calls files ..."
   prbProg.Value = 54
   SQL = "DELETE * FROM Attach WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact ToDo's
   lblProgress.Caption = "Deleting Contact To Do files ..."
   prbProg.Value = 63
   SQL = "DELETE * FROM ToDo WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact appts
   lblProgress.Caption = "Deleting Contact Appointment files ..."
   prbProg.Value = 72
   SQL = "DELETE * FROM Appts WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact email's
   lblProgress.Caption = "Deleting Contact E-Mail files ..."
   prbProg.Value = 81
   SQL = "DELETE * FROM CEMail WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact related contact info
   lblProgress.Caption = "Deleting Related Contact files ..."
   prbProg.Value = 90
   SQL = "DELETE * FROM RelateCont WHERE MasterContID = " & g_lngContID
   dbContact.Execute (SQL)
   SQL = "DELETE * FROM RelateCont WHERE SubContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   'delete contact related project info
   lblProgress.Caption = "Deleting Related Project files ..."
   prbProg.Value = 100
   SQL = "DELETE * FROM RelateProject WHERE fkContID = " & g_lngContID
   dbContact.Execute (SQL)
   
   Screen.MousePointer = vbDefault
   MsgBar "Contact Sucessfully Deleted", False
   
   MsgBox "Delete Complete!", , APP_MSG_NAME
   
   MsgBar vbNullString, False
   
   Me.Hide
   
   g_lngContID = 0
   UnloadAllForms
   Load frmHome
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   Unload Me
End Sub

Private Sub DeleteProject()
   'delete the select contact from the database
   Const sMOD_NAME As String = "frmDelete.DeleteProject"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Deleting Project Information", True
   prbProg.Visible = True
   
   Dim SQL As String
   
   'delete project
   lblProgress.Caption = "Deleting Project file ..."
   prbProg.Value = 14
   SQL = "DELETE * FROM Projects WHERE ProjID = " & g_lngProjID
   dbContact.Execute (SQL)
   
   'delete project comments
   lblProgress.Caption = "Deleting Project Comment files ..."
   prbProg.Value = 28
   SQL = "DELETE * FROM PComments WHERE fkProjID = " & g_lngProjID
   dbContact.Execute (SQL)
   
   'delete project user fields
   lblProgress.Caption = "Deleting Project User Field files ..."
   prbProg.Value = 42
   SQL = "DELETE * FROM PUFldValues WHERE fkProjID = " & g_lngProjID
   dbContact.Execute (SQL)
   
   'delete project notes & calls
   lblProgress.Caption = "Deleting Project Notes/Calls files ..."
   prbProg.Value = 56
   SQL = "DELETE * FROM Attach WHERE fkProjID = " & g_lngProjID
   dbContact.Execute (SQL)
   
   'delete project ToDo's
   lblProgress.Caption = "Deleting Project To Do files ..."
   prbProg.Value = 70
   SQL = "DELETE * FROM ToDo WHERE fkProjID = " & g_lngProjID
   dbContact.Execute (SQL)
   
   'delete project appts
   lblProgress.Caption = "Deleting Project Appointment files ..."
   prbProg.Value = 84
   SQL = "DELETE * FROM Appts WHERE fkProjID = " & g_lngProjID
   dbContact.Execute (SQL)
   
   'delete Project related project info
   lblProgress.Caption = "Deleting Related Project files ..."
   prbProg.Value = 100
   SQL = "DELETE * FROM RelateProject WHERE fkProjID = " & g_lngProjID
   dbContact.Execute (SQL)
   
   Screen.MousePointer = vbDefault
   MsgBar "Project Sucessfully Deleted", False
   
   MsgBox "Delete Complete!", , APP_MSG_NAME
   
   MsgBar vbNullString, False
   
   Me.Hide
   
   g_lngProjID = 0
   UnloadAllForms
   Load frmHome
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   Unload Me
End Sub

Private Sub Timer1_Timer()
   Text1.Text = Text1.Text + 1
   
   If (Text1.Text = 2) Then
      Select Case m_intType
         Case 1 'contact
            Call DeleteContact
         Case 2 'project
            Call DeleteProject
      End Select
   End If
End Sub
