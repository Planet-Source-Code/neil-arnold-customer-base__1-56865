VERSION 5.00
Begin VB.Form frmNewProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Up New Project"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   4275
      TabIndex        =   3
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   4275
      TabIndex        =   2
      Top             =   675
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   750
      MaxLength       =   255
      TabIndex        =   1
      Top             =   225
      Width           =   3390
   End
   Begin VB.Label Label2 
      Caption         =   "Example : Customer Audit"
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
      Height          =   165
      Index           =   2
      Left            =   1575
      TabIndex        =   4
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Project:"
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   225
      Width           =   615
   End
End
Attribute VB_Name = "frmNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim rsProj As Recordset 'main recordset
Dim rsList As Recordset 'all other data work

Dim m_strProjName As String 'project name

Dim m_lngNewID As Long 'for new contact id

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmNewProject.cmdOpt_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'Next>
         If (Not ValidateEntry()) Then Exit Sub
         
         Call PostEntry
      Case 1 'cancel
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmNewProject.Form_Load"
   On Error GoTo Error_Handler
   
   Set rsProj = dbContact.OpenRecordset("Projects", dbOpenTable)
   
   'flatten all needed borders
   FlatBorder Text1.hWnd
   
   Call GetNewProjID
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsProj.Close
   Set rsProj = Nothing
   
   Set frmNewProject = Nothing
End Sub

Private Sub GetNewProjID()
   'create a new contact ID
   Const sMOD_NAME As String = "frmNewProject.GetNewProjID"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT MAX(ProjID)AS MAXID FROM Projects"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!MAXID)) Then
            m_lngNewID = !MAXID + 1
         Else
            m_lngNewID = 1
         End If
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub Text1_GotFocus()
   highLight
End Sub

Private Function ValidateEntry() As Boolean
   'make sure some text was entered
   ValidateEntry = True
   
   If (Len(Text1) < 1) Then
      MsgBox "You must enter a new project name.", _
         vbInformation + vbOKOnly, "Validate : Project Name"
      Text1.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   'post the new record to the database
   Const sMOD_NAME As String = "frmNewProject.PostEntry"
   On Error GoTo Error_Handler
   
   Dim strSetting As String
   
   strSetting = "Default"
   
   rsProj.AddNew
   
   With rsProj
      !ProjID = m_lngNewID
      !Setting = strSetting
      
      If (Len(Text1)) Then !PName = Text1
      
      .Update
   End With
   
   g_lngProjID = m_lngNewID
   
   Me.Hide
   
   UnloadAllForms
   Load frmProjEntry
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

