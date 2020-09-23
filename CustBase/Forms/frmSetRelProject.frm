VERSION 5.00
Begin VB.Form frmSetRelProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Related Project"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetRelProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   2625
      TabIndex        =   1
      Top             =   4125
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   4125
      Width           =   1215
   End
   Begin VB.ListBox lstProject 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   75
      TabIndex        =   2
      Top             =   375
      Width           =   3765
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   3825
      Y1              =   3885
      Y2              =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   75
      X2              =   3825
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
      Caption         =   "All default Projects listed"
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
      Top             =   75
      Width           =   3765
   End
End
Attribute VB_Name = "frmSetRelProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsRelProj As Recordset 'main recordset
Dim rsList As Recordset

Dim m_blnCancelled As Boolean
Dim m_lngProjID As Long

Private Sub cmdOpts_Click(Index As Integer)
   Select Case Index
      Case 0 'OK
         If (m_lngProjID = 0) Then
            MsgBox "You must select a Project from the list.", , APP_MSG_NAME
            Exit Sub
         End If
         
         Call PostEntry
      Case 1 'Cancel
         m_blnCancelled = True
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmSetRelProject.Form_Load"
   On Error GoTo Error_Handler
   
   'set  main recordset
   Set rsRelProj = dbContact.OpenRecordset("RelateProject", dbOpenTable)
   
   'load the list
   Call LoadAllProjects
   
   m_blnCancelled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsRelProj.Close
   Set rsRelProj = Nothing
   
   If (m_blnCancelled = False) Then
      Call frmContEntry.LoadRelatedProjects
   End If
   
   Set frmSetRelProject = Nothing
End Sub

Private Sub LoadAllProjects()
   'load all projects listed in the databse
   Const sMOD_NAME As String = "frmSetRelProject.Form_Load"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT ProjID, PName FROM Projects "
   SQL = SQL & "ORDER BY PName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ProjID)) Then
               If (Not IsNull(!PName)) Then lstProject.AddItem !PName
               lstProject.ItemData(lstProject.NewIndex) = !ProjID
            End If
            .MoveNext
         Wend
      Else
         MsgBox "There are no Projects entered to select.", , APP_MSG_NAME
         Unload Me
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub lstProject_Click()
   Const sMOD_NAME As String = "frmSetRelProject.lstProject_Click"
   On Error GoTo Error_Handler
   
   m_lngProjID = lstProject.ItemData(lstProject.ListIndex)
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while selecting a Project Name!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub PostEntry()
   'post the project relation to the database
   Const sMOD_NAME As String = "frmSetRelProject.PostEntry"
   On Error GoTo Error_Handler
   
   rsRelProj.AddNew
   
   With rsRelProj
      !fkProjID = m_lngProjID
      !fkContID = g_lngContID
      !ContShowName = frmContEntry.Text1(1).Text
      
      .Update
   End With
   
   Me.Hide
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting this record!" & vbCrLf & _
      "Sorry for the inconvenience", , APP_MSG_NAME
   Unload Me
End Sub
