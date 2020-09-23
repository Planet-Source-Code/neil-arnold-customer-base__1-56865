VERSION 5.00
Begin VB.Form frmSetRelContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Related Contact"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetRelContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstContact 
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
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   3825
      Y1              =   3885
      Y2              =   3885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
      Caption         =   "All default Contacts listed"
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   75
      X2              =   3825
      Y1              =   3900
      Y2              =   3900
   End
End
Attribute VB_Name = "frmSetRelContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsRelCont As Recordset 'main recordset
Dim rsList As Recordset

Dim m_blnCancelled As Boolean
Dim m_lngContID As Long

Private Sub cmdOpts_Click(Index As Integer)
   Select Case Index
      Case 0 'OK
         If (m_lngContID = 0) Then
            MsgBox "You must select a Contact from the list.", , APP_MSG_NAME
            Exit Sub
         End If
         
         Call PostEntry
      Case 1 'Cancel
         m_blnCancelled = True
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmSetRelContact.Form_Load"
   On Error GoTo Error_Handler
   
   'set  main recordset
   Set rsRelCont = dbContact.OpenRecordset("RelateProject", dbOpenTable)
   
   'load the list
   Call LoadAllContacts
   
   m_blnCancelled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsRelCont.Close
   Set rsRelCont = Nothing
   
   If (m_blnCancelled = False) Then
      Call frmProjEntry.LoadRelContactInfo
   End If
   
   Set frmSetRelContact = Nothing
End Sub

Private Sub LoadAllContacts()
   'load all contacts listed in the databse
   Const sMOD_NAME As String = "frmSetRelContact.Form_Load"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strSetting As String
   
   strSetting = "Default"
   
   SQL = "SELECT ContID, Setting, ShownName FROM Contacts "
   SQL = SQL & "WHERE Setting = '" & strSetting & "' "
   SQL = SQL & "ORDER BY ShownName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ContID)) Then
               If (Not IsNull(!ShownName)) Then lstContact.AddItem !ShownName
               lstContact.ItemData(lstContact.NewIndex) = !ContID
            End If
            .MoveNext
         Wend
      Else
         MsgBox "There are no Contacts entered to select.", , APP_MSG_NAME
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

Private Sub lstContact_Click()
   Const sMOD_NAME As String = "frmSetRelContact.lstContact_Click"
   On Error GoTo Error_Handler
   
   m_lngContID = lstContact.ItemData(lstContact.ListIndex)
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while selecting a Contact Name!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub PostEntry()
   'post the project relation to the database
   Const sMOD_NAME As String = "frmSetRelContact.PostEntry"
   On Error GoTo Error_Handler
   
   rsRelCont.AddNew
   
   With rsRelCont
      !fkProjID = g_lngProjID
      !fkContID = m_lngContID
      !ContShowName = lstContact.Text
      
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

