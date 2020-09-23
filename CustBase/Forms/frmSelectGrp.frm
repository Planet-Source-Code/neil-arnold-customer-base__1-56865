VERSION 5.00
Begin VB.Form frmSelectGrp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Group(s)"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectGrp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTemp 
      Height          =   240
      Left            =   150
      TabIndex        =   6
      Top             =   300
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   3
      Left            =   2925
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   2
      Left            =   1575
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   2925
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4050
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&New ..."
      Height          =   390
      Index           =   0
      Left            =   2925
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3525
      Width           =   1215
   End
   Begin VB.ListBox lstGroup 
      Height          =   3885
      Left            =   150
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   600
      Width           =   2565
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   4125
      Y1              =   4635
      Y2              =   4635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   75
      X2              =   4125
      Y1              =   4650
      Y2              =   4650
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
      Caption         =   "Click to select groups to assign"
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
      TabIndex        =   5
      Top             =   150
      Width           =   3990
   End
End
Attribute VB_Name = "frmSelectGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Const GTYPE = "GRP"

Dim m_strSelection As String
Dim m_blnCancelled As Boolean

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmSelectGrp.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'new
         Load frmAddGroup
         frmAddGroup.Show vbModeless, frmSelectGrp
      Case 1 'remove
         Call DeleteGroup
      Case 2 'OK
         Call GetString
         Call UpdateContactRecord
      Case 3 'cancel
         m_blnCancelled = True
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmSelectGrp.Form_Load"
   On Error GoTo Error_Handler
   
   'flatten all needed borders
   FlatBorder lstGroup.hWnd
   
   'setup groups list
   Call InitializeScreen
   
   m_blnCancelled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (m_blnCancelled = False) Then
      Call frmContEntry.LoadMainContactInfo
   End If
   
   'remove data & form reference
   Set frmSelectGrp = Nothing
End Sub

Public Sub InitializeScreen()
   'setup the opening screen
   Const sMOD_NAME As String = "frmSelectGrp.InitializeScreen"
   On Error GoTo Error_Handler
   
   Call LoadAllGroups
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub LoadAllGroups()
   'load all listed groups
   Const sMOD_NAME As String = "frmSelectGrp.LoadAllGroups"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT ItemID, Description FROM Lookup "
   SQL = SQL & "WHERE ItemID = '" & GTYPE & "' "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lstGroup.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then lstGroup.AddItem !Description
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

Private Sub lstGroup_Click()
   cmdOpts(1).Enabled = True
End Sub

Private Sub GetString()
   'get all items selected in lstGroup and insert a comma between items
   Const sMOD_NAME As String = "frmSelectGrp.GetString"
   On Error GoTo Error_Handler
   
   Dim sTemp As String
   Dim iCtr As Integer
   
   txtTemp.Text = ""
   
   For iCtr = 0 To lstGroup.ListCount - 1
      If lstGroup.Selected(iCtr) Then
         sTemp = txtTemp.Text
         If Len(sTemp) = 0 Then
            txtTemp.Text = sTemp & lstGroup.List(iCtr)
         Else
            txtTemp.Text = sTemp & "," & lstGroup.List(iCtr)
         End If
      End If
   Next
   
   m_strSelection = txtTemp.Text
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub UpdateContactRecord()
   'save the selection(s) to the contact file
   Const sMOD_NAME As String = "frmSelectGrp.UpdateContactRecord"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   If (Len(m_strSelection) <= 0) Then
      MsgBox "You must select at least one item from the list.", , APP_MSG_NAME
      Exit Sub
   End If
   
   SQL = "SELECT * FROM Contacts WHERE ContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   rsList.Edit
   
   With rsList
      If (Len(m_strSelection) > 0) Then !Group = m_strSelection
      
      .Update
   End With
   
   Me.Hide
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub DeleteGroup()
   'delete selected group item
   Const sMOD_NAME As String = "frmSelectGrp.DeleteGroup"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim sMsg As String
   Dim iMsg As VbMsgBoxResult
   Dim strChoice As String
   
   strChoice = lstGroup.Text
   
   sMsg = "Are you sure you want to Delete " & vbCrLf
   sMsg = sMsg & strChoice & " as one of the Contact Groups?"
   
   iMsg = MsgBox(sMsg, vbQuestion + vbYesNo, "Warning : Deleting Group")
   
   If (iMsg <> vbYes) Then Exit Sub
   
   SQL = "DELETE * FROM Lookup WHERE ItemID = '" & GTYPE & "' "
   SQL = SQL & "AND Description = '" & strChoice & "' "
   
   dbContact.Execute (SQL)
   
   Call InitializeScreen
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Deleting the group record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
