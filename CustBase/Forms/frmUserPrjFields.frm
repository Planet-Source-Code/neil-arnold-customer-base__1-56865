VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserPrjFields 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New / Edit User Project Fields"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserPrjFields.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstUserFld 
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
      Left            =   150
      TabIndex        =   0
      Top             =   525
      Width           =   3315
   End
   Begin VB.TextBox txtValue 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4125
      MaxLength       =   100
      TabIndex        =   1
      Top             =   2025
      Width           =   4440
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Remove"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   3675
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3525
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&New ..."
      Height          =   390
      Index           =   0
      Left            =   3675
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   390
      Index           =   2
      Left            =   7725
      TabIndex        =   3
      Top             =   4275
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar tbrSave 
      Height          =   330
      Left            =   8625
      TabIndex        =   2
      Top             =   2010
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save The Entered Value"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   75
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPrjFields.frx":000C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   8925
      Y1              =   4110
      Y2              =   4110
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H00DEE3E6&
      Caption         =   "Add / Edit / Modify Project User Defined Fields"
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
      TabIndex        =   7
      Top             =   150
      Width           =   8790
   End
   Begin VB.Label Label1 
      Caption         =   "Value:"
      Height          =   240
      Left            =   3600
      TabIndex        =   6
      Top             =   2055
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   150
      X2              =   8925
      Y1              =   4125
      Y2              =   4125
   End
End
Attribute VB_Name = "frmUserPrjFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsUserFld As Recordset 'main recordset
Dim rsPUFld As Recordset 'for user field values
Dim rsList As Recordset 'all other data work

Dim m_lngFldID As Long 'for field selected from list
Dim m_lngRefNum As Long 'for editing previous saved value
Dim m_strFldDesc As String 'for selected field description

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmUserPrjFields.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'new
         Call AddNewUserField
      Case 1 'remove
         Call DeleteUserField
      Case 2 'cancel
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmUserPrjFields.Form_Load"
   On Error GoTo Error_Handler
   
   'flatten all needed borders
   FlatBorder lstUserFld.hWnd
   FlatBorder txtValue.hWnd
   
   'setup the screen
   Call InitializeScreen
   
   'set main recordsets
   Set rsUserFld = dbContact.OpenRecordset("PUserFields", dbOpenTable)
   Set rsPUFld = dbContact.OpenRecordset("PUFldValues", dbOpenTable)
   
   'disable toolbar button
   tbrSave.Buttons(1).Enabled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Const sMOD_NAME As String = "frmUserPrjFields.Form_Unload"
   On Error GoTo Error_Handler
   
   'remove data & form reference
   rsUserFld.Close
   Set rsUserFld = Nothing
   rsPUFld.Close
   Set rsPUFld = Nothing
   
   Call frmProjEntry.LoadUserDefInfo
   
   Set frmUserPrjFields = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub InitializeScreen()
   'Set up the opening screen
   Const sMOD_NAME As String = "frmUserPrjFields.InitializeScreen"
   On Error GoTo Error_Handler
   
   Call LoadUserFields
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub LoadUserFields()
   Const sMOD_NAME As String = "frmUserPrjFields.LoadUserFields"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT RefNum, Description FROM PUserFields "
   SQL = SQL & "ORDER BY Description"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lstUserFld.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then lstUserFld.AddItem !Description
            lstUserFld.ItemData(lstUserFld.NewIndex) = !RefNum
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

Private Sub AddNewUserField()
   'add a new user defined field to the database
   Const sMOD_NAME As String = "frmUserPrjFields.AddNewUserField"
   On Error GoTo Error_Handler
   
   Dim strField As String
   
   strField = InputBox("Enter the new User Field to be used.", _
      "Enter New User Field")
   
   If (Len(strField) > 100) Then
      MsgBox "The field entered is too long. (100 characters max)", , APP_MSG_NAME
      Exit Sub
   ElseIf (Len(strField) <= 0) Then
      Exit Sub
   End If
   
   rsUserFld.AddNew
   
   With rsUserFld
      !Description = strField
      
      .Update
   End With
   
   Call LoadUserFields
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Adding the new User Field!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub lstUserFld_Click()
   On Error Resume Next
   
   m_lngFldID = lstUserFld.ItemData(lstUserFld.ListIndex)
   m_strFldDesc = lstUserFld.Text
   
   If (m_lngFldID > 0) Then
      txtValue.Enabled = True
      txtValue.BackColor = vbWhite
      
      cmdOpts(1).Enabled = True
      
      txtValue.SetFocus
      
      Call LoadSavedValue
   End If
End Sub

Private Sub tbrSave_ButtonClick(ByVal Button As MSComctlLib.Button)
   Const sMOD_NAME As String = "frmUserPrjFields.tbrSave_ButtonClick"
   On Error GoTo Error_Handler
   
   Select Case Button.Key
      Case "Save"
         If (Not ValidateEntry()) Then Exit Sub
         
         Call PostEntry
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Saving the record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub txtValue_Change()
   tbrSave.Buttons(1).Enabled = True
End Sub

Private Sub txtValue_GotFocus()
   highLight
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
   Const sMOD_NAME As String = "frmUserPrjFields.txtValue_KeyPress"
   On Error GoTo Error_Handler
   
   If KeyAscii = vbKeyReturn Then
      If (Not ValidateEntry()) Then Exit Sub
         
      Call PostEntry
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while trying to Post this record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub LoadSavedValue()
   'if the selected contact already has a value for this field, look it up
   'so that it can be modified
   Const sMOD_NAME As String = "frmUserPrjFields.LoadSavedValue"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT RefNum, fkProjID, fkUserFld, Value "
   SQL = SQL & "FROM PUFldValues "
   SQL = SQL & "WHERE fkProjID = " & g_lngProjID
   SQL = SQL & " AND fkUserFld = '" & m_strFldDesc & "' "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         icurState = NOW_EDITING
         txtValue.Text = ""
         .MoveFirst
         If (Not IsNull(!RefNum)) Then m_lngRefNum = !RefNum
         If (Not IsNull(!Value)) Then txtValue.Text = !Value
      Else
         icurState = NOW_ADDING
         txtValue.Text = "" 'modified 11/9/04
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(txtValue) < 1) Then
      Indx = MsgBox("You Must Enter A User Field Value", _
         vbInformation + vbOKOnly, "Validate : User Field Value")
      txtValue.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmUserPrjFields.PostEntry"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting User Field Value Entry", True
   
   If (icurState = NOW_ADDING) Then
      rsPUFld.AddNew
   Else
      With rsPUFld
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngRefNum
            If Not .NoMatch Then
               rsPUFld.Edit
            Else
               MsgBox "Error Posting Record", , "Posting Error"
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
      End With
   End If
   
   With rsPUFld
      !fkProjID = g_lngProjID
      !fkUserFld = m_strFldDesc
      If (Len(txtValue)) Then !Value = txtValue
      
      .Update
   End With
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   'clear variables
   m_lngFldID = 0
   m_strFldDesc = ""
   m_lngRefNum = 0
   
   'disable all unneeded controls
   txtValue.Enabled = False
   txtValue.BackColor = vbButtonFace
   txtValue.Text = ""
   tbrSave.Buttons(1).Enabled = False
   cmdOpts(1).Enabled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting the record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Private Sub DeleteUserField()
   Const sMOD_NAME As String = "frmUserPrjFields.DeleteUserField"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim sMsg As String
   Dim iMsg As VbMsgBoxResult
   Dim DeleteSQL As String
   
   SQL = "SELECT fkUserFld FROM PUFldValues "
   SQL = SQL & "WHERE fkUserFld = '" & m_strFldDesc & "' "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         MsgBox "That field is in use in a Project Record and cannot be removed.", _
            vbInformation + vbOKOnly, "Cannot Remove Used Field"
         rsList.Close
         Set rsList = Nothing
         
         txtValue.Enabled = False
         txtValue.BackColor = vbButtonFace
         m_lngFldID = 0
         m_strFldDesc = ""
         cmdOpts(1).Enabled = False
         
         Exit Sub
      Else
         GoTo Delete
      End If
   End With
   
Delete:
   sMsg = "Are you sure you want to DELETE User Field" & vbCrLf
   sMsg = sMsg & "[ " & m_strFldDesc & " ] from the User Field List?"
   
   iMsg = MsgBox(sMsg, vbCritical + vbYesNo, "Warning : Delete User Field")
   
   If (iMsg <> vbYes) Then Exit Sub
   
   DeleteSQL = "DELETE * FROM PUserFields WHERE RefNum = " & m_lngFldID
   
   dbContact.Execute (DeleteSQL)
   
   Call LoadUserFields
   
   txtValue.Enabled = False
   txtValue.BackColor = vbButtonFace
   m_lngFldID = 0
   m_strFldDesc = ""
   cmdOpts(1).Enabled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An An un-known error occurred while Deleting the record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
