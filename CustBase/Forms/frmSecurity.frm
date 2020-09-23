VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSecurity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Security"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   390
      Index           =   4
      Left            =   6000
      TabIndex        =   9
      Top             =   3300
      Width           =   1290
   End
   Begin VB.CommandButton cmdOpts 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   6975
      Picture         =   "frmSecurity.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   315
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   4575
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2025
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4575
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1650
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   4575
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1275
      Width           =   2715
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "Change &Password"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1515
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Remove User"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   1500
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1365
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Add New User"
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   3150
      Width           =   1365
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1500
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":06CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvUser 
      Height          =   1890
      Left            =   75
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   3334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Names"
         Object.Width           =   4260
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check to Enable Security"
      Height          =   240
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   450
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Password Max. Length is 15 characters (letters or numbers)"
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
      Left            =   4575
      TabIndex        =   15
      Top             =   2400
      Width           =   2115
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   2925
      X2              =   7275
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Validate Password:"
      Height          =   240
      Index           =   2
      Left            =   3150
      TabIndex        =   14
      Top             =   2062
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Password:"
      Height          =   240
      Index           =   1
      Left            =   3150
      TabIndex        =   13
      Top             =   1687
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
      Height          =   240
      Index           =   0
      Left            =   3150
      TabIndex        =   12
      Top             =   1312
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   2925
      X2              =   7275
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Authorized Users"
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
      Left            =   75
      TabIndex        =   11
      Top             =   825
      Width           =   7215
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Enable / Disable Program Security"
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
      TabIndex        =   10
      Top             =   75
      Width           =   7215
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsUser As Recordset 'main recordset
Dim rsList As Recordset 'all other data work

Dim m_strSelUserName As String 'for selected user name
Dim m_lngSelUserID As Long 'for selected user ID
Dim m_strFstPass As String 'first pwd entered
Dim m_strVerPass As String 'verified pwd
Dim m_strExistPwd As String 'existing password

Dim m_blnChanged As Boolean

Private Sub Check1_Click()
   If (Check1.Value = 1) Then
      g_blnIsSecure = True
   ElseIf (Check1.Value = 0) Then
      g_blnIsSecure = False
   End If
   
   SaveSetting APP_CATEGORY, APPNAME, "Security", IIf(g_blnIsSecure, "-1", "0")
End Sub

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmSecurity.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Dim iCtr As Integer
   
   Select Case Index
      Case 0 'add user
         icurState = NOW_ADDING
         For iCtr = 0 To 2
            Text1(iCtr).Enabled = True
            Text1(iCtr).BackColor = vbWhite
         Next
         Text1(0).SetFocus
         m_blnChanged = False
      Case 1 'remove user
         If (m_lngSelUserID = 0) Then
            MsgBox "You must select a User Name from the list first.", _
               vbInformation + vbOKOnly, APP_MSG_NAME
            Exit Sub
         End If
         
         Call DeleteUserProfile
      Case 2 'change password
         If (m_lngSelUserID = 0) Then
            MsgBox "You must select a User Name from the list first.", _
               vbInformation + vbOKOnly, APP_MSG_NAME
            Exit Sub
         End If
         
         Call GetUserPassword
         Call VerifyUserPassword
      Case 3 'save
         If (Not ValidateEntry()) Then Exit Sub
         
         Call PostEntry
      Case 4 'close
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmSecurity.Form_Load"
   On Error GoTo Error_Handler
   
   Dim iCtr As Integer
   
   'set main recordset
   Set rsUser = dbContact.OpenRecordset("Security", dbOpenTable)
   
   'flatten all needed items
   FlatBorder lvUser.hWnd
   For iCtr = 0 To 2
      FlatBorder Text1(iCtr).hWnd
   Next
   
   'set opening screen
   Call InitializeScreen
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim iMsg As VbMsgBoxResult
   
   If (m_blnChanged = True) Then
      iMsg = MsgBox("The User Information has changed." & vbCrLf & "Would you like to save the changes?", _
         vbQuestion + vbYesNo, "Verify Changes")
      If (iMsg <> vbYes) Then
         Cancel = False
      Else
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove any data & form reference
   rsUser.Close
   Set rsUser = Nothing
   
   Set frmSecurity = Nothing
End Sub

Public Sub InitializeScreen()
   Const sMOD_NAME As String = "frmSecurity.InitializeScreen"
   On Error GoTo Error_Handler
   
   'check global "use security" variable
   If (g_blnIsSecure = True) Then
      Check1.Value = 1
   Else
      Check1.Value = 0
   End If
   
   'set the screen for operation upon opening
   Call LoadAllUsers
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub LoadAllUsers()
   'load all users listed in the system
   Const sMOD_NAME As String = "frmSecurity.LoadAllUsers"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, UserName FROM Security ORDER BY UserName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         Check1.Enabled = True
         lvUser.ListItems.Clear
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               Set Item = lvUser.ListItems.Add(, "ID" & !RefNum, !UserName, , 1)
            End If
            .MoveNext
         Wend
      Else
         Check1.Enabled = False
         lvUser.ListItems.Clear
         'reset security settings in the case of no users
         SaveSetting APP_CATEGORY, APPNAME, "Security", "0"
         g_blnIsSecure = False
         Check1.Value = 0
         Exit Sub
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub lvUser_Click()
   Const sMOD_NAME As String = "frmSecurity.lvUser_Click"
   On Error GoTo Error_Handler
   
   m_lngSelUserID = CLng(Mid$(lvUser.SelectedItem.Key, 3, Len(lvUser.SelectedItem.Key)))
   m_strSelUserName = lvUser.SelectedItem
   
   cmdOpts(1).Enabled = True
   cmdOpts(2).Enabled = True
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub Text1_Change(Index As Integer)
   m_blnChanged = True
   cmdOpts(3).Enabled = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   highLight
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   Select Case Index
      Case 1 'first entry
         m_strFstPass = Text1(1).Text
      Case 2 'ver entry
         m_strVerPass = Text1(2).Text
   End Select
End Sub

Private Function ValidateEntry() As Boolean
   Dim iMsg As Integer
   
   ValidateEntry = True
   
   If (Len(Text1(0)) < 1) Then
      iMsg = MsgBox("You Must Enter A User Name", _
         vbInformation + vbOKOnly, "Validate : User Name")
      Text1(0).SetFocus
      ValidateEntry = False
      Exit Function
   End If
   If (Len(Text1(1)) < 1) Then
      iMsg = MsgBox("You Must Enter A Password (15 chrs. max.)", _
         vbInformation + vbOKOnly, "Validate : Password")
      Text1(1).SetFocus
      ValidateEntry = False
      Exit Function
   End If
   If (Len(Text1(2)) < 1) Then
      iMsg = MsgBox("You Must Re-Enter The Password (15 chrs. max.)", _
         vbInformation + vbOKOnly, "Validate : Re-Enter The Password")
      Text1(2).SetFocus
      ValidateEntry = False
      Exit Function
   End If
   If (StrComp(m_strFstPass, m_strVerPass) <> 0) Then
      iMsg = MsgBox("The two password entries do not match. Please Re-Enter", _
         vbInformation + vbOKOnly, "Verify Password Entries")
      Text1(1).SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmSecurity.PostEntry"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting User Security Entry", True
   
   If (icurState = NOW_ADDING) Then
      rsUser.AddNew
   Else
      With rsUser
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngSelUserID
            If Not .NoMatch Then
               rsUser.Edit
            Else
               MsgBox "Error Posting Record", , "Posting Error"
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
      End With
   End If
   
   With rsUser
      If (Len(Text1(0))) Then !UserName = Text1(0)
      If (Len(Text1(2))) Then !Password = Base64Encode(Text1(2).Text)
      
      .Update
   End With
   
   'reset the screen for another entry
   Call ResetScreen
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Private Sub ResetScreen()
   'clear and reset the screen for another action
   Dim iCtr As Integer
   
   'set edit mode
   icurState = NOW_ADDING
   
   'clear variables
   m_strSelUserName = ""
   m_lngSelUserID = 0
   m_strFstPass = ""
   m_strVerPass = ""
   
   'clear screen
   For iCtr = 0 To 2
      Text1(iCtr).Text = ""
      Text1(iCtr).Enabled = False
      Text1(iCtr).BackColor = vbButtonFace
   Next
   m_blnChanged = False
   
   'reset buttons
   For iCtr = 1 To 3
      cmdOpts(iCtr).Enabled = False
   Next
   
   'refresh the listview
   Call LoadAllUsers
   
   cmdOpts(0).SetFocus
End Sub

Private Sub DeleteUserProfile()
   Const sMOD_NAME As String = "frmSecurity.DeleteUserProfile"
   On Error GoTo Error_Handler
   
   'remove a user entry
   Dim SQL As String
   Dim iMsg As VbMsgBoxResult
   Dim sMsg As String
   
   sMsg = "Are you sure you want to remove [ " & m_strSelUserName & " ]" & vbCrLf
   sMsg = sMsg & "as a User Name from the system?"
   
   iMsg = MsgBox(sMsg, vbQuestion + vbYesNo, "Verify User Name Delete")
   
   If (iMsg <> vbYes) Then
      m_strSelUserName = ""
      m_lngSelUserID = 0
      
      cmdOpts(1).Enabled = False
      cmdOpts(2).Enabled = False
      Exit Sub
   End If
   
   SQL = "DELETE * FROM Security WHERE RefNum = " & m_lngSelUserID
   
   dbContact.Execute (SQL)
   
   Call ResetScreen
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Deleting the User Profile!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub GetUserPassword()
   'get the selected users password for verification
   Const sMOD_NAME As String = "frmSecurity.GetUserPassword"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT RefNum, Password FROM Security "
   SQL = SQL & "WHERE RefNum = " & m_lngSelUserID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!Password)) Then
            m_strExistPwd = Base64Decode(!Password)
         End If
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub VerifyUserPassword()
   'verify that this user is able to change this password
   Const sMOD_NAME As String = "frmSecurity.VerifyUserPassword"
   On Error GoTo Error_Handler
   
   Dim strVerifyPwd As String
   Dim iCtr As Integer
   
   strVerifyPwd = InputBox("Enter your current password." & vbCrLf & "Remember, it is Case Sensitive", _
      "Verify User Password", , Me.Left + 1500, Me.Top + 1500)
   
   If (StrComp(m_strExistPwd, strVerifyPwd) <> 0) Then
      MsgBox "You are not authorized to change this password.", _
         vbInformation + vbOKOnly, "Dis-allow Password Change"
         'reset variables
      m_strSelUserName = ""
      m_lngSelUserID = 0
      m_strExistPwd = ""
      cmdOpts(1).Enabled = False
      cmdOpts(2).Enabled = False
      Exit Sub
   End If
   
   'allow password modification
   icurState = NOW_EDITING
   
   For iCtr = 0 To 2
      Text1(iCtr).Enabled = True
      Text1(iCtr).BackColor = vbWhite
   Next
   
   Text1(0) = m_strSelUserName
   Text1(1) = m_strExistPwd
   Text1(0).SetFocus
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Verifying the User Password!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
