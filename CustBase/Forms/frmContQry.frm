VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContQry 
   Caption         =   "Contacts Query"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmContQry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmContQry.frx":0442
   ScaleHeight     =   7740
   ScaleWidth      =   11325
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00F3F3ED&
      Height          =   390
      Left            =   4275
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   10
      Top             =   2475
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H00666633&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11325
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   11325
      Begin VB.OptionButton optWhere 
         BackColor       =   &H00666633&
         Caption         =   "<>"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   6000
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Not Equal To"
         Top             =   98
         Width           =   615
      End
      Begin VB.OptionButton optWhere 
         BackColor       =   &H00666633&
         Caption         =   ">"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   5550
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Greater Than"
         Top             =   98
         Width           =   465
      End
      Begin VB.OptionButton optWhere 
         BackColor       =   &H00666633&
         Caption         =   "<"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   5100
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Less Than"
         Top             =   98
         Width           =   465
      End
      Begin VB.OptionButton optWhere 
         BackColor       =   &H00666633&
         Caption         =   "="
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   4650
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Equals"
         Top             =   98
         Value           =   -1  'True
         Width           =   465
      End
      Begin VB.ComboBox cboUserDef 
         BackColor       =   &H00666633&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1650
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   55
         Width           =   2940
      End
      Begin VB.ComboBox cboValue 
         BackColor       =   &H00666633&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   6675
         TabIndex        =   5
         Top             =   75
         Width           =   2715
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   " Names -"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   450
         TabIndex        =   9
         Top             =   98
         Width           =   840
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmContQry.frx":058C
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Field:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   1275
         TabIndex        =   8
         Top             =   105
         Width           =   390
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   9825
         Picture         =   "frmContQry.frx":06D6
         Stretch         =   -1  'True
         Top             =   105
         Width           =   240
      End
      Begin VB.Label lblSaveAs 
         BackStyle       =   0  'Transparent
         Caption         =   "Save File As ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10125
         MouseIcon       =   "frmContQry.frx":0880
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   105
         Width           =   1140
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   7215
      Left            =   0
      TabIndex        =   11
      Top             =   525
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Job Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "E-Mail"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "City"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "St."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ZIP"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape shpResult 
      BorderColor     =   &H00CCCCB4&
      Height          =   990
      Left            =   1050
      Top             =   1575
      Width           =   765
   End
End
Attribute VB_Name = "frmContQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_strField As String
Dim m_strWhereOper As String
Dim m_strWhereItem As String

Private Sub cboUserDef_Click()
   m_strField = cboUserDef.Text
   
   Call LoadAllAvailValues
End Sub

Private Sub cboValue_Click()
   Call LookupUserValues
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvResult, picGrdClr
   End If
End Sub

Private Sub cboValue_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call LookupUserValues
      If (g_blnAltColors = True) Then
         AltLVBackground lvResult, picGrdClr
      End If
   End If
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmContQry.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Contact Query Screen", True
   frmMain.picStatus.BackColor = &H666633
   
   'load groups combo
   Call LoadAllFields
   
   'set the default where operator
   m_strWhereOper = "="
   
   'set global form identifier
   g_strFormFlag = ""
   
   'set gridline preference
   lvResult.GridLines = g_blnShowLines
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
   
   On Error Resume Next
   
   LockWindowUpdate frmContQry.hWnd
   
   'adjust lvResult
   lvResult.Move 15, 480, Me.ScaleWidth - 30, Me.ScaleHeight - 495
   '***adjust lvResult column widths
   lvResult.ColumnHeaders(1).Width = lvResult.Width * 0.17 'name
   lvResult.ColumnHeaders(2).Width = lvResult.Width * 0.1  'job title
   lvResult.ColumnHeaders(3).Width = lvResult.Width * 0.12  'phone
   lvResult.ColumnHeaders(4).Width = lvResult.Width * 0.15 'e-mail
   lvResult.ColumnHeaders(5).Width = lvResult.Width * 0.17 'address
   lvResult.ColumnHeaders(6).Width = lvResult.Width * 0.15  'city
   lvResult.ColumnHeaders(7).Width = lvResult.Width * 0.04  'state
   lvResult.ColumnHeaders(8).Width = (lvResult.Width * 0.1) - 265  'zip
   shpResult.Move lvResult.Left - 15, lvResult.Top - 15, lvResult.Width + 30, lvResult.Height + 30
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   Set frmContQry = Nothing
End Sub

Private Sub LoadAllFields()
   'load all user defined fields stored
   Const sMOD_NAME As String = "frmContQry.LoadAllFields"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT Description FROM CUserFields ORDER BY Description"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then cboUserDef.AddItem !Description
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

Private Sub LoadAllAvailValues()
   'load all values from contact user field records for entered field
   Const sMOD_NAME As String = "frmContQry.LoadAllFields"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT DISTINCT fkUserFld, Value FROM CUFldValues "
   SQL = SQL & "WHERE fkUserFld = '" & m_strField & "' "
   SQL = SQL & "ORDER BY Value"
   
   cboValue.Clear
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Value)) Then cboValue.AddItem !Value
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the Values list!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub optWhere_Click(Index As Integer)
   m_strWhereOper = optWhere(Index).Caption
End Sub

Private Sub LookupUserValues()
   'lookup all user ID's that match the criteria
   Dim SQL As String
   
   m_strWhereItem = cboValue.Text
   
   SQL = "SELECT fkContID, fkUserFld, Value FROM CUFldValues "
   SQL = SQL & "WHERE fkUserFld = '" & m_strField & "' "
   
   Select Case m_strWhereOper
      Case "="
         SQL = SQL & "AND Value = '" & m_strWhereItem & "' "
      Case "<"
         SQL = SQL & "AND Value < '" & m_strWhereItem & "' "
      Case ">"
         SQL = SQL & "AND Value > '" & m_strWhereItem & "' "
      Case "<>"
         SQL = SQL & "AND Value <> '" & m_strWhereItem & "' "
   End Select
   
   SQL = SQL & " ORDER BY fkContID"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvResult.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkContID)) Then
               Call LoadContactInfo(!fkContID)
            End If
            .MoveNext
         Wend
      Else
         MsgBar vbNullString, False
         lvResult.ListItems.Clear
      End If
   End With
   
   MsgBar "Showing: " & rsList.RecordCount & " Records (of " & rsList.RecordCount & " total)", False
   
   rsList.Close
   Set rsList = Nothing
End Sub

Private Sub LoadContactInfo(lngContID As Long)
   'load the selected contact info into the grid
   Const sMOD_NAME As String = "frmContQry.LoadContactInfo"
   On Error GoTo Error_Handler
   
   Dim strPhone As String
   Dim strEmail As String
   Dim SQL As String
   Dim NameSQL As String
   Dim strType As String
   Dim Item As ListItem
   Dim rsListing As Recordset
   Dim rsAddr As Recordset
   
   strType = "Home"
   
   NameSQL = "SELECT ContID, ShownName, JobTitle FROM Contacts "
   NameSQL = NameSQL & "WHERE ContID = " & lngContID
   NameSQL = NameSQL & " ORDER BY ContID"
        
   Set rsListing = dbContact.OpenRecordset(NameSQL)
   
   With rsListing
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ContID)) Then
               If (Not IsNull(!ShownName)) Then
                  Set Item = lvResult.ListItems.Add(, "ID" & !ContID, !ShownName)
               End If
               If (Not IsNull(!JobTitle)) Then Item.SubItems(1) = !JobTitle
               
               strPhone = GetPhoneNum(!ContID)
               If (Not IsNull(strPhone)) Then Item.SubItems(2) = strPhone
               
               strEmail = GetEMail(!ContID)
               If (Not IsNull(strEmail)) Then Item.SubItems(3) = strEmail
               
               'Get address info
                  SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip "
                  SQL = SQL & "FROM CAddress WHERE fkContID = " & !ContID
                  SQL = SQL & " AND fkLookup = '" & strType & "' "
                  
                  Set rsAddr = dbContact.OpenRecordset(SQL)
                  
                  With rsAddr
                     If (.RecordCount > 0) Then
                        .MoveFirst
                        While Not .EOF
                           If (Not IsNull(!Street)) Then Item.SubItems(4) = !Street
                           If (Not IsNull(!City)) Then Item.SubItems(5) = !City
                           If (Not IsNull(!State)) Then Item.SubItems(6) = !State
                           If (Not IsNull(!Zip)) Then Item.SubItems(7) = !Zip
                           .MoveNext
                        Wend
                     End If
                  End With
                  rsAddr.Close
                  Set rsAddr = Nothing
               .MoveNext
            End If
         Wend
      End If
   End With
   
   rsListing.Close
   Set rsListing = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub lblSaveAs_Click()
   'add code to save as CSV file
   Const sMOD_NAME As String = "frmContQry.lblSaveAs_Click"
   On Error GoTo Error_Handler
   
   Dim strFilePath As String
   Dim iCtr As Integer
   
   If (lvResult.ListItems.Count <= 0) Then
      MsgBox "There are no item's in the list to save.", , APP_MSG_NAME
      Exit Sub
   End If
   
   With frmMain.cdlMain
      .DialogTitle = "Save file as CSV File"
      .Filter = "Comma Separated Values (*.csv)|*.csv"
      .Flags = cdlOFNOverwritePrompt
      .InitDir = App.Path & "\Templates"
      .CancelError = True
      .ShowSave
   End With
   
   strFilePath = frmMain.cdlMain.FileName
   
   
   Open strFilePath For Output As #1
   
   Print #1, Chr(34) & "Name" & Chr(34); Chr(44) & _
         Chr(34) & "Job Title" & Chr(34); Chr(44) & _
         Chr(34) & "Phone" & Chr(34); Chr(44) & _
         Chr(34) & "E-Mail" & Chr(34); Chr(44) & _
         Chr(34) & "Address" & Chr(34); Chr(44) & _
         Chr(34) & "City" & Chr(34); Chr(44) & _
         Chr(34) & "St." & Chr(34); Chr(44) & _
         Chr(34) & "ZIP" & Chr(34)
   For iCtr = 1 To lvResult.ListItems.Count
      lvResult.SelectedItem = lvResult.ListItems(iCtr)
      Print #1, Chr(34) & lvResult.SelectedItem & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(1) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(2) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(3) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(4) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(5) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(6) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(7) & Chr(34)
   Next
   Close #1
   
   MsgBox "The File was Sucessfully Created and Saved.", , APP_MSG_NAME
   
   Exit Sub
Error_Handler:
   If Err <> 32755 And Err <> 3049 Then   'check for common dialog cancelled
      LogErrors sMOD_NAME, Err.Number, Err.Description
      ShowError
      Exit Sub
   End If
End Sub

Private Sub lvResult_Click()
   Const sMOD_NAME As String = "frmContQry.lvResult_Click"
   On Error GoTo Error_Handler
   
   g_lngContID = CLng(Mid$(lvResult.SelectedItem.Key, 3, Len(lvResult.SelectedItem.Key)))
   
   'code to open Contact entry screen
   UnloadAllForms
   Load frmContEntry
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred loading the selected Contact information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub


