VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNames 
   Caption         =   "Names"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNames.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00F3F3ED&
      Height          =   390
      Left            =   4275
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   7
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
      ScaleWidth      =   11355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11355
      Begin VB.ComboBox cboGroup 
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
         Left            =   4950
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   75
         Width           =   2415
      End
      Begin VB.ComboBox cboView 
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
         Left            =   1875
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   55
         Width           =   2415
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
         Left            =   9300
         MouseIcon       =   "frmNames.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   105
         Width           =   1140
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   9000
         Picture         =   "frmNames.frx":074C
         Stretch         =   -1  'True
         Top             =   100
         Width           =   240
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "View:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   1275
         TabIndex        =   5
         Top             =   98
         Width           =   540
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Group:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   4350
         TabIndex        =   3
         Top             =   105
         Width           =   540
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmNames.frx":08F6
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
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
         TabIndex        =   1
         Top             =   98
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   7215
      Left            =   75
      TabIndex        =   6
      Top             =   525
      Width           =   11190
      _ExtentX        =   19738
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
Attribute VB_Name = "frmNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_strView As String
Dim m_strGroup As String

Private Sub cboGroup_Click()
   LockWindowUpdate frmNames.hWnd
   
   m_strGroup = cboGroup.Text
   Call LoadFilteredNamesList
   
   LockWindowUpdate 0
End Sub

Private Sub cboView_Click()
   LockWindowUpdate frmNames.hWnd
   
   m_strView = cboView.Text
   Call LoadContactInfo
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmNames.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Names List", True
   frmMain.picStatus.BackColor = &H666633
   
   'set up cboView
   With cboView
      .AddItem "All Contacts"
      .AddItem "Default"
      .AddItem "Hidden"
      .AddItem "Private"
   End With
   
   cboGroup.AddItem "All"
   
   m_strView = "Default"
   m_strGroup = "All"
   
   'load groups combo
   Call LoadAllGroups
   
   cboView.Text = "Default"
   
   'set global form identifier
   g_strFormFlag = ""
   
   'set gridline preference
   lvResult.GridLines = g_blnShowLines
   
   Screen.MousePointer = vbDefault
   'MsgBar vbNullString, False
   
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
   
   LockWindowUpdate frmNames.hWnd
   
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
   Set frmNames = Nothing
End Sub

Private Sub LoadContactInfo()
   'load the selected contact info into the grid
   Const sMOD_NAME As String = "frmNames.LoadContactInfo"
   On Error GoTo Error_Handler
   
   Dim strPhone As String
   Dim strEmail As String
   Dim SQL As String
   Dim NameSQL As String
   Dim strType As String
   Dim Item As ListItem
   Dim rsAddr As Recordset
   
   strType = "Home"
   
   Select Case m_strView
      Case "All Contacts"
         NameSQL = "SELECT ContID, ShownName, JobTitle FROM Contacts "
         NameSQL = NameSQL & "ORDER BY ShownName"
      Case Else
         NameSQL = "SELECT ContID, Setting, ShownName, JobTitle "
         NameSQL = NameSQL & "FROM Contacts "
         NameSQL = NameSQL & "WHERE Setting = '" & m_strView & "' "
         NameSQL = NameSQL & "ORDER BY ShownName"
   End Select
   
   Set rsList = dbContact.OpenRecordset(NameSQL)
   
   lvResult.ListItems.Clear
   
   With rsList
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
         If (g_blnAltColors = True) Then
            AltLVBackground lvResult, picGrdClr
         End If
      Else
         MsgBar vbNullString, False
         rsList.Close
         Set rsList = Nothing
         Exit Sub
      End If
   End With
   
   MsgBar "Found: " & rsList.RecordCount & " Records", False
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub lblSaveAs_Click()
   'add code to save as CSV file
   Const sMOD_NAME As String = "frmNames.lblSaveAs_Click"
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
   Const sMOD_NAME As String = "frmNames.lvResult_Click"
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
      MsgBox "An un-known error occurred while opening the Contacts screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub LoadAllGroups()
   'load all entered groups into cboGroup
   Const sMOD_NAME As String = "frmNames.LoadAllGroups"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strItemID
   
   strItemID = "GRP"
   
   SQL = "SELECT ItemID, Description FROM Lookup "
   SQL = SQL & "WHERE ItemID = '" & strItemID & "' "
   SQL = SQL & "ORDER BY Description"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then cboGroup.AddItem !Description
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub LoadFilteredNamesList()
   'load the selected contact info into the grid
   Const sMOD_NAME As String = "frmNames.LoadFilteredNamesList"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strPhone As String
   Dim strEmail As String
   Dim Item As ListItem
   Dim AddrSQL As String
   Dim rsAddr As Recordset
   Dim strType As String
   Dim strSetting As String
   
   strSetting = "Default"
   strType = "Home"
   
   SQL = "SELECT ContID, Setting, ShownName, JobTitle, Group "
   SQL = SQL & "FROM Contacts "
   SQL = SQL & "WHERE Setting = '" & strSetting & "' "
   SQL = SQL & "ORDER BY ShownName"
   
   lvResult.ListItems.Clear
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Group)) Then
               If InStr(1, !Group, m_strGroup) Then
                  If (Not IsNull(!ContID)) Then
                     If (Not IsNull(!ShownName)) Then
                        Set Item = lvResult.ListItems.Add(, "ID" & !ContID, !ShownName)
                     End If
                  End If
                  If (Not IsNull(!JobTitle)) Then Item.SubItems(1) = !JobTitle
                  
                  strPhone = GetPhoneNum(!ContID)
                  If (Not IsNull(strPhone)) Then Item.SubItems(2) = strPhone
                  
                  strEmail = GetEMail(!ContID)
                  If (Not IsNull(strEmail)) Then Item.SubItems(3) = strEmail
                  
                  'Get address info
                     AddrSQL = "SELECT fkContID, fkLookup, Street, City, State, Zip "
                     AddrSQL = AddrSQL & "FROM CAddress WHERE fkContID = " & !ContID
                     AddrSQL = AddrSQL & " AND fkLookup = '" & strType & "' "
                     
                     Set rsAddr = dbContact.OpenRecordset(AddrSQL)
                     
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
               Else
                  .MoveNext
               End If
            End If
         Wend
         If (g_blnAltColors = True) Then
            AltLVBackground lvResult, picGrdClr
         End If
      Else
         MsgBar vbNullString, False
         rsList.Close
         Set rsList = Nothing
         Exit Sub
      End If
   End With
   
   MsgBar "Found: " & lvResult.ListItems.Count & " Records", False
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub
