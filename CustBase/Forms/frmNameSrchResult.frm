VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmNameSrchResult 
   Caption         =   "Name Search Results"
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
   Icon            =   "frmNameSrchResult.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvResult 
      Height          =   6840
      Left            =   150
      TabIndex        =   3
      Top             =   675
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   12065
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
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00EDDFE5&
      Height          =   390
      Left            =   4050
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   2
      Top             =   2475
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H007E5669&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   0
      Width           =   11355
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   " Results for Entered Name Search"
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
         Left            =   450
         TabIndex        =   1
         Top             =   105
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmNameSrchResult.frx":0442
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
   End
   Begin VB.Shape shpResult 
      BorderColor     =   &H00EDDFE5&
      Height          =   990
      Left            =   825
      Top             =   1575
      Width           =   765
   End
End
Attribute VB_Name = "frmNameSrchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmNameSrchResult.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Name Search Results", True
   frmMain.picStatus.BackColor = &H7E5669
   
   'Load all needed data
   Call LoadContactInfo
   
   'set global form identifier
   g_strFormFlag = ""
   
   'set gridline preference
   lvResult.GridLines = g_blnShowLines
   
   Screen.MousePointer = vbDefault
   
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
   
   LockWindowUpdate frmNameSrchResult.hWnd
   
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

Private Sub LoadContactInfo()
   'load the selected contact info into the grid
   Const sMOD_NAME As String = "frmNameSrchResult.LoadContactInfo"
   On Error GoTo Error_Handler
   
   Dim strPhone As String
   Dim strEmail As String
   Dim SQL As String
   Dim strType As String
   Dim Item As ListItem
   Dim rsAddr As Recordset
   
   strType = "Home"
   
   Set rsList = dbContact.OpenRecordset(g_strNameSQL)
   
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
      End If
   End With
   
   MsgBar "Found: " & rsList.RecordCount & " Records", False
   
   rsList.Close
   Set rsList = Nothing
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvResult, picGrdClr
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub lvResult_Click()
   Const sMOD_NAME As String = "frmNameSrchResult.lvResult_Click"
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
