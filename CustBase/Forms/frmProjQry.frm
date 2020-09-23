VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjQry 
   Caption         =   "Projects Query"
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
   Icon            =   "frmProjQry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   11325
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00DEE6F0&
      Height          =   390
      Left            =   3600
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2775
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H002A59A0&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11325
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   11325
      Begin VB.ComboBox cboValue 
         BackColor       =   &H002A59A0&
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
         Left            =   6750
         TabIndex        =   5
         Top             =   75
         Width           =   2715
      End
      Begin VB.ComboBox cboUserDef 
         BackColor       =   &H002A59A0&
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
         Left            =   1725
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   55
         Width           =   2940
      End
      Begin VB.OptionButton optWhere 
         BackColor       =   &H002A59A0&
         Caption         =   "="
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   4725
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Equals"
         Top             =   98
         Value           =   -1  'True
         Width           =   465
      End
      Begin VB.OptionButton optWhere 
         BackColor       =   &H002A59A0&
         Caption         =   "<"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   5175
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Less Than"
         Top             =   98
         Width           =   465
      End
      Begin VB.OptionButton optWhere 
         BackColor       =   &H002A59A0&
         Caption         =   ">"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   5625
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Greater Than"
         Top             =   98
         Width           =   465
      End
      Begin VB.OptionButton optWhere 
         BackColor       =   &H002A59A0&
         Caption         =   "<>"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   6075
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Not Equal To"
         Top             =   98
         Width           =   615
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
         MouseIcon       =   "frmProjQry.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   105
         Width           =   1140
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   9825
         Picture         =   "frmProjQry.frx":074C
         Stretch         =   -1  'True
         Top             =   105
         Width           =   240
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Field:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   1350
         TabIndex        =   8
         Top             =   105
         Width           =   390
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmProjQry.frx":08F6
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   " Projects-"
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
         TabIndex        =   7
         Top             =   105
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   7290
      Left            =   0
      TabIndex        =   11
      Top             =   450
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   12859
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Start Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "End Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape shpResult 
      BorderColor     =   &H00B0C0D6&
      Height          =   990
      Left            =   975
      Top             =   1650
      Width           =   765
   End
End
Attribute VB_Name = "frmProjQry"
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
   Const sMOD_NAME As String = "frmProjQry.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Project Query Screen", True
   frmMain.picStatus.BackColor = &H2A59A0
   
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
   
   LockWindowUpdate frmProjQry.hWnd
   
   'adjust lvResult
   lvResult.Move 15, 480, Me.ScaleWidth - 30, Me.ScaleHeight - 495
   '***adjust lvResult column widths
   lvResult.ColumnHeaders(1).Width = lvResult.Width * 0.25 'project name
   lvResult.ColumnHeaders(2).Width = lvResult.Width * 0.16  'start date
   lvResult.ColumnHeaders(3).Width = lvResult.Width * 0.16  'end date
   lvResult.ColumnHeaders(4).Width = lvResult.Width * 0.17 'status
   lvResult.ColumnHeaders(5).Width = (lvResult.Width * 0.26) - 265  'type
   shpResult.Move lvResult.Left - 15, lvResult.Top - 15, lvResult.Width + 30, lvResult.Height + 30
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   Set frmProjQry = Nothing
End Sub

Private Sub LoadAllFields()
   'load all user defined fields stored
   Const sMOD_NAME As String = "frmProjQry.LoadAllFields"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT Description FROM PUserFields ORDER BY Description"
   
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
   Const sMOD_NAME As String = "frmProjQry.LoadAllFields"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT DISTINCT fkUserFld, Value FROM PUFldValues "
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
   Dim rsListing As Recordset
   
   m_strWhereItem = cboValue.Text
   
   SQL = "SELECT fkProjID, fkUserFld, Value FROM PUFldValues "
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
   
   SQL = SQL & " ORDER BY fkProjID"
   
   Set rsListing = dbContact.OpenRecordset(SQL)
   
   lvResult.ListItems.Clear
   
   With rsListing
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkProjID)) Then
               Call LoadProjectInfo(!fkProjID)
            End If
            .MoveNext
         Wend
      Else
         MsgBar vbNullString, False
         lvResult.ListItems.Clear
      End If
   End With
   
   rsListing.Close
   Set rsListing = Nothing
End Sub

Private Sub LoadProjectInfo(lngProjID As Long)
   'load all stored projects
   Const sMOD_NAME As String = "frmProjQry.LoadProjectInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT ProjID, PName, Status, PrjType, StartDate, EndDate "
   SQL = SQL & "FROM Projects "
   SQL = SQL & "WHERE ProjID = " & lngProjID
   SQL = SQL & " ORDER BY PName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ProjID)) Then
               Set Item = lvResult.ListItems.Add(, "ID" & !ProjID, !PName)
            End If
            If (Not IsNull(!StartDate)) Then Item.SubItems(1) = Format(!StartDate, "mm/dd/yyyy")
            If (Not IsNull(!EndDate)) Then Item.SubItems(2) = Format(!EndDate, "mm/dd/yyyy")
            If (Not IsNull(!Status)) Then Item.SubItems(3) = !Status
            If (Not IsNull(!PrjType)) Then Item.SubItems(4) = !PrjType
            .MoveNext
         Wend
         MsgBar "Showing: " & .RecordCount & " Records (of " & .RecordCount & " total)", False
      Else
         MsgBar vbNullString, False
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub lvResult_Click()
   Const sMOD_NAME As String = "frmProjQry.lvResult_Click"
   On Error GoTo Error_Handler
   
   g_lngProjID = CLng(Mid$(lvResult.SelectedItem.Key, 3, Len(lvResult.SelectedItem.Key)))
   
   'code to open Contact entry screen
   UnloadAllForms
   Load frmProjEntry
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while loading the requested Project information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lblSaveAs_Click()
   'save the list as a CSV file
   Const sMOD_NAME As String = "frmProjQry.lblSaveAs_Click"
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
   
   Print #1, Chr(34) & "Project Name" & Chr(34); Chr(44) & _
         Chr(34) & "Job Title" & Chr(34); Chr(44) & _
         Chr(34) & "Start Date" & Chr(34); Chr(44) & _
         Chr(34) & "End Date" & Chr(34); Chr(44) & _
         Chr(34) & "Status" & Chr(34); Chr(44) & _
         Chr(34) & "Type" & Chr(34)
   For iCtr = 1 To lvResult.ListItems.Count
      lvResult.SelectedItem = lvResult.ListItems(iCtr)
      Print #1, Chr(34) & lvResult.SelectedItem & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(1) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(2) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(3) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(4) & Chr(34)
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
