VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjects 
   Caption         =   "Projects"
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
   Icon            =   "frmProjects.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00DEE6F0&
      Height          =   390
      Left            =   3600
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2925
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
      ScaleWidth      =   11355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11355
      Begin VB.ComboBox cboView 
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
         Left            =   2175
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   50
         Width           =   2565
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "- View:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   1575
         TabIndex        =   5
         Top             =   115
         Width           =   540
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   " All Projects"
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
         TabIndex        =   2
         Top             =   105
         Width           =   1140
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmProjects.frx":0442
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   9000
         Picture         =   "frmProjects.frx":66CC
         Stretch         =   -1  'True
         Top             =   100
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
         Left            =   9300
         MouseIcon       =   "frmProjects.frx":6876
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   105
         Width           =   1140
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   7215
      Left            =   0
      TabIndex        =   4
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
      Top             =   1800
      Width           =   765
   End
End
Attribute VB_Name = "frmProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_strView As String

Private Sub cboView_Click()
   m_strView = cboView.Text
   
   Call LoadProjectInfo
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmNames.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Projects List", True
   frmMain.picStatus.BackColor = &H2A59A0
   
   'set global form identifier
   g_strFormFlag = ""
   
   'load view combo
   With cboView
      .AddItem "All Projects"
      .AddItem "Default"
      .AddItem "Hidden"
      .AddItem "Private"
   End With
   
   'set default view setting
   m_strView = "Default"
   
   cboView.Text = m_strView
   
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
   
   LockWindowUpdate frmProjects.hWnd
   
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
   Set frmProjects = Nothing
End Sub

Private Sub LoadProjectInfo()
   'load all stored projects
   Const sMOD_NAME As String = "frmProjects.LoadProjectInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   Select Case m_strView
      Case "All Projects"
         SQL = "SELECT ProjID, PName, Status, PrjType, StartDate, EndDate "
         SQL = SQL & "FROM Projects ORDER BY PName"
      Case Else
         SQL = "SELECT ProjID, PName, Status, Setting, PrjType, "
         SQL = SQL & "StartDate, EndDate FROM Projects "
         SQL = SQL & "WHERE Setting = '" & m_strView & "' "
         SQL = SQL & "ORDER BY PName"
   End Select
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvResult.ListItems.Clear
   
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
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvResult, picGrdClr
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub lblSaveAs_Click()
   'save the list as a CSV file
   Const sMOD_NAME As String = "frmProjects.lblSaveAs_Click"
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

Private Sub lvResult_Click()
   Const sMOD_NAME As String = "frmProjects.lvResult_Click"
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
      MsgBox "An un-known error occurred while loading the Projects screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub
