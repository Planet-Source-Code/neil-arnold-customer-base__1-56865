VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToDoList 
   Caption         =   "To Do List"
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
   Icon            =   "frmToDoList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00EEE7E0&
      Height          =   390
      Left            =   3675
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   5
      Top             =   2475
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H00896A4B&
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
         BackColor       =   &H00896A4B&
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   55
         Width           =   2415
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "To Do -"
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
         TabIndex        =   4
         Top             =   105
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmToDoList.frx":0442
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "View:"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   1200
         TabIndex        =   3
         Top             =   105
         Width           =   540
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   9000
         Picture         =   "frmToDoList.frx":0884
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
         MouseIcon       =   "frmToDoList.frx":0A2E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   105
         Width           =   1140
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   7215
      Left            =   0
      TabIndex        =   6
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Project"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Due"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape shpResult 
      BorderColor     =   &H00CCCCB4&
      Height          =   990
      Left            =   900
      Top             =   1650
      Width           =   765
   End
End
Attribute VB_Name = "frmToDoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_strView As String

Private Sub cboView_Click()
   m_strView = cboView.Text
   Call LoadToDoInfo
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvResult, picGrdClr
   End If
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmToDoList.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading To Do List", True
   frmMain.picStatus.BackColor = &H896A4B
   
   'set up cboView
   With cboView
      .AddItem "All To Do's"
      .AddItem "Completed To Do's"
      .AddItem "Private To Do's"
      .AddItem "Urgent To Do's"
      .Text = "All To Do's"
   End With
   
   m_strView = "All To Do's"
   
   'set global form identifier
   g_strFormFlag = "ToDo"
   
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
   
   LockWindowUpdate frmToDoList.hWnd
   
   'adjust lvResult
   lvResult.Move 15, 480, Me.ScaleWidth - 30, Me.ScaleHeight - 495
   '***adjust lvResult column widths
   lvResult.ColumnHeaders(1).Width = lvResult.Width * 0.25 'name
   lvResult.ColumnHeaders(2).Width = lvResult.Width * 0.25  'job title
   lvResult.ColumnHeaders(3).Width = lvResult.Width * 0.25  'phone
   lvResult.ColumnHeaders(4).Width = (lvResult.Width * 0.25) - 265  'zip
   shpResult.Move lvResult.Left - 15, lvResult.Top - 15, lvResult.Width + 30, lvResult.Height + 30
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   Set frmToDoList = Nothing
End Sub

Public Sub LoadToDoInfo()
   'load the desired to do info into lvResult
   Const sMOD_NAME As String = "frmToDoList.LoadToDoInfo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strWhere As String
   Dim strOrder As String
   Dim strContact As String
   Dim strProject As String
   Dim Item As ListItem
   
   SQL = "SELECT RefNum, Subject, fkContID, fkProjID, DueDate, "
   SQL = SQL & "Private, Urgent, Completed FROM ToDo "
   
   Select Case m_strView
      Case "All To Do's"
         'do nothing
      Case "Completed To Do's"
         strWhere = "WHERE Completed = True "
         SQL = SQL & strWhere
      Case "Private To Do's"
         strWhere = "WHERE Private = True "
         SQL = SQL & strWhere
      Case "Urgent To Do's"
         strWhere = "WHERE Urgent = True "
         SQL = SQL & strWhere
   End Select
   
   strOrder = "ORDER BY DueDate"
   SQL = SQL & strOrder
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvResult.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               If (Not IsNull(!Subject)) Then
                  Set Item = lvResult.ListItems.Add(, "ID" & !RefNum, !Subject)
               End If
            End If
            If (Not IsNull(!fkContID)) Then
               strContact = ConvertContactName(!fkContID)
               Item.SubItems(1) = strContact
            End If
            If (Not IsNull(!fkProjID)) Then
               strProject = ConvertProjectName(!fkProjID)
               Item.SubItems(2) = strProject
            End If
            If (Not IsNull(!DueDate)) Then
               Item.SubItems(3) = Format(!DueDate, "ddd mm/dd/yyyy")
            End If
            
            .MoveNext
         Wend
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
   'save the list as a CSV file
   Const sMOD_NAME As String = "frmToDoList.lblSaveAs_Click"
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
   
   Print #1, Chr(34) & "Subject" & Chr(34); Chr(44) & _
         Chr(34) & "Name" & Chr(34); Chr(44) & _
         Chr(34) & "Project" & Chr(34); Chr(44) & _
         Chr(34) & "Due Date" & Chr(34)
   For iCtr = 1 To lvResult.ListItems.Count
      lvResult.SelectedItem = lvResult.ListItems(iCtr)
      Print #1, Chr(34) & lvResult.SelectedItem & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(1) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(2) & Chr(34); Chr(44) & _
      Chr(34) & lvResult.SelectedItem.SubItems(3) & Chr(34)
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
   Const sMOD_NAME As String = "frmHome.lvToDo_Click"
   On Error GoTo Error_Handler
   
   Dim lngToDo As Long
   
   lngToDo = CLng(Mid$(lvResult.SelectedItem.Key, 3, Len(lvResult.SelectedItem.Key)))
   
   'code to open ToDo entry screen
   icurState = NOW_EDITING
   frmToDo.m_lngToDoID = lngToDo
   Load frmToDo
   frmToDo.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while loading the requested To Do information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub


