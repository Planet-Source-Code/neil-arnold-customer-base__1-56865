VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHome 
   Caption         =   "Home Page"
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
   Icon            =   "frmHome.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvNotes 
      Height          =   1965
      Left            =   4275
      TabIndex        =   18
      Top             =   3825
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   3466
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Content"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvToDo 
      Height          =   2340
      Left            =   4275
      TabIndex        =   16
      Top             =   1050
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   4128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Due"
         Object.Width           =   1482
      EndProperty
   End
   Begin MSComctlLib.ListView lvAppts 
      Height          =   2490
      Left            =   7350
      TabIndex        =   12
      Top             =   3750
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   4877
      EndProperty
   End
   Begin VB.PictureBox picCal 
      BorderStyle     =   0  'None
      Height          =   2490
      Left            =   7350
      ScaleHeight     =   2490
      ScaleWidth      =   3990
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   3990
      Begin MSComCtl2.MonthView CalMain 
         Height          =   2460
         Left            =   225
         TabIndex        =   9
         Top             =   0
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   4339
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16448230
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MonthBackColor  =   14806501
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   53346305
         TitleBackColor  =   12046014
         TrailingForeColor=   12632256
         CurrentDate     =   38247
      End
   End
   Begin MSComctlLib.ListView lvProjects 
      Height          =   1665
      Left            =   75
      TabIndex        =   7
      Top             =   5850
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   2937
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project"
         Object.Width           =   4260
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2275
      EndProperty
   End
   Begin MSComctlLib.ListView lvRecNames 
      Height          =   4365
      Left            =   75
      TabIndex        =   5
      Top             =   1050
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   7699
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Contact Name"
         Object.Width           =   4260
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Phone"
         Object.Width           =   2275
      EndProperty
   End
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00EEF4F0&
      Height          =   390
      Left            =   10800
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   4
      Top             =   7275
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H00336600&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   0
      Width           =   11355
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmHome.frx":0442
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Base / Contact Manager - Home Page"
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
   End
   Begin VB.Shape shpNotes 
      BorderColor     =   &H00B7CEBE&
      Height          =   1215
      Left            =   4125
      Top             =   3975
      Width           =   615
   End
   Begin VB.Label lblNotes 
      BackColor       =   &H00B7CEBE&
      Caption         =   " Notes"
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
      Left            =   4275
      TabIndex        =   17
      Top             =   3600
      Width           =   2865
   End
   Begin VB.Shape shpToDo 
      BorderColor     =   &H00B7CEBE&
      Height          =   1365
      Left            =   6675
      Top             =   1125
      Width           =   615
   End
   Begin VB.Label lblToDoDueHdr 
      BackColor       =   &H00E1EDE5&
      Caption         =   " Due:"
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
      Left            =   6075
      TabIndex        =   15
      Top             =   825
      Width           =   1065
   End
   Begin VB.Label lblToDoSubjHdr 
      BackColor       =   &H00E1EDE5&
      Caption         =   " Subject:"
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
      Left            =   4275
      TabIndex        =   14
      Top             =   825
      Width           =   1815
   End
   Begin VB.Label lblToDo 
      BackColor       =   &H00B7CEBE&
      Caption         =   " To Do List"
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
      Left            =   4275
      TabIndex        =   13
      Top             =   600
      Width           =   2865
   End
   Begin VB.Shape shpAppts 
      BorderColor     =   &H00B7CEBE&
      Height          =   1065
      Left            =   7200
      Top             =   3900
      Width           =   465
   End
   Begin VB.Label lblApptsHdr 
      BackColor       =   &H00E1EDE5&
      Caption         =   " Date:             Subject:"
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
      Left            =   7350
      TabIndex        =   11
      Top             =   3525
      Width           =   3990
   End
   Begin VB.Label lblAppts 
      BackColor       =   &H00B7CEBE&
      Caption         =   " Appointments"
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
      Left            =   7350
      TabIndex        =   10
      Top             =   3300
      Width           =   3990
   End
   Begin VB.Shape shpProjects 
      BorderColor     =   &H00B7CEBE&
      Height          =   1140
      Left            =   3375
      Top             =   5850
      Width           =   840
   End
   Begin VB.Label lblProjects 
      BackColor       =   &H00B7CEBE&
      Caption         =   " All Projects"
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
      TabIndex        =   6
      Top             =   5625
      Width           =   3990
   End
   Begin VB.Shape shpRecNames 
      BorderColor     =   &H00B7CEBE&
      Height          =   1065
      Left            =   3300
      Top             =   1125
      Width           =   915
   End
   Begin VB.Label lblRecNameHdr 
      BackColor       =   &H00E1EDE5&
      Caption         =   " Contact Name:                             Phone:"
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
      Top             =   825
      Width           =   3990
   End
   Begin VB.Label lblRecNames 
      BackColor       =   &H00B7CEBE&
      Caption         =   " Recent Names"
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
      TabIndex        =   2
      Top             =   600
      Width           =   3990
   End
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_lngToDo As Long
Dim m_lngNotes As Long
Dim m_lngAppts As Long

Private Sub CalMain_DateClick(ByVal DateClicked As Date)
   frmCalDay.m_blnIsSystem = False
   frmCalDay.m_vSrchDate = CalMain.Value
   UnloadAllForms
   Load frmCalDay
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmHome.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Home Screen", True
   frmMain.picStatus.BackColor = &H336600
   
   'Set calendar date
   CalMain.Value = Date
   
   'Load all needed data
   Call LoadContactInfo
   Call LoadProjects
   Call LoadToDo
   Call LoadNotes
   Call LoadAppts
   
   'set screen flag
   g_strFormFlag = "Home"
   
   'set gridline preference
   lvAppts.GridLines = g_blnShowLines
   lvNotes.GridLines = g_blnShowLines
   lvProjects.GridLines = g_blnShowLines
   lvRecNames.GridLines = g_blnShowLines
   lvToDo.GridLines = g_blnShowLines
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   ShowError
End Sub

Private Sub Form_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
   
   On Error Resume Next
   
   LockWindowUpdate frmHome.hWnd
   
   'adjust recent names items
   lvRecNames.Height = (Me.ScaleHeight - 1095) * 0.7
   shpRecNames.Move lvRecNames.Left - 15, lvRecNames.Top - 15, lvRecNames.Width + 30, lvRecNames.Height + 30
   'adjust project items
   lblProjects.Top = lvRecNames.Top + lvRecNames.Height + 240
   lvProjects.Top = lblProjects.Top + lblProjects.Height
   lvProjects.Height = Me.ScaleHeight - 1710 - lvRecNames.Height
   shpProjects.Move lvProjects.Left - 15, lvProjects.Top - 15, lvProjects.Width + 30, lvProjects.Height + 30
   'adjust calendar items
   picCal.Move Me.ScaleWidth - picCal.Width - 75, 615
   CalMain.Move (picCal.ScaleWidth - CalMain.Width) / 2, (picCal.ScaleHeight - CalMain.Height) / 2
   'adjust appointment items
   lblAppts.Left = picCal.Left
   lblApptsHdr.Left = picCal.Left
   lvAppts.Left = picCal.Left
   shpAppts.Move lvAppts.Left - 15, lvAppts.Top - 15, lvAppts.Width + 30, lvAppts.Height + 30
   'adjust to do items
   lblToDo.Move lblRecNames.Left + lblRecNames.Width + 225, lblToDo.Top, Me.ScaleWidth - lblRecNames.Width - picCal.Width - 600
   lblToDoDueHdr.Move lblToDo.Left + lblToDo.Width - 1065, lblToDoDueHdr.Top
   lblToDoSubjHdr.Move lblToDo.Left, lblToDoDueHdr.Top, lblToDo.Width - 1065
   lvToDo.Move lblToDo.Left, lblToDoSubjHdr.Top + 240, lblToDo.Width, Me.ScaleHeight * 0.35
   shpToDo.Move lvToDo.Left - 15, lvToDo.Top - 15, lvToDo.Width + 30, lvToDo.Height + 30
   '   adjust lvToDo column header
   lvToDo.ColumnHeaders(1).Width = lvToDo.Width - 1125
   'adjust notes items
   lblNotes.Move lblToDo.Left, lvToDo.Top + lvToDo.Height + 225, lblToDo.Width
   lvNotes.Move lblNotes.Left, lblNotes.Top + 240, lblNotes.Width, Me.ScaleHeight * 0.35
   shpNotes.Move lvNotes.Left - 15, lvNotes.Top - 15, lvNotes.Width + 30, lvNotes.Height + 30
   '   adjust lvNotes column header
   lvNotes.ColumnHeaders(1).Width = lvNotes.Width - 285
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   Set frmHome = Nothing
End Sub

Public Sub LoadContactInfo()
   'load contact information into lvRecNames
   Const sMOD_NAME As String = "frmHome.LoadContactInfo"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim strPhone As String
   Dim strSetting As String
   
   strSetting = "Default"
   
   SQL = "SELECT TOP 22 ContID, Setting, ShownName FROM Contacts " 'added 10.28.04 [TOP 22]
   SQL = SQL & "WHERE Setting = '" & strSetting & "' "
   SQL = SQL & "ORDER BY ContID DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ContID)) Then
               Set Item = lvRecNames.ListItems.Add(, "ID" & !ContID, !ShownName)
               strPhone = GetPhoneNum(!ContID)
               Item.SubItems(1) = strPhone
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvRecNames, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadProjects()
   'load all open projects
   Const sMOD_NAME As String = "frmHome.LoadProjects"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim strStatus As String
   Dim strComp As String
   
   SQL = "SELECT ProjID, PName, Status "
   SQL = SQL & "FROM Projects "
   SQL = SQL & "ORDER BY PName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ProjID)) Then
               If (Not IsNull(!Status)) Then
                  strStatus = !Status
               End If
               Select Case strStatus
                  Case "Complete"
                     .MoveNext
                  Case "Closed"
                     .MoveNext
                  Case Else
                     Set Item = lvProjects.ListItems.Add(, "ID" & !ProjID, !PName)
                     Item.SubItems(1) = "* " & !Status
                     .MoveNext 'added 10.23.04 from user input (josimar silva)
               End Select
            Else 'added 10.23.04 from user input (josimar silva)
               .MoveNext 'added 10.23.04 from user input (josimar silva)
            End If
            '.MoveNext 'removed 10.23.04 from user input (josimar silva)
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvProjects, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadToDo()
   'load all to do items not completed
   Const sMOD_NAME As String = "frmHome.LoadToDo"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   
   SQL = "SELECT RefNum, Subject, DueDate, Completed, Private "
   SQL = SQL & "FROM ToDo WHERE Completed = False "
   SQL = SQL & "AND Private = False "
   SQL = SQL & "ORDER BY DueDate DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvToDo.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               Set Item = lvToDo.ListItems.Add(, "ID" & !RefNum, !Subject)
               Item.SubItems(1) = Format(!DueDate, "mm/dd")
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvToDo, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadNotes()
   'load all notes into lvNotes
   Const sMOD_NAME As String = "frmHome.LoadNotes"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim strType As String
   
   strType = "N"
   
   SQL = "SELECT RefNum, fkContID, fkProjID, NType, TextBody FROM Attach "
   SQL = SQL & "WHERE NType = '" & strType & "' "
   SQL = SQL & " AND fkContID <= " & 0
   SQL = SQL & " AND fkProjID <= " & 0
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvNotes.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               Set Item = lvNotes.ListItems.Add(, "ID" & !RefNum, !TextBody)
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvNotes, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub LoadAppts()
   'load all appointments in lvAppts
   Const sMOD_NAME As String = "frmHome.LoadAppts"
   On Error GoTo Error_Handler
   
   Dim Item As ListItem
   Dim SQL As String
   Dim vDate As Variant
   
   'vDate = Format(Date, "mm/dd/yyyy")
   'vDate = "#" & vDate & "#"
   'updated code 10.23.04 thanks to submission by michael doering
   vDate = Format$(CVDate(Date), "\#mm\/dd\/yyyy\#")
   
   SQL = "SELECT RefNum, Subject, DateFrom, Private FROM Appts "
   SQL = SQL & "WHERE DateFrom >= " & vDate
   SQL = SQL & " AND Private = False "
   SQL = SQL & "ORDER BY DateFrom" ' DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvAppts.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               Set Item = lvAppts.ListItems.Add(, "ID" & !RefNum, Format(!DateFrom, "ddd mm/dd"))
               Item.SubItems(1) = !Subject
            End If
            .MoveNext
         Wend
      End If
   End With
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvAppts, picGrdClr
   End If
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub lvAppts_Click()
   Const sMOD_NAME As String = "frmHome.lvAppts_Click"
   On Error GoTo Error_Handler
   
   m_lngAppts = CLng(Mid$(lvAppts.SelectedItem.Key, 3, Len(lvAppts.SelectedItem.Key)))
   
   'code to open Appts entry screen
   icurState = NOW_EDITING
   frmAppt.m_lngApptID = m_lngAppts
   Load frmAppt
   frmAppt.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while opening the Appointments dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvNotes_Click()
   Const sMOD_NAME As String = "frmHome.lvNotes_Click"
   On Error GoTo Error_Handler
   
   m_lngNotes = CLng(Mid$(lvNotes.SelectedItem.Key, 3, Len(lvNotes.SelectedItem.Key)))
   
   'code to open Notes entry screen
   icurState = NOW_EDITING
   frmNotes.m_lngNoteID = m_lngNotes
   Load frmNotes
   frmNotes.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while opening the Notes/Calls dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvProjects_Click()
   Const sMOD_NAME As String = "frmHome.lvProjects_Click"
   On Error GoTo Error_Handler
   
   g_lngProjID = CLng(Mid$(lvProjects.SelectedItem.Key, 3, Len(lvProjects.SelectedItem.Key)))
   
   'code to open Project entry screen
   UnloadAllForms
   Load frmProjEntry
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while opening the Projects screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub lvRecNames_Click()
   Const sMOD_NAME As String = "frmHome.lvRecNames_Click"
   On Error GoTo Error_Handler
   
   g_lngContID = CLng(Mid$(lvRecNames.SelectedItem.Key, 3, Len(lvRecNames.SelectedItem.Key)))
   
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

Private Sub lvToDo_Click()
   Const sMOD_NAME As String = "frmHome.lvToDo_Click"
   On Error GoTo Error_Handler
   
   m_lngToDo = CLng(Mid$(lvToDo.SelectedItem.Key, 3, Len(lvToDo.SelectedItem.Key)))
   
   'code to open ToDo entry screen
   icurState = NOW_EDITING
   frmToDo.m_lngToDoID = m_lngToDo
   Load frmToDo
   frmToDo.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while opening the To Do dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub
