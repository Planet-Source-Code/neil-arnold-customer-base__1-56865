VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmToDo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New/ Edit To Do"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmToDo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   4
      Left            =   6450
      TabIndex        =   23
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   3
      Left            =   5100
      TabIndex        =   22
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   2775
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Delete"
      Height          =   390
      Index           =   1
      Left            =   1425
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&New"
      Height          =   390
      Index           =   0
      Left            =   75
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3900
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00EAFFFF&
      ForeColor       =   &H00696969&
      Height          =   3015
      Index           =   1
      Left            =   4200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   525
      Width           =   3390
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Completed"
      Height          =   240
      Index           =   3
      Left            =   2700
      TabIndex        =   17
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Urgent"
      Height          =   240
      Index           =   2
      Left            =   1425
      TabIndex        =   16
      Top             =   3300
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Private"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   15
      Top             =   3300
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      ItemData        =   "frmToDo.frx":000C
      Left            =   1350
      List            =   "frmToDo.frx":003D
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2775
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remind Me:"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   12
      Top             =   2812
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      ItemData        =   "frmToDo.frx":00AF
      Left            =   2625
      List            =   "frmToDo.frx":0143
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1875
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   900
      TabIndex        =   9
      Top             =   1875
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Format          =   53739521
      CurrentDate     =   38252
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1425
      Width           =   3090
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmToDo.frx":0303
      Left            =   900
      List            =   "frmToDo.frx":0305
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   975
      Width           =   3090
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   900
      MaxLength       =   255
      TabIndex        =   6
      Top             =   525
      Width           =   3090
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   7650
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   75
      X2              =   7650
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Label Label1 
      Caption         =   "before"
      Height          =   240
      Index           =   5
      Left            =   2775
      TabIndex        =   14
      Top             =   2812
      Width           =   690
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Settings"
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
      Left            =   150
      TabIndex        =   11
      Top             =   2400
      Width           =   3840
   End
   Begin VB.Label Label1 
      Caption         =   "at:"
      Height          =   240
      Index           =   4
      Left            =   2325
      TabIndex        =   5
      Top             =   1912
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Due:"
      Height          =   240
      Index           =   3
      Left            =   150
      TabIndex        =   4
      Top             =   1912
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Project:"
      Height          =   240
      Index           =   2
      Left            =   150
      TabIndex        =   3
      Top             =   1462
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   1012
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Subject:"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   562
      Width           =   690
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H00DEE3E6&
      Caption         =   " To Do"
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
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   7440
   End
End
Attribute VB_Name = "frmToDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const TextNote = " [enter description / notes]"

Dim rsToDo As Recordset 'main recordset
Dim rsRemind As Recordset 'for reminders
Dim rsList As Recordset 'all other data work

Dim m_strOnEnter As String
Dim m_strOnExit As String
Dim m_lngContID As Long
Dim m_lngProjID As Long
Dim m_vStDate As Variant
Dim m_vStTime As Variant
Dim m_vRemDate As Variant
Dim m_vRemTime As Variant
Dim m_strRemInt As String
Dim m_lngRemindID As Long
Dim m_vCurrentDate As Variant
Dim m_intMinBefore As Integer
Dim m_blnCancelled As Boolean
Dim m_blnIsClearing As Boolean
Dim m_lngNewID As Long

Public m_lngToDoID As Long 'for id transfer

Private Sub Check1_Click(Index As Integer)
   If Index = 0 Then
      If ((Combo1(2).Enabled = False) And (Check1(0).Value = 1)) Then
         MsgBox "You must select a Due Time first", , APP_MSG_NAME
         Check1(0).Value = 0
         Exit Sub
      ElseIf (Combo1(2).Enabled = True) Then
         Combo1(3).Enabled = Check1(0).Value
         If (Combo1(3).Enabled = True) Then
            Combo1(3).BackColor = vbWhite
            Combo1(3).Text = "5 Min"
         Else
            Combo1(3).BackColor = vbButtonFace
         End If
      End If
   End If
End Sub

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmToDo.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Dim iMsg As VbMsgBoxResult
   
   Select Case Index
      Case 0 'New
         Call SetupNewRecord
      Case 1 'Delete
         Call DeleteCurrentRecord
      Case 2 'Print
         iMsg = MsgBox("Print this To Do item on printer " & Printer.DeviceName, _
            vbQuestion + vbYesNo, "Confirm Print Record")
         
         If (iMsg <> vbYes) Then Exit Sub
         
         Call PrintPage
      Case 3 'OK
         If (Not ValidateEntry()) Then Exit Sub
         
         'if no date is selected use current date
         If (dtpDate.Value = Format(m_vCurrentDate, "mm/dd/yyyy")) Then
            m_vStDate = Format(m_vCurrentDate, "mm/dd/yyyy")
         End If
         
         Call PostEntry
      Case 4 'Cancel
         m_blnCancelled = True
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Combo1_Click(Index As Integer)
   Const sMOD_NAME As String = "frmToDo.Combo1_Click"
   On Error GoTo Error_Handler
   
   If Index = 0 Then
      m_lngContID = Combo1(0).ItemData(Combo1(0).ListIndex)
   End If
   If Index = 1 Then
      m_lngProjID = Combo1(1).ItemData(Combo1(1).ListIndex)
   End If
   If Index = 2 Then
      If (m_blnIsClearing = False) Then
         If (Combo1(2).Text = "<Select>") Then
            MsgBox "You must select a Valid Time", , APP_MSG_NAME
            Combo1(2).Text = "8:00 AM"
            Combo1(2).SetFocus
         Else
            m_vStTime = Format(Combo1(2).Text, "h:nn AMPM")
         End If
      End If
   End If
   If Index = 3 Then
      m_strRemInt = Combo1(3).Text
      If (m_blnIsClearing = False) Then
         If (m_strRemInt = "<Select>") Then
            MsgBox "You must select a Valid Time Interval", , APP_MSG_NAME
            m_strRemInt = ""
            Combo1(3).Text = "5 Min"
            Combo1(3).SetFocus
            Exit Sub
         End If
      End If
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub dtpDate_CloseUp()
   m_vStDate = dtpDate.Value
   Combo1(2).Enabled = True
   Combo1(2).BackColor = vbWhite
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmToDo.Form_Load"
   On Error GoTo Error_Handler
   
   'set main recordsets
   Set rsToDo = dbContact.OpenRecordset("ToDo", dbOpenTable)
   Set rsRemind = dbContact.OpenRecordset("Remind", dbOpenTable)
   
   'set date picker to today
   dtpDate.Value = Date
   
   'set current date
   m_vCurrentDate = Date
   m_vCurrentDate = Format(m_vCurrentDate, "mm/dd/yyyy")
   m_vStDate = m_vCurrentDate
   
   'setup all combo's
   m_blnIsClearing = True
   Combo1(0).AddItem " "
   Combo1(1).AddItem " "
   Combo1(2).AddItem "<Select>"
   Combo1(2).Text = "<Select>"
   Combo1(3).AddItem "<Select>"
   Combo1(3).Text = "<Select>"
   m_blnIsClearing = False
   
   'flatten all needed borders
   Dim Indx As Integer
   
   For Indx = 0 To 3
      FlatBorder Combo1(Indx).hWnd
   Next
   For Indx = 0 To 1
      FlatBorder Text1(Indx).hWnd
   Next
   FlatBorder dtpDate.hWnd
   
   Text1(1).Text = TextNote
   
   'set up screen
   Call InitializeScreen
   
   m_blnCancelled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Const sMOD_NAME As String = "frmToDo.Form_Unload"
   On Error GoTo Error_Handler
   
   'remove data & form reference
   rsToDo.Close
   Set rsToDo = Nothing
   rsRemind.Close
   Set rsRemind = Nothing
   
   If (m_blnCancelled = False) Then
      Select Case g_strFormFlag
         Case "Home"
            Call frmHome.LoadToDo
         Case "CEnt"
            Call frmContEntry.LoadToDoInfo
         Case "PEnt"
            Call frmProjEntry.LoadToDoInfo
         Case "ToDo"
            Call frmToDoList.LoadToDoInfo
      End Select
   End If
   
   Set frmToDo = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Public Sub InitializeScreen()
   'set up the opening screen
   Const sMOD_NAME As String = "frmToDo.InitializeScreen"
   On Error GoTo Error_Handler
   
   Call LoadContactNames(Combo1(0))
   Call LoadProjectNames(Combo1(1))
   
   If (icurState = NOW_ADDING) Then
      If (g_strFormFlag = "CEnt") Then
         Call GetPersonalContName
      ElseIf (g_strFormFlag = "PEnt") Then
         Call GetPersonalProjName
      End If
   End If
   If (icurState = NOW_EDITING) Then
      With rsToDo
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngToDoID
            
            Call PopulateFields
            cmdOpts(1).Enabled = True
            cmdOpts(2).Enabled = True
         End If
      End With
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub PopulateFields()
   'load the desired record onto the screen
   Const sMOD_NAME As String = "frmToDo.PopulateFields"
   On Error GoTo Error_Handler
   
   Dim strContact As String
   Dim strProject As String
   
   With rsToDo
      If (Not IsNull(!Subject)) Then Text1(0) = !Subject
      'add code to retrieve project name or contact name
      'If (Not IsNull(!fkContID)) Then
      If (!fkContID > 0) Then
         m_lngContID = !fkContID
         strContact = ConvertContactName(m_lngContID)
         Combo1(0).Text = Trim(strContact)
      End If
      'If (Not IsNull(!fkProjID)) Then
      If (!fkProjID > 0) Then
         m_lngProjID = !fkProjID
         strProject = ConvertProjectName(m_lngProjID)
         Combo1(1).Text = strProject
      End If
      If (Not IsNull(!DueDate)) Then
         m_vStDate = !DueDate
         dtpDate = !DueDate
      End If
      If (Not IsNull(!DueTime)) Then
         Combo1(2).Text = Format(!DueTime, "h:nn AMPM")
         m_vStTime = !DueTime
         Combo1(2).Enabled = True
         Combo1(2).BackColor = vbWhite
      End If
      If (Not IsNull(!TextBody)) Then
         Text1(1) = !TextBody
         Text1(1).ForeColor = vbBlack
      End If
      If (!Remind = True) Then
         Check1(0).Value = 1
         If (Not IsNull(!RemAmt)) Then Combo1(3).Text = !RemAmt
         Call GetReminderID
      Else
         Check1(0).Value = 0
      End If
      If (!Private = True) Then
         Check1(1).Value = 1
      Else
         Check1(1).Value = 0
      End If
      If (!Urgent = True) Then
         Check1(2).Value = 1
      Else
         Check1(2).Value = 0
      End If
      If (!Completed = True) Then
         Check1(3).Value = 1
      Else
         Check1(3).Value = 0
      End If
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the records!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   m_strOnEnter = Text1(Index).Text
   If Index = 0 Then
      highLight
   End If
   If Index = 1 Then
      If (m_strOnEnter = TextNote) Then
         Text1(1).Text = ""
         Text1(1).ForeColor = vbBlack
      End If
   End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   m_strOnExit = Text1(Index).Text
   
   If Index = 1 Then
      If (m_strOnExit = "") Then
         Text1(Index).Text = TextNote
         Text1(Index).ForeColor = &H696969
      End If
   End If
End Sub

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmToDo.PostEntry"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting To Do Entry", True
   
   If (icurState = NOW_ADDING) Then
      rsToDo.AddNew
   Else
      With rsToDo
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngToDoID
            If Not .NoMatch Then
               rsToDo.Edit
            Else
               MsgBox "Error Posting Record", , "Posting Error"
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
      End With
   End If
   
   With rsToDo
      If (Len(Text1(0))) Then !Subject = Text1(0)
      If (Len(Text1(1))) Then
         If (Text1(1).Text <> TextNote) Then
            !TextBody = Text1(1)
         End If
      End If
      
      If (Not IsNull(m_vStDate)) Then !DueDate = m_vStDate
      If (m_lngContID > 0) Then !fkContID = m_lngContID
      If (m_lngProjID > 0) Then !fkProjID = m_lngProjID
      If (Len(Combo1(2).Text)) Then
         If (Combo1(2).Text <> "<Select>") Then
            !DueTime = Format(Combo1(2).Text, "hh:nn AMPM")
         End If
      End If
      
      If (Check1(0).Value = 1) Then
         !Remind = True
         If (m_strRemInt <> "") Then !RemAmt = m_strRemInt
      Else
         !Remind = False
      End If
      If (Check1(1).Value = 1) Then
         !Private = True
      Else
         !Private = False
      End If
      If (Check1(2).Value = 1) Then
         !Urgent = True
      Else
         !Urgent = False
      End If
      If (Check1(3).Value = 1) Then
         !Completed = True
      Else
         !Completed = False
      End If
      
      .Update
   End With
   
   If (Check1(0).Value = 1) Then 'modified 10.23.04
      Call GetRemindTimes(m_strRemInt)
      If (icurState = NOW_ADDING) Then
         Call GetLatestToDoID
         Call PostReminder
      ElseIf ((icurState = NOW_EDITING) And (m_lngRemindID = 0)) Then
         icurState = NOW_ADDING
         m_lngNewID = m_lngToDoID
         Call PostReminder
      Else
         Call PostReminder
      End If
      'Call PostReminder
   End If
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Me.Hide
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting the records!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(Text1(0)) < 1) Then
      Indx = MsgBox("You Must Enter A To Do Subject", _
         vbInformation + vbOKOnly, "Validate : To Do Subject")
      Text1(0).SetFocus
      ValidateEntry = False
      Exit Function
   End If
   If (m_vStDate = vbNullString) Then
      Indx = MsgBox("You Must Select A Due Date", _
         vbInformation + vbOKOnly, "Validate : Due Date")
      dtpDate.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub SetupNewRecord()
   'reset the screen for adding a new record
   Dim Indx As Integer
   
   m_blnIsClearing = True
   
   'clear textboxes
   Text1(0).Text = ""
   Text1(1).Text = TextNote
   Text1(1).ForeColor = &H696969
   'clear combos
   Combo1(0).Text = " "
   Combo1(1).Text = " "
   Combo1(2).Text = "<Select>"
   Combo1(2).Enabled = False
   Combo1(2).BackColor = vbButtonFace
   Combo1(3).Text = "<Select>"
   Combo1(3).Enabled = False
   Combo1(3).BackColor = vbButtonFace
   'set date picker to today
   dtpDate.Value = Date
   'reset checkboxes
   For Indx = 0 To 3
      Check1(Indx).Value = 0
   Next
   'reset editmode
   icurState = NOW_ADDING
   cmdOpts(1).Enabled = False
   cmdOpts(2).Enabled = False
   Text1(0).SetFocus
   
   m_blnIsClearing = False
End Sub

Private Sub DeleteCurrentRecord()
   'delete the current to do item
   Const sMOD_NAME As String = "frmToDo.DeleteCurrentRecord"
   On Error GoTo Error_Handler
   
   Dim iMsg As VbMsgBoxResult
   Dim sMsg As String
   Dim SQL As String
   
   If (icurState = NOW_ADDING) Then
      MsgBox "There is no current record to delete.", , APP_MSG_NAME
      Exit Sub
   End If
   
   sMsg = "Are you sure you want to DELETE this To Do item?"
   
   iMsg = MsgBox(sMsg, vbCritical + vbYesNo, "Warning : Record Deletion")
   
   If (iMsg <> vbYes) Then Exit Sub
   
   'delete this To Do item
   SQL = "DELETE * FROM ToDo WHERE RefNum = " & m_lngToDoID
   dbContact.Execute (SQL)
   
   'delete the reminder, if there is one
   SQL = "DELETE * FROM Remind WHERE RefNum = " & m_lngRemindID
   dbContact.Execute (SQL)
   
   Me.Hide
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Deleting the record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub GetPersonalContName()
   'if adding an appt from the Contact Entry screen lookup the current
   'contacts name
   Dim strContact As String
   
   strContact = ConvertContactName(g_lngContID)
   Combo1(0).Text = Trim(strContact)
End Sub

Private Sub GetPersonalProjName()
   'if adding an appt from the Project Entry screen lookup the current
   'projects name
   Dim strProject As String
   
   strProject = ConvertProjectName(g_lngProjID)
   Combo1(1).Text = Trim(strProject)
End Sub

Private Sub PrintPage()
   'print the current note/call
   Dim strTitle As String
   Dim vCurDate As Variant, vCurTime As Variant
   Dim strGrdProj As String, strGrdCont As String, strGrdDesc As String
   Dim strGrdDue As String, strGrdUrgent As String, strGrdComp As String
   Dim strUrgent As String, strComplete As String, strDueTime As String
   
   'set page title
   If (Text1(0).Text = "") Then
      strTitle = "To Do"
   Else
      strTitle = Text1(0).Text
   End If
   'set date and time to current Date/Time
   vCurDate = Format(Date, "m/dd/yy")
   vCurTime = Format(Time, "h:nn AMPM")
   
   'set left header captions
   strGrdProj = "Project"
   strGrdCont = "Contact"
   strGrdDue = "Due"
   strGrdUrgent = "Urgent"
   strGrdComp = "Completed"
   strGrdDesc = "Description"
   
   'set urgent & complete factors
   If (Check1(2).Value = 1) Then
      strUrgent = "Yes"
   Else
      strUrgent = "No"
   End If
   If (Check1(3).Value = 1) Then
      strComplete = "Yes"
   Else
      strComplete = "No"
   End If
   
   Printer.ScaleMode = vbCentimeters
   
   Printer.FontName = "Tahoma"
   Printer.FontSize = 10
   Printer.FontBold = False
   Printer.CurrentX = 1.3
   Printer.CurrentY = 1.5
   Printer.Print vCurTime;
   Printer.CurrentY = 1.7
   Printer.FontSize = 14
   Printer.FontBold = False
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTitle)) / 2
   Printer.Print strTitle
   Printer.CurrentX = 1.3
   Printer.CurrentY = 1.9
   Printer.FontSize = 10
   Printer.FontBold = False
   Printer.Print vCurDate
   
   Printer.FontSize = 8
   Printer.FontBold = False
   
   'print row 1 cell (project)
   Printer.Line (1.3, 2.6)-(20.3, 2.6) 'top line
   Printer.Line (1.3, 2.6)-(1.3, 3.08) 'left line (2.6 + .48)
   Printer.Line (2.95, 2.6)-(2.95, 3.08) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 2.6)-(20.3, 3.08) '3rd left line
   Printer.Line (1.3, 3.08)-(20.3, 3.08) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 2.65 '(2.6 + .05)
   Printer.Print strGrdProj;
   Printer.CurrentX = 3 '(2.95 + .05)
   Printer.Print Combo1(1).Text
   'print row 2 cell (contact)
   Printer.Line (1.3, 3.08)-(1.3, 3.56) 'left line (3.08 + .48)
   Printer.Line (2.95, 3.08)-(2.95, 3.56) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 3.08)-(20.3, 3.56) '3rd left line
   Printer.Line (1.3, 3.56)-(20.3, 3.56) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 3.13 '(3.08 + .05)
   Printer.Print strGrdCont;
   Printer.CurrentX = 3 '(2.95 + .05)
   Printer.Print Combo1(0).Text
   'print row 3 cell (due)
   Printer.Line (1.3, 3.56)-(1.3, 4.04) 'left line (3.56 + .48)
   Printer.Line (2.95, 3.56)-(2.95, 4.04) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 3.56)-(20.3, 4.04) '3rd left line
   Printer.Line (1.3, 4.04)-(20.3, 4.04) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 3.61 '(3.56 + .05)
   Printer.Print strGrdDue;
   Printer.CurrentX = 3 '(2.95 + .05)
   If (Combo1(2).Text = "<Select>") Then
      strDueTime = "[No Time Selected]"
   Else
      strDueTime = Combo1(2).Text
   End If
   Printer.Print dtpDate.Value & " at " & strDueTime
   'print row 4 cell (urgent)
   Printer.Line (1.3, 4.04)-(1.3, 4.52) 'left line (4.04 + .48)
   Printer.Line (2.95, 4.04)-(2.95, 4.52) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 4.04)-(20.3, 4.52) '3rd left line
   Printer.Line (1.3, 4.52)-(20.3, 4.52) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 4.09 '(4.04 + .05)
   Printer.Print strGrdUrgent;
   Printer.CurrentX = 3 '(2.95 + .05)
   Printer.Print strUrgent
   'print row 5 cell (completed)
   Printer.Line (1.3, 4.52)-(1.3, 5) 'left line (4.52 + .48)
   Printer.Line (2.95, 4.52)-(2.95, 5) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 4.52)-(20.3, 5) '3rd left line
   Printer.Line (1.3, 5)-(20.3, 5) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 4.57 '(4.52 + .05)
   Printer.Print strGrdComp;
   Printer.CurrentX = 3 '(2.95 + .05)
   Printer.Print strComplete
   'print row 6 cell (description)
   Printer.Line (1.3, 5)-(1.3, 14) 'left line (4.52 + .48)[13.52 was 5]
   Printer.Line (2.95, 5)-(2.95, 14) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 5)-(20.3, 14) '3rd left line
   Printer.Line (1.3, 14)-(20.3, 14) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 5.05 '(5 + .05)
   Printer.Print strGrdDesc;
   Printer.CurrentX = 3 '(2.95 + .05)
   Call WrapPrintText(Text1(1).Text)
   
   Dim strFoot As String
   
   Printer.CurrentY = 26.3
   Printer.CurrentX = 18.5
   strFoot = "Page " & CStr(Printer.Page)
   Printer.Print strFoot
   
   Printer.EndDoc
End Sub

Private Sub GetReminderID()
   'get the reminder record ID in case it needs to be updated or removed
   Const sMOD_NAME As String = "frmToDo.GetReminderID"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT RefNum, fkToDoID FROM Remind "
   SQL = SQL & "WHERE fkToDoID = " & m_lngToDoID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!RefNum)) Then m_lngRemindID = !RefNum
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub PostReminder()
   'post the reminder record
   Const sMOD_NAME As String = "frmToDo.PostReminder"
   On Error GoTo Error_Handler
   
   Dim strType As String
   
   strType = "TD"
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting To Do Reminder Entry", True
   
   If (icurState = NOW_ADDING) Then
      rsRemind.AddNew
   Else
      With rsRemind
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngRemindID
            If Not .NoMatch Then
               rsRemind.Edit
            Else
               MsgBox "Error Posting Record", , "Posting Error"
               LogErrors sMOD_NAME, Err.Number, Err.Description
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
      End With
   End If
   
   With rsRemind
      If (Not IsNull(m_vRemDate)) Then !RemDate = m_vRemDate
      If (Not IsNull(m_vRemTime)) Then !RemTime = m_vRemTime
      
      If (icurState = NOW_ADDING) Then
         If (m_lngNewID > 0) Then !fkToDoID = m_lngNewID
      Else
         !fkToDoID = m_lngToDoID
      End If
      
      !Type = strType
      
      If (icurState = NOW_EDITING) Then
         !Completed = False
      End If
      
      .Update
   End With
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Private Sub GetRemindTimes(strRemInt As String)
   Const sMOD_NAME As String = "frmToDo.GetRemindTimes"
   On Error GoTo Error_Handler
   
   Dim intHr As Integer
   Dim intMin As Integer
   
   Select Case strRemInt
      Case "5 Min"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If ((intHr = 0) And (intMin < 30)) Then '12 midnight
            intHr = 23
            intMin = 55
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intMin = intMin - 5
            m_vRemDate = m_vStDate
         End If
      Case "10 Min"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If ((intHr = 0) And (intMin < 30)) Then '12 midnight
            intHr = 23
            intMin = 50
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intMin = intMin - 10
            m_vRemDate = m_vStDate
         End If
      Case "15 Min"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If ((intHr = 0) And (intMin < 30)) Then '12 midnight
            intHr = 23
            intMin = 45
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intMin = intMin - 15
            m_vRemDate = m_vStDate
         End If
      Case "20 Min"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If ((intHr = 0) And (intMin < 30)) Then '12 midnight
            intHr = 23
            intMin = 40
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intMin = intMin - 20
            m_vRemDate = m_vStDate
         End If
      Case "30 Min"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If ((intHr = 0) And (intMin < 30)) Then '12 midnight
            intHr = 23
            intMin = 30
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf ((intHr = 0) And (intMin = 30)) Then '12:30 am
            intHr = 0
            intMin = 0
            m_vRemDate = m_vStDate
         Else
            intMin = intMin - 30
            m_vRemDate = m_vStDate
         End If
      Case "1 Hr"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If (intHr = 0) Then '12 midnight
            intHr = 23
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intHr = intHr - 1
            m_vRemDate = m_vStDate
         End If
      Case "2 Hr"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If (intHr = 0) Then '12 midnight
            intHr = 22
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 1) Then
            intHr = 23
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intHr = intHr - 2
            m_vRemDate = m_vStDate
         End If
      Case "3 Hr"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If (intHr = 0) Then '12 midnight
            intHr = 21
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 1) Then
            intHr = 22
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 2) Then
            intHr = 23
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intHr = intHr - 3
            m_vRemDate = m_vStDate
         End If
      Case "6 Hr"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         If (intHr = 0) Then '12 midnight
            intHr = 18
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 1) Then
            intHr = 19
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 2) Then
            intHr = 20
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 3) Then
            intHr = 21
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 4) Then
            intHr = 22
            m_vRemDate = DateValue(m_vStDate) - 1
         ElseIf (intHr = 5) Then
            intHr = 23
            m_vRemDate = DateValue(m_vStDate) - 1
         Else
            intHr = intHr - 6
            m_vRemDate = m_vStDate
         End If
      Case "12 Hr"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         Select Case intHr
            Case 0 '12 midnight
               intHr = 12
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 1
               intHr = 13
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 2
               intHr = 14
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 3
               intHr = 15
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 4
               intHr = 16
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 5
               intHr = 17
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 6
               intHr = 18
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 7
               intHr = 19
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 8
               intHr = 20
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 9
               intHr = 21
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 10
               intHr = 22
               m_vRemDate = DateValue(m_vStDate) - 1
            Case 11
               intHr = 23
               m_vRemDate = DateValue(m_vStDate) - 1
            Case Else '12 noon
               intHr = intHr - 12
               m_vRemDate = m_vStDate
         End Select
      Case "1 Day"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         m_vRemDate = DateValue(m_vStDate) - 1
      Case "2 Days"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         m_vRemDate = DateValue(m_vStDate) - 2
      Case "3 Days"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         m_vRemDate = DateValue(m_vStDate) - 3
      Case "1 Week"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         m_vRemDate = DateValue(m_vStDate) - 7
      Case "2 Weeks"
         intHr = Hour(m_vStTime)
         intMin = Minute(m_vStTime)
         m_vRemDate = DateValue(m_vStDate) - 14
   End Select
   
   m_vRemTime = TimeSerial(intHr, intMin, 0)
   m_vRemTime = Format(m_vRemTime, "h:nn AMPM")
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub GetLatestToDoID()
   'get the ID number of the To Do item just saved
   Const sMOD_NAME As String = "frmToDo.GetLatestToDoID"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT MAX(RefNum) AS RefID FROM ToDo"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!RefID)) Then m_lngNewID = !RefID
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub
