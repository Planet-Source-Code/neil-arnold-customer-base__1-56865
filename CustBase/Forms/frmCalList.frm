VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCalList 
   Caption         =   "Appointment List"
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
   Icon            =   "frmCalList.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00EFF2F2&
      Height          =   390
      Left            =   3075
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   12
      Top             =   4050
      Visible         =   0   'False
      Width           =   390
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   6240
      Left            =   150
      TabIndex        =   11
      Top             =   1425
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   11007
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Project"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox cboFiltOpt 
      BackColor       =   &H00B1CBD4&
      Height          =   315
      Left            =   2100
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   2565
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      BackColor       =   &H004A4A4A&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   11355
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11355
      Begin MSComctlLib.TabStrip tbsMain 
         Height          =   315
         Left            =   8850
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   150
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Day"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Week"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Month"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "List"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblBanner 
         BackStyle       =   0  'Transparent
         Caption         =   "Calendar"
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
         Left            =   450
         TabIndex        =   2
         Top             =   75
         Width           =   840
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   75
         Picture         =   "frmCalList.frx":0442
         Stretch         =   -1  'True
         Top             =   75
         Width           =   315
      End
   End
   Begin VB.Shape shpResult 
      BorderColor     =   &H00E3E9EB&
      Height          =   690
      Left            =   3825
      Top             =   3525
      Width           =   1065
   End
   Begin VB.Label lblHdrPrj 
      BackColor       =   &H00E3E9EB&
      Caption         =   " Project"
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
      Left            =   8400
      TabIndex        =   10
      Top             =   1050
      Width           =   2415
   End
   Begin VB.Label lblHdrName 
      BackColor       =   &H00E3E9EB&
      Caption         =   " Name"
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
      Left            =   6000
      TabIndex        =   9
      Top             =   1050
      Width           =   2415
   End
   Begin VB.Label lblHdrTo 
      BackColor       =   &H00E3E9EB&
      Caption         =   " To"
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
      Left            =   4575
      TabIndex        =   8
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Label lblHdrFrom 
      BackColor       =   &H00E3E9EB&
      Caption         =   " From"
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
      Left            =   3150
      TabIndex        =   7
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Label lblHdrEvent 
      BackColor       =   &H00E3E9EB&
      Caption         =   " Event"
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
      TabIndex        =   6
      Top             =   1050
      Width           =   3015
   End
   Begin VB.Label lblSubHeader 
      BackColor       =   &H00B1CBD4&
      Caption         =   "View:"
      Height          =   315
      Index           =   1
      Left            =   1575
      TabIndex        =   4
      Top             =   600
      Width           =   465
   End
   Begin VB.Label lblSubHeader 
      BackColor       =   &H00B1CBD4&
      Caption         =   " Appointments -"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   600
      Width           =   11190
   End
End
Attribute VB_Name = "frmCalList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_strSrchOpt As String

Private Sub cboFiltOpt_Click()
   m_strSrchOpt = cboFiltOpt.Text
   
   Call LoadAppointments
   
   If (g_blnAltColors = True) Then
      AltLVBackground lvResult, picGrdClr
   End If
End Sub

Private Sub Form_Activate()
   tbsMain.Tabs(4).Selected = True
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmCalMnth.Form_Load"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar " Loading Calendar List View", True
   frmMain.picStatus.BackColor = &H4A4A4A
   
   'load cboFiltOpt
   With cboFiltOpt
      .AddItem "Upcoming Appointments"
      .AddItem "Past Appointments"
      .AddItem "All Appointments"
      .Text = "Upcoming Appointments"
   End With
   
   'set global from identifier
   g_strFormFlag = "CList"
   
   'set gridline preference
   lvResult.GridLines = g_blnShowLines
   
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
   
   LockWindowUpdate frmCalList.hWnd
   
   'adjust sub header
   lblSubHeader(0).Move 150, 615, Me.ScaleWidth - 300
   lblSubHeader(1).Move lblSubHeader(0).Left + 1425, lblSubHeader(0).Top
   cboFiltOpt.Move lblSubHeader(0).Left + 1950, lblSubHeader(0).Top
   'adjust header labels
   lblHdrEvent.Move lblSubHeader(0).Left, lblSubHeader(0).Top + 465, (Me.ScaleWidth - 300) * 0.24
   lblHdrFrom.Move lblHdrEvent.Left + lblHdrEvent.Width, lblHdrEvent.Top, (Me.ScaleWidth - 300) * 0.15
   lblHdrTo.Move lblHdrFrom.Left + lblHdrFrom.Width, lblHdrFrom.Top, (Me.ScaleWidth - 300) * 0.15
   lblHdrName.Move lblHdrTo.Left + lblHdrTo.Width, lblHdrTo.Top, (Me.ScaleWidth - 300) * 0.23
   lblHdrPrj.Move lblHdrName.Left + lblHdrName.Width, lblHdrName.Top, (Me.ScaleWidth - 300) * 0.23
   'adjust lvResult
   lvResult.Move lblHdrEvent.Left, lblHdrEvent.Top + 240, Me.ScaleWidth - 300, Me.ScaleHeight - 1785
   '***adjust lvResult column widths
   lvResult.ColumnHeaders(1).Width = (Me.ScaleWidth - 565) * 0.24
   lvResult.ColumnHeaders(2).Width = (Me.ScaleWidth - 565) * 0.15
   lvResult.ColumnHeaders(3).Width = (Me.ScaleWidth - 565) * 0.15
   lvResult.ColumnHeaders(4).Width = (Me.ScaleWidth - 565) * 0.23
   lvResult.ColumnHeaders(5).Width = (Me.ScaleWidth - 565) * 0.23
   'adjust shpResult
   shpResult.Move lvResult.Left - 15, lvResult.Top - 255, lvResult.Width + 30, lvResult.Height + 270
   
   LockWindowUpdate 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   Set frmCalList = Nothing
End Sub

Private Sub lvResult_Click()
   Const sMOD_NAME As String = "frmCalList.lvResult_Click"
   On Error GoTo Error_Handler
   
   Dim lngAppts As Long
   
   lngAppts = CLng(Mid$(lvResult.SelectedItem.Key, 3, Len(lvResult.SelectedItem.Key)))
   
   'code to open Appts entry screen
   icurState = NOW_EDITING
   frmAppt.m_lngApptID = lngAppts
   Load frmAppt
   frmAppt.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   If Err.Number = 91 Then
      MsgBox "There is nothing from the list to select", , APP_MSG_NAME
      Exit Sub
   Else
      LogErrors sMOD_NAME, Err.Number, Err.Description
      MsgBox "An un-known error occurred while populating the grid!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   End If
End Sub

Private Sub picBanner_Resize()
   tbsMain.Move picBanner.ScaleWidth - tbsMain.Width
End Sub

Private Sub tbsMain_Click()
   Select Case tbsMain.SelectedItem.Index
      Case 1 'Day
         UnloadAllForms
         frmCalDay.m_blnIsSystem = True
         Load frmCalDay
      Case 2 'Week
         UnloadAllForms
         Load frmCalWeek
      Case 3 'Month
         UnloadAllForms
         Load frmCalMnth
      Case 4 'List
         'take no action
   End Select
End Sub

Public Sub LoadAppointments()
   'load all desired appointments into the grid
   Dim SQL As String
   Dim strOrder As String
   Dim vDate As Variant
   Dim Item As ListItem
   Dim strContact As String
   Dim strProject As String
   
   vDate = Date
   vDate = "#" & vDate & "#"
   
   SQL = "SELECT RefNum, fkContID, fkProjID, Subject, DateFrom, DateTo "
   SQL = SQL & "FROM Appts "
   
   Select Case m_strSrchOpt
      Case "Upcoming Appointments"
         SQL = SQL & "WHERE DateFrom > " & vDate & " "
      Case "Past Appointments"
         SQL = SQL & "WHERE DateFrom < " & vDate & " "
      Case "All Appointments"
         SQL = SQL
   End Select
   
   strOrder = "ORDER BY DateFrom DESC"
   
   SQL = SQL & strOrder
   
   lvResult.ListItems.Clear
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!RefNum)) Then
               If (Not IsNull(!Subject)) Then
                  Set Item = lvResult.ListItems.Add(, "ID" & !RefNum, !Subject)
               End If
            End If
            If (Not IsNull(!DateFrom)) Then Item.SubItems(1) = !DateFrom
            If (Not IsNull(!DateTo)) Then Item.SubItems(2) = !DateTo
            If (Not IsNull(!fkContID)) Then
               strContact = ConvertContactName(!fkContID)
               Item.SubItems(3) = strContact
            End If
            If (Not IsNull(!fkProjID)) Then
               strProject = ConvertProjectName(!fkProjID)
               Item.SubItems(4) = strProject
            End If
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
End Sub

