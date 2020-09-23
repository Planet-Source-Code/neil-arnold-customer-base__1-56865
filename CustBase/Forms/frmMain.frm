VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Customer Base - Contact Management System"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5850
      Top             =   3450
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BackColor       =   &H00336600&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   11850
      TabIndex        =   7
      Top             =   8175
      Width           =   11850
      Begin VB.PictureBox picTray 
         Height          =   390
         Left            =   10500
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   330
         ScaleWidth      =   330
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   " Ready ..."
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
         Left            =   75
         TabIndex        =   8
         Top             =   75
         Width           =   11715
      End
   End
   Begin MSComDlg.CommonDialog cdlMain 
      Left            =   6525
      Top             =   4050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   5775
      Top             =   4125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C12
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8758
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A14E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB60
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      _ExtentX        =   20902
      _ExtentY        =   794
      BandCount       =   1
      _CBWidth        =   11850
      _CBHeight       =   450
      _Version        =   "6.0.8169"
      Child1          =   "Picture1"
      MinHeight1      =   390
      Width1          =   3000
      NewRow1         =   0   'False
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   30
         ScaleHeight     =   390
         ScaleWidth      =   11730
         TabIndex        =   1
         Top             =   30
         Width           =   11730
         Begin VB.TextBox txtSrch 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   8175
            TabIndex        =   2
            Top             =   0
            Width           =   3315
         End
         Begin MSComctlLib.Toolbar tbrGo 
            Height          =   330
            Left            =   11475
            TabIndex        =   3
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   582
            ButtonWidth     =   1138
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imlMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   1
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Find"
                  Object.ToolTipText     =   "Find Entered Name"
                  ImageIndex      =   10
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar tbrMain 
            Height          =   330
            Left            =   450
            TabIndex        =   4
            Top             =   0
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   582
            ButtonWidth     =   1614
            ButtonHeight    =   582
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imlMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "New"
                  Key             =   "New"
                  Object.ToolTipText     =   "Add A New Item"
                  ImageIndex      =   4
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   6
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "drpName"
                        Text            =   "Name ..."
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "drpPrj"
                        Text            =   "Project ..."
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "drpSep1"
                        Text            =   "-"
                     EndProperty
                     BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "drpNote"
                        Text            =   "Note ..."
                     EndProperty
                     BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "drpToDo"
                        Text            =   "To Do ..."
                     EndProperty
                     BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "drpAppt"
                        Text            =   "Appointment ..."
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Names"
                  Key             =   "Names"
                  Object.ToolTipText     =   "View The Names Listing Screen"
                  ImageIndex      =   5
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Project"
                  Key             =   "Prj"
                  Object.ToolTipText     =   "View The Projects Screen"
                  ImageIndex      =   6
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Sched"
                  Key             =   "Cal"
                  Object.ToolTipText     =   "View The Calendar Screen"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "To Do"
                  Key             =   "ToDo"
                  Object.ToolTipText     =   "View The To Do List Screen"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Print"
                  Key             =   "Print"
                  Object.ToolTipText     =   "Print The Currently Active Item"
                  ImageIndex      =   9
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar tbrNav 
            Height          =   330
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            Style           =   1
            ImageList       =   "imlMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Home"
                  Object.ToolTipText     =   "Go To Home Screen"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Name Search :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7050
            TabIndex        =   6
            Top             =   45
            Width           =   1065
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileDBUtil 
         Caption         =   "Database Utilities"
         Begin VB.Menu mnuFileCompDB 
            Caption         =   "Compact Database"
         End
         Begin VB.Menu mnuFileRepDB 
            Caption         =   "Repair Database"
         End
         Begin VB.Menu mnuFileSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileBackupDB 
            Caption         =   "Backup Data ..."
         End
         Begin VB.Menu mnuFileSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileRestoreDB 
            Caption         =   "Restore Data ..."
         End
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete Item"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditNew 
         Caption         =   "New"
         Begin VB.Menu mnuEditNewName 
            Caption         =   "Name ..."
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuEditNewProject 
            Caption         =   "Project ..."
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuEditSep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditNewNote 
            Caption         =   "Note ..."
            Shortcut        =   ^O
         End
         Begin VB.Menu mnuEditNewToDo 
            Caption         =   "To Do ..."
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuEditNewAppt 
            Caption         =   "Appointment ..."
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu mnuEditSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPrefer 
         Caption         =   "Preferences"
      End
      Begin VB.Menu mnuEditPwd 
         Caption         =   "Set/Edit Password"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewHome 
         Caption         =   "Home"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewNames 
         Caption         =   "Names"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewProj 
         Caption         =   "Projects"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewCal 
         Caption         =   "Calendar"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuViewToDo 
         Caption         =   "To Do"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewCQry 
         Caption         =   "Contact Query (User Defined)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuViewPQry 
         Caption         =   "Project Query (User Defined)"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpFiles 
         Caption         =   "Contact Management Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpErrLog 
         Caption         =   "Error Log"
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About ..."
      End
   End
   Begin VB.Menu mnuDelContFld 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDelContactFld 
         Caption         =   "Delete..."
      End
   End
   Begin VB.Menu mnuDelProjFld 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuDelProjectFld 
         Caption         =   "Delete ..."
      End
   End
   Begin VB.Menu mnuTrayPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuTrayPopRestore 
         Caption         =   "Restore Customer Base - Main Screen"
      End
      Begin VB.Menu mnuTrayPopSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayPopClose 
         Caption         =   "Close Customer Base"
      End
   End
   Begin VB.Menu mnuRCMemoPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuRCMemoPopEdit 
         Caption         =   "Edit Related Contact Memo ?"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Public m_lngContUserFld 'for deleting contact user fields
Public m_strContUserFld 'for deleting contact user fields
Public m_lngProjUserFld 'for deleting project user fields
Public m_strProjUserFld 'for deleting project user fields
Public m_blnEmerShutdown As Boolean 'if password fails, allow to close
                                    'with no stops
Public m_lngRelContID As Long 'for editing related contact memo
Public m_strOldRCMemo As String 'for existing related contact memo

'used for reminder function
Dim m_blnAllowClose As Boolean
Dim m_vSysTime As Variant
Dim m_vSysDate As Variant
Dim m_vDBTime As Variant
Dim m_vDBDate As Variant

Private Sub MDIForm_Load()
   'Load form
   Const sMOD_NAME As String = "frmMain.MDIForm_Load"
   On Error GoTo Error_Handler
   
   Dim X As Integer
   
   'check to see if program is already running, if it is just END
   If App.PrevInstance = True Then
      MsgBox "Customer Base - Contact Manager is already running.", _
         vbCritical + vbOKOnly, APP_MSG_NAME
      End
   End If
   
   App.TaskVisible = False
   
   'get form coordinates
   X = Val(GetRegistryString("WindowState", "2"))
   If X = vbMaximized Then
      Show
   ElseIf X <> vbMinimized Then
      frmMain.WindowState = X
   Else
      frmMain.WindowState = 0
   End If
   If frmMain.WindowState = 0 Then
      frmMain.Left = Val(GetRegistryString("WindowLeft", "1500"))
      frmMain.Top = Val(GetRegistryString("WindowTop", "1500"))
      Show
      frmMain.Width = Val(GetRegistryString("WindowWidth", "12000"))
      frmMain.Height = Val(GetRegistryString("WindowHeight", "9000"))
   End If
   
   On Error GoTo Error_Handler
   
   'login to Jet
   On Error Resume Next
   Set g_wsWorkSpc = DBEngine.CreateWorkspace("MainWS", "admin", vbNullString)
   If Err <> 0 Then
      ShowError
      Unload Me
      Exit Sub
   End If
   
   On Error GoTo Error_Handler
   
   'add the workspace to the collection to bump the count
   Workspaces.Append g_wsWorkSpc
   Me.Show
   LoadRegistrySettings
   
   HideMainTools
   
   m_vSysDate = Date
   m_vSysDate = Format(m_vSysDate, "mm/dd/yyyy")
   
   Call OpenLocalDB
   
   'check for database updates
   Call CheckForCurrentDB
   
   'check for password protection
   If (g_blnIsSecure = True) Then
      Load frmPassLogon
      frmPassLogon.Show vbModal, frmMain
      If (m_blnEmerShutdown = True) Then
         Unload Me
         Exit Sub
      End If
   End If
   
   'check for any reminders for this date, and set cancel close flag
   'if any exist
   Call CheckReminderList
   'load Home screen
   Load frmHome
   
   'start checking for reminders
   Timer1.Enabled = True
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   On Error Resume Next
   
   Dim iMsg As VbMsgBoxResult
   
   If (m_blnEmerShutdown = True) Then
      Cancel = False
      Exit Sub
   End If
      
   If (m_blnAllowClose = True) Then
      iMsg = MsgBox("Close Customer Base?", vbQuestion + vbYesNo, "Confirm Close")
      
      If (iMsg <> vbYes) Then
         Cancel = True
         Exit Sub
      Else
         ShutDownCustBase
      End If
   ElseIf (m_blnAllowClose = False) Then
      Cancel = True
      Me.Hide
      Call AddToTray
   End If
End Sub

Private Sub MDIForm_Resize()
   If Me.WindowState = vbMinimized Then Exit Sub
   
   On Error Resume Next
   
   If Me.Width < 12000 Then
      Me.Width = 12000
      Exit Sub
   End If
   
   If Me.Height < 9000 Then
      Me.Height = 9000
      Exit Sub
   End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   'Unload form
   Set frmMain = Nothing
End Sub

Private Sub mnuDelContactFld_Click()
   'delete a user contact field
   Const sMOD_NAME As String = "frmMain.mnuDelContactFld_Click"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim sMsg As String
   Dim iMsg As VbMsgBoxResult
   Dim strContact As String
   
   strContact = ConvertContactName(g_lngContID)
   
   sMsg = "Are you sure you want to DELETE [ " & m_strContUserFld & " ]" & vbCrLf
   sMsg = sMsg & "as a User Defined Field from contact " & strContact & " ?"
   iMsg = MsgBox(sMsg, vbQuestion + vbYesNo, APP_MSG_NAME)
   
   If (iMsg <> vbYes) Then Exit Sub
   
   Dim DeleteSQL As String
   
   DeleteSQL = "DELETE * FROM CUFldValues WHERE RefNum = " & m_lngContUserFld
   
   dbContact.Execute (DeleteSQL)
   
   Call frmContEntry.LoadUserDefInfo
   
   'clear variables
   m_strContUserFld = ""
   m_lngContUserFld = 0
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Deleting the Contact Field!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuDelProjectFld_Click()
   'delete the selected project field from the database
   Const sMOD_NAME As String = "frmMain.mnuDelProjectFld_Click"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim sMsg As String
   Dim iMsg As VbMsgBoxResult
   Dim strProject As String
   
   strProject = ConvertProjectName(g_lngProjID)
   
   sMsg = "Are you sure you want to DELETE [ " & m_strProjUserFld & " ]" & vbCrLf
   sMsg = sMsg & "as a User Defined Field from project " & strProject & " ?"
   iMsg = MsgBox(sMsg, vbQuestion + vbYesNo, APP_MSG_NAME)
   
   If (iMsg <> vbYes) Then Exit Sub
   
   Dim DeleteSQL As String
   
   DeleteSQL = "DELETE * FROM PUFldValues WHERE RefNum = " & m_lngProjUserFld
   
   dbContact.Execute (DeleteSQL)
   
   Call frmProjEntry.LoadUserDefInfo
   
   'clear variables
   m_strProjUserFld = ""
   m_lngProjUserFld = 0
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Deleting the Project Field!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

'******************************START EDIT MENU******************************
Private Sub mnuEditDelete_Click()
   'Delete item
   Dim iMsg As VbMsgBoxResult
   Dim sMsg As String
   
   Select Case g_strFormFlag
      Case "CEnt" 'contact
         sMsg = "Are you sure you want to DELETE this Contact File" & vbCrLf
         sMsg = sMsg & "and all of their related components and files?" & vbCrLf
         sMsg = sMsg & "This deletion is final and cannot be reversed!"
         
         iMsg = MsgBox(sMsg, vbCritical + vbYesNo, "Warning : Deleting Contact Files")
         
         If (iMsg <> vbYes) Then Exit Sub
         
         frmDelete.m_intType = 1
         Load frmDelete
         frmDelete.Show vbModeless, frmMain
      Case "PEnt"
         sMsg = "Are you sure you want to DELETE this Project File" & vbCrLf
         sMsg = sMsg & "and all of it's related components and files?" & vbCrLf
         sMsg = sMsg & "This deletion is final and cannot be reversed!"
         
         iMsg = MsgBox(sMsg, vbCritical + vbYesNo, "Warning : Deleting Project Files")
         
         If (iMsg <> vbYes) Then Exit Sub
         
         frmDelete.m_intType = 2
         Load frmDelete
         frmDelete.Show vbModeless, frmMain
   End Select
End Sub

Private Sub mnuEditNewAppt_Click()
   'Add new appointment
   Const sMOD_NAME As String = "frmMain.mnuEditNewAppt_Click"
   On Error GoTo Error_Handler
   
   icurState = NOW_ADDING
   Load frmAppt
   frmAppt.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the Appointments dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuEditNewName_Click()
   'Add new name
   Const sMOD_NAME As String = "frmMain.mnuEditNewName_Click"
   On Error GoTo Error_Handler
   
   Load frmNewName
   frmNewName.Show vbModeless, Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the New Contact Name dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuEditNewNote_Click()
   'Add new note
   Const sMOD_NAME As String = "frmMain.mnuEditNewNote_Click"
   On Error GoTo Error_Handler
   
   icurState = NOW_ADDING
   Load frmNotes
   frmNotes.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred  while opening the Notes/Calls dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuEditNewProject_Click()
   'Add new project
   Const sMOD_NAME As String = "frmMain.mnuEditNewProject_Click"
   On Error GoTo Error_Handler
   
   Load frmNewProject
   frmNewProject.Show vbModeless, Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the Projects dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuEditNewToDo_Click()
   'Add new ToDo item
   Const sMOD_NAME As String = "frmMain.mnuEditNewToDo_Click"
   On Error GoTo Error_Handler
   
   icurState = NOW_ADDING
   Load frmToDo
   frmToDo.Show vbModeless, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the To Do dialog!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuEditPrefer_Click()
   'Program preferences
   Load frmPreference
   frmPreference.Show vbModeless, frmMain
End Sub

Private Sub mnuEditPwd_Click()
   'Set/Edit password
   icurState = NOW_ADDING
   Load frmSecurity
   frmSecurity.Show vbModeless, frmMain
End Sub
'******************************END EDIT MENU********************************

'*****************************START FILE MENU*******************************
Private Sub mnuFileBackupDB_Click()
   'Backup data
   Const sMOD_NAME As String = "frmMain.mnuFileBackupDB_Click"
   On Error GoTo Error_Handler
   
   frmBackRest.m_strType = "Backup"
   frmBackRest.Show vbModal, frmMain
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Backing Up the Database!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuFileCompDB_Click()
   'Compact database
   Call CompactDB
End Sub

Private Sub mnuFileExit_Click()
   'Close program
   Unload Me
End Sub

Private Sub mnuFilePrint_Click()
   'Print
   Const sMOD_NAME As String = "frmMain.mnuFilePrint_Click"
   On Error GoTo Error_Handler
   
   Select Case g_strFormFlag
      Case "CEnt" 'contact entry
         Load frmPrintContact
         frmPrintContact.Show vbModeless, frmMain
      Case "PEnt" 'project ebtry
         Load frmPrintProject
         frmPrintProject.Show vbModeless, frmMain
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Printing!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuFileRepDB_Click()
   'Repair database
   Const sMOD_NAME As String = "frmMain.mnuFileRepDB_Click"
   On Error GoTo Error_Handler
   
   Dim sNewName As String
   
   'the file name to repair
   sNewName = App.Path & "\Data\CBaseMgr.mdb"
   
   Screen.MousePointer = vbHourglass
   MsgBar "Repairing " & sNewName, True
   
   'unload all forms & close the database
   UnloadAllForms
   dbContact.Close
   Set dbContact = Nothing
   
   DBEngine.RepairDatabase sNewName
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   're-open the compacted database
   Call OpenLocalDB
   Load frmHome
   
   MsgBox "The Database was sucessfully repaired.", , APP_MSG_NAME
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Repairing the Database!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Private Sub mnuFileRestoreDB_Click()
   'Restore data
   Const sMOD_NAME As String = "frmMain.mnuFileRestoreDB_Click"
   On Error GoTo Error_Handler
   
   If FileExists(App.Path & "\Backup\CBaseMgr.mdb") = False Then
      MsgBox "There is no Backup Data to Restore from", , APP_MSG_NAME
      Exit Sub
   End If
   
   If MsgBox("Restoring the Database from backup files will replace the existing database." & vbCrLf & _
         "Are you sure you want to Contunue?", vbYesNo, "Restore From Backup Data") = vbYes Then
      'close the running database
      UnloadAllForms
      dbContact.Close
      Set dbContact = Nothing
      
      frmBackRest.m_strType = "Restore"
      frmBackRest.Show vbModal, frmMain
   Else
      MsgBox "Database Restore Canceled", , APP_MSG_NAME
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Restoring the Database!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
'*******************************END FILE MENU*******************************

'******************************START HELP MENU******************************
Private Sub mnuHelpAbout_Click()
   'Help about
   frmAbout.Show vbModal
End Sub

Private Sub mnuHelpErrLog_Click()
   'View error log
   frmErrLog.Show vbModal
End Sub

Private Sub mnuHelpFiles_Click()
   'View help files
   Dim strHFile As String
   
   strHFile = App.Path & "\Help\CBHelp.chm"
   
   ShellExecute Me.hWnd, "open", strHFile, "", "", vbNormalFocus
End Sub
'*******************************END HELP MENU*******************************

Private Sub mnuRCMemoPopEdit_Click()
   'edit the related contact memo field
   Const sMOD_NAME As String = "frmMain.mnuRCMemoPopEdit_Click"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strNewMemo As String
   
   SQL = "SELECT MasterContID, LinkMemo, SubContID FROM RelateCont "
   SQL = SQL & "WHERE MasterContID = " & g_lngContID
   SQL = SQL & " AND SubContID = " & m_lngRelContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         strNewMemo = InputBox("Enter the new Related Contact Memo", _
            "Edit Related Contact Memo", m_strOldRCMemo)
         If (strNewMemo = "") Then
            rsList.Close
            Set rsList = Nothing
            Exit Sub
         End If
         
         .Edit
         !LinkMemo = strNewMemo
         .Update
      End If
   End With
   
   Call frmContEntry.LoadRelContactInfo
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while editing the Related Contact Memo!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuTrayPopClose_Click()
   m_blnAllowClose = True
   Call RemoveFromTray
   Unload Me
End Sub

Private Sub mnuTrayPopRestore_Click()
   frmMain.Show
   Call RemoveFromTray
End Sub

'*****************************START VIEW MENU*******************************
Private Sub mnuViewCal_Click()
   'View calendar screen
   Const sMOD_NAME As String = "frmMain.mnuViewCal_Click"
   On Error GoTo Error_Handler
   
   UnloadAllForms
   Load frmCalMnth
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the Calendar Screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuViewCQry_Click()
   'view contact user defined field query
   Const sMOD_NAME As String = "frmMain.mnuViewCQry_Click"
   On Error GoTo Error_Handler
   
   UnloadAllForms
   Load frmContQry
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the Contact Query Screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuViewHome_Click()
   'View home screen
   Const sMOD_NAME As String = "frmMain.mnuViewHome_Click"
   On Error GoTo Error_Handler
   
   UnloadAllForms
   Load frmHome
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred  while opening the Home Screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuViewNames_Click()
   'View all names screen
   Const sMOD_NAME As String = "frmMain.mnuViewNames_Click"
   On Error GoTo Error_Handler
   
   UnloadAllForms
   Load frmNames
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred  while opening the Names Screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuViewPQry_Click()
   'view project user defined field query
   Const sMOD_NAME As String = "frmMain.mnuViewPQry_Click"
   On Error GoTo Error_Handler
   
   UnloadAllForms
   Load frmProjQry
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the Projects Query Screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuViewProj_Click()
   'View projects screen
   Const sMOD_NAME As String = "frmMain.mnuViewProj_Click"
   On Error GoTo Error_Handler
   
   UnloadAllForms
   Load frmProjects
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the Projects Screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub mnuViewToDo_Click()
   'View To Do list screen
   Const sMOD_NAME As String = "frmMain.mnuViewToDo_Click"
   On Error GoTo Error_Handler
   
   UnloadAllForms
   Load frmToDoList
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while opening the To Do Screen!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
'*******************************END VIEW MENU*******************************

Private Sub picStatus_Resize()
   On Error Resume Next
   
   lblStatus.Left = 75
   lblStatus.Width = picStatus.ScaleWidth - 150
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'for popup menu at system tray
   '///adapted from example by Gary Lantz///
   Const sMOD_NAME As String = "frmMain.picTray_MouseMove"
   On Error GoTo Error_Handler
   
   Dim lRet As Long
   
   If picTray.ScaleMode = vbPixels Then
      lRet = X
   Else
      lRet = X / Screen.TwipsPerPixelX
   End If
   
   Select Case lRet
      Case WM_RBUTTONUP
         PopupMenu mnuTrayPop
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Picture1_Resize()
   tbrGo.Move Picture1.ScaleWidth - tbrGo.Width - 100
   txtSrch.Move Picture1.ScaleWidth - tbrGo.Width - txtSrch.Width - 75
   Label1.Move txtSrch.Left - Label1.Width - 75
End Sub

'***************************START GO TOOLBAR********************************
Private Sub tbrGo_ButtonClick(ByVal Button As MSComctlLib.Button)
   Const sMOD_NAME As String = "frmMain.tbrGo_ButtonClick"
   On Error GoTo Error_Handler
   
   Select Case Button.Key
      Case "Find" 'find entered name
         Call GetNamesList
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while retrieving the Contact Name!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
'*****************************END GO TOOLBAR********************************

'***************************START MAIN TOOLBAR******************************
Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
   Const sMOD_NAME As String = "frmMain.tbrMain_ButtonClick"
   On Error GoTo Error_Handler
   
   Select Case Button.Key
      Case "Names" 'view name list
         Call mnuViewNames_Click
      Case "Prj" 'view project screen
         Call mnuViewProj_Click
      Case "Cal" 'view calendar screen
         Call mnuViewCal_Click
      Case "ToDo" 'view todo screen
         Call mnuViewToDo_Click
      Case "Print" 'print active item
         Call mnuFilePrint_Click
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
'*****************************END MAIN TOOLBAR******************************

'***********************START MAIN TOOLBAR DROPDOWN*************************
Private Sub tbrMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
   Const sMOD_NAME As String = "frmMain.tbrMain_ButtonMenuClick"
   On Error GoTo Error_Handler
   
   Select Case ButtonMenu.Key
      Case "drpName" 'add new name
         Call mnuEditNewName_Click
      Case "drpPrj" 'add new project
         Call mnuEditNewProject_Click
      Case "drpNote" 'add new note
         Call mnuEditNewNote_Click
      Case "drpToDo" 'add new to do item
         Call mnuEditNewToDo_Click
      Case "drpAppt" 'add new appointment
         Call mnuEditNewAppt_Click
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
'***********************END MAIN TOOLBAR DROPDOWN***************************

'************************START NAVIGATION TOOLBAR***************************
Private Sub tbrNav_ButtonClick(ByVal Button As MSComctlLib.Button)
   Const sMOD_NAME As String = "frmMain.tbrNav_ButtonClick"
   On Error GoTo Error_Handler
   
   Select Case Button.Key
      Case "Home" 'go to home screen
         Call mnuViewHome_Click
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
'**************************END NAVIGATION TOOLBAR***************************

Private Sub Timer1_Timer()
   'set code to scan for reminders every minute
   Const sMOD_NAME As String = "frmMain.Timer1_Timer"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim lngRefNum As Long
   Dim lngApptID As Long
   Dim lngToDoID As Long
   Dim strType As String
   Dim vDate As Variant
   Dim vTime As Variant
   Dim rsRemind As Recordset
   
   vDate = "#" & m_vSysDate & "#"
   
   m_vSysTime = Format(Time, "h:nn AMPM")
   vTime = "#" & m_vSysTime & "#"
   
   SQL = "SELECT RefNum, RemDate, RemTime, fkToDoID, fkApptID, "
   SQL = SQL & "Type, Completed FROM Remind "
   SQL = SQL & "WHERE RemDate = " & vDate
   SQL = SQL & " AND RemTime = " & vTime
   SQL = SQL & " AND Completed = False"
   
   Set rsRemind = dbContact.OpenRecordset(SQL)
   
   With rsRemind
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!RefNum)) Then lngRefNum = !RefNum
         If (Not IsNull(!fkToDoID)) Then lngToDoID = !fkToDoID
         If (Not IsNull(!fkApptID)) Then lngApptID = !fkApptID
         If (Not IsNull(!Type)) Then strType = !Type
      End If
   End With
   
   If (strType <> "") Then
      Select Case strType
         Case "AP" 'appointment
            frmReminder.m_lngRemindID = lngRefNum
            frmReminder.m_lngApptID = lngApptID
            frmReminder.m_strType = strType
            Load frmReminder
            frmReminder.Show vbModal
         Case "TD" 'to do
            frmReminder.m_lngRemindID = lngRefNum
            frmReminder.m_lngToDoID = lngToDoID
            frmReminder.m_strType = strType
            Load frmReminder
            frmReminder.Show vbModal
      End Select
   End If
   
   rsRemind.Close
   Set rsRemind = Nothing
            
   'check to see if there is any more reminders to set/unset allow close flag
   Call CheckReminderList
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub txtSrch_GotFocus()
   highLight
End Sub

Private Sub txtSrch_KeyPress(KeyAscii As Integer)
   Const sMOD_NAME As String = "frmMain.txtSrch_KeyPress"
   On Error GoTo Error_Handler
   
   If KeyAscii = vbKeyReturn Then
      Call GetNamesList
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub GetNamesList()
   Const sMOD_NAME As String = "frmMain.GetNamesList"
   On Error GoTo Error_Handler
   
   Dim strEntered As String
   
   strEntered = txtSrch.Text
   If (strEntered = "") Then Exit Sub
   strEntered = strEntered & "*"
   
   If InStr(1, strEntered, "'") Then
      strEntered = SrchReplace(strEntered)
   End If
   
   g_strNameSQL = "SELECT ContID, ShownName, JobTitle FROM Contacts "
   g_strNameSQL = g_strNameSQL & "WHERE ShownName LIKE '" & strEntered & "' "
   
   Set rsList = dbContact.OpenRecordset(g_strNameSQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveLast
         .MoveFirst
      End If
      If (.RecordCount = 0) Then
         MsgBox "No names match your criteria.", , APP_MSG_NAME
      ElseIf (.RecordCount = 1) Then
         g_lngContID = !ContID
         UnloadAllForms
         Load frmContEntry
      ElseIf (.RecordCount >= 2) Then
         UnloadAllForms
         Load frmNameSrchResult
         frmNameSrchResult.Show
      'Else
         'MsgBox "No names match your criteria.", , APP_MSG_NAME
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   g_strNameSQL = ""
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub CheckReminderList()
   'check and see if there are any reminders set for this day, if there
   'are any reminders set, set m_blnAllowClose to False
   Const sMOD_NAME As String = "frmMain.CheckReminderList"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim vCheckDate As Variant
   
   'vCheckDate = "#" & m_vSysDate & "#"
   'updated code 10.23.04 thanks to submission by michael doering
   vCheckDate = Format$(CVDate(m_vSysDate), "\#mm\/dd\/yyyy\#")
   
   SQL = "SELECT RemDate, Completed FROM Remind "
   SQL = SQL & "WHERE RemDate = " & vCheckDate
   SQL = SQL & " AND Completed = False"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         m_blnAllowClose = False
      Else
         m_blnAllowClose = True
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub
