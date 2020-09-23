VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPrintContact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Contact Report"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintContact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   30
      ScaleHeight     =   1035
      ScaleWidth      =   4575
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   4605
      Begin MSComctlLib.ProgressBar prbPrint 
         Height          =   240
         Left            =   75
         TabIndex        =   12
         Top             =   525
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Gathering Data for Profile Printing ... Please Wait ..."
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
         TabIndex        =   11
         Top             =   75
         Width           =   4440
      End
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   3375
      TabIndex        =   9
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   2025
      TabIndex        =   8
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sections"
      Height          =   1065
      Left            =   75
      TabIndex        =   4
      Top             =   1350
      Width           =   4515
      Begin VB.CheckBox Check1 
         Caption         =   "Related Projects"
         Height          =   240
         Index           =   2
         Left            =   2250
         TabIndex        =   7
         Top             =   300
         Width           =   1890
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Related Contacts"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   675
         Width           =   1890
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Profile"
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   300
         Value           =   1  'Checked
         Width           =   1890
      End
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1050
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   3540
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Include"
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
      Left            =   75
      TabIndex        =   3
      Top             =   1050
      Width           =   4515
   End
   Begin VB.Label Label1 
      Caption         =   "Report Title:"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   450
      Width           =   915
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Create and Print Contact Report For -"
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
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4515
   End
End
Attribute VB_Name = "frmPrintContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_sngCurY As Single

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmPrintContact.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'ok
         Call PrintPage
      Case 1 'cancel
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmPrintContact.Form_Load"
   On Error GoTo Error_Handler
   
   'get current contact name
   txtTitle.Text = frmContEntry.Text1(0).Text
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmPrintContact = Nothing
End Sub

Private Sub txtTitle_GotFocus()
   highLight
End Sub

Private Sub PrintPage()
   'print the current note/call
   Dim sngLine As Single, sngPrtLine As Single
   Dim iCtr As Integer
   Dim jCtr As Integer
   
   Dim strTitle As String
   Dim vCurDate As Variant, vCurTime As Variant
   Dim strLeftHdrs As Variant, strData As Variant
   
   strLeftHdrs = Array("Name", "Shown As", "Setting", "Company", "Job Title", _
                       "Home Phone", "Work Phone", "Mobile Phone", "Other Phone", _
                       "Fax Phone", "Personal E-Mail", "Work E-Mail", _
                       "Other E-Mail")
   
   With frmContEntry
      strData = Array(.Text1(0), .Text1(1), .Text1(17), .Text1(3), .Text1(4), _
                      .Text1(9), .Text1(10), .Text1(11), .Text1(12), .Text1(13), _
                      .Text1(14), .Text1(15), .Text1(16))
   End With
   
   'set page title
   If (txtTitle.Text = "") Then
      strTitle = "Contact Profile Report"
   Else
      strTitle = txtTitle.Text
   End If
   'set date and time to current Date/Time
   vCurDate = Format(Date, "m/dd/yy")
   vCurTime = Format(Time, "h:nn AMPM")
   
   'dis-able command buttons
   cmdOpts(0).Enabled = False: cmdOpts(1).Enabled = False
   'show prog msg (just a visual aid)
   picMsg.Visible = True
   prbPrint.Value = 10
   DoEvents
   
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
   Printer.Print
   
   prbPrint.Value = 20
   
   Printer.FontSize = 8
   Printer.FontBold = True
   Printer.CurrentX = 1.4
   Printer.Print " Name Profile"
   Printer.Line (1.3, 3.1)-(20.3, 3.1) 'top line
   
   'print left box outline
   Printer.Line (1.3, 3.15)-(10.6, 9.39), , B
   'print left divider line
   Printer.Line (3.65, 3.15)-(3.65, 9.39)
   'print gridlines & header text
   Printer.FontBold = False
   sngLine = 3.63
   sngPrtLine = 3.15
   
   prbPrint.Value = 30
   
   For iCtr = 1 To 12
      Printer.Line (1.3, sngLine)-(10.6, sngLine)
         'print header text
         Printer.CurrentX = 1.4
         Printer.CurrentY = sngPrtLine + 0.05
         Printer.Print strLeftHdrs(iCtr - 1)
         sngPrtLine = sngPrtLine + 0.48
      sngLine = sngLine + 0.48
   Next iCtr
   Printer.CurrentX = 1.4
   Printer.CurrentY = sngPrtLine + 0.05
   Printer.Print strLeftHdrs(12)
   'print grid data
   sngPrtLine = 3.15
   For iCtr = 1 To 13
      Printer.CurrentX = 3.75
      Printer.CurrentY = sngPrtLine + 0.05
      Printer.Print strData(iCtr - 1)
      sngPrtLine = sngPrtLine + 0.48
   Next
   
   prbPrint.Value = 50
   
   'print right box outline(s)
   'home address
   Printer.Line (11.1, 3.15)-(20.3, 4.59), , B '3.15 top * 3 lines(.48)
   Printer.CurrentX = 11.2
   Printer.CurrentY = 3.2
   Printer.Print "Address: Home"
   'work address
   Printer.Line (11.1, 4.59)-(20.3, 6.03), , B '4.59 top * 3 lines(.48)
   Printer.CurrentX = 11.2
   Printer.CurrentY = 4.64
   Printer.Print "Address: Work"
   'bill to address
   Printer.Line (11.1, 6.03)-(20.3, 7.47), , B '6.03 top * 3 lines(.48)
   Printer.CurrentX = 11.2
   Printer.CurrentY = 6.08
   Printer.Print "Address: Bill To"
   'ship to address
   Printer.Line (11.1, 7.47)-(20.3, 8.91), , B '7.47 top * 3 lines(.48)
   Printer.CurrentX = 11.2
   Printer.CurrentY = 7.52
   Printer.Print "Address: Ship To"
   'print left grid line
   Printer.Line (13.45, 3.15)-(13.45, 8.91)
   
   prbPrint.Value = 60
   
   Call PopulateAddressFields
   
   'add user defined fields
   Printer.FontSize = 8
   Printer.FontBold = True
   Printer.CurrentX = 1.4
   Printer.CurrentY = 9.79
   Printer.Print " User Defined Contact Fields"
   Printer.Line (1.3, 10.19)-(20.3, 10.19) 'top line
   Printer.Line (1.3, 10.25)-(20.3, 10.25) 'top line
   Printer.CurrentX = 1.4
   Printer.CurrentY = 10.24
   Printer.Print "User Field Description";
   Printer.CurrentX = 10.3
   Printer.Print "Field Value"
   Printer.Line (1.3, 10.64)-(20.3, 10.64)
   
   prbPrint.Value = 70
   Printer.FontBold = False
   Call PopulateUserFields
   
   m_sngCurY = Printer.CurrentY
   Printer.FontBold = True
   
   If (Check1(1).Value = 1) Then
      Printer.FontBold = True
      Printer.CurrentX = 1.4
      m_sngCurY = m_sngCurY + 0.6
      Printer.CurrentY = m_sngCurY
      Printer.Print " All Related Contacts"
      m_sngCurY = m_sngCurY + 0.4
      Printer.Line (1.3, m_sngCurY)-(20.3, m_sngCurY) 'top line
      m_sngCurY = m_sngCurY + 0.06
      Printer.Line (1.3, m_sngCurY)-(20.3, m_sngCurY) 'top line
      Printer.CurrentX = 1.4
      m_sngCurY = m_sngCurY - 0.01
      Printer.CurrentY = m_sngCurY
      Printer.Print "Name";
      Printer.CurrentX = 6.47
      Printer.Print "Memo";
      Printer.CurrentX = 11.54
      Printer.Print "Phone";
      Printer.CurrentX = 16.61
      Printer.Print "E-Mail"
      m_sngCurY = m_sngCurY + 0.4
      Printer.Line (1.3, m_sngCurY)-(20.3, m_sngCurY)
      
      Printer.FontBold = False
      Call PopulateRelContacts
   End If
   
   prbPrint.Value = 80
   
   If (Check1(2).Value = 1) Then
      Printer.FontBold = True
      Printer.CurrentX = 1.4
      m_sngCurY = m_sngCurY + 0.6
      Printer.CurrentY = m_sngCurY
      Printer.Print " All Related Projects"
      m_sngCurY = m_sngCurY + 0.4
      Printer.Line (1.3, m_sngCurY)-(20.3, m_sngCurY) 'top line
      m_sngCurY = m_sngCurY + 0.06
      Printer.Line (1.3, m_sngCurY)-(20.3, m_sngCurY) 'top line
      Printer.CurrentX = 1.4
      m_sngCurY = m_sngCurY - 0.01
      Printer.CurrentY = m_sngCurY
      Printer.Print "Project Name";
      Printer.CurrentX = 10.3
      Printer.Print "Project Status"
      m_sngCurY = m_sngCurY + 0.4
      Printer.Line (1.3, m_sngCurY)-(20.3, m_sngCurY)
      
      Printer.FontBold = False
      Call PopulateRelProjects
   End If
   
   Dim strFoot As String
   
   prbPrint.Value = 90
   
   Printer.CurrentY = 26.3
   Printer.CurrentX = 18.5
   Printer.FontBold = False
   strFoot = "Page " & CStr(Printer.Page)
   Printer.Print strFoot
   
   Printer.EndDoc
   
   'en-able command buttons
   cmdOpts(0).Enabled = True: cmdOpts(1).Enabled = True
   prbPrint.Value = 0
   picMsg.Visible = False
End Sub

Private Sub PopulateAddressFields()
   'load any addresses listed for this contact
   Dim SQL As String
   Dim strType As String
   Dim strAddr As String
   Dim strCity As String
   Dim strState As String
   Dim strZip As String
   Dim strCtry As String
   
   SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip, Country "
   SQL = SQL & "FROM CAddress WHERE fkContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkLookup)) Then
               strType = !fkLookup
            End If
            If (Not IsNull(!Street)) Then strAddr = !Street
            If (Not IsNull(!City)) Then strCity = !City
            If (Not IsNull(!State)) Then strState = !State
            If (Not IsNull(!Zip)) Then strZip = !Zip
            If (Not IsNull(!Country)) Then strCtry = !Country
            Select Case strType
               Case "Home"
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 3.2
                  Printer.Print strAddr
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 3.68
                  Printer.Print strCity & ", " & strState & " " & strZip
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 4.16
                  Printer.Print strCtry
               Case "Work"
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 4.64
                  Printer.Print strAddr
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 5.12
                  Printer.Print strCity & ", " & strState & " " & strZip
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 5.6
                  Printer.Print strCtry
               Case "Bill To"
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 6.08
                  Printer.Print strAddr
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 6.56
                  Printer.Print strCity & ", " & strState & " " & strZip
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 7.04
                  Printer.Print strCtry
               Case "Ship To"
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 7.52
                  Printer.Print strAddr
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 8
                  Printer.Print strCity & ", " & strState & " " & strZip
                  Printer.CurrentX = 13.47
                  Printer.CurrentY = 8.48
                  Printer.Print strCtry
            End Select
            strType = ""
            strAddr = ""
            strCity = ""
            strState = ""
            strZip = ""
            strCtry = ""
            .MoveNext
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
End Sub

Private Sub PopulateUserFields()
   'print all user defined fields for this contact
   Dim SQL As String
   Dim sngLine As Single
   
   sngLine = 10.74
   
   SQL = "SELECT fkContID, fkUserFld, Value FROM CUFldValues "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkUserFld)) Then
               Printer.CurrentX = 1.4
               Printer.CurrentY = sngLine
               Printer.Print !fkUserFld;
               Printer.CurrentX = 10.3
               Printer.Print !Value
            End If
            .MoveNext
            sngLine = sngLine + 0.35
         Wend
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
End Sub

Private Sub PopulateRelContacts()
   'print all related contacts
   Const sMOD_NAME As String = "frmPrintContact.PopulateRelContacts"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strPhone As String
   Dim strEmail As String
   
   SQL = "SELECT MasterContID, LinkMemo, SubContID, SubContShowName "
   SQL = SQL & "FROM RelateCont "
   SQL = SQL & "WHERE MasterContID = " & g_lngContID
   SQL = SQL & " ORDER BY SubContShowName"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   m_sngCurY = m_sngCurY + 0.1
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!SubContID)) Then
               If (Not IsNull(!SubContShowName)) Then
                  Printer.CurrentX = 1.4
                  Printer.CurrentY = m_sngCurY
                  Printer.Print !SubContShowName;
                  If (Not IsNull(!LinkMemo)) Then
                     Printer.CurrentX = 6.47
                     Printer.Print !LinkMemo;
                  End If
                  'code for phone & email
                  strPhone = GetPhoneNum(!SubContID)
                  If (Not IsNull(strPhone)) Then
                     Printer.CurrentX = 11.54
                     Printer.Print strPhone;
                  End If
                  strEmail = GetEMail(!SubContID)
                  If (Not IsNull(strEmail)) Then
                     Printer.CurrentX = 16.61
                     Printer.Print strEmail
                  End If
               End If
            End If
            m_sngCurY = m_sngCurY + 0.35
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

Public Sub PopulateRelProjects()
   'load all projects related to this contact
   Const sMOD_NAME As String = "frmPrintContact.PopulateRelProjects"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT fkProjID, fkContID FROM RelateProject "
   SQL = SQL & "WHERE fkContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   m_sngCurY = m_sngCurY + 0.1
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!fkProjID)) Then
               Call GetProjectInfo(!fkProjID)
            End If
            m_sngCurY = m_sngCurY + 0.35
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

Private Sub GetProjectInfo(lngPrjID As Long)
   'get the project info from the proj link in contact table
   Const sMOD_NAME As String = "frmPrintContact.GetProjectInfo"
   On Error GoTo Error_Handler
   
   Dim cSQL As String
   Dim rsProj As Recordset
   Dim Item As ListItem
   
   cSQL = "SELECT ProjID, PName, Status FROM Projects "
   cSQL = cSQL & "WHERE ProjID = " & lngPrjID
   
   Set rsProj = dbContact.OpenRecordset(cSQL)
   
   With rsProj
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!PName)) Then
            Printer.CurrentX = 1.4
            Printer.CurrentY = m_sngCurY
            Printer.Print !PName;
         End If
         Printer.CurrentX = 10.3
         Printer.Print "* " & !Status
      End If
   End With
   
   rsProj.Close
   Set rsProj = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub
