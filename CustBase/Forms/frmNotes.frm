VERSION 5.00
Begin VB.Form frmNotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New / Edit Note"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&New"
      Height          =   390
      Index           =   0
      Left            =   150
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   1500
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   2850
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   3
      Left            =   6450
      TabIndex        =   10
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   4
      Left            =   7800
      TabIndex        =   11
      Top             =   4725
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Private"
      Height          =   240
      Index           =   1
      Left            =   6450
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Favorite (*)"
      Height          =   240
      Index           =   0
      Left            =   5175
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.OptionButton optCall 
      Caption         =   "Call"
      Height          =   240
      Left            =   6825
      TabIndex        =   4
      Top             =   1950
      Width           =   690
   End
   Begin VB.OptionButton optNote 
      Caption         =   "Note"
      Height          =   240
      Left            =   6000
      TabIndex        =   3
      Top             =   1950
      Value           =   -1  'True
      Width           =   690
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmNotes.frx":000C
      Left            =   5850
      List            =   "frmNotes.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   525
      Width           =   3090
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   5850
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   975
      Width           =   3090
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00EAFFFF&
      Height          =   3915
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   525
      Width           =   4815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   9000
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   150
      X2              =   9000
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "This is a:"
      Height          =   240
      Index           =   2
      Left            =   5175
      TabIndex        =   16
      Top             =   1950
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
      Left            =   5175
      TabIndex        =   15
      Top             =   1500
      Width           =   3765
   End
   Begin VB.Label Label1 
      Caption         =   "Project:"
      Height          =   240
      Index           =   1
      Left            =   5175
      TabIndex        =   14
      Top             =   1012
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   240
      Index           =   0
      Left            =   5175
      TabIndex        =   13
      Top             =   562
      Width           =   615
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Note"
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
      TabIndex        =   12
      Top             =   150
      Width           =   8790
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsNote As Recordset 'main recordset
Dim rsList As Recordset 'all other data work

Public m_lngNoteID As Long 'for id transfer

Dim m_strType As String '"C" = Call, "N" = Note
Dim m_lngContID As Long 'for contact id
Dim m_lngProjID As Long 'for project id
Dim m_blnCancelled As Boolean
Dim m_vSaveDate As Variant 'for record timestamp date

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmNotes.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Dim iMsg As VbMsgBoxResult
   
   Select Case Index
      Case 0 'New
         Call SetupNewRecord
         cmdOpts(2).Enabled = False
      Case 1 'Delete
         Call DeleteCurrentRecord
      Case 2 'Print
         iMsg = MsgBox("Print this Note/Call on printer " & Printer.DeviceName, _
            vbQuestion + vbYesNo, "Confirm Print Record")
         
         If (iMsg <> vbYes) Then Exit Sub
         
         Call PrintPage
      Case 3 'OK
         If (Not ValidateEntry()) Then Exit Sub
         
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
   If Index = 0 Then
      m_lngContID = Combo1(0).ItemData(Combo1(0).ListIndex)
   End If
   If Index = 1 Then
      m_lngProjID = Combo1(1).ItemData(Combo1(1).ListIndex)
   End If
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmNotes.Form_Load"
   On Error GoTo Error_Handler
   
   'set main recordset
   Set rsNote = dbContact.OpenRecordset("Attach", dbOpenTable)
   
   'set note type
   m_strType = "N"
   
   'set a blank line in the combo boxes
   Combo1(0).AddItem " "
   Combo1(1).AddItem " "
   
   'flatten all needed borders
   Dim Indx As Integer
   
   For Indx = 0 To 1
      FlatBorder Combo1(Indx).hWnd
   Next
   FlatBorder Text1.hWnd
   
   'set up screen
   Call InitializeScreen
   
   m_blnCancelled = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsNote.Close
   Set rsNote = Nothing
   
   If (m_blnCancelled = False) Then
      Select Case g_strFormFlag
         Case "Home"
            Call frmHome.LoadNotes
         Case "CEnt"
            Call frmContEntry.LoadContactHistory
         Case "PEnt"
            Call frmProjEntry.LoadProjectHistory
      End Select
   End If
   
   Set frmNotes = Nothing
End Sub

Public Sub InitializeScreen()
   'set up the opening screen
   Const sMOD_NAME As String = "frmNotes.InitializeScreen"
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
      With rsNote
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngNoteID
            
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
   Const sMOD_NAME As String = "frmNotes.PopulateFields"
   On Error GoTo Error_Handler
   
   Dim strContact As String
   Dim strProject As String
   
   With rsNote
      If (Not IsNull(!TextBody)) Then Text1 = !TextBody
      'add code to retrieve project name or contact name
      If (!fkContID > 0) Then
         m_lngContID = !fkContID
         strContact = ConvertContactName(m_lngContID)
         Combo1(0).Text = Trim(strContact)
      End If
      If (!fkProjID > 0) Then
         m_lngProjID = !fkProjID
         strProject = ConvertProjectName(m_lngProjID)
         Combo1(1).Text = strProject
      End If
      
      If (Not IsNull(!NType)) Then
         If (!NType = "C") Then
            optCall.Value = True
            m_strType = "C"
         ElseIf (!NType = "N") Then
            optNote.Value = True
            m_strType = "N"
         End If
      End If
      
      If (!Fav = True) Then
         Check1(0).Value = 1
      Else
         Check1(0).Value = 0
      End If
      If (!Priv = True) Then
         Check1(1).Value = 1
      Else
         Check1(1).Value = 0
      End If
      
      m_vSaveDate = !DateStamp
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Loading the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub SetupNewRecord()
   'reset the screen for adding a new record
   Dim Indx As Integer
   
   'clear textboxes
   Text1.Text = ""
   'clear combos
   Combo1(0).Text = " "
   Combo1(1).Text = " "
   'reset checkboxes
   For Indx = 0 To 1
      Check1(Indx).Value = 0
   Next
   'reset type variable
   m_strType = "N"
   optNote.Value = True
   'reset editmode
   icurState = NOW_ADDING
   cmdOpts(1).Enabled = False
   cmdOpts(2).Enabled = False
   Text1.SetFocus
End Sub

Private Sub DeleteCurrentRecord()
   'delete the current to do item
   Const sMOD_NAME As String = "frmNotes.DeleteCurrentRecord"
   On Error GoTo Error_Handler
   
   Dim iMsg As VbMsgBoxResult
   Dim sMsg As String
   Dim SQL As String
   
   If (icurState = NOW_ADDING) Then
      MsgBox "There is no current record to delete.", , APP_MSG_NAME
      Exit Sub
   End If
   
   sMsg = "Are you sure you want to DELETE this Call/Notes item?"
   
   iMsg = MsgBox(sMsg, vbCritical + vbYesNo, "Warning : Record Deletion")
   
   If (iMsg <> vbYes) Then Exit Sub
   
   SQL = "DELETE * FROM Attach WHERE RefNum = " & m_lngNoteID
   
   dbContact.Execute (SQL)
   
   Me.Hide
   
   Select Case g_strFormFlag
      Case "Home"
         Call frmHome.LoadNotes
      Case "CEnt"
         Call frmContEntry.LoadContactHistory
   End Select
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Deleting the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmNotes.PostEntry"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting Notes/Calls Entry", True
   
   If (icurState = NOW_ADDING) Then
      rsNote.AddNew
   Else
      With rsNote
         If (.RecordCount > 0) Then
            .MoveFirst
            .Index = "PrimaryKey"
            .Seek "=", m_lngNoteID
            If Not .NoMatch Then
               rsNote.Edit
            Else
               MsgBox "Error Posting Record", , "Posting Error"
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
      End With
   End If
   
   With rsNote
      If (Len(Text1)) Then !TextBody = Text1
      
      !NType = m_strType
      
      If (m_lngContID > 0) Then !fkContID = m_lngContID
      If (m_lngProjID > 0) Then !fkProjID = m_lngProjID
      
      If (Check1(0).Value = 1) Then
         !Fav = True
      Else
         !Fav = False
      End If
      If (Check1(1).Value = 1) Then
         !Priv = True
      Else
         !Priv = False
      End If
      
      .Update
   End With
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   Me.Hide
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(Text1) < 1) Then
      Indx = MsgBox("You Must Enter Some Subject Text", _
         vbInformation + vbOKOnly, "Validate : Subject Text")
      Text1.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub optCall_Click()
   m_strType = "C"
End Sub

Private Sub optNote_Click()
   m_strType = "N"
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
   Dim vDate As Variant, vTime As Variant
   Dim strGrdProj As String, strGrdCont As String, strGrdDesc As String
   
   'set page title
   If (m_strType = "C") Then
      strTitle = "Call"
   ElseIf (m_strType = "N") Then
      strTitle = "Note"
   End If
   'set date and time to file saved DateStamp
   vDate = Format(m_vSaveDate, "m/dd/yy")
   vTime = Format(m_vSaveDate, "h:nn AMPM")
   
   'set left header captions
   strGrdProj = "Project"
   strGrdCont = "Contact"
   strGrdDesc = "Description"
   
   Printer.ScaleMode = vbCentimeters
   
   Printer.FontName = "Tahoma"
   Printer.FontSize = 10
   Printer.FontBold = False
   Printer.CurrentX = 1.3
   Printer.CurrentY = 1.5
   Printer.Print vTime;
   Printer.CurrentY = 1.7
   Printer.FontSize = 14
   Printer.FontBold = False
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strTitle)) / 2
   Printer.Print strTitle
   Printer.CurrentX = 1.3
   Printer.CurrentY = 1.9
   Printer.FontSize = 10
   Printer.FontBold = False
   Printer.Print vDate
   
   Printer.FontSize = 8
   Printer.FontBold = False
   
   'print row 1 cell
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
   'print row 2 cell
   Printer.Line (1.3, 3.08)-(1.3, 3.56) 'left line (3.08 + .48)
   Printer.Line (2.95, 3.08)-(2.95, 3.56) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 3.08)-(20.3, 3.56) '3rd left line
   Printer.Line (1.3, 3.56)-(20.3, 3.56) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 3.13 '(3.08 + .05)
   Printer.Print strGrdCont;
   Printer.CurrentX = 3 '(2.95 + .05)
   Printer.Print Combo1(0).Text
   'print row 3 cell
   Printer.Line (1.3, 3.56)-(1.3, 13.04) 'left line (3.56 + .48)[13.04 was 4.04]
   Printer.Line (2.95, 3.56)-(2.95, 13.04) '2nd left line (1.3 + 1.65)
   Printer.Line (20.3, 3.56)-(20.3, 13.04) '3rd left line
   Printer.Line (1.3, 13.04)-(20.3, 13.04) 'bottom line
   Printer.CurrentX = 1.4 '(1.3 + .01)
   Printer.CurrentY = 3.61 '(3.56 + .05)
   Printer.Print strGrdDesc;
   Printer.CurrentX = 3 '(2.95 + .05)
   Call WrapPrintText(Text1.Text)
   
   Dim strFoot As String
   
   Printer.CurrentX = 1.3
   Printer.CurrentY = 26.3
   strFoot = Format$(Date$, "m/d/yyyy")
   Printer.Print strFoot
   Printer.CurrentX = 18.5
   strFoot = "Page " & CStr(Printer.Page)
   Printer.Print strFoot
   
   Printer.EndDoc
End Sub
