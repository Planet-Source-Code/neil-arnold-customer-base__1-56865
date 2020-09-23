VERSION 5.00
Begin VB.Form frmNewName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Up New Name"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4140
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpt 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   2925
      TabIndex        =   9
      Top             =   2325
      Width           =   1140
   End
   Begin VB.CommandButton cmdOpt 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   1650
      TabIndex        =   8
      Top             =   2325
      Width           =   1140
   End
   Begin VB.OptionButton optType 
      Caption         =   "A Company"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1290
   End
   Begin VB.OptionButton optType 
      Caption         =   "An Individual"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1500
      Value           =   -1  'True
      Width           =   1290
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   225
      MaxLength       =   90
      TabIndex        =   1
      Top             =   375
      Width           =   3690
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   4050
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   75
      X2              =   4050
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Example : Smith's Carpets, LLC"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   165
      Index           =   2
      Left            =   1575
      TabIndex        =   7
      Top             =   1830
      Width           =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "Example : William H. Smith"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   165
      Index           =   1
      Left            =   1575
      TabIndex        =   6
      Top             =   1530
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Is this an individual or a company ?"
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
      Top             =   1125
      Width           =   3990
   End
   Begin VB.Label Label2 
      Caption         =   "Example : William Smith or Smith's Carpets, LLC"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   165
      Index           =   0
      Left            =   375
      TabIndex        =   2
      Top             =   750
      Width           =   3390
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
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
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   3990
   End
End
Attribute VB_Name = "frmNewName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCont As Recordset 'main recordset
Dim rsList As Recordset 'all other data work

Dim m_strType As String 'I = individual, C = company
Dim m_strPre As String 'prefix
Dim m_strFName As String 'first name
Dim m_strMid As String 'middle name
Dim m_strLName As String 'last name
Dim m_strSuff As String 'suffix
Dim m_strCompName As String 'company name
Dim m_strFullName As String 'contact full name
Dim m_strShowName As String 'contact shown name
Dim m_strNameToParse As String

Dim m_lngNewID As Long 'for new contact id

Private Sub cmdOpt_Click(Index As Integer)
   Const sMOD_NAME As String = "frmNewName.cmdOpt_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'Next>
         If (Not ValidateEntry()) Then Exit Sub
         
         If (m_strType = "I") Then
            m_strNameToParse = txtName.Text
            If (m_strNameToParse = "") Then Exit Sub
            
            Call CollectNames
         ElseIf (m_strType = "C") Then
            m_strCompName = txtName.Text
            m_strShowName = txtName.Text
         End If
         
         If (Not CheckForDupNames()) Then Exit Sub
   
         Call PostEntry
      Case 1 'cancel
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmNewName.Form_Load"
   On Error GoTo Error_Handler
   
   Set rsCont = dbContact.OpenRecordset("Contacts", dbOpenTable)
   
   'flatten all needed borders
   FlatBorder txtName.hWnd
   
   Call GetNewContID
   
   m_strType = "I"
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsCont.Close
   Set rsCont = Nothing
   
   Set frmNewName = Nothing
End Sub

Private Sub GetNewContID()
   'create a new contact ID
   Const sMOD_NAME As String = "frmNewName.GetNewContID"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT MAX(ContID)AS MAXID FROM Contacts"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!MAXID)) Then
            m_lngNewID = !MAXID + 1
         Else
            m_lngNewID = 1
         End If
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub optType_Click(Index As Integer)
   Select Case Index
      Case 0 'individual
         m_strType = "I"
      Case 1 'company
         m_strType = "C"
   End Select
End Sub

Private Sub txtName_GotFocus()
   highLight
End Sub

Private Function ValidateEntry() As Boolean
   'make sure some text was entered
   ValidateEntry = True
   
   If (Len(txtName) < 1) Then
      MsgBox "You must enter an individual name or a company name.", _
         vbInformation + vbOKOnly, "Validate : Name Entry"
      txtName.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub CollectNames()
   'parse the entered name for individual names
   Const sMOD_NAME As String = "frmNewName.CollectNames"
   On Error GoTo Error_Handler
   
   Dim objParse As New clsNameParse
   
   objParse.ParseName (m_strNameToParse)
   
   m_strPre = objParse.Prefix
   m_strFName = objParse.FirstName
   m_strMid = objParse.MiddleName
   m_strLName = objParse.LastName
   m_strSuff = objParse.Suffix
   
   Set objParse = Nothing
   
   'make sure you at least have a first name and last name
   If (m_strFName = "") Then
      MsgBox "Make sure you have entered a valid name.", , APP_MSG_NAME
      Exit Sub
   End If
   If (m_strLName = "") Then
      MsgBox "Make sure you have entered a valid name.", , APP_MSG_NAME
      Exit Sub
   End If
   
   'Set full name
   'set prefix if needed
   If (m_strPre <> "") Then
      m_strFullName = m_strPre & " "
   Else
      m_strFullName = m_strFullName
   End If
   'set first name
   m_strFullName = m_strFullName & m_strFName & " "
   'set middle name if needed
   If (m_strMid <> "") Then
      m_strFullName = m_strFullName & m_strMid & " "
   Else
      m_strFullName = m_strFullName
   End If
   'set last name
   m_strFullName = m_strFullName & m_strLName
   'set suffix if needed
   If (m_strSuff <> "") Then
      m_strFullName = m_strFullName & " " & m_strSuff
   End If
   
   'Set shown name
   m_strShowName = m_strLName & ", " & m_strFName
   If (m_strMid <> "") Then
      m_strShowName = m_strShowName & " " & m_strMid
   End If
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub PostEntry()
   'post the new record to the database
   Const sMOD_NAME As String = "frmNewName.PostEntry"
   On Error GoTo Error_Handler
   
   Dim strSetting As String
   
   strSetting = "Default"
   
   rsCont.AddNew
   
   With rsCont
      !ContID = m_lngNewID
      
      Select Case m_strType
         Case "I" 'individual
            !CTYPE = m_strType
            !Setting = strSetting
            If (m_strFullName <> "") Then !FullName = m_strFullName
            If (m_strShowName <> "") Then !ShownName = m_strShowName
            If (m_strPre <> "") Then !Prefix = m_strPre
            If (m_strFName <> "") Then !FName = m_strFName
            If (m_strMid <> "") Then !Middle = m_strMid
            If (m_strLName <> "") Then !LName = m_strLName
            If (m_strSuff <> "") Then !Suffix = m_strSuff
         Case "C" 'company
            !CTYPE = "C"
            !Setting = strSetting
            If (m_strShowName <> "") Then !ShownName = m_strShowName
            If (m_strCompName <> "") Then !CompName = m_strCompName
      End Select
      
      .Update
   End With
   
   g_lngContID = m_lngNewID
   
   Me.Hide
   
   UnloadAllForms
   Load frmContEntry
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Function CheckForDupNames() As Boolean
   'check to see if this name is already in the database
   Dim SQL As String
   
   If InStr(1, m_strShowName, "'") Then
      m_strShowName = SrchReplace(m_strShowName)
   End If
   
   SQL = "SELECT ShownName FROM Contacts WHERE ShownName = '" & m_strShowName & "' "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         CheckForDupNames = False
         MsgBox "That Contact Name already exists, please re-enter another name.", , APP_MSG_NAME
         txtName.Text = ""
         txtName.SetFocus
         Exit Function
      Else
         CheckForDupNames = True
         m_strShowName = Replace(m_strShowName, "''", "'")
         m_strShowName = Trim(m_strShowName)
      End If
   End With
   
   rsList.Close
   Set rsList = Nothing
End Function
