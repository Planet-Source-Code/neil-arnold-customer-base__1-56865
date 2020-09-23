VERSION 5.00
Begin VB.Form frmNameDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify Name Details"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNameDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   5400
      TabIndex        =   8
      Top             =   3375
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   6750
      TabIndex        =   9
      Top             =   3375
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   5025
      TabIndex        =   7
      Top             =   900
      Width           =   2940
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2700
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   900
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2250
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   900
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1800
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   900
      MaxLength       =   30
      TabIndex        =   2
      Top             =   1350
      Width           =   2940
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   1890
   End
   Begin VB.OptionButton optType 
      Caption         =   "A Company"
      Height          =   240
      Index           =   1
      Left            =   4275
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   525
      Width           =   1290
   End
   Begin VB.OptionButton optType 
      Caption         =   "An Individual"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   525
      Value           =   -1  'True
      Width           =   1290
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
      TabIndex        =   18
      Top             =   555
      Width           =   1665
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
      Left            =   5700
      TabIndex        =   17
      Top             =   555
      Width           =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   7950
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   4035
      X2              =   4035
      Y1              =   525
      Y2              =   3075
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   150
      X2              =   7950
      Y1              =   3225
      Y2              =   3225
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   240
      Index           =   5
      Left            =   4425
      TabIndex        =   16
      Top             =   937
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4050
      X2              =   4050
      Y1              =   525
      Y2              =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Suffix:"
      Height          =   240
      Index           =   4
      Left            =   300
      TabIndex        =   15
      Top             =   2737
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Last:"
      Height          =   240
      Index           =   3
      Left            =   300
      TabIndex        =   14
      Top             =   2287
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Middle:"
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   13
      Top             =   1837
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "First:"
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   12
      Top             =   1387
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Title:"
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   11
      Top             =   937
      Width           =   540
   End
   Begin VB.Label lblBanner 
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
      Left            =   150
      TabIndex        =   10
      Top             =   150
      Width           =   7815
   End
End
Attribute VB_Name = "frmNameDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCont As Recordset 'main recordset
Dim rsList As Recordset 'all other data work

Dim m_strCType As String 'for contact type "I" = individual, "C" = company
Dim m_strPre As String
Dim m_strFirst As String
Dim m_strMiddle As String
Dim m_strLast As String
Dim m_strSuff As String
Dim m_strComp As String
Dim m_strFull As String
Dim m_strShown As String

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmNameDetail.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'OK
         If (Not ValidateEntry()) Then Exit Sub
         
         Call CombineNames
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

Private Sub Combo1_Click(Index As Integer)
   Select Case Index
      Case 0 'prefix
         If (Combo1(0).Text = "<Add new name: prefix") Then
            Load frmSetPrefix
            frmSetPrefix.Show vbModeless, frmMain
         End If
      Case 1 'suffix
         If (Combo1(1).Text = "<Add new name: suffix") Then
            Load frmSetSuffix
            frmSetSuffix.Show vbModeless, frmMain
         End If
   End Select
End Sub

Private Sub Form_Load()
   'flatten all needed items
   Const sMOD_NAME As String = "frmNameDetail.Form_Load"
   On Error GoTo Error_Handler
   
   Dim Indx As Integer
   
   For Indx = 0 To 3
      FlatBorder Text1(Indx).hWnd
   Next
   For Indx = 0 To 1
      FlatBorder Combo1(Indx).hWnd
   Next
   
   'set type variable
   m_strCType = "I"
   
   'set main recordset
   Set rsCont = dbContact.OpenRecordset("Contacts", dbOpenTable)
   
   'setup the screen
   Call InitializeScreen
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsCont.Close
   Set rsCont = Nothing
   
   Set frmNameDetail = Nothing
End Sub

Public Sub InitializeScreen()
   'setup the screen upon opening
   Const sMOD_NAME As String = "frmNameDetail.InitializeScreen"
   On Error GoTo Error_Handler
   
   Call LoadPrefixCombo
   Combo1(0).Text = " "
   
   Call LoadSuffixCombo
   Combo1(1).Text = " "
   
   With rsCont
      If (.RecordCount > 0) Then
         .MoveFirst
         .Index = "PrimaryKey"
         .Seek "=", g_lngContID
         
         Call PopulateFields
      End If
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub LoadPrefixCombo()
   'load all stored prefixes into combo1(0)
   Const sMOD_NAME As String = "frmNameDetail.LoadPrefixCombo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strLookup As String
   
   strLookup = "PRE"
   
   SQL = "SELECT ItemID, Description FROM Lookup "
   SQL = SQL & "WHERE ItemID = '" & strLookup & "' "
   SQL = SQL & "ORDER BY Description"
   
   Combo1(0).Clear
   Combo1(0).AddItem "<Add new name: prefix"
   Combo1(0).AddItem " "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then Combo1(0).AddItem !Description
            .MoveNext
         Wend
      End If
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub LoadSuffixCombo()
   'load all stored suffixes into combo1(1)
   Const sMOD_NAME As String = "frmNameDetail.LoadSuffixCombo"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim strLookup As String
   
   strLookup = "SUFF"
   
   SQL = "SELECT ItemID, Description FROM Lookup "
   SQL = SQL & "WHERE ItemID = '" & strLookup & "' "
   SQL = SQL & "ORDER BY Description"
   
   Combo1(1).Clear
   Combo1(1).AddItem "<Add new name: suffix"
   Combo1(1).AddItem " "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!Description)) Then Combo1(1).AddItem !Description
            .MoveNext
         Wend
      End If
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub PopulateFields()
   'load the desired record onto the screen
   Const sMOD_NAME As String = "frmNameDetail.PopulateFields"
   On Error GoTo Error_Handler
   
   With rsCont
      If (Not IsNull(!Prefix)) Then Combo1(0).Text = !Prefix
      If (Not IsNull(!Suffix)) Then Combo1(1).Text = !Suffix
      
      If (Not IsNull(!FName)) Then Text1(0) = !FName
      If (Not IsNull(!Middle)) Then Text1(1) = !Middle
      If (Not IsNull(!LName)) Then Text1(2) = !LName
      If (Not IsNull(!CompName)) Then Text1(3) = !CompName
      
      m_strCType = !CTYPE
      If (m_strCType = "I") Then
         optType(0).Value = True
         Call EnableControls
      Else
         optType(1).Value = True
         Call EnableControls
      End If
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub EnableControls()
   On Error Resume Next
   
   Dim Indx As Integer
   
   Select Case m_strCType
      Case "I" 'individual
         For Indx = 0 To 2
            Text1(Indx).Enabled = True
            Text1(Indx).BackColor = vbWhite
         Next
         For Indx = 0 To 1
            Combo1(Indx).Enabled = True
            Combo1(Indx).BackColor = vbWhite
         Next
         Text1(3).Enabled = False
         Text1(3).BackColor = vbButtonFace
         Combo1(0).SetFocus
      Case "C" 'company
         For Indx = 0 To 2
            Text1(Indx).Enabled = False
            Text1(Indx).BackColor = vbButtonFace
         Next
         For Indx = 0 To 1
            Combo1(Indx).Enabled = False
            Combo1(Indx).BackColor = vbButtonFace
         Next
         Text1(3).Enabled = True
         Text1(3).BackColor = vbWhite
         Text1(3).SetFocus
   End Select
End Sub

Private Sub optType_Click(Index As Integer)
   Select Case Index
      Case 0 'individual
         m_strCType = "I"
         Call EnableControls
      Case 1 'company
         m_strCType = "C"
         Call EnableControls
   End Select
End Sub

Private Function ValidateEntry() As Boolean
   'make sure some text was entered
   ValidateEntry = True
   
   If (m_strCType = "I") Then
      If (Len(Text1(0)) < 1) Then
         MsgBox "You must enter an individual first name.", _
            vbInformation + vbOKOnly, "Validate : First Name Entry"
         Text1(0).SetFocus
         ValidateEntry = False
         Exit Function
      End If
      If (Len(Text1(2)) < 1) Then
         MsgBox "You must enter an individual last name.", _
            vbInformation + vbOKOnly, "Validate : Last Name Entry"
         Text1(2).SetFocus
         ValidateEntry = False
         Exit Function
      End If
   ElseIf (m_strCType = "C") Then
      If (Len(Text1(3)) < 1) Then
         MsgBox "You must enter a company name.", _
            vbInformation + vbOKOnly, "Validate : Company Name Entry"
         Text1(3).SetFocus
         ValidateEntry = False
         Exit Function
      End If
   End If
End Function

Private Sub CombineNames()
   Const sMOD_NAME As String = "frmNameDetail.CombineNames"
   On Error GoTo Error_Handler
   
   m_strPre = Combo1(0).Text
   m_strFirst = Text1(0).Text
   m_strMiddle = Text1(1).Text
   m_strLast = Text1(2).Text
   m_strSuff = Combo1(1).Text
   m_strComp = Text1(3).Text
   
   Select Case m_strCType
      Case "I" 'individual
         'arrange full name
         If (m_strPre <> "") Then
            m_strFull = m_strPre & " "
         End If
         m_strFull = m_strFull & m_strFirst
         If (m_strMiddle <> "") Then
            m_strFull = m_strFull & " " & m_strMiddle & " "
         Else
            m_strFull = m_strFull & " "
         End If
         m_strFull = m_strFull & m_strLast
         If (m_strSuff <> "") Then
            m_strFull = m_strFull & " " & m_strSuff
         End If
         'arrange shown name
         m_strShown = m_strLast & ", "
         m_strShown = m_strShown & m_strFirst
         If (m_strMiddle <> "") Then
            m_strShown = m_strShown & " " & m_strMiddle
         End If
      Case "C" 'company
         m_strShown = m_strComp
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmNameDetail.PostEntry"
   On Error GoTo Error_Handler
   
   rsCont.Edit
   
   With rsCont
      Select Case m_strCType
         Case "I" 'individual
            !CTYPE = m_strCType
            !FullName = m_strFull
            !ShownName = m_strShown
            If (m_strPre <> "") Then !Prefix = m_strPre
            If (m_strFirst <> "") Then !FName = m_strFirst
            If (m_strMiddle <> "") Then !Middle = m_strMiddle
            If (m_strLast <> "") Then !LName = m_strLast
            If (m_strSuff <> "") Then !Suffix = m_strSuff
            
            .Update
         Case "C" 'company
            !CTYPE = m_strCType
            !ShownName = m_strShown
            !CompName = m_strComp
            
            .Update
      End Select
   End With
   
   Me.Hide
   
   Call frmContEntry.LoadMainContactInfo
   Call frmContEntry.LoadContactCombo
   frmContEntry.cboContList.Text = m_strShown
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting the information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
   Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   highLight
End Sub
