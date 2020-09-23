VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPreference 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreference.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   3900
      TabIndex        =   22
      Top             =   4575
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Height          =   390
      Index           =   0
      Left            =   2550
      TabIndex        =   21
      Top             =   4575
      Width           =   1215
   End
   Begin VB.PictureBox picComp 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   450
      ScaleHeight     =   3615
      ScaleWidth      =   4665
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   675
      Visible         =   0   'False
      Width           =   4665
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   900
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2850
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   900
         MaxLength       =   50
         TabIndex        =   19
         Top             =   2475
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   900
         MaxLength       =   25
         TabIndex        =   18
         Top             =   2100
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   900
         MaxLength       =   25
         TabIndex        =   17
         Top             =   1725
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   900
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1350
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Height          =   810
         Index           =   1
         Left            =   900
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   450
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   900
         MaxLength       =   50
         TabIndex        =   14
         Top             =   75
         Width           =   3315
      End
      Begin VB.Label Label3 
         Caption         =   "Web Site:"
         Height          =   240
         Index           =   6
         Left            =   75
         TabIndex        =   13
         Top             =   2850
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "E-Mail:"
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   12
         Top             =   2475
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Fax:"
         Height          =   240
         Index           =   4
         Left            =   75
         TabIndex        =   11
         Top             =   2100
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Phone:"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   10
         Top             =   1725
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Country:"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   9
         Top             =   1350
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   450
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Company:"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   7
         Top             =   75
         Width           =   765
      End
   End
   Begin VB.PictureBox picLists 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   300
      ScaleHeight     =   3615
      ScaleWidth      =   4665
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   525
      Width           =   4665
      Begin VB.CheckBox chkGridLines 
         Caption         =   "Check this so that all of the system grid lists will show standard grid lines"
         Height          =   390
         Left            =   150
         TabIndex        =   4
         Top             =   1500
         Width           =   4365
      End
      Begin VB.CheckBox chkAltColor 
         Caption         =   "Check this to have all populated lists show an alternating grid color"
         Height          =   390
         Left            =   150
         TabIndex        =   3
         Top             =   375
         Value           =   1  'Checked
         Width           =   4365
      End
      Begin VB.Label Label2 
         Caption         =   $"frmPreference.frx":000C
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
         Height          =   615
         Left            =   150
         TabIndex        =   5
         Top             =   825
         Width           =   4365
      End
      Begin VB.Label Label1 
         Caption         =   "Set the list grid properties for the entire system"
         Height          =   240
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4590
      End
   End
   Begin MSComctlLib.TabStrip tbsMain 
      Height          =   4140
      Left            =   150
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7303
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "List Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Company Info"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   150
      X2              =   5100
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   150
      X2              =   5100
      Y1              =   4425
      Y2              =   4425
   End
End
Attribute VB_Name = "frmPreference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsComp As Recordset

Dim m_blnChanged As Boolean

Private Sub chkAltColor_Click()
   If (chkAltColor.Value = 1) Then
      g_blnAltColors = True
   ElseIf (chkAltColor.Value = 0) Then
      g_blnAltColors = False
   End If
   
   SaveSetting APP_CATEGORY, APPNAME, "AltColors", IIf(g_blnAltColors, "-1", "0")
End Sub

Private Sub chkGridLines_Click()
   If (chkGridLines.Value = 1) Then
      g_blnShowLines = True
   ElseIf (chkGridLines.Value = 0) Then
      g_blnShowLines = False
   End If
   
   SaveSetting APP_CATEGORY, APPNAME, "GridLines", IIf(g_blnShowLines, "-1", "0")
End Sub

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmPreference.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'OK
         If (Not ValidateEntry()) Then Exit Sub
         
         Call PostEntry
      Case 1 'cancel
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Activate()
   'set current tab
   tbsMain.Tabs(1).Selected = True
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmPreference.Form_Load"
   On Error GoTo Error_Handler
   
   Dim iCtr As Integer
   
   'set the main recordset
   Set rsComp = dbContact.OpenRecordset("Company", dbOpenTable)
   
   'flatten all necessary borders
   For iCtr = 0 To 6
      FlatBorder Text1(iCtr).hWnd
   Next
   
   Call InitializeScreen
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim iMsg As VbMsgBoxResult
   
   If (m_blnChanged = True) Then
      iMsg = MsgBox("The Company information has changed." & vbCrLf & "Would you like to save the changes?", _
         vbQuestion + vbYesNo, "Save Changes")
      If (iMsg = vbYes) Then
         Cancel = True
         MsgBox "Press the OK button to save the changes?"
         Call cmdOpts_Click(0)
         Exit Sub
      End If
   End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   
   picLists.Move tbsMain.ClientLeft, tbsMain.ClientTop, tbsMain.ClientWidth, tbsMain.ClientHeight
   picComp.Move picLists.Left, picLists.Top, picLists.Width, picLists.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data and form reference
   rsComp.Close
   Set rsComp = Nothing
   
   Set frmPreference = Nothing
End Sub

Private Sub InitializeScreen()
   Const sMOD_NAME As String = "frmPreference.InitializeScreen"
   On Error GoTo Error_Handler
   
   'set up the opening screen
   
   'set the alternating grid-color check box
   If (g_blnAltColors = True) Then
      chkAltColor.Value = 1
   Else
      chkAltColor.Value = 0
   End If
   
   'set the show gridlines check box
   If (g_blnShowLines = True) Then
      chkGridLines.Value = 1
   Else
      chkGridLines.Value = 0
   End If
   
   'populate company fields
   With rsComp
      If (.RecordCount > 0) Then
         .MoveFirst
         
         If (Not IsNull(!CompName)) Then Text1(0) = !CompName
         If (Not IsNull(!Address)) Then Text1(1) = !Address
         If (Not IsNull(!Country)) Then Text1(2) = !Country
         If (Not IsNull(!Phone)) Then Text1(3) = !Phone
         If (Not IsNull(!Fax)) Then Text1(4) = !Fax
         If (Not IsNull(!Email)) Then Text1(5) = !Email
         If (Not IsNull(!WebSite)) Then Text1(6) = !WebSite
      End If
   End With
   
   m_blnChanged = False
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Private Sub tbsMain_Click()
   On Error Resume Next
   
   picLists.Visible = False
   picComp.Visible = False
   
   Select Case tbsMain.SelectedItem.Index
      Case 1 'listviews
         picLists.Visible = True
         chkAltColor.SetFocus
      Case 2 'comp info
         picComp.Visible = True
         Text1(0).SetFocus
   End Select
End Sub

Private Sub Text1_Change(Index As Integer)
   m_blnChanged = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   highLight
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(Text1(0)) < 1) Then
      Indx = MsgBox("You Must Enter A Company Name", _
         vbInformation + vbOKOnly, "Validate : Company Name")
      Text1(0).SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmPreference.PostEntry"
   On Error GoTo Error_Handler
   
   Screen.MousePointer = vbHourglass
   MsgBar "Posting Company Information Entry", True
   
   With rsComp
      .MoveFirst
      .Edit
   End With
   
   With rsComp
      If (Len(Text1(0))) Then !CompName = Text1(0)
      If (Len(Text1(1))) Then !Address = Text1(1)
      If (Len(Text1(2))) Then !Country = Text1(2)
      If (Len(Text1(3))) Then !Phone = Text1(3)
      If (Len(Text1(4))) Then !Fax = Text1(4)
      If (Len(Text1(5))) Then !Email = Text1(5)
      If (Len(Text1(6))) Then !WebSite = Text1(6)
      
      .Update
   End With
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
   
   m_blnChanged = False
   
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
