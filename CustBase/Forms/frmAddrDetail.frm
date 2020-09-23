VERSION 5.00
Begin VB.Form frmAddrDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify Address Details"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddrDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&Delete"
      Height          =   390
      Index           =   2
      Left            =   4800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   4800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   4800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1125
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1575
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2625
      Width           =   2865
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1575
      MaxLength       =   10
      TabIndex        =   9
      Top             =   2175
      Width           =   2865
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1575
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1725
      Width           =   2865
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1575
      MaxLength       =   35
      TabIndex        =   7
      Top             =   1275
      Width           =   2865
   End
   Begin VB.TextBox Text1 
      Height          =   510
      Index           =   0
      Left            =   1575
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   600
      Width           =   2865
   End
   Begin VB.Label Label1 
      Caption         =   "Country/Region:"
      Height          =   240
      Index           =   4
      Left            =   225
      TabIndex        =   5
      Top             =   2662
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "ZIP/Postal Code:"
      Height          =   240
      Index           =   3
      Left            =   225
      TabIndex        =   4
      Top             =   2212
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "State/Province:"
      Height          =   240
      Index           =   2
      Left            =   225
      TabIndex        =   3
      Top             =   1762
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "City:"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   1312
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "Street:"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   637
      Width           =   1290
   End
   Begin VB.Label lblBanner 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Verify that the address is correct"
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
      TabIndex        =   0
      Top             =   150
      Width           =   5865
   End
End
Attribute VB_Name = "frmAddrDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Public m_strAddrType As String

Private Sub cmdOpts_Click(Index As Integer)
   Select Case Index
      Case 0 'OK
         If (Not ValidateEntry()) Then Exit Sub
         
         Call PostEntry
      Case 1 'cancel
         Unload Me
      Case 2 'delete
         Call DeleteAddress
   End Select
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmAddrDetail.Form_Load"
   On Error GoTo Error_Handler
   
   'flatten all needed items
   Dim Indx As Integer
   
   For Indx = 0 To 4
      FlatBorder Text1(Indx).hWnd
   Next
   
   'setup the screen
   Call InitializeScreen
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsList.Close
   Set rsList = Nothing
   
   Set frmAddrDetail = Nothing
End Sub

Public Sub InitializeScreen()
   'set up the opening screen
   Call LoadCurrentAddress
End Sub

Private Sub LoadCurrentAddress()
   'load the needed address onto the screen
   Const sMOD_NAME As String = "frmAddrDetail.LoadCurrentAddress"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT fkContID, fkLookup, Street, City, State, Zip, Country "
   SQL = SQL & "FROM CAddress WHERE fkContID = " & g_lngContID
   SQL = SQL & " AND fkLookup = '" & m_strAddrType & "' "
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!Street)) Then Text1(0) = !Street
         If (Not IsNull(!City)) Then Text1(1) = !City
         If (Not IsNull(!State)) Then Text1(2) = !State
         If (Not IsNull(!Zip)) Then Text1(3) = !Zip
         If (Not IsNull(!Country)) Then Text1(4) = !Country
      End If
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading the Address Information!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   highLight
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(Text1(0)) < 1) Then
      Indx = MsgBox("You Must Enter A Street Address", _
         vbInformation + vbOKOnly, "Validate : Street Address")
      Text1(0).SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmAddrDetail.PostEntry"
   On Error GoTo Error_Handler
   
   rsList.Edit
   
   With rsList
      !Street = Text1(0)
      !City = Text1(1)
      !State = Text1(2)
      !Zip = Text1(3)
      !Country = Text1(4)
      
      .Update
   End With
   
   Unload Me
   Call frmContEntry.LoadAddressInfo
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Posting this entry!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub DeleteAddress()
   'delete the currently selected address
   Const sMOD_NAME As String = "frmAddrDetail.DeleteAddress"
   On Error GoTo Error_Handler
   
   Dim sMsg As String
   Dim iMsg As VbMsgBoxResult
   Dim SQL As String
   
   sMsg = "Are you sure you want to DELETE this " & vbCrLf
   sMsg = sMsg & "[ " & m_strAddrType & " ]" & " address listing for this contact?"
   
   iMsg = MsgBox(sMsg, vbQuestion + vbYesNo, "Warning : Delete Contact Address")
   
   If (iMsg <> vbYes) Then Exit Sub
   
   SQL = "DELETE * FROM CAddress WHERE fkContID = " & g_lngContID
   SQL = SQL & " AND fkLookup = '" & m_strAddrType & "' "
   
   dbContact.Execute (SQL)
   
   Call frmContEntry.LoadAddressInfo
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while deleting this address!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
