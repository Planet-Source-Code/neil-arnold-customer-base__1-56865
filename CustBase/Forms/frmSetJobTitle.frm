VERSION 5.00
Begin VB.Form frmSetJobTitle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Job Title"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetJobTitle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   3525
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   3525
      TabIndex        =   2
      Top             =   675
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   825
      MaxLength       =   100
      TabIndex        =   0
      Top             =   225
      Width           =   2565
   End
   Begin VB.Label Label2 
      Caption         =   "Example : Co-Owner"
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
      Left            =   1425
      TabIndex        =   4
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Job Title:"
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   225
      Width           =   690
   End
End
Attribute VB_Name = "frmSetJobTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsTitle As Recordset

Const TITLETYPE = "TITLE"

Private Sub cmdOpts_Click(Index As Integer)
   Select Case Index
      Case 0 'OK
         If (Not ValidateEntry()) Then Exit Sub
         
         Call PostEntry
      Case 1 'cancel
         frmContEntry.lstJobTitle.Text = frmContEntry.Text1(4).Text
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   'flatten all needed borders
   FlatBorder Text1.hWnd
   
   'set main recordset
   Set rsTitle = dbContact.OpenRecordset("Lookup", dbOpenTable)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsTitle.Close
   Set rsTitle = Nothing
   
   Set frmSetPrefix = Nothing
End Sub

Private Sub Text1_GotFocus()
   highLight
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(Text1) < 1) Then
      Indx = MsgBox("You Must Enter A Contact Job Title", _
         vbInformation + vbOKOnly, "Validate : Contact Job Title Description")
      Text1.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmSetJobTitle.PostEntry"
   On Error GoTo Error_Handler
   
   rsTitle.AddNew
   
   With rsTitle
      !ItemID = TITLETYPE
      !Description = Text1.Text
      
      .Update
   End With
   
   Me.Hide
   
   Call frmContEntry.LoadJobTitles
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

