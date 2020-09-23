VERSION 5.00
Begin VB.Form frmSetProjStatus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Project Status"
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
   Icon            =   "frmSetProjStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   825
      MaxLength       =   100
      TabIndex        =   0
      Top             =   225
      Width           =   2565
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
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   240
      Left            =   75
      TabIndex        =   4
      Top             =   225
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "Example : Is Working"
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
      Left            =   1500
      TabIndex        =   3
      Top             =   600
      Width           =   1290
   End
End
Attribute VB_Name = "frmSetProjStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsStatus As Recordset

Const STATTYPE = "PSTAT"

Private Sub cmdOpts_Click(Index As Integer)
   Select Case Index
      Case 0 'OK
         If (Not ValidateEntry()) Then Exit Sub
         
         Call PostEntry
      Case 1 'cancel
         frmProjEntry.lstStatus.Text = frmProjEntry.Text1(1).Text
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   'flatten all needed borders
   FlatBorder Text1.hWnd
   
   'set main recordset
   Set rsStatus = dbContact.OpenRecordset("Lookup", dbOpenTable)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'remove data & form reference
   rsStatus.Close
   Set rsStatus = Nothing
   
   Set frmSetProjStatus = Nothing
End Sub

Private Sub Text1_GotFocus()
   highLight
End Sub

Private Function ValidateEntry() As Boolean
   Dim Indx As Integer
   
   ValidateEntry = True
   
   If (Len(Text1) < 1) Then
      Indx = MsgBox("You Must Enter A Project Status Item", _
         vbInformation + vbOKOnly, "Validate : Project Status Item Description")
      Text1.SetFocus
      ValidateEntry = False
      Exit Function
   End If
End Function

Private Sub PostEntry()
   Const sMOD_NAME As String = "frmSetProjStatus.PostEntry"
   On Error GoTo Error_Handler
   
   rsStatus.AddNew
   
   With rsStatus
      !ItemID = STATTYPE
      !Description = Text1.Text
      
      .Update
   End With
   
   Me.Hide
   
   Call frmProjEntry.LoadStatusItems
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

