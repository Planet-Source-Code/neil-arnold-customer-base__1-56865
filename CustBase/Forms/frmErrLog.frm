VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmErrLog 
   BackColor       =   &H007E5669&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Error Log"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picGrdClr 
      BackColor       =   &H00EDDFE5&
      Height          =   390
      Left            =   9300
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   1
      Top             =   2700
      Visible         =   0   'False
      Width           =   390
   End
   Begin MSComctlLib.ListView lvError 
      Height          =   3165
      Left            =   225
      TabIndex        =   0
      Top             =   150
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   5583
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Module Name:"
         Object.Width           =   5980
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number:"
         Object.Width           =   1746
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Error Description:"
         Object.Width           =   8625
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date/Time:"
         Object.Width           =   3704
      EndProperty
   End
End
Attribute VB_Name = "frmErrLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmErrLog.Form_Load"
   On Error GoTo Error_Handler
   
   Call LoadErrorList
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   
   lvError.Move 75, 75, Me.ScaleWidth - 150, Me.ScaleHeight - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmErrLog = Nothing
End Sub

Private Sub LoadErrorList()
   Const sMOD_NAME As String = "frmErrLog.LoadErrorList"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim Item As ListItem
   
   SQL = "SELECT ModName, ErrNum, ErrDesc, DateStamp "
   SQL = SQL & "FROM ErrLog ORDER BY DateStamp DESC"
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   lvError.ListItems.Clear
   
   With rsList
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ModName)) Then Set Item = lvError.ListItems.Add(, , !ModName)
            If (Not IsNull(!ErrNum)) Then Item.SubItems(1) = !ErrNum
            If (Not IsNull(!ErrDesc)) Then Item.SubItems(2) = !ErrDesc
            If (Not IsNull(!DateStamp)) Then Item.SubItems(3) = Format(!DateStamp, "mm/dd/yyyy") & " " & Format(!DateStamp, "hh:nn AMPM")
            .MoveNext
         Wend
      End If
   End With
   
   AltLVBackground lvError, picGrdClr
   
   rsList.Close
   Set rsList = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub


