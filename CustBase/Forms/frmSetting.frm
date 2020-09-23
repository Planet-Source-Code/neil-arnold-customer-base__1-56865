VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Setting"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optSet 
      Caption         =   "Mark as Private"
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   5
      Top             =   1275
      Width           =   3465
   End
   Begin VB.OptionButton optSet 
      Caption         =   "Show as Favorite (Default)"
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   4
      Top             =   900
      Value           =   -1  'True
      Width           =   3465
   End
   Begin VB.OptionButton optSet 
      Caption         =   "Hide in the drop-down Name Lists"
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   525
      Width           =   3465
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   1350
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   2775
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   75
      X2              =   3975
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   75
      X2              =   3975
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
      Caption         =   " Select the desired show/hide setting"
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
      TabIndex        =   2
      Top             =   150
      Width           =   3840
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsList As Recordset

Dim m_strSetting As String
Dim m_blnCancelled As Boolean

Private Sub cmdOpts_Click(Index As Integer)
   Const sMOD_NAME As String = "frmSetting.cmdOpts_Click"
   On Error GoTo Error_Handler
   
   Select Case Index
      Case 0 'OK
         If (g_strFormFlag = "CEnt") Then
            Call PostContactEntry
         ElseIf (g_strFormFlag = "PEnt") Then
            Call PostProjectEntry
         End If
      Case 1 'Cancel
         m_blnCancelled = True
         Unload Me
   End Select
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub Form_Load()
   Const sMOD_NAME As String = "frmSetting.Form_Load"
   On Error GoTo Error_Handler
   
   'set default setting
   m_strSetting = "Default"
   
   m_blnCancelled = False
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (m_blnCancelled = False) Then
      If (g_strFormFlag = "CEnt") Then
         Call frmContEntry.LoadMainContactInfo
      ElseIf (g_strFormFlag = "PEnt") Then
         Call frmProjEntry.LoadMainProjectInfo
      End If
   End If
   
   'remove data & form reference
   Set frmSetting = Nothing
End Sub

Private Sub optSet_Click(Index As Integer)
   Select Case Index
      Case 0 'Hidden
         m_strSetting = "Hidden"
      Case 1 'Default
         m_strSetting = "Default"
      Case 2 'Private
         m_strSetting = "Private"
   End Select
End Sub

Private Sub PostContactEntry()
   'modify the database
   Const sMOD_NAME As String = "frmSetting.PostContactEntry"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT * FROM Contacts WHERE ContID = " & g_lngContID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   rsList.Edit
   
   With rsList
      !Setting = m_strSetting
      
      .Update
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Me.Hide
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Private Sub PostProjectEntry()
   'modify the database
   Const sMOD_NAME As String = "frmSetting.PostProjectEntry"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   
   SQL = "SELECT * FROM Projects WHERE ProjID = " & g_lngProjID
   
   Set rsList = dbContact.OpenRecordset(SQL)
   
   rsList.Edit
   
   With rsList
      !Setting = m_strSetting
      
      .Update
   End With
   
   rsList.Close
   Set rsList = Nothing
   
   Me.Hide
   
   Unload Me
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while updating the record!" & vbCrLf & _
         "Sorry for the inconvenience", , APP_MSG_NAME
End Sub
