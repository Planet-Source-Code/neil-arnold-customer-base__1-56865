VERSION 5.00
Begin VB.Form frmSwtContact 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Load New Contact"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOpts 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Index           =   1
      Left            =   2250
      TabIndex        =   2
      Top             =   450
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpts 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Index           =   0
      Left            =   825
      TabIndex        =   1
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DEE3E6&
      Caption         =   "Would you like to load this contact information?"
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
      TabIndex        =   0
      Top             =   75
      Width           =   4065
   End
End
Attribute VB_Name = "frmSwtContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_lngContactID As Long

Dim m_blnDoLoad As Boolean

Private Sub cmdOpts_Click(Index As Integer)
   Select Case Index
      Case 0 'OK
         m_blnDoLoad = True
         Unload Me
      Case 1 'Cancel
         m_blnDoLoad = False
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   m_blnDoLoad = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If (m_blnDoLoad = True) Then
      UnloadAllForms
      g_lngContID = m_lngContactID
      Load frmContEntry
   End If
   
   Set frmSwtContact = Nothing
End Sub
