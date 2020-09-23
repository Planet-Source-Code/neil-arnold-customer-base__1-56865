VERSION 5.00
Begin VB.Form frmBackRest 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1740
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2415
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   2415
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   75
      Top             =   1200
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   390
      Left            =   75
      TabIndex        =   0
      Top             =   675
      Width           =   2265
   End
End
Attribute VB_Name = "frmBackRest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_strType As String
 
Dim m_intTimer As Integer

Private Sub Form_Activate()
   Select Case m_strType
      Case "Backup"
         lblProgress.Caption = "Backing Up Current Data ..."
         Call BackupDatabase
         Timer1.Enabled = True
      Case "Restore"
         lblProgress.Caption = "Restoring Program Data ..."
         Call RestoreDatabase
         Timer1.Enabled = True
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmBackRest = Nothing
End Sub

Private Sub Timer1_Timer()
   m_intTimer = m_intTimer + 1
   
   If m_intTimer = 3 Then
      Select Case m_strType
         Case "Backup"
            lblProgress.Caption = "Data Backup Complete ..."
            Me.Hide
            Unload Me
         Case "Restore"
            lblProgress.Caption = "Data Restore Complete ..."
            Timer1.Enabled = False
            
            'open restored database
            Call OpenLocalDB
            Load frmHome
   
            Me.Hide
            Unload Me
      End Select
   End If
End Sub
