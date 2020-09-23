Attribute VB_Name = "modSysTray"
Option Explicit

Public Type NOTIFYICONDATA
   cbSize           As Long
   hWnd             As Long
   uId              As Long
   uFlags           As Long
   uCallBackMessage As Long
   hIcon            As Long
   szTip            As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206

Public NIData As NOTIFYICONDATA

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias _
      "Shell_NotifyIconA" (ByVal dwMessage As Long, _
      pnid As NOTIFYICONDATA) As Boolean
      
'///Adapted from code by Gary Lantz///
Public Sub AddToTray()
   'add the system tray icon
   With NIData
      .cbSize = Len(NIData)
      .hWnd = frmMain.picTray.hWnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = frmMain.picTray.Picture
      .szTip = "Customer Base - Contact Manager" & vbCrLf & "Appointment / To Do Reminder"
   End With
   
   Shell_NotifyIcon NIM_ADD, NIData
End Sub

'///Adapted from code by Gary Lantz///
Public Sub RemoveFromTray()
   'remove the system tray icon
   Shell_NotifyIcon NIM_DELETE, NIData
End Sub
