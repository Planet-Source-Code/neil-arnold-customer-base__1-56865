Attribute VB_Name = "modAPI"
Option Explicit

'use this for flatborder & flatten sub's
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'use this for flatborder & flatten sub's
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4

'use this to aid in screen flicker when repainting
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

'for data backup & restore operations
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_SILENT = &H4
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_FILESONLY = &H80

Type SHFILEOPSTRUCT
   hWnd      As Long
   wFunc     As Long
   pFrom     As String
   pTo       As String
   fFlags    As Integer
   fAborted  As Boolean
   hNameMaps As Long
   sProgress As String
End Type

Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

'use to play a wav file
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
      (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'used to show .chm help files
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hWnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'///This code is from cEdit by Ackbar///
Public Sub FlatBorder(ByVal hWnd As Long)
   Const sMOD_NAME As String = "modAPI.FlatBorder"
   On Error GoTo FlatBorder_Error
   
   Dim TFlat As Long
   
   TFlat = GetWindowLong(hWnd, GWL_EXSTYLE)
   TFlat = TFlat And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
   
   SetWindowLong hWnd, GWL_EXSTYLE, TFlat
   SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
   
   Exit Sub
FlatBorder_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

'///This code is from cEdit by Ackbar///
Public Sub Flatten(ByVal frm As Form)
   Const sMOD_NAME As String = "modAPI.Flatten"
   On Error GoTo Flatten_Error
   
   Dim CTL As Control
   
   For Each CTL In frm.Controls
      Select Case TypeName(CTL)
         Case "CommandButton", "TextBox", "ListBox", "FileTree", "TreeView", "ProgressBar", "PictureBox"
            FlatBorder CTL.hWnd
      End Select
   Next
   
   Exit Sub
Flatten_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

'///This code is modified from Contact Control by Tertius Klopper///
Sub BackupDatabase()
   'backup the current user databsae
   On Error Resume Next
   
   Dim lFileOp  As Long
   Dim lresult  As Long
   Dim lFlags   As Long
   Dim SHFileOp As SHFILEOPSTRUCT
   
   Screen.MousePointer = vbHourglass
   MsgBar "Backing Up Data Files", True
   
   lFileOp = FO_COPY
   lFlags = lFlags And Not FOF_SILENT
   lFlags = lFlags Or FOF_NOCONFIRMATION
   lFlags = lFlags Or FOF_NOCONFIRMMKDIR
   lFlags = lFlags Or FOF_FILESONLY
   
   With SHFileOp
      .wFunc = lFileOp
      .pFrom = App.Path & "\Data\CBaseMgr.mdb" & vbNullChar
      .pTo = App.Path & "\Backup" & vbNullChar
      .fFlags = lFlags
   End With
   
   lresult = SHFileOperation(SHFileOp)
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub

'///This code is modified from Contact Control by Tertius Klopper///
Sub RestoreDatabase()
   'restore the user database
   On Error Resume Next
   
   Dim lFileOp  As Long
   Dim lresult  As Long
   Dim lFlags   As Long
   Dim SHFileOp As SHFILEOPSTRUCT
   
   Screen.MousePointer = vbHourglass
   MsgBar "Restoring Data Files", True
   
   lFileOp = FO_COPY
   lFlags = lFlags And Not FOF_SILENT
   lFlags = lFlags Or FOF_NOCONFIRMATION
   lFlags = lFlags Or FOF_NOCONFIRMMKDIR
   lFlags = lFlags Or FOF_FILESONLY
   
   With SHFileOp
      .wFunc = lFileOp
      .pFrom = App.Path & "\Backup\CBaseMgr.mdb" & vbNullChar
      .pTo = App.Path & "\Data" & vbNullChar
      .fFlags = lFlags
   End With
   
   lresult = SHFileOperation(SHFileOp)
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
End Sub
