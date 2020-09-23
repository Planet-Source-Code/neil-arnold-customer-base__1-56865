Attribute VB_Name = "modGeneral"
Option Explicit

'GLOBAL VARIABLES
Public g_wsWorkSpc As Workspace
Public dbContact As Database
Public g_strDBName As String
Public g_lngContID As Long
Public g_lngProjID As Long
Public g_strFormFlag As String
Public g_strNameSQL As String
Public g_blnAltColors As Boolean
Public g_blnShowLines As Boolean
Public g_blnIsSecure As Boolean

'Public constants
Public Const APPNAME = "CustomerBase"
Public Const APP_CATEGORY = "Application"
Public Const APP_MSG_NAME = "Customer Base - Contact Manager"

'Enumeration for setting the state of the program's
'edit mode.
Public Enum curState
   NOW_ADDING
   NOW_SAVING
   NOW_EDITING
   NOW_DELETING
   NOW_IDLE
End Enum
Public icurState As curState

'for error log
Dim rsError As Recordset

Function GetRegistryString(ByVal vsItem As String, ByVal vsDefault As String) As String
   GetRegistryString = GetSetting(APP_CATEGORY, APPNAME, vsItem, vsDefault)
End Function

Sub Main()
   'frmMain.Show
   On Error Resume Next
   
   Load frmMain
End Sub

Sub ShowError()
   Dim sTemp As String
   
   Screen.MousePointer = vbDefault
   
   sTemp = "The following Error occurred:" & vbCrLf & vbCrLf
   'add the error string
   sTemp = sTemp & "Description : " & Err.Description & vbCrLf
   'add the error number
   sTemp = sTemp & "Number : " & Err
   
   Beep
   
   MsgBox sTemp
End Sub

Sub HideMainTools()
   With frmMain
      .mnuFilePrint.Enabled = False
      .mnuEditDelete.Enabled = False
      
      .tbrMain.Buttons(7).Enabled = False
   End With
End Sub

Sub OpenLocalDB()
   Const sMOD_NAME As String = "modGeneral.OpenLocalDB"
   On Error GoTo OpenLocalDB_Error
   
   MsgBar "Opening Database", True
   Screen.MousePointer = vbHourglass
  
   g_strDBName = App.Path & "\Data\CBaseMgr.mdb"
   
   Set dbContact = g_wsWorkSpc.OpenDatabase(g_strDBName, False)
   
   Screen.MousePointer = vbDefault
   MsgBar vbNullString, False
  
   Exit Sub
OpenLocalDB_Error:
   'LogErrors sMOD_NAME, Err.Number, Err.Description
   ShowError
End Sub

Sub ShutDownCustBase()
   On Error Resume Next
   
   'save all the current Registry settings
   SaveRegistrySettings
   
   UnloadAllForms
   
   dbContact.Close
   Set dbContact = Nothing
   
   'End 'removed to stop program crash when compiled version closes 10.24.04
End Sub

Sub UnloadAllForms()
   On Error Resume Next
   
   Dim i As Integer
   
   For i = Forms.Count - 1 To 1 Step -1
      Unload Forms(i)
   Next
End Sub

Sub SaveRegistrySettings()
   On Error Resume Next
   
   SaveSetting APP_CATEGORY, APPNAME, "WindowState", frmMain.WindowState
   If frmMain.WindowState = vbNormal Then
      SaveSetting APP_CATEGORY, APPNAME, "WindowTop", frmMain.Top
      SaveSetting APP_CATEGORY, APPNAME, "WindowLeft", frmMain.Left
      SaveSetting APP_CATEGORY, APPNAME, "WindowWidth", frmMain.Width
      SaveSetting APP_CATEGORY, APPNAME, "WindowHeight", frmMain.Height
   End If
End Sub

Public Sub highLight()
   Const sMOD_NAME As String = "modGeneral.highLight"
   On Error GoTo highLight_Error
   
   With Screen.ActiveForm
      If (TypeOf .ActiveControl Is TextBox) Then
         .ActiveControl.SelStart = 0
         .ActiveControl.SelLength = (Len(.ActiveControl))
      'ElseIf (TypeOf .ActiveControl Is MaskEdBox) Then
         '.ActiveControl.PromptInclude = True
         '.ActiveControl.SelStart = 0
         '.ActiveControl.SelLength = (Len(.ActiveControl))
         '.ActiveControl.PromptInclude = False
      End If
   End With
   
   Exit Sub
highLight_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Sub MsgBar(rsMsg As String, rPauseFlag As Integer)
   If Len(rsMsg) = 0 Then
      Screen.MousePointer = vbDefault
      frmMain.lblStatus.Caption = "Ready ..."
   Else
      If rPauseFlag Then
         frmMain.lblStatus.Caption = rsMsg & ", please wait..."
      Else
         frmMain.lblStatus.Caption = rsMsg
      End If
   End If
End Sub

Sub ReloadContactForm()
   Const sMOD_NAME As String = "modGeneral.ReloadContactForm"
   On Error GoTo ReloadContactForm_Error
   
   Dim frmContEntry As New frmContEntry
   
   UnloadAllForms
   Load frmContEntry
   
   Exit Sub
ReloadContactForm_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred!" & vbCrLf & "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Sub LogErrors(strMod As String, lngErrNum As Long, strErrDesc As String)
   Const sMOD_NAME As String = "modGeneral.LogErrors"
   On Error GoTo LogErrors_Error
   
   Set rsError = dbContact.OpenRecordset("ErrLog", dbOpenTable)
   
   rsError.AddNew
   
   With rsError
      If (strMod <> "") Then !ModName = strMod
      If (Not IsNull(lngErrNum)) Then !ErrNum = lngErrNum
      If (strErrDesc <> "") Then !ErrDesc = strErrDesc
      
      .Update
   End With
   
   rsError.Close
   Set rsError = Nothing
   
   Exit Sub
ReloadContactForm_Error:
   Exit Sub
End Sub

Public Function SrchReplace(ByVal sStrToFix) As String
   Dim iPos As Integer
   Dim sCharToRepl As String
   Dim sReplWith As String
   Dim sTempString As String
   
   sCharToRepl = "'"
   sReplWith = "''"
   
   iPos = InStr(sStrToFix, sCharToRepl)
   sTempString = ""
   
   Do While iPos
      sTempString = sTempString & Left$(sStrToFix, iPos - 1)
      sTempString = sTempString & sReplWith
      sTempString = sTempString & _
         Mid$(sStrToFix, iPos + 1, Len(sStrToFix))
      iPos = InStr(iPos + 1, sStrToFix, sCharToRepl)
   Loop
   
   SrchReplace = sTempString
End Function

Function ConvertContactName(lngContID As Long) As String
   'lookup the contact name from the passed contact id
   Const sMOD_NAME As String = "modGeneral.ConvertContactName"
   On Error GoTo Error_Handler
   
   Dim cSQL As String
   Dim rsConv As Recordset
   
   cSQL = "SELECT ContID, ShownName FROM Contacts "
   cSQL = cSQL & "WHERE ContID = " & lngContID
   
   Set rsConv = dbContact.OpenRecordset(cSQL)
   
   With rsConv
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!ShownName)) Then
            ConvertContactName = !ShownName
         End If
      End If
   End With
   
   rsConv.Close
   Set rsConv = Nothing
   
   Exit Function
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

Function ConvertProjectName(lngProjID As Long) As String
   'lookup the contact name from the passed contact id
   Const sMOD_NAME As String = "modGeneral.ConvertProjectName"
   On Error GoTo Error_Handler
   
   Dim cSQL As String
   Dim rsConv As Recordset
   
   cSQL = "SELECT ProjID, PName FROM Projects "
   cSQL = cSQL & "WHERE ProjID = " & lngProjID
   
   Set rsConv = dbContact.OpenRecordset(cSQL)
   
   With rsConv
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!PName)) Then
            ConvertProjectName = !PName
         End If
      End If
   End With
   
   rsConv.Close
   Set rsConv = Nothing
   
   Exit Function
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

Sub LoadContactNames(objCombo As Object)
   'load all names into objCombo
   Const sMOD_NAME As String = "modGeneral.LoadAllNames"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim rsListing
   
   SQL = "SELECT ContID, ShownName FROM Contacts "
   SQL = SQL & "ORDER BY ShownName"
   
   Set rsListing = dbContact.OpenRecordset(SQL)
   
   With rsListing
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ContID)) Then
               If (Not IsNull(!ShownName)) Then objCombo.AddItem Trim(!ShownName)
               objCombo.ItemData(objCombo.NewIndex) = !ContID
            End If
            .MoveNext
         Wend
      End If
   End With
   
   rsListing.Close
   Set rsListing = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Sub LoadProjectNames(objCombo As Object)
   'load all projects into objCombo
   Const sMOD_NAME As String = "modGeneral.LoadAllProjects"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim rsListing
   
   SQL = "SELECT ProjID, PName FROM Projects "
   SQL = SQL & "ORDER BY PName"
   
   Set rsListing = dbContact.OpenRecordset(SQL)
   
   With rsListing
      If (.RecordCount > 0) Then
         .MoveFirst
         While Not .EOF
            If (Not IsNull(!ProjID)) Then
               If (Not IsNull(!PName)) Then objCombo.AddItem !PName
               objCombo.ItemData(objCombo.NewIndex) = !ProjID
            End If
            .MoveNext
         Wend
      End If
   End With
   
   rsListing.Close
   Set rsListing = Nothing
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

'//The following code adapted from sample authored by Harald W. Genauck (http://www.aboutvb.de)
Sub AltLVBackground(lv As ListView, pic As PictureBox, _
         Optional ByVal StartAtOddRow As Boolean = False, _
         Optional ByVal AltBackColor As OLE_COLOR = -1)
        
   Const sMOD_NAME As String = "modGeneral.AltLVBackground"
   On Error GoTo Error_Handler
   
   Dim h               As Single
   Dim sw              As Single
   Dim oAltBackColor   As OLE_COLOR
   
   '//If AltBackColor is not passed, default to picturebox backcolor
   If AltBackColor = -1 Then
      oAltBackColor = pic.BackColor
   Else
      oAltBackColor = AltBackColor
   End If
   
   With lv
      If .View = lvwReport Then
         If .ListItems.Count Then
            .PictureAlignment = lvwTile
            h = .ListItems(1).Height
            With pic
               .Visible = False
               .BackColor = lv.BackColor
               .BorderStyle = 0
               .Height = h * 2
               .Width = 10 * Screen.TwipsPerPixelX
               sw = .ScaleWidth
               .AutoRedraw = True
               If StartAtOddRow Then
                  pic.Line (0, 0)-Step(sw, h - Screen.TwipsPerPixelY), oAltBackColor, BF
               Else
                  pic.Line (0, h)-Step(sw, h), oAltBackColor, BF
               End If
               Set lv.Picture = .Image
               .AutoRedraw = False
               .BackColor = oAltBackColor
            End With
            .Refresh
            Exit Sub
         End If
      End If
      Set .Picture = Nothing
   End With
   
   Exit Sub
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

Function GetPhoneNum(lngContID As Long) As String
   Const sMOD_NAME As String = "modGeneral.GetPhoneNum"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim rsConv As Recordset
   Dim strType As String
   
   strType = "Home"
   
   SQL = "SELECT fkContID, fkLookup, PhoneNum FROM CPhone "
   SQL = SQL & "WHERE fkContID = " & lngContID
   SQL = SQL & " AND fkLookup = '" & strType & "' "
   
   Set rsConv = dbContact.OpenRecordset(SQL)
   
   With rsConv
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!PhoneNum)) Then
            GetPhoneNum = !PhoneNum
         End If
      End If
   End With
   
   rsConv.Close
   Set rsConv = Nothing
   
   Exit Function
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

Function GetEMail(lngContID As Long) As String
   Const sMOD_NAME As String = "modGeneral.GetEMail"
   On Error GoTo Error_Handler
   
   Dim SQL As String
   Dim rsConv As Recordset
   Dim strType As String
   
   strType = "Personal"
   
   SQL = "SELECT fkContID, fkLookup, Email FROM CEMail "
   SQL = SQL & "WHERE fkContID = " & lngContID
   SQL = SQL & " AND fkLookup = '" & strType & "' "
   
   Set rsConv = dbContact.OpenRecordset(SQL)
   
   With rsConv
      If (.RecordCount > 0) Then
         .MoveFirst
         If (Not IsNull(!Email)) Then
            GetEMail = !Email
         End If
      End If
   End With
   
   rsConv.Close
   Set rsConv = Nothing
   
   Exit Function
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

Sub CompactDB()
   Const sMOD_NAME As String = "modGeneral.CompactDB"
   On Error GoTo Error_Handler
   
   Dim sOldName As String
   Dim sNewName As String
   Dim sNewName2 As String
   Dim nEncrypt As Integer
   
   'the file name to compact
   sOldName = App.Path & "\Data\CBaseMgr.mdb"
   
   'the file name to compact to
   sNewName = App.Path & "\Data\CBaseMgr.mdb"
   
   Screen.MousePointer = vbHourglass
   MsgBar "Compacting " & APP_MSG_NAME & " database.", True
   'we are going to overwrite the same file, so we need to create a new MDB
   'and rename after the compact is successful
   If sOldName = sNewName Then
      sNewName2 = sNewName 'save the new name
      sNewName = Left(sNewName, Len(sNewName) - 1) & "N"
   End If
   
   'unload all forms & close the database
   UnloadAllForms
   dbContact.Close
   Set dbContact = Nothing
   
   DBEngine.CompactDatabase sOldName, sNewName, dbLangGeneral, dbVersion30
   
   'check for an overwrite of the original mdb
   If VBA.Right(sNewName, 1) = "N" Then
      Kill sNewName2             'nuke the old one
      Name sNewName As sNewName2 'rename the new one to the original name
      sNewName = sNewName2       'reset to the correct name
   End If
   
   're-open the compacted database
   Call OpenLocalDB
   Load frmHome
   
   MsgBar vbNullString, False
   Screen.MousePointer = vbDefault
   
   MsgBox "The database was sucessfully compacted", , APP_MSG_NAME
   
   Exit Sub
   
Error_Handler:
   'LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while Compacting the Database!" & vbCrLf _
      & "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Sub WrapPrintText(strText As String)
   'word wrap the large portion of text memos
   Dim i As Integer, sCurrWord As String
   
   i = 1
   
   Do Until i > Len(strText)
      sCurrWord = ""
      Do Until i > Len(strText) Or Mid$(strText, i, 1) <= " "
         sCurrWord = sCurrWord & Mid$(strText, i, 1)
         i = i + 1
      Loop
      
      If (Printer.CurrentX + Printer.TextWidth(sCurrWord)) > 19.8 Then
         Printer.CurrentX = 3
         Printer.CurrentY = Printer.CurrentY + 0.38
      End If
      
      Printer.Print sCurrWord;
      
      Do Until i > Len(strText) Or Mid$(strText, i, 1) > " "
         Select Case Mid$(strText, i, 1)
            Case " "
               Printer.Print " ";
            Case Chr$(10) 'line-feed
               Printer.CurrentX = 3
               Printer.CurrentY = Printer.CurrentY + 0.38
            Case Else
               'take no action
         End Select
         i = i + 1
      Loop
   Loop
End Sub

Public Sub LoadRegistrySettings()
   Const sMOD_NAME As String = "modGeneral.LoadRegistrySettings"
   On Error GoTo Error_Handler
   
   g_blnAltColors = Val(GetRegistryString("AltColors", "-1"))
   g_blnShowLines = Val(GetRegistryString("GridLines", "0"))
   g_blnIsSecure = Val(GetRegistryString("Security", "0"))
   
   Exit Sub
   
Error_Handler:
   LogErrors sMOD_NAME, Err.Number, Err.Description
   MsgBox "An un-known error occurred while loading Saved Settings!" & vbCrLf _
      & "Sorry for the inconvenience", , APP_MSG_NAME
End Sub

Sub CheckForCurrentDB()
   'check for any database updates
   Dim tblDef As TableDef
   Dim blnDBIsCurrent As Boolean
   
   blnDBIsCurrent = False
   
   For Each tblDef In dbContact.TableDefs
      If tblDef.Name = "Security" Then
         blnDBIsCurrent = True
      End If
   Next
   
   If (blnDBIsCurrent = False) Then
      MsgBox "The database is not current." & _
         vbCrLf & "It will be updated for you automatically", , _
         "Updating Database"
      Call CreateSecurityTable
      MsgBox "Update Complete.", , APP_MSG_NAME
      Exit Sub
   Else
      Exit Sub
   End If
End Sub

Sub CreateSecurityTable()
   Dim tblDef As TableDef
   Dim fldDef As Field
   Dim idxIndex As Index

   On Error Resume Next

   'Create new table
   Set tblDef = dbContact.CreateTableDef("Security")

   'Create the fields in the Security table
   '--- Define A New Field ---
   Set fldDef = tblDef.CreateField("RefNum", dbLong)
   fldDef.Attributes = dbAutoIncrField
   tblDef.Fields.Append fldDef
   '--- Define A New Field ---
   Set fldDef = tblDef.CreateField("UserName", dbText, 50)
   fldDef.AllowZeroLength = True
   tblDef.Fields.Append fldDef
   '--- Define A New Field ---
   Set fldDef = tblDef.CreateField("Password", dbText, 50)
   fldDef.AllowZeroLength = True
   tblDef.Fields.Append fldDef

   'Append all fields to the newly created table
   dbContact.TableDefs.Append tblDef

   '--- Define A New Index ---
   Set idxIndex = tblDef.CreateIndex("PrimaryKey")
   With idxIndex
      .Fields = "RefNum"
      .Primary = True
   End With
   tblDef.Indexes.Append idxIndex
   '--- Define A New Index ---
   Set idxIndex = tblDef.CreateIndex("RefNum")
   With idxIndex
      .Fields = "RefNum"
   End With
   tblDef.Indexes.Append idxIndex
End Sub

Public Function FileExists(sFileName As String) As Boolean
   '** Description:
   '** Check to see if file exists
   On Error GoTo FExistsError
   
   Dim F As String
   
   F = FreeFile
   Open sFileName For Input As #F 'Open file
   Close #F
FExistsError:
   If Err.Number = 53 Then 'If doesn't exists
      FileExists = False 'Set FileExists to False
   ElseIf Err.Number = 0 Then 'else if exists
      FileExists = True 'Set FileExists to True
   End If
End Function
