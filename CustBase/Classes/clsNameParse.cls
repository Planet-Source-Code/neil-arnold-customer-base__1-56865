VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNameParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
' Name: Parse Person Name Class (variety of formats)
' Description:Got extremely tired trying
'     to find a quick and systematic way of parsing a field that contained a user's
'     name in a variety of formats, so I created this little class that will parse out a
'     person's name into first, middle, last, title, prefix, suffix. It can handle
'     names like Dr. John Doe - Dr. Doe, John P -Doe, John - John P. Doe, Jr. - and a few
'     more formats. Hope others will find this useful. Currently the class cannot
'     handle muliple suffixes and multiple titles.
'     If someone reworks it to make it better, please send it my way. FYI- commented
'     all over, should be easy to read.
' By: Dino Damalas

' Inputs:None

' Returns:None

'Assumes:None

'Side Effects:None
'This code is copyrighted and has limited warranties.
'Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.36765/lngWId.1/qx
'     /vb/scripts/ShowCode.htm
'for details.
'***************************************************************************

'***************************************************************************
' Module Name: clsNameParse
' Module Type: Class Module
' Filename: clsNameParse.cls
' Author:Dino Damalas
' Date: 7/10/2002
' References:Microsoft VBScript Regular Expressions 1.0
' Purpose: Use this class when dealing with inconsistent
' person name formats. This object will parse a person's name into
' - Prefix
' - Suffix
' - First Name
' - Middle Name / Middle Initial
' - Last Name
' - Title
' examples: Dr. John P Doe Jr
'  Dr. Doe, John P.
'  John Doe
'  Doe, John P.
'  John P. Doe, CEO
'  ...etc

' Example Use: Dim objParse as new clsParse
'  objParse.ParseName("Dr. Doe, John P.")
'  strFirstName= objParse.FirstName
'  strLastName = objParse.LastName
'  strMiddleName= objParse.MiddleName
'  strMiddleInit= objParse.MiddleInitial
'  strPrefix= objParse.Prefix
'  strSuffix= objParse.Suffix
'  strTitle= objparse.title
'  set objParse = nothing
'******************************************************************

Option Explicit

'--member var declaration
Private m_objRegExp As RegExp
Private m_strPrefix As String
Private m_strSuffix As String
Private m_strLastName As String
Private m_strFirstName As String
Private m_strMiddleName As String
Private m_strMiddleInitial As String
Private m_strTitle As String
Private m_strFullName As String
Private m_blnHasError As Boolean
Private m_strErrorMessage As String

Private Sub Class_Initialize()
   m_blnHasError = False
   m_strErrorMessage = ""
   Set m_objRegExp = New RegExp
   m_objRegExp.IgnoreCase = True
End Sub

Private Sub Class_Terminate()
   Set m_objRegExp = Nothing
End Sub

'Purpose: Main sub to initiate parsing of name
Public Sub ParseName(ByVal strName As String)
   Const sMOD_NAME As String = "clsNameParse.ParseName"
   On Error GoTo ParseName_Error
   
   '-- pick apart name by removing prefix, suffix, and title
   strName = Trim(fncExtractSuffix(strName))
   strName = Trim(fncExtractPrefix(strName))
   strName = Trim(fncExtractTitle(strName))
   m_objRegExp.Global = True
   '-- check for last, first combo (Doe, John) ----
   m_objRegExp.Pattern = "[^ \f\n\r\t\v\,]+\,\s+\S+ "
   
   If m_objRegExp.Test(strName) = True Then
      Call subParseLastFirst(strName)
   Else
      '-- check if first middle last combo (John P. Doe) ---
      m_objRegExp.Pattern = "^\S+\s+\S+\s+\S+$"
      If m_objRegExp.Test(strName) Then
         Call subParseFirstMiddleLast(strName)
      Else
         '-- check if first last combo (John Doe)--
         m_objRegExp.Pattern = "^\S+\s+\S+$"
         If m_objRegExp.Test(strName) Then
            Call subParseFirstLast(strName)
         Else
            '--if does not fit in this format tell user
            '   we have a problem
            m_blnHasError = True
            m_strErrorMessage = "Unable To Parse"
         End If
      End If
   End If
   
   Exit Sub
ParseName_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

' Purpose: Removes the prefix from the name and sets the
'          Prefix property of the class
Private Function fncExtractPrefix(ByVal strName As String) As String
   Const sMOD_NAME As String = "clsNameParse.fncExtractPrefix"
   On Error GoTo fncExtractPrefix_Error
   
   '--declare vars
   Dim aryPrefix As Variant
   Dim intCounter As Integer
   Dim strReturn As String
   Dim objMatches As MatchCollection
   '--initialize vars
   strReturn = strName
   '--populate array with a bunch of possible prefixes
   aryPrefix = Array("mr", "mrs", "miss", "dr", "prof", "pvt", "pfc", _
                     "lcpl", "cpl", "spc", "sgt", "ssgt", "gysgt", "msgt", _
                     "mgysgt", "lt", "capt", "col", "ltcol", "gen", "adm", "rdm")
   '--loop through the array looking for matches using regexp
   m_objRegExp.Global = False
   
   
   For intCounter = 0 To UBound(aryPrefix)
      m_objRegExp.Pattern = "^" & aryPrefix(intCounter) & "\.?\s+"
      
      If m_objRegExp.Test(strName) Then
         '-- if found, replace with empty string
         strReturn = Trim(m_objRegExp.Replace(strName, ""))
         Set objMatches = m_objRegExp.Execute(strName)
         '--set prefix property
         Me.Prefix = Trim(objMatches(0).Value)
         Set objMatches = Nothing
         Exit For
      End If
   Next
   
   fncExtractPrefix = strReturn
   
   Exit Function
fncExtractPrefix_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

'Purpose: Removes the suffix from the
'     name and sets the Suffix property of the class
Private Function fncExtractSuffix(ByVal strName As String) As String
   Const sMOD_NAME As String = "clsNameParse.fncExtractSuffix"
   On Error GoTo fncExtractSuffix_Error
   
   '--declare vars
   Dim arySuffix As Variant
   Dim intCounter As Integer
   Dim strReturn As String
   Dim objMatches As MatchCollection
   '--initialize vars
   strReturn = strName
   '--populate array with a bunch of possible suffixes
   arySuffix = Array("md", "i", "ii", "iid", "iii", "iv", "jr", _
                     "sr", "v", "vi", "vii", "do", "dds", "np", "pa", _
                     "phd", "ph d", "esq")
   '--loop through the array looking for matches using regexp
   m_objRegExp.Global = False
   
   For intCounter = 0 To UBound(arySuffix)
      m_objRegExp.Pattern = "\b" & arySuffix(intCounter) & "\.?(\s+|$)"
      
      If m_objRegExp.Test(strName) Then
         '-- if found, replace with empty string
         strReturn = Trim(m_objRegExp.Replace(strName, ""))
         Set objMatches = m_objRegExp.Execute(strName)
         '--set prefix property
         Me.Suffix = Trim(objMatches(0).Value)
         Set objMatches = Nothing
         Exit For
      End If
   Next
   fncExtractSuffix = strReturn
   
   Exit Function
fncExtractSuffix_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

'Purpose: Removes title from name and sets the title property of the class
Private Function fncExtractTitle(ByVal strName As String) As String
   Const sMOD_NAME As String = "clsNameParse.fncExtractTitle"
   On Error GoTo fncExtractTitle_Error
   
   '--delcare vars
   Dim strReturn As String
   Dim intCommaPos As Integer
   Dim objMatches As MatchCollection
   Dim objMatch As Match
   '--initialize vars
   strReturn = strName
   '--get the first position of a comma
   intCommaPos = InStr(1, strName, ",", vbTextCompare)
   '--see if we have a comma in the name
   
   
   If intCommaPos > 0 Then
      m_objRegExp.Pattern = "[^ \f\n\r\t\v\,]+\,\s+\S+"
      '--check to see if this comma is lastname, firstname format
      
      If m_objRegExp.Test(strName) = True Then
         '--check to see if there is another comma
         '   since first is a last, first name seperator
         If InStr(intCommaPos + 1, strName, ",", vbTextCompare) > 0 Then
            '--if the last character is not a comma
            '     then parse out the title
            If Right(Trim(strName), 1) <> "," Then
               m_objRegExp.Pattern = "\,\s+\S+\s*$"
               Set objMatches = m_objRegExp.Execute(strName)
               
               For Each objMatch In objMatches
                  '--set the title
                  Me.Title = fncScrubString(objMatch.Value)
               Next
               Set objMatches = Nothing
               strReturn = m_objRegExp.Replace(strName, "")
            End If
         End If
      End If
   End If
   
   fncExtractTitle = strReturn
   
   Exit Function
fncExtractTitle_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Function

'Purpose: removes any commas or extra spacings from name part
Private Function fncScrubString(ByVal strNamePart As String) As String
   fncScrubString = Trim(Replace(strNamePart, ",", ""))
End Function

'Purpose: Parses a name that is in LastName, FirstName format
Private Sub subParseLastFirst(ByVal strName As String)
   Const sMOD_NAME As String = "clsNameParse.subParseLastFirst"
   On Error GoTo subParseLastFirst_Error
   
   '--declare vars
   Dim objMatches As MatchCollection
   Dim objMatch As Match
   Dim intCounter As Integer
   '--initialize
   intCounter = 1
   m_objRegExp.Global = True
   m_objRegExp.Pattern = "\S+"
   Set objMatches = m_objRegExp.Execute(strName)
   
   For Each objMatch In objMatches
      Select Case intCounter
         Case 1 '-- first time around is last name
            Me.LastName = fncScrubString(objMatch.Value)
         Case 2 '-- second time around is first name
            Me.FirstName = fncScrubString(objMatch.Value)
         Case 3 '-- if there is a third than its the middlename
            Me.MiddleName = fncScrubString(objMatch.Value)
      End Select
      intCounter = intCounter + 1
   Next
   
   Exit Sub
subParseLastFirst_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

'Purpose: Parses a name in FirstName LastName format (no middle name)
Private Sub subParseFirstLast(ByVal strName As String)
   Const sMOD_NAME As String = "clsNameParse.subParseFirstLast"
   On Error GoTo subParseFirstLast_Error
   
   '--declare vars
   Dim objMatches As MatchCollection
   Dim objMatch As Match
   Dim intCounter As Integer
   '--initialize
   intCounter = 1
   '--set up regexp object
   m_objRegExp.Global = True
   m_objRegExp.Pattern = "\S+"
   Set objMatches = m_objRegExp.Execute(strName)
   '--run through matches
   
   For Each objMatch In objMatches
      Select Case intCounter
         Case 1 '-- first time around we set first name
            Me.FirstName = fncScrubString(objMatch.Value)
         Case 2 '-- second time we set last name
            Me.LastName = fncScrubString(objMatch.Value)
      End Select
      intCounter = intCounter + 1
   Next
   
   Exit Sub
subParseFirstLast_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

'Purpose: Parses a name in FirstName Middlename LastName format
Private Sub subParseFirstMiddleLast(ByVal strName As String)
   Const sMOD_NAME As String = "clsNameParse.subParseFirstMiddleLast"
   On Error GoTo subParseFirstMiddleLast_Error
   
   '--declare vars
   Dim objMatches As MatchCollection
   Dim objMatch As Match
   Dim intCounter As Integer
   '--initialize vars
   intCounter = 1
   '--set up regexp object
   m_objRegExp.Global = True
   m_objRegExp.Pattern = "\S+"
   Set objMatches = m_objRegExp.Execute(strName)
   '--loop thorough matches
   
   For Each objMatch In objMatches
      Select Case intCounter
         Case 1 '-- first time is firstname
            Me.FirstName = fncScrubString(objMatch.Value)
         Case 2 '-- second time around is middlename
            Me.MiddleName = fncScrubString(objMatch.Value)
         Case 3 '-- last time around is last name
            Me.LastName = fncScrubString(objMatch.Value)
      End Select
      intCounter = intCounter + 1
   Next
   
   Exit Sub
subParseFirstMiddleLast_Error:
   LogErrors sMOD_NAME, Err.Number, Err.Description
End Sub

'Purpose: Use this sub to clear out members when you implementing
'   in code where you don't reinstantiate the object again
Public Sub Clear()
   Me.FirstName = ""
   Me.MiddleInitial = ""
   Me.MiddleName = ""
   Me.LastName = ""
   Me.Suffix = ""
   Me.Prefix = ""
   Me.Title = ""
   m_blnHasError = False
   m_strErrorMessage = ""
End Sub

Public Property Get Prefix() As String
   Prefix = m_strPrefix
End Property

Public Property Let Prefix(ByVal strPrefix As String)
   m_strPrefix = strPrefix
End Property

Public Property Get Suffix() As String
   Suffix = m_strSuffix
End Property

Public Property Let Suffix(ByVal strSuffix As String)
   m_strSuffix = strSuffix
End Property

Public Property Get LastName() As String
   LastName = m_strLastName
End Property

Public Property Let LastName(ByVal strLastName As String)
   m_strLastName = strLastName
End Property

Public Property Get FirstName() As String
   FirstName = m_strFirstName
End Property

Public Property Let FirstName(ByVal strFirstName As String)
   m_strFirstName = strFirstName
End Property

Public Property Get MiddleName() As String
   MiddleName = m_strMiddleName
End Property

Public Property Let MiddleName(ByVal strMiddleName As String)
   m_strMiddleName = strMiddleName
   '--set up middle initial while we're here
   
   If Len(strMiddleName) > 1 Then
      Me.MiddleInitial = Left(strMiddleName, 1)
   Else
      Me.MiddleInitial = ""
   End If
End Property

Public Property Get MiddleInitial() As String
   MiddleInitial = m_strMiddleInitial
End Property

Public Property Let MiddleInitial(ByVal strMiddleInitial As String)
   m_strMiddleInitial = strMiddleInitial
End Property

Public Property Get Title() As String
   Title = m_strTitle
End Property

Public Property Let Title(ByVal strTitle As String)
   m_strTitle = strTitle
End Property

Public Property Get HasError() As Boolean
   HasError = m_blnHasError
End Property

Public Property Get ErrorMessage() As String
   ErrorMessage = m_strErrorMessage
End Property

      

