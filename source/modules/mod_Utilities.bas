''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                     UTILITY FUNCTIONS                        '
'                                                              '
'       This module contains useful functions that you         '
'       can use in expressions on your forms and reports.      '
'                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Database   'Use database order for string comparisons
Option Explicit

Function IsLoaded(ByVal strFormName As String) As Integer
' From Northwind sample database:
' Returns True if the specified form is open in Form view or Datasheet view.
 
' These variables are used to test the return values of the SysCmd function
'  and the CurrentView property of the requested form.
    Const conObjStateClosed = 0
    Const conDesignView = 0

    'The first If statement uses the SysCmd function to check the current
    'state of the requested form. It can be in one of four states: not open
    'or nonexistent, open, new, or changed but not saved.
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        
        'The second If statement checks for the current view of the requested
        'form, assuming the previous If statement found it to be open. If
        'the form is currently open in Form View, the function will return
        'true. If the form is in Design View, the function will return false.
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    
    End If
    
End Function


Function IsNothing(varToTest As Variant) As Integer
  ' Comments: Tests for a "logical" nothing based on data type
  '           Empty and Null = Nothing
  '           Number = 0 is Nothing
  '           Zero length string is Nothing
  '           Date/Time is never Nothing
  ' Parameters: varToTest = Variant data to test
  ' Return: True/false
  ' Dependencies: No
  ' Created: 7/27/00 MAW
  ' Modified:
  '
  ' --------------------------------------------------------
On Error GoTo Err_IsNothing

    IsNothing = True

    Select Case varType(varToTest)
        Case vbEmpty
            Exit Function
        Case vbNull
            Exit Function
        Case vbBoolean
            If varToTest Then IsNothing = False
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
            If varToTest <> 0 Then IsNothing = False
        Case vbDate
            IsNothing = False
        Case vbString
            If (Len(varToTest) <> 0 And varToTest <> " ") Then IsNothing = False
    End Select

Exit_IsNothing:
    On Error GoTo 0
    Exit Function

Err_IsNothing:
    MsgBox "Error#" & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "IsNothing"
    Resume Exit_IsNothing

End Function

Function Larger(lngA() As Long) As Integer
'  Input:  An array of Long values of any length
'
'  Output: Integer index of the largest value
'
    
    Dim lngBig As Long, intX As Integer, intI As Integer

    intI = LBound(lngA)
    lngBig = lngA(intI)
    Larger = intI

    For intX = intI + 1 To UBound(lngA)
        If lngA(intX) > lngBig Then
            lngBig = lngA(intX)
            Larger = intX
        End If
    Next intX

End Function

Function Smaller(lngA() As Long) As Integer
'  Input:  An array of Long values of any length
'
'  Output: Integer index of the smallest value
'
    
    Dim lngSmall As Long, intX As Integer, intI As Integer

    intI = LBound(lngA)
    lngSmall = lngA(intI)
    Smaller = intI

    For intX = intI + 1 To UBound(lngA)
        If lngA(intX) < lngSmall Then
            lngSmall = lngA(intX)
            Smaller = intX
        End If
    Next intX

End Function

Public Function FormAssist()
' Generic function to respond to OnAction on
'   custom menu or toolbar help command
' Function checks for an active form, then
'   looks for a "FormHelp" handler subroutine
'   on that form.
Dim frm As Form

    ' Set an error trap
    On Error Resume Next

    ' Try to locate a form that has the focus
    Set frm = Screen.ActiveForm
    If Err <> 0 Then
        ' Error means no active form,
        '  so open standard Office Assistant
        Application.Assistant.Help
        Exit Function
    End If
    
    ' No error, so try to call the FormHelp
    '   method of the active form
    frm.FormHelp
    If Err <> 0 Then
        ' Error means no FormHelp method for
        '  the current form,
        '  so open standard Office Assistant
        Application.Assistant.Help
    End If
End Function

Public Sub FindMatch(rstTemp As DAO.Recordset, _
    strFind As String)

    Dim varBookmark As Variant
    Dim strMessage As String

    With rstTemp
        ' Store current record location.
        varBookmark = .Bookmark
        .FindFirst strFind

        ' If Find method fails, notify user and return to the
        ' last current record.
        If .NoMatch Then
            strMessage = _
                "Not found! Returning to current record." & _
                vbCr & vbCr & "NoMatch = " & .NoMatch

MsgBox strMessage
            .Bookmark = varBookmark
        End If
    End With
End Sub

Public Function FileExists(varFile As Variant) As Boolean
'Return whether a file exists
On Error GoTo Err_FileExists

If IsNull(varFile) Then
    FileExists = False
    Exit Function
End If
FileExists = (Len(Dir(varFile)) > 0)

Exit_FileExists:
    Exit Function

Err_FileExists:
    FileExists = False
    Resume Exit_FileExists

End Function

Public Function FolderExists(varPath As Variant) As Boolean
    'Return whether a folder exists
    'Added 6/14/2011 Based on Code from Allen Browne
    On Error Resume Next
    If Len(varPath) > 0 Then
        FolderExists = (Len(Dir$(varPath, vbDirectory)) > 0&)
    End If
End Function

Public Function GetPath(ByVal strFilePath As String) As String

On Error GoTo Err_Handler

Dim strTemp As String

Do While (InStr(strFilePath, "\") > 0)
    strTemp = strTemp & Left(strFilePath, InStr(strFilePath, "\"))
    strFilePath = Mid(strFilePath, InStr(strFilePath, "\") + 1)
Loop

Exit_Handler:
    GetPath = strTemp
    Exit Function

Err_Handler:
    strTemp = ""
    Resume Exit_Handler

End Function

Function ReplaceChars_TSB(strIn As String, strFind As String, strReplace As String) As String
  ' Comments  : Replaces characters in a string
  ' Parameters: strIn - string to replace in
  '             strFind - character to find
  '             strReplace - character to replace with
  ' Returns   : modified string
  '
  Dim intCounter As Integer
  Dim strTmp As String
  Dim chrTmp As String * 1

  For intCounter = 1 To Len(strIn)
    chrTmp = Mid$(strIn, intCounter)
    If chrTmp <> strFind Then
      strTmp = strTmp & chrTmp
    Else
      strTmp = strTmp & strReplace
    End If
  Next intCounter

  ReplaceChars_TSB = strTmp

End Function

Public Function TitleCaseNameSplit(strIn As String) As String

Dim strOut As String
Dim intI As Integer

For intI = 1 To Len(strIn)
    If IsCap(Mid(strIn, intI, 1)) Then
        Select Case intI
            Case 2 To (Len(strIn) - 1)  'middle letters
                'if the previous letter was a capital letter, don't put a space before this one
                'unless the next letter is lowercase
                If IsCap(Mid(strIn, intI - 1, 1)) Then
                    If IsCap(Mid(strIn, intI + 1, 1)) Then
                        strOut = strOut & Mid(strIn, intI, 1)
                    Else
                        strOut = strOut & " " & Mid(strIn, intI, 1)
                    End If
                Else
                    'if the previous letter was lowercase, put a space
                    strOut = strOut & " " & Mid(strIn, intI, 1)
                End If
            Case 1  'first letter
                strOut = UCase(Left(strIn, 1))
            Case Len(strIn) 'last letter
                'if the previous letter was a capital, don't put a space
                If IsCap(Mid(strIn, intI - 1, 1)) Then
                    strOut = strOut & Mid(strIn, intI, 1)
                Else
                    strOut = strOut & " " & Mid(strIn, intI, 1)
                End If
        End Select
    Else
        strOut = strOut & Mid(strIn, intI, 1)
    End If
Next

TitleCaseNameSplit = Capitalizer(Trim(strOut))

End Function

Function ReplaceString_TSB(strTextIn As String, strFind As String, strReplace As String, fCaseSensitive As Boolean) As String
  ' Comments   : replaces a substring in a string with another
  ' Parameters : strTextIn - string to work on
  '              strFind - string to find
  '              strReplace - string to replace with
  '              fCaseSensitive - True for case sensitive search, False for case-insensitive search
  ' Returns    : modified string
  '
  Dim strTmp As String
  Dim intPos As Integer
  Dim intCaseSensitive As Integer

  intCaseSensitive = IIf(fCaseSensitive, 0, 1)

  strTmp = strTextIn
  intPos = InStr(1, strTmp, strFind, intCaseSensitive)
  
  Do While intPos > 0
    strTmp = Left$(strTmp, intPos - 1) & strReplace & Mid$(strTmp, intPos + Len(strFind))
    intPos = InStr(intPos + Len(strReplace), strTmp, strFind, intCaseSensitive)
  Loop

  ReplaceString_TSB = strTmp
  
End Function

Public Function ReplaceListItem(strList As String, strFind As String, strReplace As String, strDelimiter As String, Optional booCaseSensitive As Boolean = False, Optional booTrim As Boolean = False) As String
Dim strItem As String
Dim intCompare As Integer
Dim strResult As String
Dim strChar As String
Dim booSemi As Boolean
Dim intI As Integer
Dim strNewList As String

intCompare = 1
If booCaseSensitive = True Then intCompare = 0
If booTrim Then strFind = Trim(strFind)

'Loop through items in the list
Do Until InStr(strList, strDelimiter) = 0
    'Get each item in the list
    If booTrim Then
        strItem = Trim(Left(strList, InStr(strList, strDelimiter) - 1))
    Else
        strItem = Left(strList, InStr(strList, strDelimiter) - 1)
    End If
        
    strList = Mid(strList, InStr(strList, strDelimiter) + 1)

    'Compare the item to the string we wish to replace
    If StrComp(strItem, strFind, intCompare) = 0 Then
        'If they're the same, then replace the item
        strResult = strResult & strReplace & strDelimiter
    Else
        strResult = strResult & strItem & strDelimiter
    End If
Loop

    'Do the last item in the list
    If StrComp(strList, strFind, intCompare) = 0 Then
        'If they're the same, then replace the item
        strResult = strResult & strReplace
    Else
        strResult = strResult & strList
    End If

'Clean up the semicolons
    'First eliminate any leading semicolons
    Do Until Left(strResult, 1) <> strDelimiter
        strResult = Mid(strResult, 2)
    Loop
    'Next eliminate any trailing semicolons
    Do Until Right(strResult, 1) <> strDelimiter
        strResult = Left(strResult, Len(strResult) - 1)
    Loop
    'Finally, eliminate grouped semicolons in the list
    For intI = 1 To Len(strResult)
        strChar = Mid(strResult, intI, 1)
        If strChar = strDelimiter Then
            If booSemi = True Then
            Else
                strNewList = strNewList & strChar
            End If
            booSemi = True
        Else
            strNewList = strNewList & strChar
            booSemi = False
        End If
    Next intI

ReplaceListItem = strNewList

End Function

Public Sub ControlHandler(strFormName As String, strTag As String, strOperation As String)
Dim frm As Form
Dim ctl As Control

On Error GoTo Err_ControlHandler

Set frm = Forms(strFormName)

For Each ctl In frm.Controls
    If ctl.Tag = strTag Then
        Select Case strOperation
            Case "Hide"
                ctl.Visible = False
            Case "Unhide"
                ctl.Visible = True
            Case "Disable"
                ctl.Enabled = False
            Case "Enable"
                ctl.Enabled = True
            Case "Lock"
                ctl.Locked = True
            Case "Unlock"
                ctl.Locked = False
        End Select
    End If
Next

Exit_ControlHandler:
    On Error Resume Next
    Set frm = Nothing
    Exit Sub

Err_ControlHandler:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_ControlHandler

End Sub

Public Function UnderscoreNameSplit(strNameIn As String) As String
UnderscoreNameSplit = Capitalizer(ReplaceChars_TSB(strNameIn, "_", " "))
End Function

Public Function Capitalizer(strIn As String) As String
Dim strWorking As String
Dim strWord As String
Dim booLastWord As Boolean

strIn = Trim(strIn)

Do
    If InStr(strIn, " ") = 0 Then
        booLastWord = True
        strWord = Trim(strIn)
    Else
        strWord = Left(strIn, InStr(strIn, " ") - 1)
        strIn = Mid(strIn, InStr(strIn, " ") + 1)
    End If
        Select Case strWord
            Case "id", "tsn", "nps"
                strWord = UCase(strWord)
            Case Else
                strWord = UCase(Left(strWord, 1)) & Mid(strWord, 2)
        End Select
    strWorking = strWorking & " " & strWord
Loop Until booLastWord

Capitalizer = Trim(strWorking)

End Function

Public Function IsCap(strChar As String) As Boolean
Select Case Asc(strChar)
    Case 65 To 90
        IsCap = True
    Case Else
        IsCap = False
End Select
End Function

Public Function NumSpaces(varIn As Variant) As Integer
Dim intSpaceCount As Integer
Dim strWorking As String

If Not IsNothing(varIn) Then
    strWorking = varIn
    
    Do Until (InStr(strWorking, " ")) = 0
        'we have at least one space
        intSpaceCount = intSpaceCount + 1
        strWorking = Mid(strWorking, InStr(strWorking, " ") + 1)
    Loop
End If

NumSpaces = intSpaceCount
End Function

Public Function GetFileName(ByVal strFilePath As String) As String
Dim strTemp As String

Do While (InStr(strFilePath, "\") > 0)
    strTemp = strTemp & Left(strFilePath, InStr(strFilePath, "\"))
    strFilePath = Mid(strFilePath, InStr(strFilePath, "\") + 1)
Loop

GetFileName = strFilePath
End Function

Public Function FileIsReadOnly(strFileName As String) As Boolean
On Error GoTo Err_FileIsReadOnly

FileIsReadOnly = ((GetAttr(strFileName) And vbReadOnly) <> 0)

Exit_FileIsReadOnly:
    Exit Function

Err_FileIsReadOnly:
    Select Case Err.Number
        Case 76  'file not found
            MsgBox "Unable to locate file " & strFileName & "."
        Case Else
            MsgBox Err.Number & " - " & Err.Description
            Resume Exit_FileIsReadOnly
    End Select

End Function

Public Function ListCompareRemove(strListMain As String, ByVal strListToKeep As String, strDelimiter As String) As String
Dim strItem As String
Dim intI As Integer
Dim strNewList As String

If Not Right(strListToKeep, 1) = strDelimiter Then
    strListToKeep = strListToKeep & strDelimiter
End If

If Not Left(strListToKeep, 1) = strDelimiter Then
    strListToKeep = strDelimiter & strListToKeep
End If

If Not Right(strListMain, 1) = strDelimiter Then
    strListMain = strListMain & strDelimiter
End If

Do Until InStr(strListMain, strDelimiter) = 0
    strItem = strDelimiter & Trim(Left(strListMain, InStr(strListMain, strDelimiter)))
    strListMain = Mid(strListMain, InStr(strListMain, strDelimiter) + 1)

    If InStr(strListToKeep, strItem) > 0 Then
        strNewList = strNewList & Mid(strItem, 2)
    End If
Loop

'Clean up the delimiters
strNewList = DelimiterCleanup(strNewList, strDelimiter)

ListCompareRemove = strNewList

End Function

Public Function DelimiterCleanup(strList As String, strDelimiter As String) As String
Dim strNewList As String
Dim strChar As String
Dim booSemi As Boolean
Dim intI As Integer

    'First eliminate any leading delimiters
    Do Until Left(strList, 1) <> strDelimiter
        strList = Mid(strList, 2)
    Loop
    'Next eliminate any trailing delimiters
    Do Until Right(strList, 1) <> strDelimiter
        strList = Left(strList, Len(strList) - 1)
    Loop
    'Finally, eliminate grouped delimiters in the list
    For intI = 1 To Len(strList)
        strChar = Mid(strList, intI, 1)
        If strChar = strDelimiter Then
            If booSemi = True Then
            Else
                strNewList = strNewList & strChar
            End If
            booSemi = True
        Else
            strNewList = strNewList & strChar
            booSemi = False
        End If
    Next intI

DelimiterCleanup = strNewList
End Function

Public Function ListCompare(strListMain As String, ByVal strListToRemove As String, strDelimiter As String) As String
'Compares two semicolon-delimited lists and eliminates items from strListMain that are in strListToRemove
Dim strItem As String
Dim intI As Integer
Dim strNewList As String

Do Until InStr(strListToRemove, strDelimiter) = 0
    strItem = Trim(Left(strListToRemove, InStr(strListToRemove, strDelimiter) - 1))
    strListToRemove = Mid(strListToRemove, InStr(strListToRemove, strDelimiter) + 1)

    'Remove the item from inside the body of the Main List
    strListMain = ReplaceListItem(strListMain, strItem, "", strDelimiter, False, True)
Loop

    'Do the last item in the list
    strListMain = ReplaceListItem(strListMain, strListToRemove, "", strDelimiter, False, True)

'Clean up the semicolons
strNewList = DelimiterCleanup(strListMain, strDelimiter)

ListCompare = strNewList
End Function

Public Function UnrecognizedDatabaseFormat(strFileName As String) As Boolean
Dim db As Database

On Error GoTo Err_UnrecognizedDatabaseFormat

Set db = OpenDatabase(strFileName)

UnrecognizedDatabaseFormat = False

Exit_UnrecognizedDatabaseFormat:
    On Error Resume Next
    Set db = Nothing
    Exit Function

Err_UnrecognizedDatabaseFormat:
    Select Case Err.Number
        Case 3343 'unrecognized database format
            UnrecognizedDatabaseFormat = True
        Case Else
            UnrecognizedDatabaseFormat = True
    End Select
    Resume Exit_UnrecognizedDatabaseFormat

End Function

Public Function TableNamesAndRecordCounts() As String
Dim db As Database
Dim tdf As TableDef

Set db = CurrentDb

For Each tdf In db.TableDefs
    If Not Left(tdf.Name, 4) = "MSys" Then
        Debug.Print tdf.Name & ": " & tdf.RecordCount
    End If
Next

Set tdf = Nothing
Set db = Nothing
End Function

Public Function FiscalYear(datDate As Date) As Integer
Dim intYear As Integer

intYear = Year(datDate)
If Month(datDate) >= 10 Then
    intYear = intYear + 1
End If

FiscalYear = intYear
End Function

Public Function CharacterCount(strInput As String, strChar As String) As Integer
'Returns the number of times a specified character occurs in a string
Dim i As Integer
Dim intCount As Integer

For i = 1 To Len(strInput)
    If Mid(strInput, i, 1) = strChar Then
        intCount = intCount + 1
    End If
Next
CharacterCount = intCount
End Function

Function InStrAtPos(ByVal strString As String, ByVal strSearchChar As String, iOccurrence As Integer) As Integer
'Returns the location of the specified occurrence of a character. If specified occurrence = 0, the location of the last occurrence is returned
Dim iNext As Integer
If iOccurrence = 0 Then iOccurrence = CharacterCount(strString, strSearchChar)
  Do While iNext < iOccurrence
    InStrAtPos = InStr(InStrAtPos + 1, strString, strSearchChar)
    If InStrAtPos > 0 Then iNext = iNext + 1
    If iNext = iOccurrence Or InStrAtPos = 0 Then Exit Do
  Loop
End Function

Public Function InStrReverse(strString As String, strSearchString As String) As Integer
'*******************************************
'Name:      InStrReverse (Function)
'Purpose:   Return location of last instance of a character or phrase.
'Inputs:    Call InStrReverse("the quick brown fox jumped the lazy dog","the")
'Output:    28 - Location of last occurence of "the"
'Modified by M Lehman based on code by Imboden on Google Access Forum
'*******************************************
Dim i As Integer, n As Integer, tlen As Integer

n = 0
tlen = Len(strSearchString)
For i = Len(RTrim(strString)) To 1 Step -1
  If Mid(strString, i, tlen) = strSearchString Then
      n = i
      Exit For
  End If
Next i
InStrReverse = n
End Function


Public Sub PrintFields(strObjectName)
Dim db As Database
Dim tdf As TableDef
Dim qdf As QueryDef
Dim fld As Field
Dim strOutput As String

On Error Resume Next

Set db = CurrentDb
Set tdf = db.TableDefs(strObjectName)

If Err.Number = 3265 Then
    Err.Clear
    On Error GoTo Err_PrintFields
    Set qdf = db.QueryDefs(strObjectName)
    
    For Each fld In qdf.Fields
        Debug.Print fld.Name
    Next
Else
    For Each fld In tdf.Fields
        Debug.Print fld.Name
    Next
End If

Exit_PrintFields:
    Set tdf = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub

Err_PrintFields:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_PrintFields

End Sub

Function GetLastWord_TSB(strIn As String, strRest As String, chrDelimit As String) As String
  ' Comments   : returns the last word in delimited string strIn, puts strRest in strRest
  ' Parameters : strIn - string to search
  '              chrDelimit - character used as the delimiter
  ' Set        : strRest - set to the rest of the string
  ' Returns    : last word of string
  '
  Dim strTmp As String
  Dim intI As Integer
  Dim intP As Integer

  strTmp = Trim$(strIn)
  intP = 1

  For intI = Len(strTmp) To 1 Step -1
    If Mid$(strTmp, intI, 1) = chrDelimit Then
      intP = intI + 1
      Exit For
    End If
  Next intI

  If intP = 1 Then
    GetLastWord_TSB = strTmp
    strRest = ""
  Else
    GetLastWord_TSB = Mid$(strTmp, intP)
    strRest = Trim$(Left$(strTmp, intP - 1))
  End If

End Function

Function CountOccurrences_TSB(strIn As String, strFind As String) As Integer
  ' Comments  : returns the number of times a string appears in a string
  ' Parameters: strIn - string to search in
  '             strFind - string to search for
  ' Returns   : number of occurrences
  '
  Dim intPos As Integer
  Dim intWordCount As Integer

  intPos = InStr(strIn, strFind)
  
  If intPos > 0 Then
    intWordCount = 1
  End If
  
  Do While intPos > 0
    
    intPos = InStr(intPos + 1, strIn, strFind)
    
    If intPos > 0 Then
      intWordCount = intWordCount + 1
    End If

  Loop

  CountOccurrences_TSB = intWordCount

End Function

Public Function ListItemGet(strList As String, bytItemCount As Byte, strDelimiter As String) As String
Dim bytCount As Byte
Dim strOutput As String
Dim strRemainder As String
Dim intDelimPos As Integer

strRemainder = strList

Do
    bytCount = bytCount + 1
    intDelimPos = InStr(strRemainder, strDelimiter)
    If intDelimPos > 0 Then
        strOutput = strOutput & strDelimiter & Left(strRemainder, intDelimPos - 1)
        strRemainder = Mid(strRemainder, intDelimPos + 1)
    Else
        If Len(strRemainder) > 0 Then
            strOutput = strOutput & strDelimiter & strRemainder
        End If
    End If
Loop Until bytCount = bytItemCount Or intDelimPos = 0

If Len(strOutput) > 0 Then
    ListItemGet = Mid(strOutput, Len(strDelimiter) + 1)
Else
    ListItemGet = strOutput
End If

End Function

Public Function TableNamesAndIndexes(Optional varFileType As Variant) As String
Dim db As Database
Dim tdf As TableDef
Dim idx As Index
Dim fld As Field
Dim rst As Recordset

Set db = CurrentDb

If IsMissing(varFileType) Then
    For Each tdf In db.TableDefs
        If Not Left(tdf.Name, 4) = "MSys" Then
            Debug.Print "Table: " & tdf.Name
            For Each idx In tdf.Indexes
                Debug.Print vbTab & "Index: " & idx.Name
                For Each fld In idx.Fields
                    Debug.Print vbTab & vbTab & "Field: " & fld.Name
                Next fld
            Next idx
        End If
    Next tdf
Else
    Set rst = db.OpenRecordset("SELECT LinkTableName FROM tblLinkedTables WHERE LinkCategory='" & varFileType & "' ORDER BY LinkTableName;", dbOpenForwardOnly)
    
    Do Until rst.EOF
        Set tdf = db.TableDefs(rst!LinkTableName)
        Debug.Print "Table: " & tdf.Name
        For Each idx In tdf.Indexes
            Debug.Print vbTab & "Index: " & idx.Name
            For Each fld In idx.Fields
                Debug.Print vbTab & vbTab & "Field: " & fld.Name
            Next fld
        Next idx
        rst.MoveNext
    Loop
End If

On Error Resume Next

rst.Close
Set rst = Nothing
Set fld = Nothing
Set idx = Nothing
Set tdf = Nothing
Set db = Nothing
End Function

Public Function CorrectText(strInputText As String, Optional strDelimiter As String = "'") As String
Dim strTemp As String

strTemp = strDelimiter
strTemp = strTemp & ReplaceString_TSB(strInputText, strDelimiter, strDelimiter & strDelimiter, False)
strTemp = strTemp & strDelimiter
CorrectText = strTemp
End Function

Public Function CorrectLikeText(strInputText As String, Optional booFrontWildcard As Boolean = False, Optional booEndWildcard As Boolean = False, Optional strDelimiter As String = "'") As String
Dim strTemp As String

strTemp = strDelimiter
If booFrontWildcard Then
    strTemp = strTemp & "*"
End If
strTemp = strTemp & ReplaceString_TSB(strInputText, strDelimiter, strDelimiter & strDelimiter, False)
If booEndWildcard Then
    strTemp = strTemp & "*"
End If
strTemp = strTemp & strDelimiter
CorrectLikeText = strTemp
End Function

Public Sub ExtraSpaceRemoval(strTableName As String, strFieldName As String)
Dim strSQL As String

'eliminate leading and trailing spaces
strSQL = "UPDATE " & strTableName & " SET " & strFieldName & " = Trim([" & strFieldName & "]) WHERE " & strFieldName & " LIKE ' *' OR " & strFieldName & " LIKE '* ';"
DoCmd.SetWarnings False
DoCmd.RunSQL strSQL
DoCmd.SetWarnings True

'eliminate internal spaces except for single spaces
Do Until DCount(strFieldName, strTableName, strFieldName & " LIKE '*  *'") = 0
    strSQL = "UPDATE " & strTableName & " SET " & strFieldName & " = ReplaceString_TSB([" & strFieldName & "],'  ',' ', False) WHERE " & strFieldName & " LIKE '*  *';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
Loop

End Sub

Public Function LastPhrase(strIn As String, strPhrase As String) As Integer
'Author: Simon Kingston
'Purpose: Locates the position of a substring within another string, starting from the end of the word.
'         - basically an Instr() that starts from the end of the word instead of the beginning
'Parameters:    strIn - the string to be searched
'               strPhrase - the string to find
'Output: the position of strPhrase within strIn
'Example: LastPhrase("Find the final pint","in") = 17
Dim intPos As Integer
Dim intPhraseLength As Integer

intPos = Len(strIn)
intPhraseLength = Len(strPhrase)

Do Until Mid(strIn, intPos, intPhraseLength) = strPhrase
    intPos = intPos - 1
    If intPos = 0 Then
        Exit Do
    End If
Loop

LastPhrase = intPos
End Function

Public Function HasWeirdChar(strIn As String) As Boolean
'Author: Simon Kingston
'Purpose: Find scientific names with unusual characters
'Description: Returns false if string strIn contains only letters, periods, spaces,
'             and/or hyphens.  Returns true otherwise.  Note: legitimate subgenus
'             names can contain parentheses.
'Examples: HasWeirdChar("Canis lupus") = False
'          HasWeirdChar("Canis & lupus") = True
Dim i As Integer
Dim intLength As Integer
Dim strChar As String

intLength = Len(strIn)

Do Until i = intLength
    i = i + 1
    strChar = Mid(strIn, i, 1)
    Select Case Asc(strChar)
        Case 65 To 90, 97 To 122, 46, 32, 45
            'upper and lower case letters, period, space, and hyphen
        Case Else
            HasWeirdChar = True
            Exit Do
    End Select
Loop
End Function

Public Function EliminateDataBetweenChars(strNameIn As String, strCharStart As String, strCharEnd As String) As String
'Author: Simon Kingston
'Description: Eliminates strings within strNameIn that start with strCharStart and end with strCharEnd.
'             Useful for stripping HTML tags and authorities that are in parentheses.
'Example: EliminateDataBetweenChars("<i>Canis lupus</i><br>","<",">") = "Canis lupus"
Dim strWorking As String
Dim intParenStart As Integer
Dim intParenEnd As Integer
Dim intCurrentPos As Integer
Dim strChar As String

strWorking = strNameIn

Do Until InStr(strWorking, strCharStart) = 0 Or InStr(strWorking, strCharEnd) = 0
    intParenStart = InStr(strWorking, strCharStart)
    intCurrentPos = intParenStart

    Do
        intParenEnd = 0
        intCurrentPos = intCurrentPos + 1
        strChar = Mid(strWorking, intCurrentPos, 1)
        Select Case strChar
            Case strCharStart
                intParenStart = intCurrentPos
            Case strCharEnd
                intParenEnd = intCurrentPos
            Case Else
        End Select
    Loop Until intParenEnd > 0
    
    strWorking = Left(strWorking, intParenStart - 1) & Mid(strWorking, intParenEnd + 1)
Loop

EliminateDataBetweenChars = strWorking
End Function

Function IsObjectInDB_TSB(strDatabase As String, strType As String, strName As String) As Boolean
  ' Comments  : determines if the named object is in the named database
  ' Parameters: strDatabase - path and name of the database to look in or "" (blank string) for the current database
  '             strType - type of object: "table", "query", "form", "report", "macro" or "module"
  '             strName - name of object
  ' Returns   : True-object exists, False-object does not exist
  '
  Dim dbsTemp As Database
  Dim strCon As String
  Dim varDummy As Variant

  On Error GoTo Proc_Err
  
  If strDatabase = "" Then
    Set dbsTemp = CurrentDb()
  Else
    Set dbsTemp = DBEngine.Workspaces(0).OpenDatabase(strDatabase)
  End If
  
  Select Case strType
    Case "Table", "Tables": strCon = "Tables"
    Case "Query", "Queries": strCon = "Tables"
    Case "Form", "Forms": strCon = "Forms"
    Case "Report", "Reports": strCon = "Reports"
    Case "Macro", "Macros", "Scripts": strCon = "Scripts"
    Case "Module", "Modules": strCon = "Modules"
  End Select
  
  varDummy = dbsTemp.Containers(strCon).Documents(strName).Name
  IsObjectInDB_TSB = True

  dbsTemp.Close

Proc_Exit:
  Exit Function

Proc_Err:
  IsObjectInDB_TSB = False
  Resume Proc_Exit
  
End Function

Function DeleteObject_TSB(intType As Integer, strName As String) As Boolean
  ' Comments  : Deletes an object (table, query, form, report, macro, or module)
  ' Parameters: intType - an Access constant for the type of object to delete
  '                       (acTable, acQuery, acForm, acReport, acMacro, acModule)
  '             strName - name of the object to delete
  ' Returns   : True if the object was deleted or does not exist,
  '             False if the object could not be deleted
  Dim fOK As Boolean
  
  Const errNotFound = 3011

  fOK = True

  On Error Resume Next
  DoCmd.DeleteObject intType, strName

  If Err.Number <> 0 Then
    ' Error 3011 occurs if the object could not be found (already deleted).
    ' If this error occurs, consider the deletion successful.
    ' For more information, check the Error$ value.
    fOK = (Err.Number = errNotFound)
  End If

  DeleteObject_TSB = fOK

End Function

Public Function fPathParsing(fullPath As String, PathFormat As String) As String
' Edited by: Alan Williams 11/5/2002
' Parses fullpath into Dir, filename, and extension.
' Example calls:
            '? fPathParsing("C:\work\Seasonals.xls", "D")
            'C:\work\
            '? fPathParsing("C:\work\Seasonals.xls", "DN")
            'C:\work\Seasonals
            '? fPathParsing("C:\work\Seasonals.xls", "NE")
            'Seasonals.xls
            '? fPathParsing("C:\work\Seasonals.xls", "N")
            'Seasonals
            '? fPathParsing("C:\work\Seasonals.xls", "E")
            '.xls
Dim i As Integer, f As String, Found As Integer
Dim DirName As String, FName As String, Ext As String
  
  fullPath = Trim$(fullPath)
'
' Get directory name
'
  f = ""
  Found = False
  For i = Len(fullPath) To 1 Step -1
    If Mid$(fullPath, i, 1) = "\" Then
      f = Mid$(fullPath, i + 1)
      DirName = Left$(fullPath, i)
      Found = True
      Exit For
    End If
  Next i
  If Not Found Then
    f = fullPath
  End If
'
' Get File name and extension
'
  If f = "." Or f = ".." Then
    FName = f
  Else
    i = InStr(f, ".")
    If i > 0 Then
      FName = Left$(f, i - 1)
      Ext = Mid$(f, i)
    Else
      FName = f
    End If
  End If
Select Case PathFormat
    Case "D"
        fPathParsing = DirName
    Case "N"
        fPathParsing = FName
    Case "E"
        fPathParsing = Ext
    Case "DN"
        fPathParsing = DirName & FName
    Case "NE"
        fPathParsing = FName & Ext
    Case "DNE"
        fPathParsing = DirName & FName & Ext
    Case Else
        fPathParsing = fullPath

End Select

End Function

Public Sub Create_File(strFullFileName As String, text As String)
  Dim fso, txtfile
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set txtfile = fso.CreateTextFile(strFullFileName, True)
  txtfile.Write (text)
  txtfile.Close
  
  Set txtfile = Nothing
  Set fso = Nothing
End Sub

Public Function XML_Tag(strTag As String, strValue As String) As String
'Authored: ?, Simon Kingston
'Purpose: Tags a value (strValue) with opening and closing XML tags specified (strTag)
'Parameters: strTag = XML tag to use for opening and closing tag
'            strValue = string to put between XML tags
'Returns: Text string of value (strValue) between XML tags (strTag)
'Procedure calls: None
'Edits:

XML_Tag = "<" & strTag & ">" & strValue & "</" & strTag & ">"
End Function

Public Function XML_Read(strTag As String, strIn As String) As String
'Authored: ?, Simon Kingston
'Purpose: Searches a string (strIn) for an XML tag (strTag) and retrieves the data found inside the first opening and closing tag found
'Parameters: strTag = XML tag to search for
'            strIn = string to search for XML data
'Returns: Text string of data found between XML tags, or if tags not foung, an empty string
'Procedure calls: None
'Edits:

Dim strOut As String
Dim strLeadTag As String
Dim strEndTag As String
Dim lngLeadTagPosition As Long
Dim lngEndTagPosition As Long

strLeadTag = "<" & strTag & ">"
strEndTag = "</" & strTag & ">"
lngLeadTagPosition = InStr(strIn, strLeadTag)
lngEndTagPosition = InStr(strIn, strEndTag)

If lngLeadTagPosition > 0 And lngEndTagPosition > lngLeadTagPosition Then
    strOut = Mid(strIn, lngLeadTagPosition + Len(strLeadTag), lngEndTagPosition - lngLeadTagPosition - Len(strLeadTag))
End If

XML_Read = strOut

End Function

Public Function GetDataType(strTableName As String, strFieldName As String) As Integer
Dim intResult As Integer

On Error Resume Next

intResult = CurrentDb.TableDefs(strTableName)(strFieldName).Type
GetDataType = intResult
End Function

Public Function GetCriteriaString(strCriteriaStart As String, strTableName As String, strFieldName As String, strFormName As String, strControlName As String) As String
Dim strResult As String

On Error Resume Next

strResult = strCriteriaStart

Select Case GetDataType(strTableName, strFieldName)
    Case dbText
        strResult = strResult & CorrectText(Forms(strFormName)(strControlName))
    Case dbGUID
        strResult = strResult & StringFromGUID(Nz(Forms(strFormName)(strControlName), ""))
    Case Else
        strResult = strResult & Forms(strFormName)(strFieldName)
End Select

GetCriteriaString = strResult
End Function

Public Function NothingZ(varTest As Variant, varDefault As Variant) As Variant
'Created: 10/3/2006, Simon Kingston
'Purpose: Analagous to nz function, but instead of just testing for null, it tests varTest to see if it is nothing (see IsNothing function).
'         If it is, return a default, otherwise, return varTest.
Dim varResult As Variant

If IsNothing(varTest) Then
    varResult = varDefault
Else
    varResult = varTest
End If
NothingZ = varResult
End Function

' =================================
' FUNCTION:     fxnReplaceString
' Description:  Replaces a substring in a string with another
' Parameters:   strTextIn - string to work on
'               strFind - string to find
'               strReplace - string to replace with
'               fCaseSensitive - True for case sensitive search, False otherwise
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, date unknown
' Revisions:    John R. Boetsch, May 17, 2006 - error trapping, documentation
' =================================

Function fxnReplaceString(strTextIn As String, strFind As String, _
    strReplace As String, fCaseSensitive As Boolean) As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    Dim intPos As Integer
    Dim intCaseSensitive As Integer

    ' Convert the case-sensitive boolean to the comparison constant (1=binary, 2=textual)
    intCaseSensitive = fCaseSensitive + 1

    strTemp = strTextIn
    intPos = InStr(1, strTemp, strFind, intCaseSensitive)

    Do While intPos > 0
        strTemp = Left$(strTemp, intPos - 1) & strReplace & Mid$(strTemp, intPos + Len(strFind))
        intPos = InStr(intPos + Len(strReplace), strTemp, strFind, intCaseSensitive)
    Loop

    fxnReplaceString = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnReplaceString)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnChangeDelimiter
' Description:  Replaces delimiters in an input string; default is to change double-quotes
'               to single quotes
' Parameters:   strInputText - string to work on
'               strCurrDelimiter - current delimiter in the string (default: double-quote)
'               strNewDelimiter - desired replacement delimiter (default: single-quote)
' Returns:      modified string
' Throws:       none
' References:   fxnReplaceString
' Source/date:  John R. Boetsch, May 17, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnChangeDelimiter(strInputText As String, _
    Optional strCurrDelimiter As String = """", _
    Optional strNewDelimiter As String = "'") As String

    On Error GoTo Err_Handler

    Dim strTemp As String
    
    ' Call the replace string function, specifying the delimiter and no case-sensitive search
    strTemp = fxnReplaceString(strInputText, strCurrDelimiter, strNewDelimiter, False)
    fxnChangeDelimiter = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnChangeDelimiter)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnTrimSpaces
' Description:  Removes leading and trailing space characters from a string
' Parameters:   strInputText - string to work on
' Returns:      modified string
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 25, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnTrimSpaces(strInputText As String) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    strTemp = strInputText

    ' First trim leading spaces
    Do While Left(strTemp, 1) = " "
        strTemp = Right(strTemp, Len(strTemp) - 1)
    Loop
    ' Then trim trailing spaces
    Do While Right(strTemp, 1) = " "
        strTemp = Left(strTemp, Len(strTemp) - 1)
    Loop

    fxnTrimSpaces = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnTrimSpaces)"
            Resume Exit_Procedure
    End Select

End Function

Public Sub HyperlinkFormat(booHyperlink As Boolean, ctl As Control)
Const cBlue As Long = 16711680
Const cBlack As Long = 0

ctl.FontUnderline = booHyperlink

If booHyperlink Then
    ctl.ForeColor = cBlue
Else
    ctl.ForeColor = cBlack
End If
    
End Sub

Public Function fxnGetDbFileName(strTableName As String) As String
Dim strFileName As String

On Error GoTo Error_Handler

strFileName = CurrentDb.TableDefs(strTableName).Connect
If Len(strFileName) > 0 Then
    'linked table
    strFileName = ReplaceString_TSB(strFileName, ";DATABASE=", "", False)
Else
    'local table
    strFileName = CurrentDb.Name
End If

Exit_Handler:
    fxnGetDbFileName = strFileName
    Exit Function

Error_Handler:
    strFileName = ""
    Resume Exit_Handler

End Function


Function fxnGetSummaryInfo(strPropName As String, strDatabasePath As String) As String
    On Error GoTo Err_Handler

' =================================
' FUNCTION:     fxnGetSummaryInfo
' Description:  Provides information from a database property as specified
' Note:         Use a call like: GetSummaryInfo("strTitle", Currentdb.name)
' Parameters:   strPropName (property), strDatabasePath (db path)
' Returns:      value of the input db property; if no value has been set, returns
'               "None"; "Error" returned if unknown error
' Throws:       none
' References:   none
' Source/date:  MS Access help files, courtesy of Alan Williams (SHEN)
' Revisions:    John R. Boetsch, May 15, 2006 - error trap, simplified, updated documentation
'               JRB, 6/22/2006 - modified error trap for error 3051 to eliminate message
'               JRB, 12/29/2009 - added error trap for 3024, fixed messages to "Error"
' =================================

    Dim db As DAO.Database, cnt As Container
    Dim doc As Document, prp As Property
    Dim wsp As Workspace

    Set wsp = DBEngine.Workspaces(0)
    Set db = wsp.OpenDatabase(strDatabasePath)
    Set cnt = db.Containers!Databases
    Set doc = cnt.Documents!SummaryInfo

    doc.Properties.Refresh
    fxnGetSummaryInfo = doc.Properties(strPropName)

Exit_Procedure:
    On Error Resume Next
    Set wsp = Nothing
    Set db = Nothing
    Set cnt = Nothing
    Set doc = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case 3024   ' ADP file
            fxnGetSummaryInfo = "Error"
            Resume Exit_Procedure
        Case 3270
            fxnGetSummaryInfo = "Error"
            Resume Exit_Procedure
        Case 3265
            fxnGetSummaryInfo = "Error"
            Resume Exit_Procedure
        Case 3051   ' Read-only or no rights
            fxnGetSummaryInfo = "Error"
            Resume Exit_Procedure
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Startup message error (fxnGetSummaryInfo)"
            fxnGetSummaryInfo = "Error"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnTableExists
' Description:  Returns whether the specified table exists in the current database collection
' Parameters:   strTableName - string for the name of the table to check
' Returns:      True if the specified table exists in the master systems table, or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/29/2009
' Revisions:    JRB, 12/29/2009 - updated to determine project type first
' =================================

Public Function fxnTableExists(ByVal strTableName As String) As Boolean
    On Error GoTo Err_Handler

    Select Case CurrentProject.ProjectType
      Case acADP
        fxnTableExists = DCount("*", "sysobjects", "(([xtype] In ('U')) AND ([name]='" & _
            strTableName & "'))")
      Case acMDB
        fxnTableExists = DCount("*", "MSysObjects", "(([Type] In (1,4,6)) AND ([Name]=""" & _
            strTableName & """))")
      Case Else
        fxnTableExists = False
    End Select

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnTableExists)"
    End Select
    Resume Exit_Procedure

End Function

Public Function fxnQueryExists(QueryName As String) As Boolean

'=================================================  ============================
' hlfUtils.TableExists
'-----------------------------------------------------------------------------
' Copyright by Heather L. Floyd - Floyd Innovations - www.floydinnovations.com
' Created 08-01-2005
'-----------------------------------------------------------------------------
' Purpose:  Checks to see whether the named query exists in the database
'-----------------------------------------------------------------------------
' Parameters:
' ARGUEMENT             :   DESCRIPTION
'-----------------------------------------------------------------------------
' TableName (String)    :   Name of query to check for
'-----------------------------------------------------------------------------
' Returns:  True, if table found in current db, False if not found.
'=================================================  ============================

Dim strQueryNameCheck
On Error GoTo ErrorCode

'try to assign queryname value
strQueryNameCheck = CurrentDb.QueryDefs(QueryName)

'If no error and we get to this line, true
fxnQueryExists = True

ExitCode:
    On Error Resume Next
    Exit Function

ErrorCode:
    Select Case Err.Number
        Case 3265  'Item not found in this collection
            fxnQueryExists = False
            Resume ExitCode
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "hlfUtils.TableExists"
            'Debug.Print "Error " & Err.number & ": " & Err.Description & "hlfUtils.TableExists"
            Resume ExitCode
    End Select

End Function