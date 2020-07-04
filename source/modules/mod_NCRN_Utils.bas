Option Compare Database

Private Declare PtrSafe Function apiGetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'From Steve Potts posting on Google Access Forum
Declare PtrSafe Function GetUsername _
   Lib "advapi32.dll" _
   Alias "GetUserNameA" _
   (ByVal lpBuffer As String, _
   nSize As Long) As Long

Public Function WindowsName() As String
   On Error Resume Next
   WindowsName = " "
   Dim strName As String
   strName = Space(8)
   Call GetUsername(strName, 8)
   WindowsName = strName
End Function

Function NetworkUsername() As String
'On Error GoTo Err_Handler
    'Purpose:    Returns the network login name
    Dim lngLen As Long
    Dim lngX As Long
    Dim strUserName As String

    NetworkUsername = "Unknown"

    strUserName = String$(254, 0)
    lngLen = 255&
    lngX = apiGetUserName(strUserName, lngLen)
    If (lngX > 0&) Then
        NetworkUsername = Left$(strUserName, lngLen - 1&)
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    'Call LogError(Err.Number, Err.Description, conMod & ".NetworkUserName", , False)
    Resume Exit_Handler
End Function

'Provided by Allen Browne, allen@allenbrowne.com. Updated June 2006.
'TableInfo() function
'This function displays in the Immediate Window (Ctrl+G) the structure of any table in the current database.
'For Access 2000 or 2002, make sure you have a DAO reference.
'The Description property does not exist for fields that have no description, so a separate function handles that error.

Function TableInfo(strTableName As String)
On Error GoTo TableInfoErr
   ' Purpose:   Display the field names, types, sizes and descriptions for a table.
   ' Argument:  Name of a table in the current database.
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   Dim fld As DAO.field
   
   Set db = CurrentDb()
   Set tdf = db.TableDefs(strTableName)
   Debug.Print "FIELD NAME", "FIELD TYPE", "SIZE", "DESCRIPTION"
   Debug.Print "==========", "==========", "====", "==========="

   For Each fld In tdf.Fields
      Debug.Print fld.Name,
      Debug.Print FieldTypeName(fld),
      Debug.Print fld.Size,
      Debug.Print GetDescrip(fld)
   Next
   Debug.Print "==========", "==========", "====", "==========="

TableInfoExit:
   Set db = Nothing
   Exit Function

TableInfoErr:
   Select Case Err
   Case 3265&  'Table name invalid
      MsgBox strTableName & " table doesn't exist"
   Case Else
      Debug.Print "TableInfo() Error " & Err & ": " & Error
   End Select
   Resume TableInfoExit
End Function


Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function

Function GetQueryDescription(strQryName As String) As String
   On Error Resume Next
   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef
   Set db = CurrentDb()
   Set qdf = db.QueryDefs(strQryName)
   GetQueryDescription = qdf.Properties("Description")
End Function

'----------------------------------------------
' RETIRED - 7/1/2020 - covered in dev_mod_Git
'----------------------------------------------
'Function FieldTypeName(fld As DAO.field) As String
'    'Purpose: Converts the numeric results of DAO Field.Type to text.
'    Dim strReturn As String    'Name to return
'
'    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
'        Case dbBoolean: strReturn = "Yes/No"            ' 1
'        Case dbByte: strReturn = "Byte"                 ' 2
'        Case dbInteger: strReturn = "Integer"           ' 3
'        Case dbLong                                     ' 4
'            If (fld.Attributes And dbAutoIncrField) = 0& Then
'                strReturn = "Long Integer"
'            Else
'                strReturn = "AutoNumber"
'            End If
'        Case dbCurrency: strReturn = "Currency"         ' 5
'        Case dbSingle: strReturn = "Single"             ' 6
'        Case dbDouble: strReturn = "Double"             ' 7
'        Case dbDate: strReturn = "Date/Time"            ' 8
'        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
'        Case dbText                                     '10
'            If (fld.Attributes And dbFixedField) = 0& Then
'                strReturn = "Text"
'            Else
'                strReturn = "Text (fixed width)"        '(no interface)
'            End If
'        Case dbLongBinary: strReturn = "OLE Object"     '11
'        Case dbMemo                                     '12
'            If (fld.Attributes And dbHyperlinkField) = 0& Then
'                strReturn = "Memo"
'            Else
'                strReturn = "Hyperlink"
'            End If
'        Case dbGUID: strReturn = "GUID"                 '15
'
'        'Attached tables only: cannot create these in JET.
'        Case dbBigInt: strReturn = "Big Integer"        '16
'        Case dbVarBinary: strReturn = "VarBinary"       '17
'        Case dbChar: strReturn = "Char"                 '18
'        Case dbNumeric: strReturn = "Numeric"           '19
'        Case dbDecimal: strReturn = "Decimal"           '20
'        Case dbFloat: strReturn = "Float"               '21
'        Case dbTime: strReturn = "Time"                 '22
'        Case dbTimeStamp: strReturn = "Time Stamp"      '23
'
'        'Constants for complex types don't work prior to Access 2007.
'        Case 101&: strReturn = "Attachment"         'dbAttachment
'        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
'        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
'        Case 104&: strReturn = "Complex Long"       'dbComplexLong
'        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
'        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
'        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
'        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
'        Case 109&: strReturn = "Complex Text"       'dbComplexText
'        Case Else: strReturn = "Field type " & fld.Type & " unknown"
'    End Select
'
'    FieldTypeName = strReturn
'End Function