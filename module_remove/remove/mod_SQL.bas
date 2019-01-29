Attribute VB_Name = "mod_SQL"
Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_SQL
' Level:        Framework module
' VERSION:      1.10
' Description:  Database/SQL properties, functions & subroutines
'
' Source/date:  Bonnie Campbell, 7/24/2014
' Revisions:    BLC, 7/24/2014 - 1.00 - initial version
'               BLC, 8/19/2014 - 1.01 - added versioning
'               BLC, 5/26/2015 - 1.02 - added mod_db_Templates subs/functions - GetQuerySQL, GetSQLDbTemplate
'               BLC, 6/30/2015 - 1.03 - combined GetDbQuerySQL with GetQuerySQL, renamed get... to Get... functions
'               BLC, 8/21/2015 - 1.04 - added ConcatRelated notes for Error 3048 using linked tables
'               BLC, 3/16/2016 - 1.05 - added PrepareWhereClause() [Uplands 2016 preseason mods]
' --------------------------------------------------------------------
'               BLC, 4/18/2017          added updated version to Invasives db
' --------------------------------------------------------------------
'               BLC, 4/18/2017 - 1.06 - adjusted for invasives
'               BLC, 4/28/2017 - 1.07 - added SQL_encode(), GetParamsFromSQL() moved from mod_Db
'               BLC, 10/4/2017 - 1.08 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC, 10/5/2017 - 1.09 - moved DbCurrent property to mod_Db
'               BLC, 10/6/2017 - 1.10 - update DbCurrent > CurrDb
' =================================

' ---------------------------------
'   Retrieve SQL
' ---------------------------------

' ---------------------------------
' FUNCTION:     GetSQL
' Description:  Retrieve query SQL string using query name
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:
'   Albert D. Kallal  (Access MVP) Edmonton, Alberta Canada kallal@msn.com - Sept 8, 2010
'   http://social.msdn.microsoft.com/Forums/office/en-US/3a26a941-b75b-49e4-bfe8-10c152f2b6c0/sql-or-querydef-in-vba-code?forum=accessdev
'   Daniel Pineault, CARDA Consultants Inc. - June 10, 2010
'   http://www.devhut.net/2010/06/10/ms-access-vba-edit-a-querys-sql-statement/
' Adapted:      Bonnie Campbell, July, 2014 for NCPN tools
' Revisions:    BLC, 7/23/2014 - initial version
'               BLC, 6/30/2015 - rename get... to Get...
'               BLC, 10/6/2017 - update dbCurrent > CurrDb
' ---------------------------------
Public Function GetSQL(strQuery As String) As String
On Error GoTo Err_Handler:

   GetSQL = CurrDb.QueryDefs(strQuery).SQL
   
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSQL[mod_Point_Intercept])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     GetWhereSQL
' Description:  Prepare a SQL WHERE clause based on the parameters, parameter types, fields, and
'               current WHERE clause (strWhere) passed into the function
' Assumptions:  Assumes parameters passed through params will each have the parameter name, type, and field name
'                   params(x,0) = parameter value
'                   params(x,1) = parameter type
'                   params(x,2) = database field name
'               NOTE: The function does not currently handle dependent parameters which require
'                     the presence of other parameters to be included in the WHERE clause
'                     These have to be accommodated separately.
' Parameters:   Completed SQL WHERE clause (string)
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, August, 2014 for NCPN tools
' Adapted:      Bonnie Campbell, July, 2014 for NCPN tools
' Revisions:    BLC, 8/11/2014 - initial version
'               BLC, 6/30/2015 - rename from get... to Get...
' ---------------------------------
Public Function GetWhereSQL(strWHERE As String, Params As Variant) As String
On Error GoTo Err_Handler:
Dim blnCheck As Boolean
Dim strParam As String
Dim i As Integer

    'default
    blnCheck = False

    For i = 0 To UBound(Params) - 1
    
        'handle empty field values
        If Len(Params(i, 2)) > 0 Then
    
            'handle when param isn't the only parameter (need ' AND ' in SQL WHERE clause)
            If Len(strWHERE) > 0 Then strWHERE = strWHERE & " AND"
    
            'check if parameter is is non-empty (string) or non-zero (integer)
            Select Case Params(i, 1)
                Case "string"
                    If Len(Trim(Params(i, 0))) > 0 Then blnCheck = True
                    strParam = "'" & Params(i, 0) & "'"
                Case "integer"
                    If Params(i, 0) > 0 Then blnCheck = True
                    strParam = Params(i, 0)
            End Select
        
            'prepare SQL
            If Not IsNull(Params(i, 0)) And blnCheck Then
             strWHERE = strWHERE & " " & Params(i, 2) & " = " & strParam
            End If
        
        Else
            Exit For 'done
        End If
    Next
    
   GetWhereSQL = strWHERE
   
Exit_Function:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetWhereSql[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     GetQuerySQL
' Description:  Get SQL for a query
' Assumptions:  -
' Parameters:   strQueryName - Name of query to fetch SQL for (string)
' Returns:      full SQL for the query (string)
' Throws:       none
' References:   -
' Source/date:
' S. Phinney, July 13, 2009
' http://bytes.com/topic/access/answers/871500-getting-sql-string-query
' Adapted:      Bonnie Campbell, February 23, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/23/2015 - initial version
'   BLC, 5/1/2015 - moved from mod_App_Data to mod_SQL
'   ----------------- GetDbQuerySQL revisions -----------
'   BLC, 6/16/2014 - initial version
'   BLC, 5/26/2015 - moved from mod_db_Templates to mod_SQL, added error handling
'   ------------------------------------------------------
'   BLC, 6/30/2015 - combined with GetDbQuerySQL (similar functions)
'   BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Private Function GetQuerySQL(strQueryName As String) As String
Dim qdf As DAO.QueryDef
 
    'fetch query
    Set qdf = CurrDb.QueryDefs(strQueryName)
    
    'return SQL
    GetQuerySQL = qdf.SQL
 
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetQuerySQL[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:     GetSQLTemplate
' Description:  loads SQL templates (queries as SQL string) into memory as a dictionary object
'               with query SQL strings available without querying the db tsys_SQL_templates table
' Parameters:
' Returns:      dictionary object stored in tempVars.Item("SQL")
' Assumptions:  placing
' Throws:       none
' References:   tsys_SQL_templates, Microsoft Scripting Runtime (dictionary object)
' Source/date:  Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - initial version
'               BLC, 5/26/2015 - moved from mod_db_Templates to mod_SQL, added error handling
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Sub GetSQLTemplates(Optional strVersion As String = "")
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, Key As String, Value As String
    
    'handle default
    strSQLWhere = " WHERE Is_Supported > 0"
    
    If Len(strVersion) > 0 Then
        strSQLWhere = " AND LCase(versionID) = LCase(" & strVersion & " )"
    End If
    
    'sql
    strSQL = "SELECT * FROM tsys_Db_Templates" & strSQLWhere
    
    Set db = CurrDb
    Set rst = db.OpenRecordset(strSQL)
    
    'handle no records
    If rst.EOF Then
        MsgBox "Sorry, no templates were found for this database version.", vbExclamation, _
            "Linked Database Templates Not Found"
        DoCmd.CancelEvent
        GoTo Exit_Sub
    End If
    
    'prepare dictionary
    Dim dict As New Scripting.Dictionary
    Dim ary(1 To 4) As String
    Dim i As Integer
    
    'prepare the dictionary key array
    ary(1) = "context"
    ary(2) = "template_Name"
    ary(3) = "SQLstring" 'template
    ary(4) = "var_list"
    
    rst.MoveFirst
    Do Until rst.EOF
        'populate the dictionary
        For i = 1 To UBound(ary)
            Key = ary(i)
            If (ary(i) = "SQLstring") Then
                Value = rst!template
            Else
                Value = rst.Fields(ary(i))
            End If
            If Not dict.Exists(Key) Then
                dict.Add Key, Value
            End If
        Next
        rst.MoveNext
    Loop
    
    TempVars.Add "SQL", dict

    'cleanup
    Set dict = Nothing
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSQLTemplates[mod_SQL])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
'   Alter SQL
' ---------------------------------

' ---------------------------------
' FUNCTION:     SQLencode
' Description:  sanitizes SQL to remove special characters
' Parameters:   strSQL - SQL to sanitize (string)
' Returns:      strSanitized - sanitized SQL (string)
' Assumptions:
' Throws:       none
' References:
'   Susan Harkins, March 2, 2011
'   http://www.techrepublic.com/blog/microsoft-office/5-rules-for-embedding-strings-in-vba-code/
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/6/2016 - initial version
' ---------------------------------
Public Function SQLencode(strSQL)
On Error GoTo Err_Handler
    
    Dim aryReplace(1, 2) As String
    Dim i As Integer
    Dim strNewSQL As String
    
    'default
    strNewSQL = ""
    
    'exit if no description
    If Len(strSQL) = 0 Then GoTo Exit_Handler
    
    '--------------------------
    ' replacement characters
    '--------------------------
    '   "   Chr(34)
    '   '   Chr(39)
    '--------------------------
    aryReplace(0, 0) = """"
    aryReplace(0, 1) = 34
    aryReplace(1, 0) = "'"
    aryReplace(1, 1) = 39
    
    For i = 0 To UBound(aryReplace, 1)
        strNewSQL = Replace(strSQL, aryReplace(i, 0), "Chr(" & aryReplace(i, 1) & ")")
    Next

    SQLencode = strNewSQL
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SQLencode[mod_SQL])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'   SQL Parameters
' ---------------------------------

' ---------------------------------
' FUNCTION:     SetParam
' Description:  Set a parameter value (useful for parameter queries)
' Assumptions:  Companion GetParam() function exists & param is publicly defined
' Parameters:   paramValue - parameter name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 24, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/24/2015 - initial version
'   BLC, 5/1/2015  - moved from mod_App_Data to mod_SQL
' ---------------------------------
Public Function SetParam(paramValue As Variant)

On Error GoTo Err_Handler
Dim param As Variant
    
    param = paramValue
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetParam[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     GetParam
' Description:  Get a parameter value (useful for parameter queries)
' Assumptions:  Companion GetParam() function exists & param is publicly defined
' Parameters:   paramValue - parameter name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 24, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/24/2015 - initial version
'   BLC, 5/1/2015  - moved from mod_App_Data to mod_SQL
' ---------------------------------
Public Function GetParam()

On Error GoTo Err_Handler
Dim param As Variant

    GetParam = param
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParam[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     GetParamsFromSQL
' Description:  extracts parameters from SQL string
' Assumptions:  -
' Parameters:   sql - SQL to retrieve parameters from(string)
' Returns:      params - delimited string of parameters and parameter types (string)
' References:   -
' Source/date:  Bonnie Campbell, September 20 2016
' Revisions:    BLC, 9/20/2016 - initial version
' ---------------------------------
Public Function GetParamsFromSQL(SQL As String) As String
On Error GoTo Err_Handler

    Dim Params As String
    
    'default
    Params = ""
    
    If Len(SQL) > 0 Then
        If InStr(SQL, "PARAMETERS ") Then
            Dim delimPos As Integer
            
            Params = Replace(SQL, "PARAMETERS ", "")
            delimPos = InStr(Params, ";")
            Params = Left(Params, delimPos - 1)
            Params = Replace(Params, ", ", "|")
            Params = Replace(Params, " ", ":")
            
            'convert TEXT(#) values to STRING
            If InStr(Params, "TEXT(") Then
                'remove TEXT( )
                Params = Replace(Params, "TEXT(", "STRING")
                Params = Replace(Params, ")", "")
                'remove numerics
                Params = RemoveChars(Params, False)
            End If
            
        End If
    End If
    
Exit_Handler:
    GetParamsFromSQL = Params
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParamsFromSQL[mod_SQL])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'   SQL Functions
' ---------------------------------

' ---------------------------------
' SUB:          ConcatRelated
' Description:  Used in SQL queries to generate concatenated string of related records
' Assumptions:  used in Access SQL or control
' Parameters:   strField - field to retrieve results from & concatenate (string)
'               strTable - table or query name (string)
'               strWHERE - limiting WHERE clause (string)
'               strOrderBy - sorting ORDER BY clause (string)
'               strSeparator - character to use between concatenated values (string)
' Returns:      SQL (string, variant, or NULL if no matches)
' Notes:        1. Use square brackets around field/table names with spaces or odd characters.
'               2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
'               3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
'               4. Returning more than 255 characters to a recordset triggers this Access bug:
'                  http://allenbrowne.com/bug-16.html
'               -------------------------------------------------------------
'                IMPORTANT:
'               -------------------------------------------------------------
'                 ConcatRelated should NOT be used with linked tables or
'                 queries with linked tables. Doing so results in
'                 Error 3048: "Can't open any more databases"
'                 when the recordset is instantiated
'                       Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
'                 DBEngine(0)(0) is significantly faster than CurrentDb and is not
'                 the problem. The issue is that linked tables increase the number of databases which
'                 must be opened.
'                 If the SQL statement using ConcatRelated has many records (hundreds, thousands)
'                 Access must be crashed using Task Manager (& compacted/repaired on re-opening)
'                 to stop the process.
'                 See:
'                   Keri Hardwic, July 26, 2002
'                   http://computer-programming-forum.com/1-vba/56b6e02cf3f9c2f7.htm
'               -------------------------------------------------------------
' Usage:        SQL string:
'                SELECT CompanyName,  ConcatRelated("OrderDate", "tblOrders", "CompanyID = "
'                   & [CompanyID]) FROM tblCompany;
'               Access textbox control:
'                =ConcatRelated("OrderDate", "tblOrders", "CompanyID = " & [CompanyID])
' Throws:       none
' References:   none
' Source/date:
' Allen Browne, June, 2008
' http://allenbrowne.com/func-concat.html
' Adapted:      Bonnie Campbell, April 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/7/2015 - initial version
'   BLC - 5/1/2015 - integrated into Invasives Reporting tool
'   BLC - 8/21/2015 - added notation re: linked tables
' ---------------------------------
Public Function ConcatRelated(strField As String, _
    strTable As String, _
    Optional strWHERE As String, _
    Optional strOrderBy As String, _
    Optional strSeparator = ", ") As Variant
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelated = Null
    
    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWHERE <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWHERE
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).Type > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).Value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Function:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConcatRelated[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          Coalesce
' Description:  Used in SQL queries to generate concatenated string of records
' Assumptions:  used in Access SQL or control
' Parameters:   strSQL - field to retrieve results from & concatenate (string)
'               NameList() - list of items to concatenate (string)
'               strDelim - character to use between concatenated values (string)
' Returns:      SQL (string, variant, or NULL if no matches)
' Usage:        SQL string:
'               SELECT documents.MembersOnly, Coalsce("SELECT FName From Persons WHERE Member=True",":") AS Who,
'               Coalsce("", ":", "Mary", "Joe", "Pat?") As Others FROM documents;
' Throws:       none
' References:   none
' Source/date:
' Fionuala, September 18, 2008
' http://stackoverflow.com/questions/92698/combine-rows-concatenate-rows?lq=1
' Adapted:      Bonnie Campbell, April 8, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/8/2015  - initial version
'   BLC - 5/1/2015 - integrated into Invasives Reporting tool
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Function Coalsce(strSQL As String, strDelim, ParamArray NameList() As Variant)
On Error GoTo Err_Handler

Dim db As Database
Dim rs As DAO.Recordset
Dim strList As String

    Set db = CurrDb

    If strSQL <> "" Then
        Set rs = db.OpenRecordset(strSQL)

        Do While Not rs.EOF
            strList = strList & strDelim & rs.Fields(0)
            rs.MoveNext
        Loop

        strList = mid(strList, Len(strDelim))
    Else

        strList = Join(NameList, strDelim)
    End If

    Coalsce = strList

Exit_Function:
    'Clean up
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Coalesce[mod_SQL])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' SUB:          PrepareWhereClause
' Description:  Prepares a where clause from multiple params
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 16, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/8/2016  - initial version
' ---------------------------------
Public Function PrepareWhereClause(Params() As String) As String
On Error GoTo Err_Handler
    
    Dim strWHERE As String
    Dim i As Integer
    
    'default
    strWHERE = ""

    'check all params for length, then insert an " AND " if there's a new non-empty clause
    For i = 0 To UBound(Params)
        
        'add to clause
        If Len(strWHERE) > 0 And Len(Params(i)) > 0 Then
            strWHERE = strWHERE & " AND "
        End If
        
        'add param to where clause
        If Len(Params(i)) > 0 Then
            strWHERE = strWHERE & Params(i)
        End If
    Next
    

Exit_Handler:
    PrepareWhereClause = strWHERE
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrepareWhereClause[mod_SQL])"
    End Select
    Resume Exit_Handler
End Function

