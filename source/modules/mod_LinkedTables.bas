' =================================
' MODULE:       basLinkedTables
' Description:  Standard module for verifying and updating links to back-end tables
'
'   The functions in this module require that the database contain the following two tables:
'
'   tsys_Link_Files:  Link_type (txt 50), Link_file_name (txt 100), Link_file_path (txt 255);
'       optional fields:  Link_description (txt 255).  [Link_type] should be 'Back-end data'
'       for the primary back-end database (in case of multiple back-ends).
'
'   tsys_Link_Tables:  Link_type (txt 50), Link_table (txt 100), Table_type (txt 50),
'       Description_text (txt 255).
'
' Source/date:  John R. Boetsch, May 24, 2006
' Revisions:    <name, date, desc - add lines as you go>

Option Compare Database
Option Explicit

' =================================
' FUNCTION:     fxnVerifyLinks
' Description:  Loops through all of the linked tables to verify valid links
' Parameters:   none
' Returns:      True (no link errors) or False
' Throws:       none
' References:   fxnCheckLink
' Source/date:  John R. Boetsch, May 24, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Public Function fxnVerifyLinks() As Boolean
Dim rst As DAO.Recordset
Dim intNumTables As Integer
Dim intI As Integer
Dim varReturn As Variant
Dim strLinkTableName As String

On Error GoTo Err_Handler

fxnVerifyLinks = False  ' Default unless successful

' Set the recordset to the system table joined with the Access MSysObjects table
'   This recordset looks for only linked tables starting with "t", and has joins
'   to allow all actual tables to show up in case the system link table is missing
'   any information.
Set rst = CurrentDb.OpenRecordset("SELECT MSysObjects.Name, MSysObjects.Database, " & _
    "tsys_Link_Files.Link_file_path FROM tsys_Link_Files RIGHT JOIN " & _
    "(MSysObjects LEFT JOIN tsys_Link_Tables ON MSysObjects.Name = " & _
    "tsys_Link_Tables.Link_table) ON tsys_Link_Files.Link_type = " & _
    "tsys_Link_Tables.Link_type WHERE (((MSysObjects.Name) Like 't*') " & _
    "And ((MSysObjects.Type) = 6)) ORDER BY MSysObjects.Name;", dbOpenSnapshot)

' Counts the number of linked tables in the recordset
rst.MoveLast    ' Need to do this to make the record count accurate
intNumTables = rst.RecordCount
If intNumTables = 0 Then    ' No linked tables in the recordset
    MsgBox "There are no linked tables found in the systems tables." & _
        vbCrLf & "Please contact the database administrator before " & _
        "using this application.", vbCritical, "Missing db links (fxnVerifyLinks)"
    GoTo Exit_Procedure
End If

'   Initialize the system meter to indicate progress
varReturn = SysCmd(acSysCmdInitMeter, "Verifying tables", intNumTables)
intI = 0
rst.MoveFirst

' Loop through each record and check for bad links
'   Send to error handler if a bad link is encountered
Do Until rst.EOF
    intI = intI + 1
    varReturn = SysCmd(acSysCmdUpdateMeter, intI)
    strLinkTableName = rst![Name]
    ' Make sure the linked table opens properly
    If fxnCheckLink(strLinkTableName) = False Then
        ' Unable to open a linked table (not a critical error)
        MsgBox "Unable to open the following table:" & vbCrLf & vbCrLf & _
            strLinkTableName, vbExclamation, "Broken table links"
        GoTo Exit_Procedure
    ' Check for linked tables that are not in the system table
    ElseIf IsNull(rst![Link_file_path]) Then
        ' Actual linked table not contained in the system links table
        MsgBox "The following table is not found in the system linking table." & _
            vbCrLf & "Please contact the database administrator before using " & _
            "this application." & vbCrLf & vbCrLf & strLinkTableName, _
            vbCritical, "Missing db links (fxnVerifyLinks)"
        GoTo Exit_Procedure
    ' Make sure the actual linked database matches the system table
    ElseIf rst![Link_file_path] <> rst![Database] Then
        ' The database linking tools are not functioning properly - the
        '   information in the system table does not match the actual linked db
        MsgBox "The actual linked database does not match the information " & _
            "in the system linking table." & vbCrLf & "Please contact the " & _
            "database administrator before using this application.", _
            vbCritical, "Database link update error (fxnVerifyLinks)"
        GoTo Exit_Procedure
    Else
    ' Table link is valid
        rst.MoveNext
    End If
Loop

' If no bad links were encountered
fxnVerifyLinks = True

Exit_Procedure:
    On Error Resume Next
    varReturn = SysCmd(acSysCmdRemoveMeter)
    rst.Close
    Set rst = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case 3135
            MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
                "database administrator before using this application.", vbCritical, _
                "SQL String Error (fxnVerifyLinks)"
        Case 3061   ' Bad parameters for the SQL string
            MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
                "database administrator before using this application.", vbCritical, _
                "SQL String Error (fxnVerifyLinks)"
        Case 3078   ' Missing table from the SQL string
            MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
                "database administrator before using this application.", vbCritical, _
                "SQL String Error (fxnVerifyLinks)"
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnVerifyLinks)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnCheckLink
' Description:  Checks the status of the link for the specified table
' Parameters:   strTable - name of the table to check
' Returns:      True (valid link) or False
' Throws:       none
' References:   none
' Source/date:  From Access97 Developer's Handbook by Litwin, Getz and Gilbert (Sybex)
'               Copyright 1997.  All Rights Reserved
'               Created 09/13/94 pel; Last modified 07/10/96 pel.
' Revisions:    John R. Boetsch, May 17, 2006 - updated documentation, added error traps
' =================================

Public Function fxnCheckLink(strTable As String) As Boolean
Dim varRet As Variant

On Error Resume Next
' Check for failure.  If can't determine the name of
' the first field in the table, the link must be bad.
varRet = CurrentDb.TableDefs(strTable).Fields(0).Name
fxnCheckLink = (Err = 0)
    
Exit_Procedure:
    Exit Function

End Function

' =================================
' FUNCTION:     fxnGetLinkFile
' Description:  Opens the open file dialog and returns the file name
' Parameters:   varInitialDir - the directory to start searching in
' Returns:      the file name, or Null if none is specified
' Throws:       none
' References:   adhAddFilterItem, adhCommonFileOpenSave
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 17, 2006 - updated documentation and error trap
' =================================

Public Function fxnGetLinkFile(Optional ByVal varInitialDir As Variant) As Variant
    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim lngFlags As Long

    ' Use the open file dialog to interactively browse to and select the desired file
    strFilter = adhAddFilterItem(strFilter, "Access (*.*db)", "*.*db")
    
    lngFlags = adhOFN_HIDEREADONLY Or _
        adhOFN_HIDEREADONLY Or adhOFN_NOCHANGEDIR
    
    fxnGetLinkFile = adhCommonFileOpenSave( _
        InitialDir:=varInitialDir, _
        OpenFile:=True, _
        Filter:=strFilter, _
        Flags:=lngFlags, _
        DialogTitle:="Locate data file")

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnGetLinkFile)"
            Resume Exit_Procedure
    End Select

End Function

' =================================
' FUNCTION:     fxnRefreshLinks
' Description:  Updates the link to the specified back-end database tables after first
'               verifying that the tables exist in the specified link file
' Parameters:   strSQL (query listing the tables to re-link), varFileName
' Returns:      True (successfully relinked) or False
' Throws:       none
' References:   none
' Source/date:  Susan Huse, fall 2004 and Mark A. Wotawa, 02/08/2000
' Revisions:    John R. Boetsch, May 22, 2006 - combined verify and refresh functions
'               for table links, fixed meter increment problem updated documentation and
'               error traps
' =================================

Public Function fxnRefreshLinks(strSQL As String, varFileName As Variant) As Boolean
    On Error GoTo Err_Handler

    Dim dbGet As DAO.Database
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim intNumTables As Integer
    Dim varReturn As Variant
    Dim intI As Integer
    Dim strLinkTableName As String

    fxnRefreshLinks = False   ' Default unless all tables verified

    ' Opens the target database and the current system table containing the list
    '   of tables for refreshing links
    Set dbGet = DBEngine.OpenDatabase(varFileName)
    Set db = CurrentDb
    Set rst = db.OpenRecordset(strSQL, dbOpenSnapshot)

    ' Counts the number of tables in the system table associated with this db
    rst.MoveLast    ' Need to do this to make the record count accurate
    intNumTables = rst.RecordCount
    If intNumTables = 0 Then    ' No linked tables in the recordset
        MsgBox "There are no linked tables associated with one or more of " & _
            "these database files." & vbCrLf & "Please contact the database " & _
            "administrator before using this application.", vbCritical, _
            "Missing db links (fxnRefreshLinks)"
        GoTo Exit_Procedure
    End If

    ' First pass to verify the tables in the specified database
    '   Initialize the system meter to indicate progress
    varReturn = SysCmd(acSysCmdInitMeter, "Verifying tables", intNumTables)
    intI = 0
    rst.MoveFirst
    Do Until rst.EOF
        intI = intI + 1
        varReturn = SysCmd(acSysCmdUpdateMeter, intI)
        strLinkTableName = rst![Link_table]
        varReturn = dbGet.TableDefs(strLinkTableName).Fields(0).Name
        rst.MoveNext
    Loop

    ' Second pass to refresh links now that they are validated
    '   Reinitialize the system meter to indicate progress
    varReturn = SysCmd(acSysCmdInitMeter, "Updating table links", intNumTables)
    intI = 0
    rst.MoveFirst
    Do Until rst.EOF
        intI = intI + 1
        varReturn = SysCmd(acSysCmdUpdateMeter, intI)
        strLinkTableName = rst![Link_table]
        Set tdf = db.TableDefs(strLinkTableName)
        tdf.Connect = ";DATABASE=" & varFileName
        tdf.RefreshLink
        rst.MoveNext
    Loop
    
    fxnRefreshLinks = True    ' Links successfully updated

Exit_Procedure:
    On Error Resume Next
    varReturn = SysCmd(acSysCmdRemoveMeter)
    dbGet.Close
    Set dbGet = Nothing
    rst.Close
    Set tdf = Nothing
    Set rst = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    fxnRefreshLinks = False
    Select Case Err.Number
        Case 3021
            MsgBox "Error #" & Err.Number & ":  There are no table links associated " & _
                "with one or more of these files." & vbCrLf & "Please contact the " & _
                "database administrator before using this application.", vbCritical, _
                "Missing db links (fxnRefreshLinks)"
        Case 3024
            MsgBox "Error #" & Err.Number & ":  Cannot find the following file:" & _
                vbCrLf & vbCrLf & varFileName, vbCritical, _
                "Database file not found (fxnRefreshLinks)"
        Case 3078   ' Also got this error if the function call SQL string has a bad
                    '   reference to the system table
            MsgBox "Error #" & Err.Number & ":  The following table is not native " & _
                "to the selected database file." & vbCrLf & "Please make sure you " & _
                "browsed to to the correct file." & vbCrLf & vbCrLf & strLinkTableName, _
                vbCritical, "Incorrect link file (fxnRefreshLinks)"
        Case 3061   ' Bad parameters for the SQL string
            MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
                "database administrator before using this application.", vbCritical, _
                "SQL String Error (fxnRefreshLinks)"
        Case 3265
            MsgBox "Error #" & Err.Number & ":  The database file is missing the " & _
                "following table:" & vbCrLf & vbCrLf & strLinkTableName, _
                vbCritical, "Missing database table (fxnRefreshLinks)"
        Case 3219 ' Trying to update a link on top of an imported table
            MsgBox "Error #" & Err.Number & ":  You are trying to update a link to " & _
                "a table that has already been imported." & vbCrLf & vbCrLf & _
                strLinkTableName & vbCrLf & vbCrLf & "Please call the database " & _
                "administrator to help you relink this table manually." & vbCrLf & _
                "Afterwards you will be able to automatically update links again.", _
                vbCritical, "Link error (fxnRefreshLinks)"
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (fxnRefreshLinks)"
    End Select
    Resume Exit_Procedure

End Function