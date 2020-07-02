Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_Linked_Tables
' Level:        Framework module
' Version:      1.09
' Description:  Linked table related functions & subroutines
'
' Adapted from: John R. Boetsch, May 24, 2006
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    JRB, 7/9/2009 - simplified by moving certain functions to another module
'               JRB, 12/30/2009 - moved fxnVerifyLinks to another module
'               ------------------------------------------------------------------------
'               BLC, 4/30/2015 - 1.00 - added fxnVerifyLinks, fxnRefreshLinks, fxnVerifyLinkTableInfo,
'                                fxnMakeBackup from mod_Custom_Functions
'               BLC, 5/19/2015 - 1.01 - renamed functions, removed fxn prefix
'               ------------------------------------------------------------------------
'               BLC, 8/22/2017 - 1.07 - merge prior versions into single framework version
'               ------------------------------------------------------------------------
'                       BLC, 6/10/2015 - 1.02 - fixed VerifyLinkTableInfo to add new linked tables to tsys_Link_Tables
'                       BLC, 6/12/2015 - 1.03 - replaced TempVars.item(... with TempVars("...
'                       BLC, 9/30/2015 - 1.04 - added check & resolve double quotes in table descriptions in RefreshLinks
'                       BLC, 12/1/2015 - 1.05 - resolve issues with linked database updates to differently named backend databases
'                       BLC, 12/3/2015 - 1.06 - added UpdateTSysTableDb
'               ------------------------------------------------------------------------
'                       BLC, 6/5/2016  - 1.02 - renamed frm_Progress_Meter to ProgressMeter,
'                                           removed underscores from fields
'                       BLC, 1/24/2017 - 1.03 - revised MakeBackup() to use FilePath vs. File_path
'                                               (tsys_Link_Dbs)
'                       BLC, 2/22/2017 - 1.04 - added alternative path for new vs. legacy forms (ConnectDbs
'                                           vs. frm_Connect_Dbs), BACKEND_REQUIRED check
'               ------------------------------------------------------------------------
'               BLC, 10/4/2017 - 1.08 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC, 5/16/2019 - 1.09 - added fw_ module prefix
' =================================

' ---------------------------------
'  References
' ---------------------------------

' --------------------------------------------------------------------------------
'   Msys Objects
' --------------------------------------------------------------------------------
' Source: Pat Hartman March 13, 2006
'         http://www.access-programmers.co.uk/forums/showthread.php?t=103811
' --------------------------------------------------------------------------------
'   Type   TypeDesc           Type  TypeDesc
'  -32768  Form                 1   Table - Local Access Tables
'  -32766  Macro                2   Access Object - Database
'  -32764  Reports              3   Access Object - Containers
'  -32761  Module               4   Table - Linked ODBC Tables
'  -32758  Users                5   Queries
'  -32757  Database Document    6   Table - Linked Access Tables
'  -32756  Data Access Pages    8   SubDataSheets
' --------------------------------------------------------------------------------


' ---------------------------------
'   Database Level
' ---------------------------------

' ---------------------------------
' FUNCTION:     VerifyConnections
' Description:  Checks the status of back-end connections
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   FileExists, FormIsOpen, TestODBCConnection, VerifyLinks,
'                   VerifyLinkTableInfo
' Source/date:  Susan Huse, fall 2004
' Adapted:      Bonnie Campbell, June, 2014 for NCPN WQ Utilities tool
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
'               JRB, 5/24/2006 - updated documentation, error traps, modified backup
'                   strategy and added verification of individual table links
'               JRB, 7/27/2006 - added code to open the always-open back-end connection
'                   form upon confirming a good connection
'               JRB, 6/29/2009 - revised system table structure; default to connected=False;
'                   removed backup call; revised to work with ODBC connections
'               JRB, 10/8/2009 - added Proc_Final_Status to make verifying connections
'                   optional if there is an Access back-end file
'               BLC, 7/31/2014 - changed gvars to TempVars, shifted to initApp module
'               BLC, 9/5/2014  - added check for remote (network) backends (IsNetworkFile)
'               BLC, 4/30/2015 - switched from fxnSwitchboardIsOpen to FormIsOpen(frmSwitchboard)
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 5/22/2015 - moved from mod_Initialize_App to mod_Linked_Tables
'               BLC, 6/12/2015 - replaced TempVars.item("... with TempVars("...
'               BLC, 6/5/2016  - removed underscores from field names
'               BLC, 2/22/2017 - added alternative path for new vs. legacy forms (ConnectDbs
'                                vs. frm_Connect_Dbs), BACKEND_REQUIRED check
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Function VerifyConnections()
    On Error GoTo Err_Handler
    
    PushCallStack "VerifyConnections"

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSysTable As String
    Dim strDbName As String
    Dim strTable As String
    Dim strDbPath As String
    Dim strServer As String
    Dim strErrMsg As String
    Dim blnHasError As Boolean

    Set db = CurrDb
    TempVars.Item("Connected") = False           ' Default in case of error
    TempVars.Item("HasAccessBE") = False         ' Flag to indicate that at least 1 Access BE exists
    strSysTable = "tsys_Link_Dbs"   ' System table listing linked tables
    blnHasError = False             ' Flag to indicate error status

    ' Check the information in the application tables, exit if there is a problem
    If VerifyLinkTableInfo = False Then GoTo Exit_Procedure

    ' Set the recordset to the system table
    Set rs = db.OpenRecordset(strSysTable, dbOpenTable, dbReadOnly)

    Do Until rs.EOF
        strDbName = rs.Fields("LinkDb")
        If rs.Fields("IsODBC") = True Then
            ' ODBC connection
            If Not IsNull(rs![Server]) Then
                strServer = rs![Server]
                ' Test the firs table in the list for this back-end to test the connection
                strTable = DFirst("[LinkTable]", "tsys_Link_Tables", _
                    "[LinkDb]=""" & strDbName & """")
                If TestODBCConnection(strTable, , , False) = False Then
                    blnHasError = True
                    If strErrMsg <> "" Then strErrMsg = strErrMsg & vbCrLf & vbCrLf
                    strErrMsg = strErrMsg & _
                        "The following server connection is not working:" & _
                        vbCrLf & "  Db name: " & strDbName & _
                        vbCrLf & "  Server:  " & strServer
                End If
            Else    ' Missing server name
                If strErrMsg <> "" Then strErrMsg = strErrMsg & vbCrLf & vbCrLf
                strErrMsg = strErrMsg & _
                    "Missing the server name for the following database:" & _
                    vbCrLf & "  Db name: " & strDbName
            End If
        Else
            ' Access back-end - update the global variable
            TempVars.Item("HasAccessBE") = True
            If Not IsNull(rs![FilePath]) Then
                strDbPath = rs![FilePath]
                If FileExists(strDbPath) = False Then
                    ' Cannot find the file
                    blnHasError = True
                    If strErrMsg <> "" Then strErrMsg = strErrMsg & vbCrLf & vbCrLf
                    strErrMsg = strErrMsg & _
                        "The following database file cannot be located:" & _
                        vbCrLf & "  Db name: " & strDbName & _
                        vbCrLf & "  " & strDbPath
                'Else
                    ' Check if file is remote (network) & set bit to alert user that db (app) may be slow
                    'If IsNetworkFile(strDbPath) Then
                    '    rs![Is_Network_Db] = 1
                    'End If
                End If
            Else    ' Missing file path
                blnHasError = True
                If strErrMsg <> "" Then strErrMsg = strErrMsg & vbCrLf & vbCrLf
                strErrMsg = strErrMsg & _
                    "Missing the file path for the following database:" & _
                    vbCrLf & "  Db name: " & strDbName
            End If
        End If
        rs.MoveNext
    Loop
    
    'For applications with full DbAdmin subform (DB_ADMIN_CONTROL = True) otherwise ignore
    If DB_ADMIN_CONTROL = True Then
    
        ' Check the status of individual table links, depending on application settings
        If FormIsOpen("frmSwitchboard") And blnHasError = False Then
            If Forms!frm_Switchboard.fsub_DbAdmin.Form.chkVerifyOnStartup Then
                If TempVars.Item("HasAccessBE") = True Then
                    If MsgBox("Would you like all linked table connections to be tested?", _
                        vbYesNo + vbDefaultButton2, _
                        "Checking back-end connections ...") = vbNo Then GoTo Proc_Final_Status
                End If
                If VerifyLinks = False Then
                    blnHasError = True
                    If strErrMsg <> "" Then strErrMsg = strErrMsg & vbCrLf & vbCrLf
                    strErrMsg = strErrMsg & _
                        "One or more table connections is not working properly."
                End If
            End If
        End If

    End If
    
Proc_Final_Status:
    If blnHasError Then
        If strErrMsg <> "" Then strErrMsg = strErrMsg & vbCrLf & vbCrLf
        strErrMsg = strErrMsg & _
            "You must update the back-end database connections" & vbCrLf & _
            "by selecting 'Db connections' from the menu before" & vbCrLf & _
            "using this application." & vbCrLf & vbCrLf & _
            "Would you like to fix the connection now?"
        ' Notify the user with specific error information
        If MsgBox(strErrMsg, vbCritical + vbYesNo, "Update back-end connections") _
            = vbYes Then
            ' Open the form to reconnect back-end tables
            If DbObjectExists("frm_Connect_Db") Then
                DoCmd.OpenForm "frm_Connect_Dbs"    'legacy
            ElseIf DbObjectExists("ConnectDbs") Then
                DoCmd.OpenForm "ConnectDbs", acNormal, , , , , "PreSplash"    'new
            End If
        Else
            If BACKEND_REQUIRED Then
                'close since the back-end is required
                 If MsgBox("A viable back-end is required for this application." & _
                 vbCrLf & "So I'm closing now unless you click ""No.""" & vbCrLf & vbCrLf & "", _
                 vbCritical + vbYesNo, "Closing database...") = vbYes Then _
                   DoCmd.CloseDatabase
             End If
        End If
    Else  ' If no connection errors, then set the global variable flag to True
        TempVars.Item("Connected") = True
    End If

Exit_Procedure:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    PopCallStack "VerifyConnections"
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3135, 3061, 3078  ' Problem with SQL syntax, or ref to nonexistent object, etc.
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyConnections[fw_mod_Linked_Tables])"
      Case 3011, 7874   ' System table not found
         MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyConnections[fw_mod_Linked_Tables])"
      Case 3265   ' Field name in the system table improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyConnections[fw_mod_Linked_Tables])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyConnections[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure
End Function

' =================================
' FUNCTION:     LinkedDatabase
' Description:  Returns the database file path (Access) or the database name (ODBC) for
'                   a linked table name
' Parameters:   strTableName - the name of the linked table
' Returns:      database name for the linked table, or empty string ("") if none
' Throws:       none
' References:   ParseConnectionStr
' Source/date:  John R. Boetsch, 6/24/2009
' Revisions:    JRB, 6/24/2009 - initial version
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                 multiple open connections
' =================================
Public Function LinkedDatabase(ByVal strTableName As String) As String
    On Error GoTo Err_Handler

    Dim strTemp As String

    strTemp = ParseConnectionStr(CurrDb.TableDefs(strTableName).connect)
    LinkedDatabase = strTemp

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3265
        MsgBox "The table '" & strTableName & "' was not found in the front-end.", _
            vbCritical, "Error encountered (#" & Err.Number & " - LinkedDatabase[fw_mod_Linked_Tables])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LinkedDatabase[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     ParseConnectionStr
' Description:  Return the specified portion of the linked table connection string
' Parameters:   strConnStr - linked table connection string
'               strComponent - optional string to specify the portion to return
'                   (default "DATABASE=")
'               strDelimiter - optional string delimiter (default ";")
'               blnIsFound - optional reference variable to incidate whether the
'                   specified string component is found in the connection string
' Returns:      connection string component, or empty string ("") if not found
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/24/2009
' Revisions:    JRB, 6/24/2009 - initial version
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function ParseConnectionStr(strConnStr As String, _
    Optional strComponent As String = "DATABASE=", _
    Optional strDelimiter As String = ";", _
    Optional blnIsFound As Boolean = False) As String

    On Error GoTo Err_Handler

    Dim varStartPos As Variant
    Dim varEndPos As Variant
    Dim varLength As Variant
    Dim strResult As String

    varStartPos = InStr(1, strConnStr, strComponent, vbTextCompare)
    If IsNull(varStartPos) Or IsEmpty(varStartPos) Or varStartPos = 0 Then
        ' The component is not found in the connection string
        blnIsFound = False
    Else
        blnIsFound = True
        ' Determine the end position of the database string
        varEndPos = InStr(varStartPos, strConnStr, strDelimiter, vbTextCompare)
        If varEndPos > varStartPos Then
            ' There is a delimiter following the desired string
            varStartPos = varStartPos + Len(strComponent)
            varLength = varEndPos - varStartPos
            ParseConnectionStr = mid(strConnStr, varStartPos, varLength)
        Else
            varLength = Len(strConnStr) - varStartPos + 1 - Len(strComponent)
            ParseConnectionStr = Right(strConnStr, varLength)
        End If
    End If
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ParseConnectionStr[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     MakeBackup
' Description:  Creates a backup of linked Access back-end database files
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   CreateFolder, FolderExists, ParsePath, ParseFileName,
'                   ParseFileExt, SaveFile, ZipFiles, FileExists, Pause
' Source/date:  Susan Huse, fall 2004
' Revisions:    John R. Boetsch, May 2005 - minor revisions and documentation
'               JRB, 5/16/2006 - updated documentation, error traps; modified date/time
'                   stamp to be appended to the file name; changed varCopyFile to a Variant
'                   to accommodate nulls from the procedure call
'               JRB, 1/8/2009 - streamlined to use zip files capability
'               JRB, 6/29/2009 - additional updates to accommodate multiple back-ends and
'                   revised system table structure
'               JRB, 10/8/2009 - inserted a pause in zip file creation to avoid closing
'                   before large back-end files are fully zipped
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 1/24/2017 - revised DLookup for tsys_Link_Dbs to check FilePath vs. File_path
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' =================================
Public Function MakeBackup()
    On Error GoTo Err_Handler

    ' Prompt the user to confirm before backing up ... if no, exit function
    If MsgBox("Would you like to make a backup copy of the data?", vbYesNo, _
        "Create Backup?") = vbNo Then
        GoTo Exit_Procedure
    End If

    Dim rs As DAO.Recordset
    Dim intNRecs As Integer
    Dim strDbFile As String
    Dim fs As Variant
    Dim varCopyFile As Variant
    Dim arrFile() As String
    Dim strNewFile As String
    Dim strPath As String
    Dim strBackupDate As String
    Dim blnZipped As Boolean
    Dim strBackupFolder As String

    strBackupFolder = "Db_backups"
    strBackupDate = Format$(Now, "YYYYMMDD_HHNN")

    ' Set the recordset to the systems table, grouped by linked Access databases
    Set rs = CurrDb.OpenRecordset("SELECT Database " & _
        "FROM MSysObjects " & _
        "WHERE ((MSysObjects.Type) = 6) And ((MSysObjects.Name) Not Like '~*') " & _
        "GROUP BY MSysObjects.Database;", dbOpenSnapshot)

    ' Counts the number of linked Access back-end files in the database
    rs.MoveLast    ' Need to do this to make the record count accurate
    intNRecs = rs.RecordCount
    If intNRecs = 0 Then    ' No linked databases in the recordset
        MsgBox "There are no Access back-end files to back up ...", , _
            "No back-end file to back up"
        GoTo Exit_Procedure
    End If

    ' Loop through the recordset and back up each file as indicated in the system file
    rs.MoveFirst
    Do Until rs.EOF
        strDbFile = rs![Database]
        ' If the string is not empty and backups are indicated for this back-end ...
        If strDbFile <> "" And _
            DLookup("[Backups]", "tsys_Link_Dbs", "[FilePath]=""" & strDbFile & """") Then

            ' Remove the file name from the path
            strPath = ParsePath(strDbFile)
            ' Remove the right-most back slash if present
            If Right(strPath, 1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
            ' Update the backup folder string unless it is already the current folder
            arrFile() = Split(strPath, "\")
            If strBackupFolder <> arrFile(UBound(arrFile)) Then _
                strPath = strPath & "\" & strBackupFolder
            ' Verify the existence of the backup folder (and create it if needed)
            If FolderExists(strPath) = False Then CreateFolder (strPath)
            If FolderExists(strPath) = False Then
                MsgBox "Unable to find/create the backup folder.", , "No Backup Made"
                GoTo Exit_Procedure
            End If
            ' Create the new file string by adding the current file name to the new path
            strNewFile = strPath & "\" & ParseFileName(strDbFile)
            ' Remove the current file extension
            strNewFile = Left(strNewFile, Len(strNewFile) - Len(ParseFileExt(strDbFile)))
            ' Append the backup date/time
            strNewFile = strNewFile & "_" & strBackupDate
            ' Zip the file to the new destination file name plus the ".zip" extension
            blnZipped = ZipFiles(strDbFile, strNewFile & ".zip")
            If blnZipped Then
                Dim intCounter As Integer
                intCounter = 0
                Call Pause(1000)
                Do While intCounter < 120
                    intCounter = intCounter + 1
                    If FileExists(strNewFile & ".zip") Then
                        Exit Do
                    Else
                        ' Pause for 1000 ms before trying again
                        Call Pause(1000)
                    End If
                Loop
                MsgBox "Backup file successfully created: " & vbCrLf & vbCrLf & _
                    strNewFile & ".zip", vbOKOnly
            Else
                ' Zip operation unsuccessful, so try to make an outright copy
                ' Open the save file dialog and update to the actual name given by the user
                varCopyFile = SaveFile(strNewFile, _
                    "Microsoft Access (*.mdb, *.accdb)", "*.mdb;*.accdb")
                If IsNull(varCopyFile) Then
                    ' User canceled save operation
                    MsgBox "No backup made", vbOKOnly
                Else
                    ' Perform the actual file copy
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    fs.CopyFile strDbFile, varCopyFile
                    MsgBox "Backup file successfully created: " & vbCrLf & vbCrLf & _
                        varCopyFile, vbOKOnly
                End If
            End If
            
        End If
        rs.MoveNext
    Loop    ' To next back-end

Exit_Procedure:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set fs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MakeBackup[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function


' ---------------------------------
'   Table Level
' ---------------------------------

' =================================
' FUNCTION:     CheckLink
' Description:  Checks the status of the link for the specified table
' Parameters:   strTableName - name of the table to check
' Returns:      True (valid link) or False
' Throws:       none
' References:   none
' Source/date:  From Access97 Developer's Handbook by Litwin, Getz and Gilbert (Sybex)
'               Copyright 1997.  All Rights Reserved
'               Created 09/13/94 pel; Last modified 07/10/96 pel.
' Revisions:    John R. Boetsch, May 17, 2006 - updated documentation, added error traps
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' =================================
Public Function CheckLink(strTableName As String) As Boolean
    On Error GoTo Err_Handler

    Dim varRet As Variant

    On Error Resume Next
    ' Check for failure.  If can't determine the name of
    ' the first field in the table, the link must be bad.
    varRet = CurrDb.TableDefs(strTableName).Fields(0).Name
    If Err <> 0 Then
        CheckLink = False
    Else
        CheckLink = True
    End If
    
Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CheckLink[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     IsODBC
' Description:  Determine whether the input table is connected by ODBC
' Parameters:   strTableName - string for the table name
' Returns:      True (if table object in collection and ODBC) or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/24/2009
' Revisions:    JRB, 6/24/2009 - initial version
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function IsODBC(strTableName As String) As Boolean
    On Error GoTo Err_Handler

    Dim strCriteria As String

    strCriteria = "(([Name])=""" & strTableName & """) AND (([Type]) In (1, 4, 6))"
    If DLookup("Type", "MSysObjects", strCriteria) = 4 Then
        ' ODBC connection
        IsODBC = True
    Else
        ' Native table or linked Access table
        IsODBC = False
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsODBC[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     TestODBCConnection
' Description:  Uses a pass-through query to test an ODBC connection and to trap ODBC errors
' Parameters:   strTableName - the name of the linked table
'               strConnStr - optional linked table connection string
'               varSQL - optional SQL statement to execute
'               blnRetErrMsg - optional flag to show error msg if the test fails (default=True)
' Returns:      True if the connection returns records, otherwise False
' Throws:       none
' References:   ParseConnectionStr
' Source/date:  John R. Boetsch, 6/24/2009 (adapted from http://support.microsoft.com/kb/210319)
' Revisions:    JRB, 6/24/2009 - initial version
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                 multiple open connections
' =================================
Function TestODBCConnection(strTableName As String, _
    Optional ByVal strConnStr As String, _
    Optional varSQL As Variant, _
    Optional blnRetErrMsg As Boolean = True) As Boolean

    On Error GoTo Err_Handler

    TestODBCConnection = False   ' Default in case of error

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strDbName As String

    ' Create a blank pass-through query
    Set db = CurrDb()
    Set qdf = db.CreateQueryDef("")

    ' If no revised connection string was passed, use the current connection string
    If strConnStr = "" Then strConnStr = CurrDb.TableDefs(strTableName).connect
    strDbName = ParseConnectionStr(strConnStr)

    ' Update the connection string for the pass-through query, set to not return records
    qdf.connect = strConnStr
    qdf.ReturnsRecords = False

    If IsMissing(varSQL) Then
        ' If no query statement passed, select a few records to test the connection string
        qdf.SQL = "SELECT TOP 2 * FROM " & strTableName
    Else: qdf.SQL = varSQL
    End If
    qdf.Execute

    ' Set to true (if no errors)
    TestODBCConnection = True

Exit_Procedure:
    On Error Resume Next
    Set db = Nothing
    Set qdf = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3151       ' Connection failed
        If blnRetErrMsg Then _
        MsgBox "Cannot connect to the specified database/table:" & vbCrLf & vbCrLf & _
            "  Db: " & strDbName & vbCrLf & "  Table: " & strTableName, vbCritical, _
            "Error encountered (#" & Err.Number & " - TestODBCConnection[fw_mod_Linked_Tables])"
      Case 3146, 208  ' Connection failed
        If blnRetErrMsg Then _
        MsgBox "Cannot find the table in the specified database:" & vbCrLf & vbCrLf & _
            "  Db: " & strDbName & vbCrLf & "  Table: " & strTableName, vbCritical, _
            "Error encountered (#" & Err.Number & " - TestODBCConnection[fw_mod_Linked_Tables])"
      Case 3305       ' Invalid pass-through connection string
        MsgBox "Invalid pass-through query connection string ..." & vbCrLf & _
            strTableName & " may not be an ODBC-linked table.", vbCritical, _
            "Error encountered (#" & Err.Number & " - TestODBCConnection[fw_mod_Linked_Tables])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TestODBCConnection[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     RefreshLinks
' Description:  Updates the link to the specified back-end database tables after first
'               verifying that the tables exist in the specified link file
' Parameters:   strDbName - name of the database to refresh
'               strNewConnStr - updated connection string
'               blnIsODBC - flag to indicate that the back-end is ODBC (default = False)
' Returns:      True (successfully relinked) or False
' Throws:       none
' References:   ParseConnectionStr, TestODBCConnection
' Source/date:  Susan Huse, fall 2004 and Mark A. Wotawa, 02/08/2000
' Revisions:    John R. Boetsch, 5/22/2006 - combined verify and refresh functions
'                   for table links, fixed meter increment problem updated documentation
'                   and error traps
'               JRB, 7/9/2009 - updated to accommodate ODBC links and to update the table
'                   description in tsys_Link_Tables for Access tables
'               JRB, 12/30/2009 - updated to use the popup progress meter form
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions & renamed RefreshLinks
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 5/20/2015 - updated progress meter control naming, added connection component for non-"DATABASE="
'                                connection strings (e.g. Access 2010 w/ "Dbq=")
'               BLC, 6/4/2016  - revised tsys_Link_Tables fields to match Big Rivers field naming revisions (LinkDb vs Link_db, LinkTable vs. Link_table)
'                                renamed frm_Progress_Meter to ProgressMeter
'                               -------------------------------------------------------------------------
'                               BLC, 8/22/2017 - merged in prior work
'                       BLC, 9/30/2015 - add description parsing to avoid errors due to quotes
'                       BLC, 12/1/2015 - resolve issues with linked database updates to differently named backend databases
'                       BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                           multiple open connections
' =================================
Public Function RefreshLinks(strDbName As String, ByVal strNewConnStr As String, _
    Optional strComponent As String = "DATABASE=", _
    Optional ByVal blnIsODBC As Boolean = False, _
    Optional strNewDbName As String _
        ) As Boolean
    On Error GoTo Err_Handler

    Dim varFileName As Variant
    Dim dbGet As DAO.Database
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim intNumTables As Integer
    Dim varReturn As Variant
    Dim intI As Integer
    Dim strTable As String
    Dim strDesc As String
    Dim strSQL As String
    Dim frm As Form             ' Reference to the progress popup form
    Dim strProgForm As String   ' Name of the progress popup form
    Dim strProgress As String   ' Progress bar string

    RefreshLinks = False   ' Default unless all tables verified
    'set new db name default to current name if strNewDbName not populated
    If Len(strNewDbName) = 0 Then strNewDbName = strDbName

    Set db = CurrDb
    Set rs = db.OpenRecordset(GetTemplate("s_tsys_link_tables_by_dbname", "dbName" & PARAM_SEPARATOR & strDbName), dbOpenSnapshot)
'    Set rs = db.OpenRecordset("SELECT * FROM tsys_Link_Tables WHERE " & _
'                "[tsys_Link_Tables]![LinkDb] = """ & strDbName & """", dbOpenSnapshot)

    ' Counts the number of tables in the system table associated with this db
    rs.MoveLast    ' Need to do this to make the record count accurate
    intNumTables = rs.RecordCount

    ' Initialize the progress popup form
    strProgForm = "ProgressMeter"
    DoCmd.OpenForm strProgForm
    Set frm = Forms!ProgressMeter
    frm.Caption = " Updating table connections"
    frm!tbxPercent = 0

    If blnIsODBC = False Then   ' Access back-end
        ' Opens the target database and the current system table containing the list
        '   of tables for refreshing links
        varFileName = ParseConnectionStr(strNewConnStr, strComponent)
        Set dbGet = DBEngine.OpenDatabase(varFileName)

        ' First pass to verify the tables in the new back-end database (avoids partial updates)
        '   Initialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Verifying tables in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!tbxMsg = "Verifying tables in " & strDbName
        intI = 0
        rs.MoveFirst
        Do Until rs.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!tbxPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!tbxProgress = strProgress
            frm.Repaint
            strTable = rs![LinkTable]
            Debug.Print strTable
            varReturn = dbGet.TableDefs(strTable).Fields(0).Name
            rs.MoveNext
        Loop

        ' Second pass to refresh all links now that they are validated
        '   Reinitialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Updating table links in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!tbxMsg = "Updating table links in " & strDbName
        intI = 0
        rs.MoveFirst
        Do Until rs.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!tbxPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!tbxProgress = strProgress
            frm.Repaint
            strTable = rs![LinkTable]
Debug.Print strTable
            ' Update and refresh the table connection
            Set tdf = db.TableDefs(strTable)
            tdf.connect = strNewConnStr
            tdf.RefreshLink
            ' Update the table description & Link_db in tsys_Link_Tables
            ' Set default description in case there is none
            ' Encode SQL specials (",') in description
            strDesc = " - no description - "
            strDesc = SQLencode(tdf.Properties("Description")) ' Throws trapped error 3270 if none
Debug.Print strDesc
                        
                        ''replace double quotes with singles
            'strDesc = Replace(strDesc, """", "'")
            
                        strSQL = GetTemplate("u_tsys_link_tables_description", "descr" & PARAM_SEPARATOR & strDesc & "|tbl" & PARAM_SEPARATOR & strTable)
'
'            strSQL = "UPDATE tsys_Link_Tables " & _
'                "SET tsys_Link_Tables.DescriptionText=""" & strDesc & _
'                """ WHERE (((tsys_Link_Tables.LinkTable)=""" & strTable & """));"
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
                        'update database name & description in tsys_Link_Dbs & tsys_Link_Files
            'within form modules (frm_Connect_Tables / frm_Connect_Dbs)
            rs.MoveNext
        Loop
    Else    ' ODBC back-end
        ' First pass to verify the tables in the new back-end database (avoids partial updates)
        '   Initialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Verifying tables in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!tbxMsg = "Verifying tables in " & strDbName
        intI = 0
        rs.MoveFirst
        Do Until rs.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!txtPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!tbxProgress = strProgress
            frm.Repaint
            strTable = rs![LinkTable]
            If TestODBCConnection(strTable, strNewConnStr) = False Then GoTo Exit_Procedure
            rs.MoveNext
        Loop

        ' Second pass to refresh all links now that they are validated
        '   Reinitialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Updating table links in " & _
            strDbName, intNumTables)
        ' Update the message below the progress meter
        frm!txtMsg = "Updating table links in " & strDbName
        intI = 0
        rs.MoveFirst
        Do Until rs.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            ' Update the popup progress meter
            frm!tbxPercent = Round(100 * intI / intNumTables)
            ' Update the progress bar in the progress popup with sequential "�" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNumTables), "�")
            frm!tbxProgress = strProgress
            frm.Repaint
            strTable = rs![LinkTable]
            ' Update and refresh the table connection
            Set tdf = db.TableDefs(strTable)
            ' Use test again to trap errors
            If TestODBCConnection(strTable, strNewConnStr) = True Then
                tdf.connect = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};DATABASE=C:\___TEST_DATA\Invasives_be.accdb;" 'strNewConnStr
                tdf.RefreshLink
            Else
                GoTo Exit_Procedure
            End If
            rs.MoveNext
        Loop
    End If

    RefreshLinks = True    ' Links successfully updated

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    varReturn = SysCmd(acSysCmdRemoveMeter)
    DoCmd.Close acForm, strProgForm, acSaveNo
    Set frm = Nothing
    dbGet.Close
    Set dbGet = Nothing
    rs.Close
    Set tdf = Nothing
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    RefreshLinks = False
    Select Case Err.Number
      Case 3021
        MsgBox "Error #" & Err.Number & ":  There are no table links associated " & _
            "with one or more of these files." & vbCrLf & "Please contact the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
      Case 3024
        MsgBox "Error #" & Err.Number & ":  Cannot find the following file:" & _
            vbCrLf & vbCrLf & varFileName, vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
      Case 3061   ' Bad parameters for the SQL string
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
      Case 3074   ' Missing operator, get this error also if SQL string contains double quotes
        MsgBox "Error #" & Err.Number & ": " & Err.Description & vbCrLf & _
            "This can be caused by double quotes in the SQL string.", vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
      Case 3078   ' Also got this error if the function call SQL string has a bad
                '   reference to the system table
        MsgBox "Error #" & Err.Number & ":  The following table is not native " & _
            "to the selected database file." & vbCrLf & "Please make sure you " & _
            "browsed to to the correct file." & vbCrLf & vbCrLf & strTable, _
            vbCritical, "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
      Case 3265
        MsgBox "Error #" & Err.Number & ":  The database file is missing the " & _
            "following table:" & vbCrLf & vbCrLf & strTable, _
            vbCritical, "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
      Case 3219 ' Trying to update a link on top of an imported table
        MsgBox "Error #" & Err.Number & ":  You are trying to update a link to " & _
            "a table that has already been imported." & vbCrLf & vbCrLf & _
            strTable & vbCrLf & vbCrLf & "Please call the database " & _
            "administrator to help you relink this table manually." & vbCrLf & _
            "Afterwards you will be able to automatically update links again.", _
            vbCritical, "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
      Case 3270     ' Property not found (TableDefs description)
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshLinks[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     VerifyLinkTableInfo
' Description:  Verifies that the information in tsys_Link_Dbs and tsys_Link_Tables is
'                   complete and matches that in MSysObjects
' Note:
'       MSysObjects types include:
'           Type    TypeDesc                            Type    TypeDesc
'           1       Table - Local Access Tables         -32756  Data Access Pages
'           2       Access Object - Database            -32757  Database Document
'           3       Access Object - Containers          -32758  Users
'           4       Table - Linked ODBC Tables          -32761  Module
'           5       Queries                             -32764  Reports
'           6       Table - Linked Access Tables        -32766  Macro
'           8       SubDataSheets                       -32768  Form
'       2008 refs:
'           0       Select Query (visible)              8       Select Query (hidden)
'           16      Crosstab Query (visible)            24      Crosstab Query (hidden)
'           32      Delete Query (visible)              40      Delete Query (hidden)
'           48      Update Query (visible)              56      Update Query (hidden)
'           64      Append Query (visible)              72      Append Query (hidden)
'           80      Make table Query (visible)          88      Make table Query (hidden)
'           96      Data definition Query (visible)     104     Data definition Query (hidden)
'           112     Pass through Query (visible)        120     Pass through Query (hidden)
'           128     Union Query (visible)               136     Union Query (hidden)
'
' Parameters:   none
' Returns:      True if the information matches and there are no problems, False otherwise
' Throws:       none
' References:
'       Tom Wickerath, April 26, 2008
'       http://www.pcreview.co.uk/threads/re-help-dirk-goldgar-or-someone-familiar-with-dev-ashish-search.3482377/
'       Fionnula, February 1, 2014
'       http://stackoverflow.com/questions/3994956/meaning-of-msysobjects-values-32758-32757-and-3-microsoft-access
' Source/date:  John R. Boetsch, 7/9/2009
' Revisions:    JRB, 7/27/2009 - added a do loop to update missing table descriptions
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 5/19/2015 - added check for FIX_LINKED_DBS flag when DbAdmin is not fully implemented
'               BLC, 6/4/2016  - adapted to Big Rivers Application, adjust to renamed tsys_Link_Tables fields
'                               -------------------------------------------------------------------------
'                               BLC, 8/22/2017 - merged in prior work
'                              BLC, 6/10/2015 - updated SQL insert into tsys_Link_Tables for missing MSysObjects tables
'                                                captured by qsys_Linked_tables_not_in_tsys_Link_Tables (missing Link_type)
'                                                bug resulted in new linked tables never being inserted into tsys_Link_Tables & subsequent errors
'                               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                                multiple open connections
' =================================
Public Function VerifyLinkTableInfo() As Boolean
    On Error GoTo Err_Handler

    PushCallStack "VerifyLinkTableInfo"

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim intNRecs As Integer
    Dim strTable As String
    Dim strDesc As String
    Dim blnHasError As Boolean
    Dim strSQL As String

    Set db = CurrDb
    blnHasError = False             ' Flag to indicate error status

    ' Check if FIX_LINKED_DBS is set (usually when DbAdmin is not fully implemented)
    If FIX_LINKED_DBS Then
        FixLinkedDatabase "tbl_Target_Species"
    End If

    ' First make sure that there are linked tables
    intNRecs = DCount("*", "MSysObjects", "([Type] In (4,6)) And ([Name] Not Like '~*')")
    If intNRecs = 0 Then    ' No linked tables in the recordset
        MsgBox "There are no linked tables found in the systems tables." & _
            vbCrLf & "Please contact the database administrator before " & _
            "using this application.", vbCritical, "Application error (VerifyLinkTableInfo[fw_mod_Linked_Tables])"
        GoTo Exit_Procedure
    End If

    ' Look for linked table records that no longer actually exist in the database
    intNRecs = DCount("*", "qsys_Linked_tables_not_in_MSysObjects")
    If intNRecs > 0 Then
        Set rs = db.OpenRecordset("qsys_Linked_tables_not_in_MSysObjects", _
            dbOpenSnapshot)
        Do Until rs.EOF
            ' Delete mismatched records from tsys_Link_Tables
'            strSQL = "DELETE * FROM tsys_Link_Tables WHERE ([LinkTable]=""" & _
'                rs![Link_table] & """);"
            strSQL = GetTemplate("d_tsys_link_tables", "linktbl" & PARAM_SEPARATOR & rs![LinkTable])
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            rs.MoveNext
        Loop
        rs.Close
        ' Throw an error if there are still mismatched records
        If DCount("*", "qsys_Linked_tables_not_in_MSysObjects") > 0 Then
            blnHasError = True
            DoCmd.OpenQuery "qsys_Linked_tables_not_in_MSysObjects", , acReadOnly
        End If
    End If

    ' Look for linked tables that are not in the application table
    intNRecs = DCount("*", "qsys_Linked_tables_not_in_tsys_Link_Tables")
    If intNRecs > 0 Then
        DoCmd.SetWarnings False
        ' Run the append query to add databases not in tsys_Link_Dbs
        DoCmd.OpenQuery "qsys_Linked_dbs_not_in_tsys_Link_Dbs"
        ' Append missing table records to tsys_Link_Tables
'        strSQL = "INSERT INTO tsys_Link_Tables " & _
'            "( LinkTable, LinkDb ) " & _
'                       "( Link_table, Link_db,  Link_type ) " & _
'            "SELECT qsys_Linked_tables_not_in_tsys_Link_Tables.CurrTable, " & _
'            "qsys_Linked_tables_not_in_tsys_Link_Tables.CurrDb " & _
'            "FROM qsys_Linked_tables_not_in_tsys_Link_Tables;"
        strSQL = GetTemplate("i_tsys_link_tables")
        DoCmd.RunSQL strSQL
        DoCmd.SetWarnings True
        ' Update descriptions
'        Set rs = db.OpenRecordset("SELECT * FROM tsys_Link_Tables " & _
'            "WHERE tsys_Link_Tables.Description_text Is Null", dbOpenSnapshot)
        Set rs = db.OpenRecordset(GetTemplate("s_tsys_link_tables_no_description"), dbOpenSnapshot)
        Do Until rs.EOF
            strTable = rs![LinkTable]
            Set tdf = db.TableDefs(strTable)
            ' Update the table description in tsys_Link_Tables
            ' Set default description in case there is none
            strDesc = " - no description - "
            strDesc = tdf.Properties("Description") ' Throws trapped error 3270 if none
'            strSQL = "UPDATE tsys_Link_Tables " & _
'                "SET tsys_Link_Tables.DescriptionText=""" & strDesc & _
'                """ WHERE (((tsys_Link_Tables.LinkTable)=""" & strTable & """));"
            strSQL = GetTemplate("u_tsys_link_tables_description", "descr" & PARAM_SEPARATOR & strDesc & "|tbl" & PARAM_SEPARATOR & strTable)
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            rs.MoveNext
        Loop
        rs.Close
        ' Throw an error if there are still mismatched records
        If DCount("*", "qsys_Linked_tables_not_in_tsys_Link_Tables") > 0 Then
            blnHasError = True
            DoCmd.OpenQuery "qsys_Linked_tables_not_in_tsys_Link_Tables", , acReadOnly
        End If
    End If

    ' Look for linked db records without child table records
    intNRecs = DCount("*", "qsys_Linked_dbs_without_table_records")
    If intNRecs > 0 Then
        Set rs = db.OpenRecordset("qsys_Linked_dbs_without_table_records", _
            dbOpenSnapshot)
        Do Until rs.EOF
            ' Delete mismatched records from tsys_Link_Dbs
'            strSQL = "DELETE * FROM tsys_Link_Dbs WHERE ([LinkDb]=""" & _
'                rs![Link_db] & """);"
            strSQL = GetTemplate("d_tsys_link_tables_by_db", "link_db" & PARAM_SEPARATOR & rs![LinkDb])
            DoCmd.SetWarnings False
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
            rs.MoveNext
        Loop
        rs.Close
        ' Throw an error if there are still mismatched records
        If DCount("*", "qsys_Linked_dbs_without_table_records") > 0 Then
            blnHasError = True
            DoCmd.OpenQuery "qsys_Linked_dbs_without_table_records", , acReadOnly
        End If
    End If

    ' Look for records with mismatched db name, server, file path, or ODBC info
    intNRecs = DCount("*", "qsys_Linked_tables_mismatched_info")
    If intNRecs > 0 Then
        blnHasError = True
        DoCmd.OpenQuery "qsys_Linked_tables_mismatched_info"
    End If

    ' Warn the user if an error was found
    If blnHasError Then
        MsgBox "The application tables need to be updated with" & vbCrLf & _
            "correct information about the linked back-end" & vbCrLf & _
            "databases and tables before the application can" & vbCrLf & _
            "be used." & vbCrLf & vbCrLf & "Please contact the database administrator.", _
            vbCritical, "Application error (VerifyLinkTableInfo[fw_mod_Linked_Tables])"
    End If

    VerifyLinkTableInfo = Not blnHasError

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    rs.Close
    Set tdf = Nothing
    Set rs = Nothing
    Set db = Nothing
    PopCallStack "VerifyLinkTableInfo"
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3061, 3078, 3135  ' Problem with SQL syntax, or ref to nonexistent object, etc.
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinkTableInfo[fw_mod_Linked_Tables])"
      Case 3011, 7874   ' System table not found
         MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinkTableInfo[fw_mod_Linked_Tables])"
      Case 3265     ' Field name in the system table improperly specified
        MsgBox "Error #" & Err.Number & ":  System table field not found." & _
            vbCrLf & "Please notify the database administrator before using " & _
            "this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinkTableInfo[fw_mod_Linked_Tables])"
      Case 3270     ' Property not found (TableDefs description)
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinkTableInfo[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure
End Function

' =================================
' FUNCTION:     VerifyLinks
' Description:  Loops through all of the linked tables to verify valid links
' Parameters:   none
' Returns:      True (no link errors) or False
' Throws:       none
' References:   CheckLink
' Source/date:  John R. Boetsch, May 24, 2006
' Revisions:    JRB, 7/8/2009 - simplified recordset and error traps
'               JRB, 10/8/2009 - changed progress meter message
'               JRB, 12/30/2009 - updated to use the popup progress meter form
'               -------------------------------------------------------------------------
'               BLC, 4/30/2015 - moved to mod_Linked_Tables from mod_Custom_Functions & renamed VerifyLinks
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 6/5/2016  - adapted to Big Rivers Application, adjust to renamed tsys_Link_Tables fields
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' =================================
Public Function VerifyLinks() As Boolean
    On Error GoTo Err_Handler

    PushCallStack "VerifyLinks"
    
    Dim rs As DAO.Recordset
    Dim intNumTables As Integer
    Dim intI As Integer
    Dim varReturn As Variant
    Dim strLinkTableName As String
    Dim frm As Form             ' Reference to the progress popup form
    Dim strProgForm As String   ' Name of the progress popup form
    Dim strProgress As String   ' Progress bar string

    VerifyLinks = False  ' Default unless successful

    ' Set the recordset to the system table to show all linked tables except those
    '   that have recently been deleted (which have names starting with '~')
'    Set rs = CurrentDb.OpenRecordset("SELECT MSysObjects.Name, MSysObjects.Database " & _
'        "FROM MSysObjects " & _
'        "WHERE ((MSysObjects.Name) Not Like '~*') AND ((MSysObjects.Type) In (4,6)) " & _
'        "ORDER BY MSysObjects.Name;", dbOpenSnapshot)
    Set rs = CurrDb.OpenRecordset(GetTemplate("s_msysobjects_except_deleted"), dbOpenSnapshot)

    ' Counts the number of linked tables in the recordset
    rs.MoveLast    ' Need to do this to make the record count accurate
    intNumTables = rs.RecordCount

    ' Initialize the progress popup form
    strProgForm = "ProgressMeter"
    DoCmd.OpenForm strProgForm
    Set frm = Forms!ProgressMeter
    frm.Caption = " Verifying table connections"
    frm!txtPercent = 0
    ' Initialize the message below the progress meter
    frm!txtMsg = " ... Please wait ..."

    '   Initialize the system meter to indicate progress
    varReturn = SysCmd(acSysCmdInitMeter, "Verifying table connections", intNumTables)
    intI = 0
    rs.MoveFirst

    ' Loop through each record and check for bad links
    '   Send to error handler if a bad link is encountered
    Do Until rs.EOF
        intI = intI + 1
        ' Update the status bar progress meter
        varReturn = SysCmd(acSysCmdUpdateMeter, intI)
        ' Update the popup progress meter
        frm!txtPercent = Round(100 * intI / intNumTables)
        ' Update the progress bar in the progress popup with sequential "�" characters
        '   which look like a bar because of the font of the control (20 characters=100%)
        strProgress = String(Round(19 * intI / intNumTables), "�")
        frm!tbxProgress = strProgress
        frm.Repaint
        strLinkTableName = rs![Name]
        ' Make sure the linked table opens properly
        If CheckLink(strLinkTableName) = False Then
            ' Unable to open a linked table (not a critical error)
            MsgBox "Unable to open the following table:" & vbCrLf & vbCrLf & _
                strLinkTableName, vbExclamation, "Broken table links"
            GoTo Exit_Procedure
        Else
        ' Table link is valid
            rs.MoveNext
        End If
    Loop

    ' If no bad links were encountered
    VerifyLinks = True

Exit_Procedure:
    On Error Resume Next
    varReturn = SysCmd(acSysCmdRemoveMeter)
    DoCmd.Close acForm, strProgForm, acSaveNo
    Set frm = Nothing
    rs.Close
    Set rs = Nothing
    PopCallStack "VerifyLinks"
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3061, 3078, 3135  ' Problem with SQL syntax, or ref to nonexistent object, etc.
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinks[fw_mod_Linked_Tables])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyLinks[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     FixLinkedDatabase
' Description:  Populates the database path for the linked table until full database admin linking is in place
'               This fixes a situation where tbl_Link_Dbs is not updated when the Access Linked Table Manager
'               is used to update the location of linked tables
' Parameters:   strTableName - the name of a linked table
' Returns:      -
' Throws:       none
' References:   ParseConnectionStr
' Source/date:  BLC, 5/19/2015 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                 multiple open connections
' =================================
Public Sub FixLinkedDatabase(ByVal strTableName As String)
    On Error GoTo Err_Handler

    Dim strTemp As String, strSQL As String, strCurDb As String, strCurDbPath As String
    Dim rs As DAO.Recordset

    strTemp = ParseConnectionStr(CurrDb.TableDefs(strTableName).connect)
    
    'fetch current database location
    Set rs = CurrDb.OpenRecordset("qsys_Linked_tables_mismatched_info") '_dbs")
    
    If Not rs.EOF And rs.BOF Then
    
        rs.MoveLast
        
        'handle single db otherwise do it manually via tbl_Linked_Dbs?
        If rs.RecordCount = 1 Then
            strCurDb = rs("LinkDb") '_db")
            strCurDbPath = rs("CurrPath")
            
            'populate the current database in Link_Dbs
'            strSQL = "UPDATE tsys_Link_Dbs " & _
'                     "SET FilePath = '" & strCurDbPath & "' " & _
'                     "WHERE LinkDb = '" & strCurDb & "';"
            strSQL = GetTemplate("", "curDbPath" & PARAM_SEPARATOR & strCurDbPath & "|curDb" & PARAM_SEPARATOR & strCurDb)
        
            DoCmd.RunSQL (strSQL)
        End If
        
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FixLinkedDatabase[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          UpdateTSysTablesDb
' Description:  Update database value for a table w/in tsys_Link_Tables
' Assumptions:  Tables (tsys_Link_Tables) exist with fields as noted
'               Database file & path are valid.
' Parameters:   strNewDb - new database (e.g. "mynewdb.accdb")
'               strOrigDb - original database  (e.g. "mydb.accdb")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 3, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/3/2015 - initial version
' ---------------------------------
Public Sub UpdateTSysTablesDb(strNewDb As String, strOrigDb As String)
On Error GoTo Err_Handler
    
    Dim strSQL As String
        
    DoCmd.SetWarnings False
    
    'update tsys_Link_Tables
    strSQL = "UPDATE tsys_Link_Tables SET Link_db = '" & strNewDb & "' WHERE Link_db = '" & strOrigDb & "';"
    DoCmd.RunSQL (strSQL)
        
    DoCmd.SetWarnings True

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateTSysTablesDb[fw_mod_Linked_Tables])"
    End Select
    Resume Exit_Sub
End Sub