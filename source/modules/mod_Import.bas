Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Import
' Level:        Application module
' Version:      1.03
'
' Description:  field data import related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:    ML/GS - unknown   - 1.00 - initial version
'               BLC   - 5/21/2019 - 1.01 - added documentation, error handling, option explicit,
'                                          fixed CWD data import failure issue
'               BLC   - 8/31/2019 - 1.02 - fixed event ID tracking insert (GUID type conversion error)
'               BLC   - 9/15/2019 - 1.03 - populated recordset for accurate recordcount
' =================================

'---------------------
' Declarations
'---------------------
Public intImport2 As Integer

'---------------------
' Methods
'---------------------

'---------------------
' Functions
'---------------------

' ---------------------------------
' Function:     GetImportFile
' Description:  retrieve import file
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Public Function GetImportFile() As Variant
On Error GoTo Err_Handler

    Dim strProcName As String
    Dim varImportFileName As Variant
    Dim strFilter As String
    Dim lngFlags As Long
    
    strProcName = "GetImportFile"
        'Display the Open File dialog using the adhCommonFileOpenSave
        'function in the basCommonfile module
    strFilter = adhAddFilterItem( _
        strFilter, "Access (*.*db)", "*.*db")
        lngFlags = adhOFN_HIDEREADONLY Or _
        adhOFN_HIDEREADONLY Or adhOFN_NOCHANGEDIR
       
        varImportFileName = adhCommonFileOpenSave( _
            OpenFile:=True, _
            Filter:=strFilter, _
            flags:=lngFlags, _
            DialogTitle:="Locate data file")
        
        If IsNull(varImportFileName) Then
            'user pressed Cancel
            GetImportFile = Null
            GoTo Exit_Handler
        Else
            GetImportFile = adhTrimNull(varImportFileName)
        End If
    
Exit_Handler:
    On Error GoTo 0 'in original code
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetImportFile[mod_Import])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     VerifyImportFile
' Description:  verify table exists in import file
' Assumptions:
' Parameters:   strSQL - SQL specifying tables to verify (string)
'               varFileName - file name (variant)
' Returns:      boolean (True - if files are found, False - if not)
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Mark A. Wotawa, 2/8/2000
'               Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   MAW - 2/8/2000 - verify tables prior to link to a back-end dataset
'                    modified to check that tables existed prior to attempting to import them.
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Public Function VerifyImportFile(strSQL As String, _
    varFileName As Variant) As Boolean
On Error GoTo Err_Handler
    
    Dim dbGet As Database
    Dim db As Database
    Dim rst As DAO.Recordset
    Dim intNumTables As Integer
    Dim varReturn As Variant
    Dim intI As Integer
    Dim strImportTableName As String
    Dim strProcName As String
    Dim VerifyLinkFile As Boolean
    
    strProcName = "VerifyImportFile"
    
    VerifyImportFile = True
    
    'Check to see if selected database file is valid
    On Error Resume Next
    Set dbGet = DBEngine.OpenDatabase(varFileName)
    If Err <> 0 Then
        VerifyLinkFile = False
        GoTo Exit_Handler
    Else
        On Error GoTo 0
        Set db = CurrentDb
        Set rst = db.OpenRecordset(strSQL, dbOpenForwardOnly)
        intNumTables = rst.RecordCount
        varReturn = SysCmd(acSysCmdInitMeter, "Verifying tables", _
            intNumTables)
        intI = 0
        
        Do Until rst.EOF
            strImportTableName = rst![Link_table] '********
            On Error Resume Next
            varReturn = dbGet.TableDefs(strImportTableName).Fields(0).Name
            If Err <> 0 Then
                'MsgBox "selected file is invalid", vbCritical, "Link Table Error"
                VerifyImportFile = False
                GoTo Exit_Handler
            End If
            varReturn = SysCmd(acSysCmdUpdateMeter, intI + 1)
            rst.MoveNext
        Loop
    End If
    varReturn = SysCmd(acSysCmdRemoveMeter)
    
Exit_Handler:
'original code
    On Error Resume Next
    DoCmd.Hourglass False
    varReturn = SysCmd(acSysCmdRemoveMeter)
    On Error GoTo 0
'original code
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VerifyImportFile[mod_Import])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     AppendToTable
' Description:  append data to table
' Assumptions:
' Parameters:   rsAppend - records to append (DAO.Recordset)
'               rsMain - existing master records (DAO.Recordset)
'               strAppendTableName - name of table being appended to (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
'   BLC  - 9/3/2019 - revised to check for BOF/EOF before recordcount
'                     moved CleanUp: to Exit_Handler, added delete for pre-existing Qry_AppendRecs
' ---------------------------------
Public Function AppendtoTable(rsAppend As DAO.Recordset, rsMain As DAO.Recordset, strAppendTableName As String)
On Error GoTo Err_Handler

     Dim db As DAO.Database
     Set db = CurrentDb
     Dim rsAppendLog As DAO.Recordset
     
     Dim strMain As String
     Dim strAppend As String
     
     strMain = rsMain.Name
     strAppend = rsAppend.Name
     
     Dim qrydefAppend As QueryDef
     Dim strQryNewRecs As String
     Dim qryDefUpdate As QueryDef
     
    '------------------------------
    ' Fetch New Records
    '------------------------------
    'Select only the new records from the Append recordset.
            
    strQryNewRecs = "SELECT [" & strAppend & "].*" _
        & " FROM [" & strAppend & "] LEFT JOIN " & strMain & " ON " _
        & "[" & strAppend & "]." & rsAppend.Fields(0).Name & " = " & strMain & "." _
            & rsMain.Fields(0).Name _
        & " WHERE (((" & strMain & "." & rsMain.Fields(0).Name & ") IS NULL));"
        
Debug.Print strQryNewRecs

    'ensure Qry_AppendRecs was deleted
    Dim qry As QueryDef
    For Each qry In CurrentDb.QueryDefs
        If qry.Name = "Qry_AppendRecs" Then
            DoCmd.DeleteObject acQuery, "Qry_AppendRecs"
            Exit For
        End If
    Next

    'set the query definition for the new rec SQL
    'NOTE: Should I add delete qry if existing with a warning here?
    Set qrydefAppend = db.CreateQueryDef("Qry_AppendRecs", strQryNewRecs)
        
    'Define new records from Append data set found by query
    Dim rsNewRecs As DAO.Recordset
    Set rsNewRecs = db.OpenRecordset("Qry_AppendRecs")
    
    '------------------------------
    ' Populate the recordset
    '------------------------------
    If Not (rsNewRecs.BOF And rsNewRecs.EOF) Then
        rsNewRecs.MoveLast
        rsNewRecs.MoveFirst
    End If
        
Debug.Print strMain & " # records:" & rsNewRecs.RecordCount

    'No new records? --> notify user & exit function
    If Not rsNewRecs.RecordCount > 0 Then
        MsgBox "No new records to append in " & strMain & ".", , "Append Records"
        GoTo Exit_Handler
    End If
    
    'Append data to Master Data set table
    Dim strAppendSQL As String
    Dim strResponse As String
     
    strAppendSQL = "INSERT INTO " & strMain _
                & " SELECT [" & rsNewRecs.Name & "].*" _
                & " FROM [" & rsNewRecs.Name & "];"
Debug.Print strAppendSQL
    
    strResponse = MsgBox("You are about to APPEND " & rsNewRecs.RecordCount & " records to " & strMain & ". " _
                & "Are you sure you would like to proceed?", vbYesNo, "Append New Records")
                
    If strResponse = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL (strAppendSQL)
    Else
        MsgBox "No new records appended.", , "Append New Records"
        GoTo Exit_Handler
    End If
    DoCmd.SetWarnings True

    '------------------------------
    ' Populate the Append log table
    '------------------------------
    Dim strDate As String
    strDate = Date
    'get info about master & appending data sets
    Set rsAppendLog = db.OpenRecordset("tsys_Append_Log")
    rsAppendLog.AddNew
    rsAppendLog![Table_Name] = rsMain.Name
    rsAppendLog![Append_Date] = strDate
    rsAppendLog![Append_Table_Name] = strAppendTableName
    rsAppendLog![Record_ID] = rsAppend.Fields(0)
        
    'get info on number of records actually being imported
    rsAppendLog![Append_Records] = rsNewRecs.RecordCount
    
    'update the append log
    rsAppendLog.Update
    
Exit_Handler:
    'Clean up the variables
    DoCmd.DeleteObject acQuery, qrydefAppend.Name
    
    Set db = Nothing
    Set rsAppend = Nothing
    Set rsMain = Nothing
    Set rsNewRecs = Nothing
    Set qrydefAppend = Nothing
    Set rsAppendLog = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AppendToTable[mod_Import])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     UpdateTags
' Description:  update tag data
' Assumptions:
' Parameters:   rsMain - existing master data records (DAO.Recordset)
'               rsAppend - records to append (DAO.Recordset)
'               rsEvents - event records (DAO.Recordset)
'               strAppendTableName - name of table being appended to (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Public Function UpdateTags(rsMain As DAO.Recordset, rsAppend As DAO.Recordset, rsEvents As DAO.Recordset, strAppendTableName As String)
On Error GoTo Err_Handler

    Dim db As Database
    Set db = CurrentDb
    
    Dim strMain As String
    Dim strAppend As String
    Dim strEvents As String
    
    Dim strSQLSelectNew As String
    strMain = rsMain.Name
    strAppend = rsAppend.Name
    strEvents = rsEvents.Name
    
    Dim strIDAppend As String
    strIDAppend = rsAppend.Fields(0).Name
    Dim strIDMain As String
    strIDMain = rsMain.Fields(0).Name
    
    If fxnQueryExists("_qry_FindTreesSap") Then
        db.QueryDefs.Delete ("_qry_FindTreesSap")
    End If
    
    
    'Find all of the trees or saplings from the imported events.
    
        Dim strSQLFindTreesSap
        
        strSQLFindTreesSap = "SELECT [" & strAppend & "].* " _
                            & "FROM ([" & strEvents & "] LEFT JOIN tbl_Locations ON [" & strEvents & "].Location_ID = " _
                            & "tbl_Locations.Location_ID) LEFT JOIN [" & strAppend & "] ON tbl_Locations.Location_ID = " _
                            & "[" & strAppend & "].Location_ID " _
                            & "WHERE ((([" & strAppend & "]." & rsAppend.Fields(0).Name & ") IS NOT NULL));"
        
    'Create a dataset of all of the trees or saplings from the imported events.
        
        Dim qdef_FindTrees_Sap As QueryDef
        Set qdef_FindTrees_Sap = db.CreateQueryDef("_qry_FindTreesSap", strSQLFindTreesSap)
        
           
    '*******************************************************************************************************************************
    'UPDATE EXISTING TAG RECORDS
    
    'select only the records that need to be updated.
    'Dim strIDAppend As String
    'strIDAppend = rsAppend.Fields(0).Name
    'Dim strIDMain As String
    'strIDMain = rsMain.Fields(0).Name
    
    'Reset the Append recordset to equal only those records that were collected during this event.
    Dim rsUpdate As DAO.Recordset
    Dim strUpdate As String
    
    Set rsUpdate = db.OpenRecordset("_qry_FindTreesSap")
    strUpdate = rsUpdate.Name
    
    'Select any tree or sapling records that have updated data
    strSQLSelectNew = "SELECT [" & strUpdate & "]." & strIDAppend & " , [" & strUpdate & "].TSN, [" & strUpdate & "].Tag, [" & strMain & "].Tag, " _
                    & "[" & strUpdate & "].Tag_Notes, [" & strUpdate & "].Stop_Date, [" & strUpdate & "].Tag_Status, " _
                    & "[" & strMain & "].TSN, [" & strMain & "].Tag_Notes, [" & strMain & "].Stop_Date, [" & strMain & "].Tag_Status, " _
                    & "[" & strMain & "].Azimuth, [" & strMain & "].Distance, [" & strUpdate & "].Azimuth,[" & strUpdate & "].Distance " _
                    & "FROM " & strMain & " INNER JOIN [" & strUpdate & "] " _
                    & "ON [" & strMain & "]." & strIDMain & " = [" & strUpdate & "]." & strIDAppend & ";"
    
    'save the sql statement as a query temporarily so that it can be used in the next SQl statement
    
    If fxnQueryExists("_qry_SelectUpdatedTreeSap") Then
        db.QueryDefs.Delete ("_qry_SelectUpdatedTreeSap")
    End If
    
    Dim qdefSelectUpdated As QueryDef
    Set qdefSelectUpdated = db.CreateQueryDef("_qry_SelectUpdatedTreeSap", strSQLSelectNew)
    
    Dim strqdef As String
    strqdef = qdefSelectUpdated.Name
    
    'Create a recordset and name the query that looks for tree/sapling records that need updating
    Dim rsSelectUpdated As DAO.Recordset
    Set rsSelectUpdated = db.OpenRecordset(strqdef)
    
    'check to see if there are any tree or sapling records that need updating.  If not then clean up the variables and exit the function
    If Not rsSelectUpdated.RecordCount > 0 Then
        MsgBox "No records to update in " & strMain & ".", , "Update Records"
        GoTo Append:
    End If
    
    'if there are records that need updating then:
    'populate the recordset
    rsSelectUpdated.MoveLast
    
    Dim iUpdateCount As Long 'Integer
    iUpdateCount = rsSelectUpdated.RecordCount
    Dim strResponse As String
    
    Dim strSQLUpdate As String
    
    'Create the update query
    
    strSQLUpdate = "UPDATE " & strqdef & " INNER JOIN " & strMain _
                & " ON [" & strqdef & "]." & strIDAppend & " = [" & strMain & "]." & strIDMain & " " _
                & "SET [" & strMain & "].TSN = [" & strqdef & "].[" & strAppend & ".TSN], " _
                & "[" & strMain & "].Tag = [" & strqdef & "].[" & strAppend & ".Tag], " _
                & "[" & strMain & "].Azimuth = [" & strqdef & "].[" & strAppend & ".Azimuth], " _
                & "[" & strMain & "].Distance = [" & strqdef & "].[" & strAppend & ".Distance], " _
                & "[" & strMain & "].Tag_Notes = [" & strqdef & "].[" & strAppend & ".Tag_Notes], " _
                & "[" & strMain & "].Stop_Date = [" & strqdef & "].[" & strAppend & ".Stop_Date], " _
                & "[" & strMain & "].Tag_Status = [" & strqdef & "].[" & strAppend & ".Tag_Status];"
    'run the update query
    
    strResponse = MsgBox("You are about to UPDATE " & iUpdateCount & " records in " & strMain & "." _
                & " Are you sure you would like to continue?", vbYesNo, "Update Records.")
            
            'MsgBox strSQLUpdate
            
            If strResponse = vbYes Then
                DoCmd.SetWarnings False
                DoCmd.RunSQL (strSQLUpdate)
            Else
                MsgBox "No records updated.", vbOKOnly, "Update Records."
                GoTo Append:
            End If
            
            DoCmd.SetWarnings True
            
    'declare the variables needed to create an update log entry
    Dim rsUpdateLog As DAO.Recordset
    Set rsUpdateLog = db.OpenRecordset("tsys_Update_Log")
    
    Dim dDate As String
    dDate = Date
    
    'Update the update log table
    With rsUpdateLog
        .AddNew
        .Fields(0) = strMain
        .Fields(1) = strAppendTableName
        .Fields(2) = dDate
        .Fields(3) = iUpdateCount
        .Update
    End With
    
    '***********************************************************************************************************************************
    'APPEND NEW TREE AND SAPLING RECORDS
    
Append:
    
        If rsUpdate.RecordCount > 0 Then
 
            AppendtoTable rsUpdate, rsMain, strAppendTableName
           
        Else
            GoTo CleanUp
        End If
        
CleanUp:
    
    'delete the select query
    
    DoCmd.DeleteObject acQuery, "_qry_SelectUpdatedTreeSap"
    DoCmd.DeleteObject acQuery, "_qry_FindTreesSap"
    
    'clean up the variables.
    Set db = Nothing
    
    'Set qdefSelectNew = Nothing
    Set qdefSelectUpdated = Nothing
    Set rsUpdateLog = Nothing
    Set rsSelectUpdated = Nothing
    'Set rsAppendNewTrees_Sap = Nothing
    'Set qdef_FindNewTrees_Sap = Nothing
    Set qdef_FindTrees_Sap = Nothing
    Set rsUpdate = Nothing

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateTags[mod_Import])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     UpdateEventID
' Description:  update the event ID of the import records
' Assumptions:
' Parameters:   rsAppend - records being appended (DAO.Recordset)
'
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling, revised SQL to provide spacing to address CWD import issue
'   BLC  - 8/31/2019 - updated to use SQL vs DAO for insert to avoid GUID type conversion error 3421
'   BLC  - 9/15/2019 - populated recordset before checking recordcount
' ---------------------------------
Public Function UpdateEventID(rsAppend As DAO.Recordset, GUIDMain As Variant, GUIDReplace As Variant, strTableName As String)
On Error GoTo Err_Handler

     Dim db As Database
     Set db = CurrentDb
     
     Dim strAppend As String
     strAppend = rsAppend.Name
      
    'Create the update query to replace the old Event_ID
       
    Dim strUpdateEventSQL As String
    
    Dim strFind As String
    
    'convert the GUID to a string to insert into the search criteria
    strFind = GUIDReplace
    
    'Create the update query that will replace the wrong EventID with the correct one.
        
'    strUpdateEventSQL = "UPDATE [" & strAppend & "] " _
'                        & "SET [" & strAppend & "].Event_ID = " _
'                        & GUIDMain _
'                        & " WHERE ((([" & strAppend & "].Event_ID)= " & strFind & "));"

    strUpdateEventSQL = "UPDATE [" & strAppend & "] " _
                        & "SET [" & strAppend & "].Event_ID = '" _
                        & GUIDMain _
                        & "' WHERE ((([" & strAppend & "].Event_ID)= '" & GUIDReplace & "'));"

Debug.Print "strUpdateEventSQL: " & strUpdateEventSQL
    'Run the update query
    'rsAppend.MoveFirst
'      rsAppend.MoveLast
    '------------------------------
    ' Populate the recordset
    '------------------------------
    If Not (rsAppend.BOF And rsAppend.EOF) Then
        rsAppend.MoveLast
        rsAppend.MoveFirst
    End If
       
       Dim strResponse As String
       strResponse = MsgBox("You are about to UPDATE " & rsAppend.RecordCount & " records in " & strTableName & ". " _
                        & "Are you sure you would like to proceed?", vbYesNo, "Update Records")
        If strResponse = vbYes Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL (strUpdateEventSQL)
        Else
            MsgBox "No records updated.", , "Update Records"
            GoTo CleanUp:
        End If
        
            DoCmd.SetWarnings True
                    
       'Create a recordset from the SQL statement to calculate the number of records where
       'an EventID was updated
       
       'Dim qdefUpdate As QueryDef
       'Set qdefUpdate = db.CreateQueryDef("_qry_UpdateEvent", strUpdateEventSQL)
       'Dim rsUpdate As DAO.Recordset
       'Set rsUpdate = db.OpenRecordset("_qry_UpdateEvent")
       'Dim iCount As Integer
       'Dim strDate As String
       
       'populate the recordset
       
       'rsUpdate.MoveLast
       'rsUpdate.MoveFirst
       'iCount = rsUpdate.RecordCount
       Dim strDate As Date
       strDate = Date
         
'        Dim strGUIDMain As String
'        Dim strGUIDReplace As String
'
'        strGUIDMain = GUIDMain
'        strGUIDReplace = GUIDReplace
        
Debug.Print "GUIDMain: " & GUIDMain
Debug.Print "GUIDReplace: " & GUIDReplace
        
'Debug.Print "strGUIDMain: " & strGUIDMain
'Debug.Print "strGUIDReplace: " & strGUIDReplace

        
       'Update the tracking table with new and old event ids as well as the number of records
       'updated
              
'        Dim rsTrackEvent As DAO.Recordset
'        Set rsTrackEvent = db.OpenRecordset("tsys_xref_Event_Update_Tracker")
'            With rsTrackEvent
'                .AddNew
'                    .Fields(1) = strGUIDMain 'CStr(IIf(FormatIsGUID(GUIDMain), GUIDMain, String2GUID(GUIDMain))) 'GUIDFromString(GUIDMain) 'GUIDMain
'                    .Fields(2) = strGUIDReplace 'CStr(String2GUID(CStr(GUIDReplace))) 'GUIDFromString(GUIDReplace) 'GUIDReplace
'                    .Fields(3) = strTableName
'                    .Fields(4) = rsAppend.RecordCount
'                    .Fields(5) = strDate
'                .Update
'            End With
                
        'insert using SQL
        Dim strTrackEventInsert As String
        strTrackEventInsert = "INSERT INTO tsys_xref_Event_Update_Tracker " _
            & "(Event_ID, Import_Event_ID, AppendTableName, Record_Count, [Date]) " _
            & " SELECT '" & GUIDMain & "', '" & GUIDReplace & "', '" _
            & strTableName & "', " & rsAppend.RecordCount & ", " _
            & "Now()"
          
        Debug.Print strTrackEventInsert
            
        DoCmd.SetWarnings False
        DoCmd.RunSQL strTrackEventInsert
        DoCmd.SetWarnings True
        
    'Clean up the variables
CleanUp:
    
    'Set rsUpdate = Nothing
'    Set rsTrackEvent = Nothing
    'Set qdefUpdate = Nothing
    Set db = Nothing

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case 3421 'Data type conversion error - triggered for GUIDs
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateEventID[mod_Import])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     fxnUpdateLocInfo
' Description:  update location info
' Assumptions:
' Parameters:   rsLoc - location records (DAO.Recordset)
'               strLocTblName - location table name (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Public Function fxnUpdateLocInfo(rsLoc As DAO.Recordset, strLocTblName As String)
On Error GoTo Err_Handler

    Dim db As Database
    
    Set db = CurrentDb
    
    Dim strSQLSelect As String
    Dim strSQLUpdate As String
    
    'select the records that need updating
    
    strSQLSelect = "SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, tbl_Locations.Slope, tbl_Locations.Aspect, [" & strLocTblName & "].Slope, " _
        & "[" & strLocTblName & "].Aspect " _
        & "FROM " & strLocTblName & " INNER JOIN tbl_Locations ON [" & strLocTblName & "].Location_ID = tbl_Locations.Location_ID " _
        & "WHERE (((tbl_Locations.Slope) Is Null) AND ((tbl_Locations.Aspect) Is Null) AND (([" & strLocTblName & "].Slope) Is Not Null) " _
        & "AND (([" & strLocTblName & "].Aspect) Is Not Null)) OR (((tbl_Locations.Slope) Is Null) AND (([" & strLocTblName & "].Slope) Is Not Null)) " _
        & "OR (((tbl_Locations.Aspect) Is Null) AND (([" & strLocTblName & "].Aspect) Is Not Null));"
    
    
    strSQLUpdate = "UPDATE " & strLocTblName & " INNER JOIN tbl_Locations ON [" & strLocTblName & "].Location_ID = tbl_Locations.Location_ID " _
        & "SET [tbl_Locations].[Slope] = [" & strLocTblName & "].Slope, [tbl_Locations].[Aspect] = [" & strLocTblName & "].Aspect " _
        & "WHERE (((tbl_Locations.Slope) Is Null) AND ((tbl_Locations.Aspect) Is Null) AND (([" & strLocTblName & "].Slope) Is Not Null) AND (([" & strLocTblName & "].Aspect) Is Not Null)) " _
        & "OR (((tbl_Locations.Slope) Is Null) AND (([" & strLocTblName & "].Slope) Is Not Null)) OR (((tbl_Locations.Aspect) Is Null) AND (([" & strLocTblName & "].Aspect) Is Not Null));"
    
    
    If fxnQueryExists("_qUpdate_Locs") Then
        db.QueryDefs.Delete ("_qUpdate_Locs")
    End If
    
    
    Dim qdefLocUpdates As QueryDef
    Set qdefLocUpdates = db.CreateQueryDef("_qUpdate_Locs", strSQLSelect)
    Dim rsUpdateLocs As DAO.Recordset
    Set rsUpdateLocs = db.OpenRecordset(qdefLocUpdates.Name)
    Dim strResponse As String
    
            
    
    If rsUpdateLocs.RecordCount < 1 Then
        MsgBox "No location information to update", , "NCRN Forest Vegetation Monitoring"
        GoTo CleanUp
        
    Else
        rsUpdateLocs.MoveLast
            rsUpdateLocs.MoveFirst
        strResponse = MsgBox("You are about to update " & rsUpdateLocs.RecordCount & " location records. Do you wish to proceed?", vbYesNo, "NCRN Forest Vegetation Monitoring")
        
        If strResponse = vbYes Then
            DoCmd.SetWarnings False
            DoCmd.RunSQL (strSQLUpdate)
        Else
            MsgBox "No records updated.", , "Update Records"
            GoTo CleanUp:
        End If
        
            DoCmd.SetWarnings True
    End If
    
    
CleanUp:
    Set db = Nothing
    Set qdefLocUpdates = Nothing
    Set rsUpdateLocs = Nothing

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateLocInfo[mod_Import])"
    End Select
    Resume Exit_Handler
End Function