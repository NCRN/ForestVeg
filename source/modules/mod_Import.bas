Option Compare Database

Public intImport2 As Integer


Public Function GetImportFile() As Variant

On Error GoTo Err_GetImportFile

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
        Flags:=lngFlags, _
        DialogTitle:="Locate data file")
    
    If IsNull(varImportFileName) Then
        'user pressed Cancel
        GetImportFile = Null
        GoTo Exit_GetImportFile
    Else
        GetImportFile = adhTrimNull(varImportFileName)
    End If
    
Exit_GetImportFile:
    On Error GoTo 0
    Exit Function
    
Err_GetImportFile:
    Select Case Err
    Case Else
        MsgBox "Error#" & Err.Number & ": " & Err.Description, _
            vbOKOnly + vbCritical, strProcName
    End Select
    Resume Exit_GetImportFile
    
End Function

Public Function VerifyImportFile(strSQL As String, _
    varFileName As Variant) As Boolean
    
    'Purpose:
    '   Verify table exists in the a file to import
    'In:
    '   strSQL: SQL string specifying tables to verify
    'Out:
    '   Return value: True/False
    'History:
    '   Created 02/08/2000 Mark A. Wotawa to verify the
        'prior to link to a back-end dataset.
    'Modified to check that tables existed prior to attempting
        'to import them.
    
    
    On Error GoTo Err_VerifyImportFile
    
    Dim dbGet As Database
    Dim db As Database
    Dim rst As DAO.Recordset
    Dim intNumTables As Integer
    Dim varReturn As Variant
    Dim intI As Integer
    Dim strImportTableName As String
    Dim strProcName As String
    
    strProcName = "VerifyImportFile"
    
    VerifyImportFile = True
    
    'Check to see if selected database file is valid
    On Error Resume Next
    Set dbGet = DBEngine.OpenDatabase(varFileName)
    If Err <> 0 Then
        VerifyLinkFile = False
        GoTo Exit_VerifyImportFile
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
                GoTo Exit_VerifyImportFile
            End If
            varReturn = SysCmd(acSysCmdUpdateMeter, intI + 1)
            rst.MoveNext
        Loop
    End If
    varReturn = SysCmd(acSysCmdRemoveMeter)
    
Exit_VerifyImportFile:
    On Error Resume Next
    DoCmd.Hourglass False
    varReturn = SysCmd(acSysCmdRemoveMeter)
    On Error GoTo 0
    Exit Function
        
Err_VerifyImportFile:
    Select Case Err
    Case Else
        MsgBox "Error#" & Err.Number & ": " & Err.Description, _
            vbOKOnly + vbCritical, strProcName
    End Select
    Resume Exit_VerifyImportFile
End Function

Public Function AppendtoTable(rsAppend As DAO.Recordset, rsMain As DAO.Recordset, strAppendTableName As String)

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
 
  
'Select only the new records from the Append recordset.
        
        strQryNewRecs = "SELECT [" & strAppend & "].*" _
            & " FROM [" & strAppend & "] LEFT JOIN " & strMain & " ON " _
            & "[" & strAppend & "]." & rsAppend.Fields(0).Name & " = " & strMain & "." _
                & rsMain.Fields(0).Name _
            & " WHERE (((" & strMain & "." & rsMain.Fields(0).Name & ") is null));"
    
  
'set the query definition for the new rec SQL
'NOTE: Should I add delete qry if existing with a warning here?
Set qrydefAppend = db.CreateQueryDef("Qry_AppendRecs", strQryNewRecs)
    
'Define the new records from the Append data set that were found by the query

Dim rsNewRecs As DAO.Recordset
Set rsNewRecs = db.OpenRecordset("Qry_AppendRecs")

'If there are no new records then notify the user and exit the function.

If Not rsNewRecs.RecordCount > 0 Then
    MsgBox "No new records to append in " & strMain & ".", , "Append Records"
    GoTo CleanUp:
End If

'Populate the recordset
rsNewRecs.MoveLast
rsNewRecs.MoveFirst


'Append the data to the Master Data set table

 Dim strAppendSQL As String
 Dim strResponse As String
 
   strAppendSQL = "INSERT INTO " & strMain _
                & " SELECT [" & rsNewRecs.Name & "].*" _
                & " FROM [" & rsNewRecs.Name & "];"
                
 strResponse = MsgBox("You are about to APPEND " & rsNewRecs.RecordCount & " records to " & strMain & ". " _
                & "Are you sure you would like to proceed?", vbYesNo, "Append New Records")
                
    If strResponse = vbYes Then
        DoCmd.SetWarnings False
        DoCmd.RunSQL (strAppendSQL)
    Else
        MsgBox "No new records appended.", , "Append New Records"
        GoTo CleanUp:
    End If
    DoCmd.SetWarnings True
   
'Populate the Append log table

Dim strDate As String
strDate = Date
'get the info about the master and appending data sets
Set rsAppendLog = db.OpenRecordset("tsys_Append_Log")
rsAppendLog.AddNew
rsAppendLog![Table_Name] = rsMain.Name
rsAppendLog![Append_Date] = strDate
rsAppendLog![Append_Table_Name] = strAppendTableName
rsAppendLog![Record_ID] = rsAppend.Fields(0)
    
    'get the info on the number of records that are actually being imported.
       
    rsAppendLog![Append_Records] = rsNewRecs.RecordCount
    
   'update the append log
rsAppendLog.Update
   
'Clean up the variables

CleanUp:

DoCmd.DeleteObject acQuery, qrydefAppend.Name

Set db = Nothing
Set rsAppend = Nothing
Set rsMain = Nothing
Set rsNewRecs = Nothing
Set qrydefAppend = Nothing
Set rsAppendLog = Nothing

End Function

Public Function UpdateTags(rsMain As DAO.Recordset, rsAppend As DAO.Recordset, rsEvents As DAO.Recordset, strAppendTableName As String)

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
    
    strSQLFindTreesSap = "SELECT [" & strAppend & "].*" _
                        & "FROM ([" & strEvents & "] LEFT JOIN tbl_Locations ON [" & strEvents & "].Location_ID = " _
                        & "tbl_Locations.Location_ID) LEFT JOIN [" & strAppend & "] ON tbl_Locations.Location_ID = " _
                        & "[" & strAppend & "].Location_ID " _
                        & "WHERE ((([" & strAppend & "]." & rsAppend.Fields(0).Name & ") Is Not Null));"
    
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

Dim iUpdateCount As Integer
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

Set qdefSelectNew = Nothing
Set rsUpdateLog = Nothing
Set rsSelectUpdated = Nothing
Set rsAppendNewTrees_Sap = Nothing
Set qdef_FindNewTrees_Sap = Nothing
Set rsUpdate = Nothing


End Function
Public Function UpdateEventID(rsAppend As DAO.Recordset, GUIDMain As Variant, GUIDReplace As Variant, strTableName As String)
 
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
    
strUpdateEventSQL = "UPDATE [" & strAppend & "] " _
                    & "SET [" & strAppend & "].Event_ID = " _
                    & GUIDMain _
                    & "WHERE ((([" & strAppend & "].Event_ID)= " & strFind & "));"

'Run the update query
'rsAppend.MoveFirst
  rsAppend.MoveLast
   
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
   strDate = Date
     
   'Update the tracking table with new and old event ids as well as the number of records
   'updated
   
    Dim rsTrackEvent As DAO.Recordset
    Set rsTrackEvent = db.OpenRecordset("tsys_xref_Event_Update_Tracker")
        With rsTrackEvent
            .AddNew
                .Fields(1) = GUIDMain
                .Fields(2) = GUIDReplace
                .Fields(3) = strTableName
                .Fields(4) = rsAppend.RecordCount
                .Fields(5) = strDate
            .Update
        End With
        

'Clean up the variables
CleanUp:

Set rsUpdate = Nothing
Set rsTrackEvent = Nothing
Set qdefUpdate = Nothing
Set db = Nothing

End Function

Public Function fxnUpdateLocInfo(rsLoc As DAO.Recordset, strLocTblName As String)
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

End Function