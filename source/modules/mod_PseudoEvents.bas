Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_PseudoEvents
' Level:        Application module
' Version:      1.02
'
' Description:  Application PseudoEvent related functions & subroutines
'
' Source/date:  Bonnie Campbell, September 19, 2019
' Revisions:    BLC, 9/19/2019  - 1.00 - initial version
'               BLC, 9/23/2019  - 1.01 - save original event ID prior to iterating through tables during event ID update
'                                        add DeleteRelatedPseudoEventRecords
'               BLC, 9/24/2019  - 1.02 - add tbl_Events full archive before pseudoevents delete (ArchivePsuedoEvents)
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
' -- Constants --

' -- Values --

' -- Functions --


' ---------------------------------
'  Methods
' ---------------------------------

' *********************************
'    Common
' *********************************

' ---------------------------------
' SUB:          UpdatePseudoEventIDs
' Description:  Updates event IDs for pseudo-events
' Assumptions:  Event table name is either tbl_Events or an import of tbl_Events from primary or secondary tablet
'               tsys_Append_Tables is properly configured to identify which tables have an Event_ID that
'               should be updated at the same time (to keep data linked) by setting their HasEventID = 1
'               Current tables that have HasEventID = 1 in tsys_Append_Tables:
'                   tbl_Events          tbl_CWD_Data    tbl_Plot_Floor_Condition_Data   tbl_Quadrat_Data
'                   tbl_Sapling_Data    tbl_Tree_Data   xref_Event_Contacts
'               These tables are iterated through to update Event_ID records in each based on the newly generated eid
'
' Parameters:   EventTable - name of event table to update (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 19, 2019
' Adapted:      -
' Revisions:
'   BLC - 9/19/2019 - initial version
'   BLC - 9/23/2019 - save original event ID prior to iterating through tables
' ---------------------------------
Public Sub UpdatePseudoEventIDs(EventTable As String)
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    Dim rsEventTables As DAO.Recordset
    Dim sql As String
    Dim EventIDTable As String
    Dim prefix As String
    Dim suffix As String
    
    'determine the prefix & suffix of imported tables
    suffix = Replace(EventTable, "_tbl_Events", "")
    prefix = Replace(Replace(EventTable, suffix, ""), "tbl_Events", "")
    
    'archive pseudo events from tbl_Events
    ArchivePseudoEvents
    
    'retrieve EventTable's pseudoevents
    sql = "SELECT Event_ID FROM " & EventTable & " WHERE PseudoEvent = 1;"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    'iterate through pseudo-events and update related eventIDs
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    If rs.RecordCount > 0 Then
        
        'ensure we should update pseudoevent Event_IDs
        Dim rslt As Integer
        rslt = MsgBox("Update the " & rs.RecordCount & " PseudoEvent Event_IDs in " & EventTable & "?" _
                & vbCrLf & vbCrLf & "Choose YES >> " & rs.RecordCount & " Event IDs will be updated and their original & new Event IDs will be logged" _
                & " in the tsys_EventID_Update_History table along with who made the update & when." _
                & vbCrLf & vbCrLf & "Choose NO >> Nothing happens and you return to the append dialog.", _
                vbYesNo + vbQuestion, "Continue with PseudoEvent EventID Update?")
        
        If rslt = vbYes Then
        
            'iterate through the EventTable's pseudoevent Event_IDs
            Do While Not rs.EOF
            
                'tables w/ Event_IDs (tbl_Events first)
                sql = "SELECT Table_Name FROM tsys_Append_Tables WHERE HasEventID = 1 ORDER BY Append_Order;"
                Set rsEventTables = CurrentDb.OpenRecordset(sql)
                
                If Not (rsEventTables.BOF And rsEventTables.EOF) Then
                
                    Dim tbl As String
                    Dim eid As String
                    Dim oeid As String
                    
                    'store original EventID before update
                    oeid = rs("Event_ID")
                    
                    Do While Not rsEventTables.EOF
                            
                        DoCmd.Hourglass True
                                                    
                        'only generate eid for tbl_Events
                        If rsEventTables("Table_Name") = "tbl_Events" Then
                            'get new EventID
                            eid = fxnGUIDGen
                        End If
                        
                        'determine table name
                        EventIDTable = prefix & rsEventTables("Table_Name") & suffix
                        
                        'ensure table exists
                        If TableExists(EventIDTable) Then
                        
                            'processing
                            Application.SysCmd acSysCmdSetStatus, "Processing " & EventIDTable & " pseudoevents..."
                        
                            'update related records
                            'sql = "UPDATE " & rsEventTables("Table_Name") & " SET Event_ID = '" & eid & "' WHERE Event_ID = '" & rs("Event_ID") & "';"
                            'sql = "UPDATE " & EventIDTable & " SET Event_ID = '" & eid & "' WHERE Event_ID = '" & rs("Event_ID") & "';"
                            sql = "UPDATE " & EventIDTable & " SET Event_ID = '" & eid & "' WHERE Event_ID = '" & oeid & "';"
            
            Debug.Print sql
                            With DoCmd
                                .SetWarnings False
                                .RunSQL sql
                                .SetWarnings True
                            End With
            
                            'log the udpate
                            'LogPseudoEventIDUpdate rs("Event_ID"), eid, rsEventTables("Table_Name"), "i", TempVars("ImportContact")
                            LogPseudoEventIDUpdate oeid, eid, EventIDTable, "i", TempVars("ImportContact")
                        
                            'end processing
                            Application.SysCmd acSysCmdSetStatus, "Processing " & EventIDTable & " pseudoevents complete"
                        
                        Else
                            Debug.Print EventIDTable & " table does not exist"
                            
                            'end processing
                            Application.SysCmd acSysCmdSetStatus, "Skipping " & EventIDTable & " pseudoevents (table not present)"
                
                        End If
                        
                        'iterate to next table to update its Event_ID
                        rsEventTables.MoveNext
                    Loop
                
                    'end processing
                    Application.SysCmd acSysCmdSetStatus, "Processing " & EventIDTable & " pseudoevents complete"
        
                End If
                
                'iterate to next Event_ID in the EventTable
                rs.MoveNext
            Loop
        
        Else
            Debug.Print "skip this process"
        End If
        
    End If
    
Exit_Handler:
    DoCmd.Hourglass False
    Application.SysCmd acSysCmdClearStatus
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdatePseudoEventID[mod_PseudoEvents])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     ArchivePseudoEvents
' Description:  Archives pseudo-events to tsys_PseudoEvent_ARCHIVE table
' Assumptions:  Pseudo-events are archived from tbl_Events
' Parameters:   yr - 4 digit year to include (integer, optional, default = 0)
'                    if 0 then the current & prior year pseudo-events are archived
' Returns:      rs - pseudo-event IDs (DAO recordset)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 19, 2019
' Adapted:      -
' Revisions:
'   BLC - 9/19/2019 - initial version
'   BLC - 9/24/2019 - added full tbl_Events archive
' ---------------------------------
Public Function ArchivePseudoEvents(Optional yr As Integer = 0) As DAO.Recordset
On Error GoTo Err_Handler
    
    'defaults
    Dim sql As String
    Dim yrs As String
    Dim rs As DAO.Recordset
    
    DoCmd.Hourglass True
    
    yrs = IIf(yr = 0, "Year(Now()), Year(Now())-1", yr)
    
    'capture pseudo-event IDs
    sql = "SELECT Event_ID FROM tbl_Events WHERE PseudoEvent = 1 AND Year(Event_Date) IN (" & yrs & ");"
Debug.Print sql
    Set rs = CurrentDb.OpenRecordset(sql)
    
    'create PseudoEvent_ARCHIVE if it doesn't already exist
    If TableExists("tsys_PseudoEvent_ARCHIVE") = False Then
        DoCmd.CopyObject CurrentDb.Name, "tsys_PseudoEvent_ARCHIVE", acTable, "tbl_Events"
        'DoCmd.TransferDatabase acExport, "Microsoft Access", CurrentDb.Properties("Name"), acTable, "tbl_Events", "tsys_PseudoEvent_ARCHIVE", True
        
        'convert to local
        ConvertLinkedToLocal "tsys_PseudoEvent_ARCHIVE"
        
        'clear table of existing event data (from CopyObject)
        sql = "DELETE * FROM tsys_PseudoEvent_ARCHIVE"
        
        With DoCmd
            .SetWarnings False
            .RunSQL sql
            .SetWarnings True
        End With
        
    End If

    sql = "INSERT INTO tsys_PseudoEvent_ARCHIVE SELECT * FROM tbl_Events WHERE PseudoEvent = 1 AND Year(Event_Date) IN (" & yrs & ");"
    
    Debug.Print sql
    
    With DoCmd
        .SetWarnings False
        .RunSQL sql
        .SetWarnings True
    End With
    
    'create timestamped tbl_Events ARCHIVE before deleting pseudoevents
    Dim tblEventArchive As String
    
    tblEventArchive = "tsys_Events_ARCHIVE_" & Format(Now(), "YYYYMMDD_HHmm")
    DoCmd.CopyObject CurrentDb.Name, tblEventArchive, acTable, "tbl_Events"
    
    'convert to local
    ConvertLinkedToLocal tblEventArchive
    
    'Remove pseudo-events from MASTER database
    'processing
    Application.SysCmd acSysCmdSetStatus, "Deleting MASTER tbl_Events pseudoevents..."
    
    DeletePseudoEvents "tbl_Events", year(Now())
    DeletePseudoEvents "tbl_Events", year(Now()) - 1

    'return the pseudo-event IDs
    Set ArchivePseudoEvents = rs
    
Exit_Handler:
    'cleanup
    DoCmd.Hourglass False
    Application.SysCmd acSysCmdClearStatus
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ArchivePseudoEvents[mod_PseudoEvents])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          DeletePseudoEvents
' Description:  Deletes pseudo-events from tbl_Events (default) or EventTable if provided
' Assumptions:  -
' Parameters:   EventTable - name of event table (string, optional, default = tbl_Events)
'               yr - 4 digit year to include (integer, optional, default = 0)
'                    if 0 then the current & prior year pseudo-events are archived
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 19, 2019
' Adapted:      -
' Revisions:
'   BLC - 9/19/2019 - initial version
' ---------------------------------
Public Sub DeletePseudoEvents(Optional EventTable As String = "tbl_Events", Optional yr As Integer = 0)
On Error GoTo Err_Handler
    
    'defaults
    Dim sql As String
    Dim yrs As String
    
    yrs = IIf(yr = 0, "Year(Now()), Year(Now())-1", yr)
    
    sql = "DELETE * FROM " & EventTable & " WHERE PseudoEvent = 1 AND Year(Event_Date) IN (" & yrs & ");"
    
    With DoCmd
        .SetWarnings False
        .RunSQL sql
        .SetWarnings True
    End With
        
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeletePseudoEvents[mod_PseudoEvents])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          DeleteRelatedPseudoEventRecords
' Description:  Deletes records related to pseudo-events
' Assumptions:  pseudo-events have been deleted from the passed in tbl_Events
' Parameters:   EventTable - name of event table (string, optional, default = tbl_Events)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 23, 2019
' Adapted:      -
' Revisions:
'   BLC - 9/23/2019 - initial version
' ---------------------------------
Public Sub DeleteRelatedPseudoEventRecords(Optional EventTable As String = "tbl_Events")
On Error GoTo Err_Handler
    
    'defaults
    Dim EventIDTable As String
    Dim prefix As String
    Dim suffix As String
    Dim sql As String
    Dim sqlJOIN As String
    Dim tbl As String
    Dim rsTables As DAO.Recordset
    Dim rs As DAO.Recordset
    Dim skipDelete As Boolean
    
    'default
    skipDelete = False
    
    'determine the prefix & suffix of imported tables
    suffix = Replace(EventTable, "_tbl_Events", "")
    prefix = Replace(Replace(EventTable, suffix, ""), "tbl_Events", "")
    
    'retrieve Event_IDs
    sql = "SELECT Event_ID FROM " & EventTable
    Set rs = CurrentDb.OpenRecordset(sql)
    
    'retrieve append tables
    sql = "SELECT Table_Name FROM tsys_Append_Tables ORDER BY Append_Order;"
    Set rsTables = CurrentDb.OpenRecordset(sql)
    
    'populate recordsets
    If Not (rs.BOF = True And rs.EOF = True) Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    'EventIDs
    Dim eids As String
    eids = ""
    
    'iterate through Event_IDs
    Do While Not (rs.EOF = True)
Debug.Print rs("Event_ID")
        
        eids = eids & "'" & rs("Event_ID") & "',"
        
        rs.MoveNext
    Loop
Debug.Print eids
    
    eids = Left(eids, Len(eids) - 1)

    'reset to first table
    If Not (rsTables.BOF And rsTables.EOF) Then
        rsTables.MoveLast
        rsTables.MoveFirst
    End If

    'iterate through tables
    Do While Not rsTables.EOF
    
        'default
        skipDelete = False
    
Debug.Print rsTables("Table_Name")
        Select Case rsTables("Table_Name")
            Case "tbl_Events", "tbl_Locations", "tbl_Tags", "tbl_Tags_History"
                skipDelete = True
            'Event ID
            Case "tbl_Tree_Data", "tbl_Sapling_Data", "tbl_Quadrat_Data", "tbl_CWD_Data", "xref_Event_Contacts"
                sqlJOIN = " WHERE Event_ID NOT IN (" & eids & ");" ')" & rs("Event_ID") & "';"
            'Tree_Data_ID
            Case "tbl_Tree_DBH", "tbl_Tree_Conditions", "tbl_Tree_Foliage_Conditions", "tbl_Tree_Vines"
                sqlJOIN = " t LEFT JOIN " & prefix & "tbl_Tree_Data" & suffix & " td ON td.Tree_Data_ID = t.Tree_Data_ID " _
                            & " WHERE t.Tree_Data_ID IS NULL;"
            'Sapling_Data_ID
            Case "tbl_Sapling_DBH", "tbl_Sapling_Conditions", "tbl_Sapling_Foliage_Conditions", "tbl_Sapling_Vines"
                sqlJOIN = " s LEFT JOIN " & prefix & "tbl_Sapling_Data" & suffix & " sd ON sd.Sapling_Data_ID = s.Sapling_Data_ID " _
                            & " WHERE s.Sapling_Data_ID IS NULL;"
            'Quadrat_Data_ID
            Case "tbl_Quadrat_Seedling_Data", "tbl_Quadrat_Herbaceous_Data"
                sqlJOIN = " q LEFT JOIN " & prefix & "tbl_Quadrat_Data" & suffix & " qd ON qd.Quadrat_Data_ID = q.Quadrat_Data_ID " _
                            & " WHERE q.Quadrat_Data_ID IS NULL;"
        End Select

        tbl = prefix & rsTables("Table_Name") & suffix
        
        sql = "DELETE * FROM " & tbl & sqlJOIN
        
        If skipDelete = False Then
Debug.Print tbl & ": " & sql
           With DoCmd
               .SetWarnings False
'               .RunSQL sql
               .SetWarnings True
           End With
        End If
        
        rsTables.MoveNext
    Loop
            
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3128
        Debug.Print "ERROR 3128:"
        Debug.Print tbl
        Debug.Print sql
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeleteRelatedPseudoEventRecords[mod_PseudoEvents])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          LogPseudoEventIDUpdate
' Description:  Logs pseudo event ID updates for related tables
' Assumptions:  -
' Parameters:   OrigEID - original event ID (string)
'               NewEID - new event ID (string)
'               tbl - table being updated (string)
'               TriggerProcess - what triggered the update, i = import or c = pseudo-event ID conversion (string)
'               CID - ID for whomever triggered the update (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 19, 2019
' Adapted:      -
' Revisions:
'   BLC - 9/19/2019 - initial version
' ---------------------------------
Public Sub LogPseudoEventIDUpdate(OrigEID As String, NewEID As String, tbl As String, TriggerProcess As String, CID As String)
On Error GoTo Err_Handler
    
    Dim sql As String
    
    Select Case tbl
    '------- PRIMARY -------------
        Case "tbl_Events"
            sql = "UPDATE tbl_Events SET Event_ID = '" & fxnGUIDGen & "' WHERE Event_ID = '" & OrigEID & "';"
                    
        Case "tbl_Tree_Data"
            
        Case "tbl_Plot_Floor_Condition_Data"
        Case "tbl_Sapling_Data"
        Case "xref_Event_Contacts"
        
    '------- SECONDARY -------------
        Case "tbl_Events"
        Case "tbl_Quadrat_Data"
        Case "tbl_CWD_Data"
    End Select
    
    'create table if it doesn't exist
    If TableExists("tsys_EventID_Update_History") = False Then
        sql = "CREATE TABLE tsys_EventID_Update_History(" _
            & " Process VARCHAR(2) NOT NULL" _
            & ",OriginalEventID VARCHAR(50) NOT NULL" _
            & ",NewEventID VARCHAR(50) NOT NULL" _
            & ",UpdatedTable VARCHAR(150) NOT NULL" _
            & ",ProcessDate DATETIME NOT NULL" _
            & ",ContactID VARCHAR(50) NOT NULL" _
            & ");"
            
        With DoCmd
            .SetWarnings False
            .RunSQL sql
            .SetWarnings True
        End With
            
    End If
    
    'log the event ID change
    sql = "INSERT INTO tsys_EventID_Update_History (Process, OriginalEventID, NewEventID, UpdatedTable, ProcessDate, ContactID) " _
            & " SELECT '" & TriggerProcess & "', '" & OrigEID & "', '" & NewEID & "', '" & tbl & "', Now(), '" & CID & "';"
    Debug.Print sql
    With DoCmd
        .SetWarnings False
        .RunSQL sql
        .SetWarnings True
    End With
    
    'common sql for table updates
'    sql = "UPDATE " & tbl & " SET Event_ID = '" & fxnGUIDGen & "' WHERE Event_ID = '" & OrigEID & "';"
'    Debug.Print sql
'    With DoCmd
'        .SetWarnings False
'        .RunSQL sql
'        .SetWarnings True
'    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LogPseudoEventIDUpdate[mod_PseudoEvents])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
'  Functions
' ---------------------------------

Public Function runme()
    'ArchivePseudoEvents
    
    'ConvertLinkedToLocal "PseudoEvent_ARCHIVE"
    'LogPseudoEventIDUpdate "abc", "def", "mytbl", "i", "cid"
    'UpdatePseudoEventIDs "_tbl_Events_Import_20190531_PRIMARY" '"_tbl_Events"

    DeleteRelatedPseudoEventRecords "_tbl_Events_Import_20190531_PRIMARY"
End Function