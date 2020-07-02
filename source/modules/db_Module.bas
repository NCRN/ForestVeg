Option Compare Database
Option Explicit

' =================================
' MODULE:       db_Module
' Level:        Development module
' Version:      1.02
'
' Description:  Debugging related functions & procedures for database documentation
'
' Source/date:  Bonnie Campbell, June 19, 2019
' Revisions:    BLC - 6/19/2019 - 1.00 - initial version
'               BLC - 8/15/2019 - 1.01 - added SUB_PROTOCOL, GetSubProtocol()
'               BLC - 9/19/2019 - 1.02 - added ConvertLinkedToLocal()
' =================================
' ---------------------------------
'  Declarations
' ---------------------------------
'-----------------------------------------------------------------------
' Application
'-----------------------------------------------------------------------
Public Const DB_ADMIN As Boolean = True 'False 'True        'identifies if ADMIN functionality is available

Public FIELD_SEASON As Integer                              'identifies field season year (identified in frm_Switchboard)

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Methods
' ---------------------------------

' ---------------------------------
' Definitions
' References:   -
' Source/date:
'   DJ Steele,
'   http://www.accessmvp.com/DJSteele/DSNLessLinks.html
' Adapted:      Bonnie Campbell, June 26, 2019
' Revisions:
'   BLC - 6/26/2019 - initial version
' ---------------------------------
Type TableDetails
    TableName As String
    SourceTableName As String
    Attributes As Long
    IndexSQL As String
    Description As Variant
End Type

' ---------------------------------
' FUNCTION:     AttachDSNLessTable
' Description:  Link a remote table w/o a DSN
' Assumptions:  -
' Notes:        Call in either AutoExec macro or in Form_Open event for startup form
'               AutoExec macro:
'                    AttachDSNLessTable ("authors", "authors", "(local)", "pubs", "", "")
'               Form_Open event:
'                   Private Sub Form_Open(Cancel As Integer)
'                        If AttachDSNLessTable("authors", "authors", "(local)", "pubs", "", "") Then
'                            '// All is okay.
'                        Else
'                            '// Not okay.
'                        End If
'                    End Sub
' Parameters:   LocalTableName - table name in this database (string)
'               RemoteTableName - table name in remote database (string)
'               DbServer - name of SQL server instance (string)
'               DbName - database name to connect to (string)
'               DbUser - remote database username (string)
'               DbPassword - remote database password (string)
' Returns:      True or False depending on whether the table was successfully attached
' Throws:       none
' References:   -
' Source/date:
'   Microsoft, unknown
'   https://support.microsoft.com/en-us/help/892490/how-to-create-a-dsn-less-connection-to-sql-server-for-linked-tables-in
'   VilaRestal, May 24, 2012
'   https://access-programmers.co.uk/forums/showthread.php?t=226963
' Adapted:      Bonnie Campbell, June 19, 2019
' Revisions:
'   BLC - 6/19/2019 - initial version
' ---------------------------------
Public Function AttachDSNLessTable(LocalTableName As String, RemoteTableName As String, _
                                        DbServer As String, DbName As String, _
                                        Optional DbUser As String, Optional DbPassword As String) As Boolean
On Error GoTo Err_Handler
    
    Dim tdef As TableDef
    Dim dbConn As String
    
    For Each tdef In CurrentDb.TableDefs
        If tdef.Name = LocalTableName Then
            If MsgBox("Delete " & tdef.Name & " from database?", vbYesNo, "Delete Table?") = vbYes Then
                    CurrentDb.TableDefs.Delete LocalTableName
            Else
                Dim NewTableName As String
                
                'rename to YYYYMMDD_currentname, then delete
                NewTableName = Year(Date) & Month(Date) & Day(Date) & "_" & LocalTableName
                DoCmd.CopyObject , NewTableName, acTable, LocalTableName
                CurrentDb.TableDefs.Delete LocalTableName

            End If
            Exit For
        End If
        
    Next tdef
    
    If Len(DbUser) = 0 Then
         
        'use Trusted Connection
        dbConn = "ODBC;DRIVER=SQL Server;SERVER=" & DbServer & ";Trusted_Connection=YES"
    Else
        'WARNING: This will save username & pwd w/ the linked table info
        dbConn = "ODBC;DRIVER=SQL Server;SERVER=" & DbServer & ";DATABASE=dbo." & DbName & _
                    ";UID=" & DbUser & ";PWD=" & DbPassword
    End If
    
    Set tdef = CurrentDb.CreateTableDef(LocalTableName, dbAttachSavePWD, "dbo." & RemoteTableName, dbConn)
    'Set tdef = CurrentDb.CreateTableDef(LocalTableName, dbAttachSavePWD, RemoteTableName, dbConn)
    CurrentDb.TableDefs.Append tdef
    
    AttachDSNLessTable = True

Exit_Handler:
    Exit Function
    
Err_Handler:
    AttachDSNLessTable = False
    
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescriptions[mod_DbConn])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          FixConnections
' Description:  Looks for any TableDef objects in the db w/ a connection string (i.e. linked)
'               & changes their Connect property to use a DSN-less connection
'               Looks for any QueryDef objects in the db w/ a connection string (i.e. linked)
'               & changes the Connect property of those pass-through queries to use the same
'               DSN-less connection.
' Assumptions:  Connects to the specified SQL Server db on a specified server.
'               If user ID & pwd are supplied -> assumes SQL Server Security is being used
'               If none are supplied          -> assumes Trusted Connection (Windows Security)
' Notes:
' Parameters:   ServerName - SQL Server name (string)
'               DatabaseName - Server database name (string)
'               UID - UserID if using SQL Server Security (string)
'               PWD - Password if using SQL Server Security (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
'   Doug Steele, MVP  AccessMVPHelp@gmail.com
'   with modifications suggested by George Hepworth, MVP   ghepworth@gpcdata.com
'   http://www.accessmvp.com/DJSteele/DSNLessLinks.html
' Adapted:      Bonnie Campbell, June 26, 2019
' Revisions:
'   BLC - 6/26/2019 - initial version
' ---------------------------------
Sub FixConnections( _
    ServerName As String, _
    DatabaseName As String, _
    Optional UID As String, _
    Optional PWD As String _
)

On Error GoTo Err_Handler

    Dim dbCurrent As DAO.Database
    Dim prpCurrent As DAO.Property
    Dim tdfCurrent As DAO.TableDef
    Dim qdfCurrent As DAO.QueryDef
    Dim intLoop As Integer
    Dim intToChange As Integer
    Dim strConnectionString As String
    Dim strDescription As String
    Dim strQdfConnect As String
    Dim typNewTables() As TableDetails

    'Trusted Connection or SQL Server Security?
    If (Len(UID) > 0 And Len(PWD) = 0) Or (Len(UID) = 0 And Len(PWD) > 0) Then
        MsgBox "Must supply both User ID and Password to use SQL Server Security.", _
        vbCritical + vbOKOnly, "Security Information Incorrect."
        Exit Sub
    Else
        If Len(UID) > 0 And Len(PWD) > 0 Then
            ' Use SQL Server Security
            strConnectionString = "ODBC;DRIVER={sql server};" & _
            "DATABASE=" & DatabaseName & ";" & _
            "SERVER=" & ServerName & ";" & _
            "UID=" & UID & ";" & _
            "PWD=" & PWD & ";"
        Else
            ' Use Trusted Connection
            strConnectionString = "ODBC;DRIVER={sql server};" & _
            "DATABASE=" & DatabaseName & ";" & _
            "SERVER=" & ServerName & ";" & _
            "Trusted_Connection=YES;"
        End If
    End If

    intToChange = 0
    
    Set dbCurrent = DBEngine.Workspaces(0).Databases(0)

    'Build connected TableDefs, connected tables list
    For Each tdfCurrent In dbCurrent.TableDefs
        If Len(tdfCurrent.connect) > 0 Then
            If UCase$(Left$(tdfCurrent.connect, 5)) = "ODBC;" Then
            ReDim Preserve typNewTables(0 To intToChange)
            typNewTables(intToChange).Attributes = tdfCurrent.Attributes
            typNewTables(intToChange).TableName = tdfCurrent.Name
            typNewTables(intToChange).SourceTableName = tdfCurrent.SourceTableName
            typNewTables(intToChange).IndexSQL = GenerateIndexSQL(tdfCurrent.Name)
            typNewTables(intToChange).Description = Null
            typNewTables(intToChange).Description = tdfCurrent.Properties("Description")
            intToChange = intToChange + 1
            End If
        End If
    Next

    'Iterate through Linked Tables
    For intLoop = 0 To (intToChange - 1)

        'Delete existing TableDef objects
        dbCurrent.TableDefs.Delete typNewTables(intLoop).TableName
        
        'Create new TableDef object, using the DSN-less connection
        Set tdfCurrent = dbCurrent.CreateTableDef(typNewTables(intLoop).TableName)
        tdfCurrent.connect = strConnectionString
        
        ' Unfortunately, I'm current unable to test this code,
        ' but I've been told trying this line of code is failing for most people...
        ' If it doesn't work for you, just leave it out.
        tdfCurrent.Attributes = typNewTables(intLoop).Attributes
        
        tdfCurrent.SourceTableName = typNewTables(intLoop).SourceTableName
        dbCurrent.TableDefs.Append tdfCurrent
        
        'Add the Description property to new table if it exists
        If IsNull(typNewTables(intLoop).Description) = False Then
            strDescription = CStr(typNewTables(intLoop).Description)
            Set prpCurrent = tdfCurrent.CreateProperty("Description", dbText, strDescription)
            tdfCurrent.Properties.Append prpCurrent
        End If
        
        'Create __UniqueIndex index on new table if it existed
        If Len(typNewTables(intLoop).IndexSQL) > 0 Then
            dbCurrent.Execute typNewTables(intLoop).IndexSQL, dbFailOnError
        End If
Next

' Loop through all the QueryDef objects looked for pass-through queries to change.
' Note that, unlike TableDef objects, you do not have to delete and re-add the
' QueryDef objects: it's sufficient simply to change the Connect property.
' The reason for the changes to the error trapping are because of the scenario
' described in Addendum 6 below.

For Each qdfCurrent In dbCurrent.QueryDefs
    On Error Resume Next
    strQdfConnect = qdfCurrent.connect
    On Error GoTo Err_Handler
    If Len(strQdfConnect) > 0 Then
        If UCase$(Left$(qdfCurrent.connect, 5)) = "ODBC;" Then
            qdfCurrent.connect = strConnectionString
        End If
    End If
    strQdfConnect = vbNullString
Next qdfCurrent

Exit_Handler:
  Set tdfCurrent = Nothing
  Set dbCurrent = Nothing
  Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 3270   'Property Not Found to handle tables w/o descriptions
            Resume Next
        
        Case 3291   'Syntax error in CREATE INDEX statement
            MsgBox "Error #" & Err.Number & ": " & Err.Description & vbCrLf & _
                "Problem creating the Index using" & vbCrLf & _
                typNewTables(intLoop).IndexSQL & vbCrLf _
                , vbOKOnly + vbCritical, _
                "Error encountered (#" & Err.Number & " - FixConnections[mod_DbConn])"
        
        Case 18456
            MsgBox "Error #" & Err.Number & ": " & Err.Description & vbCrLf & _
                "Wrong User ID or Password." _
                , vbOKOnly + vbCritical, _
                "Error encountered (#" & Err.Number & " - FixConnections[mod_DbConn])"
        
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - FixConnections[mod_DbConn])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GenerateIndexSQL
' Description:  Linked tables should have an index __uniqueindex
'               Looks for that index in a given table & creates a SQL statement to re-create it
'               If no such index exists, function returns an empty string ("").
' Assumptions:  Appears to be no other way to do this.
' Notes:
' Parameters:   TableName - name of table
' Returns:      SQL string for the index re-creation or empty string if index not found (string)
' Throws:       none
' References:   -
' Source/date:
'   Doug Steele, MVP  AccessMVPHelp@gmail.com
'   with modifications suggested by George Hepworth, MVP   ghepworth@gpcdata.com
'   http://www.accessmvp.com/DJSteele/DSNLessLinks.html
' Adapted:      Bonnie Campbell, June 26, 2019
' Revisions:
'   BLC - 6/26/2019 - initial version
' ---------------------------------
Function GenerateIndexSQL(TableName As String) As String
On Error GoTo Err_Handler

    Dim dbCurr As DAO.Database
    Dim idxCurr As DAO.index
    Dim fldCurr As DAO.field
    Dim strSQL As String
    Dim tdfCurr As DAO.TableDef
    
    Set dbCurr = CurrentDb()
    Set tdfCurr = dbCurr.TableDefs(TableName)

    'Check for "__UniqueIndex" index in table
    If tdfCurr.Indexes.Count > 0 Then
        On Error Resume Next
        
        Set idxCurr = tdfCurr.Indexes("__uniqueindex")
        
        If Err.Number = 0 Then
            On Error GoTo Err_Handler
            'Iterate through table fields in index & add to SQL statement
            If idxCurr.Fields.Count > 0 Then
                strSQL = "CREATE INDEX __UniqueIndex ON [" & TableName & "] ("
                For Each fldCurr In idxCurr.Fields
                strSQL = strSQL & "[" & fldCurr.Name & "], "
                Next
            
                ' Remove trailing comma and space
                strSQL = Left$(strSQL, Len(strSQL) - 2) & ")"
            End If
        End If
    End If

Exit_Handler:
    Set fldCurr = Nothing
    Set tdfCurr = Nothing
    Set dbCurr = Nothing
    GenerateIndexSQL = strSQL
Exit Function

Err_Handler:
    Select Case Err.Number
        Case 3265   'Not found in this collection --> either tablename is invalid or no __uniqueindex index
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - GenerateIndexSQL[db_Module])"
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - GenerateIndexSQL[db_Module])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          PurgeAnnualData
' Description:  Create a BE with data purged from annual data tables
'               for beginning of monitoring season
' Assumptions:  -
' Parameters:   BackupWithPurgedTables - whether the database should be backed up
'                    after the purged tables are copied as APBU_xx tables  (boolean, optional, default = false)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 15, 2019
' Adapted:
' Revisions:
'   BLC - 7/15/2019 - initial version
'   BLC - 8/18/2019 - changed to copy as local vs linked table (avoids dropping data)
' ---------------------------------
Public Sub PurgeAnnualData(Optional BackupWithPurgedTables As Boolean = False)
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim SQL As String
    Dim Response As Integer
    
    Response = Eval(MsgBox("Pre-Season Prep will delete data from data tables (e.g. events, sapling data)." _
                    & vbCrLf & vbCrLf & "Are you sure you wish to proceed?", vbYesNo, "Confirm Delete Data"))
    
    'exit if user choses not to purge data
    If Response = vbNo Then GoTo Exit_Handler
    
    'retrieve table names for tables containing annual data
    SQL = "SELECT Link_table FROM tsys_Link_Tables WHERE AnnualDbPurge = 1;"
    
    Set rs = CurrentDb.OpenRecordset(SQL)
    
    'start @ beginning
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveLast 'get accurate recordcount
        rs.MoveFirst
    
        DoCmd.Hourglass True
        
        Dim tbl As String
        
        'Do While Not rs.EOF 'causes hang?
        Do Until rs.EOF
            'copy table to APBU_<tablename>_YYYYMMDD_hhmmss in same db
            '(APBU = annual purge backup)
            
            'status bar message
            Application.SysCmd acSysCmdSetStatus, "Copying " & rs("Link_table") & "..."
            
            'DoCmd.CopyObject , "APBU_" & rs("Link_table") & "_" & Format(Now(), "YYYYMMDD_hhmmss"), acTable, rs("Link_Table")
            
            tbl = "APBU_" & rs("Link_table") & "_" & Format(Now(), "YYYYMMDD_hhmmss")
            
            'copy as LOCAL table, not linked!
            ' NOTE: TransferDatabase REQUIRES 'optional' info including - db type & full path to work
            'DoCmd.TransferDatabase acImport, , , acTable, rs("Link_Table"), "APBU_" & rs("Link_table") & "_" & Format(Now(), "YYYYMMDD_hhmmss")
'            DoCmd.TransferDatabase acImport, "Microsoft Access", Application.CurrentDb.Name, acTable, rs("Link_Table"), "APBU_" & rs("Link_table") & "_" & Format(Now(), "YYYYMMDD_hhmmss"), False 'tbl
            
            'convert the table to local (TransferDatabase should do this but it does not)
            ' NOTE: Cannot do this as TransferDatabase & Convert only create linked table??
            '       Need to find a way of doing this
'            DoCmd.OpenTable tbl
'            DoCmd.SelectObject acTable, tbl, True
'            DoCmd.RunCommand acCmdConvertLinkedTableToLocal

            CopySchemaAndData_DAO rs("Link_table"), tbl
            
            'empty current table
            Application.SysCmd acSysCmdSetStatus, "Purging " & rs("Link_table") & "..."
            
            SQL = "DELETE * FROM " & rs("Link_table") & ";"
            
            Debug.Print SQL & vbCrLf
          
            CurrentDb.Execute SQL, dbFailOnError
            
            rs.MoveNext
        Loop
        
        DoCmd.Hourglass False
    
    End If

    If BackupWithPurgedTables = True Then
        'create Db backup w/ Purged tables
        BackupDbBE
    End If

    'complete
    Application.SysCmd acSysCmdSetStatus, "Data table copying & purging complete!"
    
    'update
    MsgBox "Pre-season backup & annual data purge is complete." & vbCrLf _
           & "Review APBU_* data tables before deleting them.", _
           vbOKOnly + vbInformation, "Pre-Season Backup & Annual Db Prep Complete"
    
Exit_Handler:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Application.SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PurgeAnnualData[db_Module])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          BackupDbBE
' Description:  Back up the current database back-end (BE) file
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, May 24, 2006
' Revisions:
'   JRB   - 5/24/2006 - initial version
'   ML/GS - unknown   - initial NCRN version
'   BLC   - 7/8/2019 - documentation
'   BLC   - 7/16/2019 - shift from frm_Switchboard.btnBackup_Click to public sub in Db_module
' ---------------------------------
Public Sub BackupDbBE()
On Error GoTo Err_Handler

    ' Start the database backup function
    If fxnVerifyLinks() Then
        'status bar message
        Application.SysCmd acSysCmdSetStatus, "Backing up back-end database..."
        
        fxnMakeBackup
    Else
        MsgBox "Cannot create a backup until the database connection is fixed", _
            vbExclamation, "Data Tables Not Connected"
    End If

Exit_Handler:
    Application.SysCmd acSysCmdClearStatus
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - BackupDbBE[Db_module])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          PreSeasonDbPrep
' Description:  Create a BE with data purged from annual data tables
'               after backing up this BE database
'               for beginning of monitoring season
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 16, 2019
' Adapted:
' Revisions:
'   BLC - 7/16/2019 - initial version
' ---------------------------------
Public Sub PreSeasonDbPrep()
On Error GoTo Err_Handler

    'backup current database
    
    
    'prior year data purge (for purgeable data tables in tsys_Link_Tables)
    PurgeAnnualData

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PreSeasonDbPrep[db_Module])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetFieldSeason
' Description:  Returns the FIELD_SEASON global public variable
'               which identifies the year the latest data were collected
' Assumptions:  -
' Parameters:   -
' Returns:      FIELD_SEASON
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 16, 2019
' Adapted:
' Revisions:
'   BLC - 8/16/2019 - initial version
' ---------------------------------
Public Function GetFieldSeason() As Integer
On Error GoTo Err_Handler

    'return the field season value
    GetFieldSeason = FIELD_SEASON

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetFieldSeason[db_Module])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     CopySchemaAndData_DAO
' Description:  Copies an existing table to a new local table
'               Linked tables are converted to local & indexes are retained
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Tim Lentine, October 20, 2009
'   https://stackoverflow.com/questions/1594096/how-to-copy-a-linked-table-to-a-local-table-in-ms-access-programmatically
' Source/date:  Bonnie Campbell, August 19, 2019
' Adapted:
' Revisions:
'   BLC - 8/19/2019 - initial version
' ---------------------------------
Public Sub CopySchemaAndData_DAO(SourceTable As String, DestinationTable As String)
On Error GoTo Err_Handler

    Dim tblSource As DAO.TableDef
    Dim fld As DAO.field
    
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Set tblSource = db.TableDefs(SourceTable)
    
    Dim tblDest As DAO.TableDef
    Set tblDest = db.CreateTableDef(DestinationTable)
    
    'Iterate over source table fields and add to new table
    For Each fld In tblSource.Fields
       Dim destField As DAO.field
       Set destField = tblDest.CreateField(fld.Name, fld.Type, fld.Size)
       If fld.Type = 10 Then
          'text, allow zero length
          destField.AllowZeroLength = True
       End If
       tblDest.Fields.Append destField
    Next fld
    
    'Handle Indexes
    Dim idx As index
    Dim iIndex As Integer
    For iIndex = 0 To tblSource.Indexes.Count - 1
       Set idx = tblSource.Indexes(iIndex)
       Dim newIndex As index
       Set newIndex = tblDest.CreateIndex(idx.Name)
       With newIndex
          .Unique = idx.Unique
          .Primary = idx.Primary
          'Some Indexes are made up of more than one field
          Dim iIdxFldCount As Integer
          For iIdxFldCount = 0 To idx.Fields.Count - 1
             .Fields.Append .CreateField(idx.Fields(iIdxFldCount).Name)
          Next iIdxFldCount
       End With
    
       tblDest.Indexes.Append newIndex
    Next iIndex
    
    db.TableDefs.Append tblDest
    
    'Finally, copy data from source to destination table
    Dim SQL As String
    SQL = "INSERT INTO " & DestinationTable & " SELECT * FROM " & SourceTable
    db.Execute SQL

Exit_Handler:
   Set fld = Nothing
   Set destField = Nothing
   Set tblDest = Nothing
   Set tblSource = Nothing
   Set db = Nothing
   Exit Sub
   
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CopySchemaAndData_DAO[db_Module])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:     ConvertLinkedToLocal
' Description:  Copies an existing table to a new local table
'               Linked tables are converted to local & indexes are retained
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   ADezii, September 12, 2019
'   https://bytes.com/topic/access/answers/973098-how-copy-linked-table-another-table
' Source/date:  Bonnie Campbell, August 19, 2019
' Adapted:
' Revisions:
'   BLC - 9/19/2019 - initial version
' ---------------------------------
Public Sub ConvertLinkedToLocal(strLinkedTbl As String)
On Error GoTo Err_Handler

    Dim strConnect As String
    Dim strPath As String
    Dim strSourceTable As String
     
    strConnect = CurrentDb.TableDefs(strLinkedTbl).connect      'Connect String
     
    If InStr(strConnect, "=") = 0 Then
      MsgBox strLinkedTbl & " is not a Linked Table!", vbCritical, "Linked Table Error"
        Exit Sub
    Else
      strPath = mid$(strConnect, InStr(strConnect, "=") + 1)        'Actual DB Path
      strSourceTable = CurrentDb.TableDefs(strLinkedTbl).SourceTableName
     
      DoCmd.DeleteObject acTable, strLinkedTbl
      DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, _
                             strSourceTable, strLinkedTbl, False
    End If
    
Exit_Handler:
   Exit Sub
   
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConvertLinkedToLocal[db_Module])"
    End Select
    Resume Exit_Handler
End Sub