Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Db
' Level:        Framework module
' Version:      1.28
' Description:  Database related functions & subroutines
' Requires:     Microsoft Scripting Runtime (scrrun.dll) for Scripting.Dictionary
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               BLC, 5/26/2015 - 1.01 - added mod_db_Templates subs/functions - qryExists
'               BLC, 5/26/2016 - 1.02 - added VirtualDAORecordset()
'               BLC, 6/6/2016  - 1.03 - added error handling for duplicate Templates, renamed global to g_AppTemplates
'                                       also added SQL sanitization (escape/replace special chars)
'               BLC, 6/9/2016  - 1.04 - added CreateTempRecords()
'               BLC, 10/4/2016 - 1.05 - added GetParamsFromSQL()
'               BLC, 10/11/2016 - 1.06 - added IsRecordset(), FieldCount(), MaxDbFieldCount()
'               BLC, 10/20/2016 - 1.07 - added IsLinked()
'               BLC, 1/9/2017 - 1.08   - added SetTempVar()
'               BLC, 1/19/2017 - 1.09  - added RetrieveTableColumnData()
'               BLC, 2/1/2017  - 1.10  - added error handling for improper GetTemplates()
'                                        parameter syntax (param name:param type)
'               BLC, 2/2/2017  - 1.11  - cleared g_AppTemplates in GetTemplates()
'                                        to allow re-definition w/o restarting db
'               BLC, 2/22/2017 - 1.12  - added DbObjectExists() to validate db objects
' --------------------------------------------------------------------
'               BLC, 3/8/2017          added to Invasives db
' --------------------------------------------------------------------
'               BLC, 3/8/2017 - 1.13a - imported into invasives,
'                                      subs/functions not available in invasives
'                                      (missing reference/function):
'                                       g_AppTemplates, GetTemplate(), GetTemplates(),
'                                       CreateTempRecords(), GetParamsFromSQL(),
' --------------------------------------------------------------------
'               BLC, 3/22/2017          added to Upland db
' --------------------------------------------------------------------
'               BLC, 3/23/2017 - 1.13  - revised GetTemplates() to use "SQL" vs. "T-SQL" syntax
'               BLC, 3/28/2017 - 1.14  - added CloseObject()
'               BLC, 3/30/2017 - 1.15  - added Template dependent query function
'                                        HandleDependentQueries(), SetQueryProperty(),
'                                        DeleteRecord() moved from mod_UI
'               BLC, 3/31/2017 - 1.16  - added g_AppTemplateIDs global for Template/ID matches
'               BLC, 4/3/2017  - 1.17  - code cleanup
' --------------------------------------------------------------------
'               BLC, 4/18/2017          added updated version to Invasives db
' --------------------------------------------------------------------
'               BLC, 4/18/2017 - 1.18 - adjusted for invasives, added Scripting.Dictionary reference,
'                                       revised GetTemplates to avoid error on dictTemplates.Add
'               BLC, 6/19/2017 - 1.19 - updated SQL for OpenRecordset to properly order by tsys_BE_Updates.Update_ID vs.
'                                       tsys_BE_Updates.ID which does not exist
'               BLC, 6/22/2017 - 1.20 - added SetColumnOrdinalPosition(), CombineTableSQL()
' --------------------------------------------------------------------
'               BLC, 8/22/2017 - 1.21 - merged prior work:
'                   Invasives db
'                       BLC, 4/28/2017 - 1.19 - added Delete_All_Records() moved from mod_Temp,
'                                               reorganized sections, moved SQL_encode(), GetParamsFromSQL()
'                                               to mod_SQL
'                                               moved SetTempVar(), GetTempVarIndex(),
'                                               CreateTempTable(), RemoveTempTable(),
'                                               CreateTempRecordset(), CreateTempRecords()
'                                               to mod_Temp
'                       BLC, 7/18/2017 - 1.20 - Add RefreshTempTable for updating usys_temp_transect,
'                                       usys_temp_speciescover & other Temp tables w/ std naming
'                   Uplands db
'                       BLC, 8/10/2017 - 1.18  - add OpenAllDatabases(), CurrDb property
' --------------------------------------------------------------------
'               BLC, 10/4/2017 - 1.22 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC, 10/5/2017 - 1.23 - updated documentation, merged DbCurrent (mod_SQL)
'                                       with CurrDb
'               BLC, 10/17/2017 - 1.24 - moved SysTablesExist() from mod_Initialize_App
'               BLC, 11/24/2017 - 1.25 - revised to ShowMsg vs displayMsg, , updated to use DisplayMsg() (DeleteRecord())
'               BLC, 12/13/2017 - 1.26 - add RefreshTemplates()
'               BLC, 1/2/2018   - 1.27 - add ConvertObjectToPointer(), ConvertPointerToObject(),
'                                        & related POINTERSIZE, ZEROPOINTER, RtlMoveMemory
'               BLC, 1/12/2018  - 1.28 - update to eliminate ~TMP... tables, refresh TableDefs (ListTables)
' =================================

' ---------------------------------
' Declarations
' ---------------------------------
'   AppTemplates global dictionary --> defined in std Template [mod_Db]
Public g_AppTemplates As Scripting.Dictionary
Public g_AppTemplateIDs As Scripting.Dictionary
Public Const PARAM_SEPARATOR As String = ">>"
Public g_OpenQueries As String                  'queries generated by Templates (close @ end)

'for passing objects via pointer & reverting back to object
Private Const POINTERSIZE As Long = 4
Private Const ZEROPOINTER As Long = 0

' to avoid sub not found errors use IF/THEN
' Reference:
'   Charles Williams, November 23, 2010
'   https://stackoverflow.com/questions/4251111/how-to-make-vba-code-compatible-for-office-2010-64-bit-version-and-older-offic
'   HansUp, August 6, 2010
'   https://stackoverflow.com/questions/3426693/tempvars-and-access-2003/3427119#3427119

'Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
'    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#Else
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

' ---------------------------------
'  Database-wide Properties
' ---------------------------------

' ---------------------------------
' PROPERTY:     CurrDb
' Description:  Gets a single instance of the current db to avoid multiple calls
'               to CurrentDb which can yield to Error 3048 "Cannot open any more databases" errors
'               due to multiple open db
' Parameters:   -
' Returns:      current database object
' Throws:       -
' References:   -
' Source/date:  Dirk Goldgar, MS Access MVP - May 22, 2013
'   http://social.msdn.microsoft.com/Forums/office/en-US/9993d229-8a00-4a59-a796-dfa2dad505bc/cannot-open-any-more-databases?forum=accessdev
'   Michael Kaplan via Darrel H. Burns, February 8, 2011
'   https://social.msdn.microsoft.com/Forums/office/en-US/7ea9506f-5e91-4896-80b9-6712762388ea/currentdbtabledefs-vs-dbtabledefs-object-invalid-or-not-set-error?forum=accessdev
' Adapted:      Bonnie Campbell, July, 2014 for NCPN Riparian tools
' Revisions:    BLC, 7/23/2014 - initial version
'               BLC, 10/5/2017 - combined with DbCurrent from mod_SQL
' ---------------------------------
Private m_db As DAO.Database

Public Property Get CurrDb() As DAO.Database

    If (m_db Is Nothing) Then
        Set m_db = CurrentDb
    End If

    Set CurrDb = m_db

End Property
  
' ---------------------------------
' Types & Type Descriptions
' ---------------------------------
' -32768  Form                    1   Table - Local Access Tables
' -32766  Macro                   2   Access Object - Database
' -32764  Reports                 3   Access Object - Containers
' -32761  Module                  4   Table - Linked ODBC Tables
' -32758  Users                   5   Queries
' -32757  Database Document       6   Table - Linked Access Tables
' -32756  Data Access Pages       8   SubDataSheets
' ---------------------------------

' ---------------------------------
'   Database Connectivity
' ---------------------------------
' ---------------------------------
' Sub:          OpenAllDatabases
' Description:  open a handle to all databases & keep it open during entire time application
'               runs to avoid closing/opening it during use to improve performance
' Assumptions:
'               Databases to connect to are known to start
'               & do not change during use
'               The following calls are made:
'                  @ Application Start >> OpenAllDatabases True
'                  @ Application Close >> OpenAllDatabases False
' Parameters:   Init - TRUE to initialize (call when application starts)
'                      FALSE to close (call when application ends)
' Returns:      -
' Throws:       none
' References:
'   FMS < Total Visual SourceBook
'   http://www.fmsinc.com/microsoftaccess/performance/linkeddatabase.html
' Source/date:  Bonnie Campbell, March 22, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/22/2017 - initial version
'   BLC - 3/24/2017 - revise to restore calling form
' ---------------------------------
Public Sub OpenAllDatabases(pfInit As Boolean)
On Error GoTo Err_Handler

  Dim X As Integer
  Dim strName As String
  Dim strMsg As String
 
  ' Maximum number of back end databases to link
  Const cintMaxDatabases As Integer = 2

  ' List of databases kept in a static array so we can close them later
  Static dbsOpen() As DAO.Database
 
  If pfInit Then
    ReDim dbsOpen(1 To cintMaxDatabases)
    For X = 1 To cintMaxDatabases
      ' Specify your back end databases
      Select Case X
        Case 1:
          strName = "H:\folder\Backend1.mdb"
        Case 2:
          strName = "H:\folder\Backend2.mdb"
      End Select
      strMsg = ""

      On Error Resume Next
      Set dbsOpen(X) = OpenDatabase(strName)
      If Err.Number > 0 Then
        strMsg = "Trouble opening database: " & strName & vbCrLf & _
                 "Make sure the drive is available." & vbCrLf & _
                 "Error: " & Err.Description & " (" & Err.Number & ")"
      End If

      On Error GoTo 0
      If strMsg <> "" Then
        MsgBox strMsg
        Exit For
      End If
    Next X
  Else
    On Error Resume Next
    For X = 1 To cintMaxDatabases
      dbsOpen(X).Close
    Next X
  End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenAllDatabases[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Database Object Methods/Functions
' ---------------------------------

' =================================
' FUNCTION:     BEUpdates
' Description:  Runs SQL statement updates from the systems table tsys_BE_Updates. Such
'               updates are sometimes necessary when there is a remote copy of the back-end
'               file that the developer cannot access, but which needs to be updated to
'               include the current release information. tsys_BE_Updates has the following
'               structure:  Update_ID (txt serial number autoincrementing), Is_done (yes/no),
'               Run_date (datetime), SQL_statement (memo), Update_desc (txt 100)
' Parameters:   bRunAll - True (default), or False if only running lines where [Is_done]=False
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/10/2008
' Revisions:    JRB, 11/21/2008 - added optional parameter to either run all update lines
'                   (default), or just one where [Is_done]=False
'               BLC, 4/30/2015  - moved to mod_Db framework module from mod_Custom_Functions
'                                 added check for BOF & EOF to avoid Error #3021 no current record on rs.MoveLast when no records exist
'               BLC, 5/18/2015 - renamed & removed fxn prefix
'               BLC, 6/5/2016  - adapted for Big Rivers App naming revisions (removed field underscores)
'               BLC, 6/19/2017 - updated SQL for OpenRecordset to properly order by tsys_BE_Updates.Update_ID vs.
'                                tsys_BE_Updates.ID which does not exist
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' =================================
Public Function BEUpdates(Optional ByVal bRunAll As Boolean = True)
    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim intNumUpdates As Integer
    Dim varReturn As Variant
    Dim intI As Integer
    Dim strSQL As String
    
    Set db = CurrDb
    Set rs = db.OpenRecordset("SELECT tsys_BE_Updates.* FROM tsys_BE_Updates " & _
        "ORDER BY tsys_BE_Updates.ID;", dbOpenDynaset)

    ' Check for BOF & EOF to avoid Error # 3021 No current record
    If Not rs.BOF And rs.EOF Then

        ' Counts the number of db update records in the system table
        rs.MoveLast    ' Need to do this to make the record count accurate
        intNumUpdates = rs.RecordCount
        If intNumUpdates = 0 Then    ' No records in the recordset
            GoTo Exit_Procedure
        End If
    
        ' First pass to verify the tables in the specified database
        '   Initialize the system meter to indicate progress
        varReturn = SysCmd(acSysCmdInitMeter, "Performing database updates", intNumUpdates)
        intI = 0
        rs.MoveFirst
        On Error Resume Next
        Do Until rs.EOF
            intI = intI + 1
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            If bRunAll = True Or rs![IsDone] = False Then
                DoCmd.SetWarnings False
                strSQL = rs![SQLStatement]
                DoCmd.RunSQL strSQL
                With rs
                    .Edit
                    ![RunDate] = Now()
                    ![IsDone] = True
                    .Update
                End With
            End If
            rs.MoveNext
        Loop
        
    End If

Exit_Procedure:
    On Error Resume Next
    DoCmd.SetWarnings True
    varReturn = SysCmd(acSysCmdRemoveMeter)
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3061   ' Bad parameters for the SQL string
        MsgBox "Error #" & Err.Number & ":  SQL syntax error. Please notify the " & _
            "database administrator before using this application.", vbCritical, _
            "Error encountered (#" & Err.Number & " - BEUpdates[mod_Db])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - BEUpdates[mod_Db])"
    End Select
    Resume Exit_Procedure

End Function

' ---------------------------------
' FUNCTION:     getAccessObjectType
' Description:  looks up object type in Access sys tables
' Parameters:   strName  - name of object w/in Access
' Returns:      long (type) or NULL if object doesn't exist
'                   ----------------
'                   1 = Access Table
'                   4 = OBDB-Linked Table / View
'                   5 = Access Query
'                   6 = Attached (Linked) File  (such as Excel, another Access Table or query, text file, etc.)
'                   -32768 = Access Form
'                   -32764 = Access Report
'                   -32761 = Access Module
'                   ----------------
' Throws:       none
' References:   Tom Davidson, April 8, 2011
'   http://stackoverflow.com/questions/2090578/ms-access-determine-object-type

' Source/date:  Bonnie Campbell August 20, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 8/20/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI
' ---------------------------------
Public Function getAccessObjectType(strObject As String) As Variant
On Error GoTo Err_Handler:

    getAccessObjectType = DLookup("Type", "MSysObjects", "NAME = '" & strObject & "'")
   
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getAccessObjectType[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Database & Recordset Actions
' ---------------------------------

' ---------------------------------
' FUNCTION:     ClearTable
' Description:  Deletes records from table
' Assumptions:  Table is in the current database (not linked)
' Parameters:   strTable - table name (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015  - initial version
' ---------------------------------
Public Sub ClearTable(strTable As String)

On Error GoTo Err_Handler
    
    Dim strSQL As String
    
    'clear table
    strSQL = "DELETE * FROM " & strTable & ";"
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearTable[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     AddField
' Description:  Adds field to recordset
' Assumptions:  Table is in the current database (not linked)
' Parameters:   tdf - table to add field to
'               fldName - field name (string)
'               fldType - vartype for field (optional)
'               fldSize - size of field (optional)
' Returns:      -
' Throws:       none
' References:
'   Microsoft, March 9, 2015
'   https://msdn.microsoft.com/en-us/library/office/ff820791.aspx
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017  - initial version
' ---------------------------------
Public Sub AddField(tdf As DAO.TableDef, fldName As String, Optional fldType, Optional fldSize)
On Error GoTo Err_Handler
    
    With tdf
        If .Updatable = False Then
        
            Dim msg As String
            
            msg = "Oops! " & tdf.Name & " cannot be updated"
            
            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                "msg" & PARAM_SEPARATOR & msg & _
                "|Type" & PARAM_SEPARATOR & "caution"
            
            GoTo Exit_Handler
            
        End If
        
        .Fields.Append .CreateField(fldName, fldType, fldSize)
    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddField[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     RemoveField
' Description:  Adds field to recordset
' Assumptions:  Table is in the current database (not linked)
' Parameters:   tdf - table to add field to
'               fldName - field name (string)
' Returns:      -
' Throws:       none
' References:
'   Microsoft, March 9, 2015
'   https://msdn.microsoft.com/en-us/library/office/ff820791.aspx
' Source/date:  Bonnie Campbell, December 8, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2017  - initial version
' ---------------------------------
Public Sub RemoveField(tdf As DAO.TableDef, fldName As String)
On Error GoTo Err_Handler
    
    With tdf
        If .Updatable = False Then
        
            Dim msg As String
    
            msg = "Oops! " & tdf.Name & " cannot be updated"
        
            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                "msg" & PARAM_SEPARATOR & msg & _
                "|Type" & PARAM_SEPARATOR & "caution"
               
            GoTo Exit_Handler
        End If
        
        .Fields.Delete fldName
    End With
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveField[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Recordset Level Methods/Functions
' ---------------------------------

' ---------------------------------
' FUNCTION:     MergeRecordsets
' Description:  Merge two recordsets into one (useful when the recordsets already exist vs. direct SQL union)
' Assumptions:  Recordsets have the same fields in the same order
' Parameters:   rsA - DAO recordset A
'               rsB - DAO recordset B to merge with A
' Returns:      DAO.Recordset
' Throws:       none
' References:   none
' Source/date:
' Chris Oswald, January 26, 2011
' http://www.mrexcel.com/forum/excel-questions/524214-visual-basic-applications-joining-multiple-recordets-multiple-databases.html
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/7/2015  - initial version
'   BLC - 5/13/2015 - moved from mod_App_Data to mod_Db
' ---------------------------------
Public Function MergeRecordsets(rsA As DAO.Recordset, rsB As DAO.Recordset) As DAO.Recordset

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rsOut As DAO.Recordset
    Dim iCount As Integer
    
    'handle empty recordsets
    If rsA Is Nothing Then
        'check rsB
        If rsB Is Nothing Then
            GoTo Exit_Handler
        Else
            Set MergeRecordsets = rsB
            GoTo Exit_Handler
        End If
    End If
    

'With rsA
    'check if rsA and rsB are both populated --> if not, exit
    If (rsA.EOF And rsA.BOF) Then
        'rsA not populated
        If (rsB.EOF And rsB.BOF) Then
            'neither is populated --> EXIT!
            GoTo Exit_Handler
        Else
            'rsB populated --> return rsB
            Set MergeRecordsets = rsB
        End If
    Else
        'rsA populated --> if rsB not populated, return rsA
        If (rsB.EOF And rsB.BOF) Then
            Set MergeRecordsets = rsA
            GoTo Exit_Handler
        End If
    'End If
    
    'create output recordset vs. just adding to rsB
    Set rsOut = rsA
    Do Until rsB.EOF
        'add rsB values as new rsOut records
        rsOut.AddNew
        For iCount = 0 To rsB.Fields.Count - 1
            rsOut.Fields(iCount).Value = rsB.Fields(iCount).Value
        Next
        rsOut.Update
        rsB.MoveNext
    Loop
    
    'rsOut.Edit
    
    'iterate through recordset
    'rsA.MoveFirst
    'Do Until rsA.EOF
        'add rsA values as new rsOut records
     '   rsOut.AddNew
     '   For iCount = 0 To rsA.Fields.count - 1
     '       rsOut.Fields(iCount).Value = rsA.Fields(iCount).Value
     '   Next
     '   rsOut.Update
     '   rsA.MoveNext
    'Loop
'End With
End If
    Set MergeRecordsets = rsOut

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MergeRecordsets[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Validate Database Objects
' ---------------------------------
' ---------------------------------
' FUNCTION:     DbObjectExists
' Description:  indicate if tables, forms, reports, modules, macros exist in the database
' Assumptions:  -
' Parameters:   strName - name of object (string)
'               oType - type of object (optional string, default - frm)
'                       frm-form, mac-macro, mod-module, qry-query, rpt-report, tbl-table
'                       erd-database diagram, funct-function
' Returns:      true - object exists | false - object doesn't exist in database (boolean)
' Throws:       none
' References:   none
' Source/date:
'   Richard (rapsr59), January 7, 2009
'   https://access-programmers.co.uk/forums/showthread.php?t=161339
'   Microsoft, unknown
'   https://msdn.microsoft.com/en-us/library/office/ff845448.aspx
' Adapted:      Bonnie Campbell, February 22, 2017 - for NCPN tools
' Revisions:
'   BLC - 2/22/2017  - initial version
' ---------------------------------
Public Function DbObjectExists(strName As String, Optional oType As String = "frm") As Boolean
On Error GoTo Err_Handler

    Dim db As Object, db2 As Object, db3 As Object
    Dim obj As AccessObject
    Dim dbColl As Object 'AccessObject 'New Collection
    
    'CurrentProject contains: AllForms, AllMacros, AllModules, AllReports objects
    Set db = Application.CurrentProject
    
    'CurrentData contains: AllTables, AllQueries, AllFunctions objects
    Set db2 = Application.CurrentData
    
    'CodeData contains: AllFunctions object
 '   Set db3 = Application.CodeData
    
    Select Case oType
        Case "erd"
            Set dbColl = db2.AllDatabaseDiagrams
        Case "frm"
            Set dbColl = db.AllForms
'        Case "funct"
'            Set dbColl = db3.AllFunctions '<-- results in 2467: expression...refers to an object that is closed or doesn't exist.
        Case "mac"
            Set dbColl = db.AllMacros
        Case "mod"
            Set dbColl = db.AllModules
        Case "qry"
            Set dbColl = db2.AllQueries
        Case "rpt"
            Set dbColl = db.AllReports
        Case "tbl"
            Set dbColl = db2.AllTables
    End Select
    
    For Each obj In dbColl
    
        If obj.Name = strName Then
            DbObjectExists = True
            GoTo Exit_Handler
        End If
        
    Next

    DbObjectExists = False
    
Exit_Handler:
    'cleanup
    Set db = Nothing
    Set db2 = Nothing
    Set db3 = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DbObjectExists[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' =================================
' FUNCTION:     TableExists
' Description:  Returns whether the specified table exists in the current database collection
' Parameters:   strTableName - string for the name of the table to check
' Returns:      True if the specified table exists in the master systems table, or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 6/29/2009
' Revisions:    JRB, 6/29/2009 - initial version
'               BLC, 4/30/2015 - moved from mod_Utilities
'               BLC, 5/18/2015 - renamed, removed fxn prefix
' =================================
Public Function TableExists(ByVal strTableName As String) As Boolean
    On Error GoTo Err_Handler

    TableExists = DCount("*", "MSysObjects", "(([Type] In (1,4,6)) AND ([Name]=""" & _
        strTableName & """))")

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TableExists[mod_Db])"
    End Select
    Resume Exit_Handler

End Function

' ---------------------------------
' FUNCTION:     DbTableExists
' Description:  determine if a table exists w/in a database
' Assumptions:  used for retrieving table field data for mapping fields, etc.
' Parameters:   tbl - name of database table (string)
'               tdfRefresh - whether to refresh table defs or not (boolean, optional)
'               db - database to reference (DAO.database object)
' Returns:      whether or not table exists (boolean)
' Throws:       none
' References:
'   David W. Fenton, June 7, 2010
'   http://stackoverflow.com/questions/2985513/check-if-access-table-exists
'   Based on testing, when passed an existing db variable, this function is fastest
'   Tony Toews, unknown
'   http://www.granite.ab.ca/access/Temptables.htm
'   David Fenton's functino originally based on Tony Toew's function in TempTables.MDB
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Function DbTableExists(tbl As String, Optional tdfRefresh As Boolean, _
                                Optional db As DAO.Database) As Boolean
On Error GoTo Err_Handler
  
  Dim tdf As DAO.TableDef

  'set db if passed
  If db Is Nothing Then Set db = CurrDb()
  
  'refresh tables
  If tdfRefresh Then db.TableDefs.Refresh
  
  Set tdf = db(tbl)
  
  DbTableExists = True

Exit_Handler:
    'cleanup
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case 3265
            DbTableExists = False
        Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DbTableExists[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SysTablesExist
' Description:  Checks if select system tables exist
' Assumptions:  -
' Parameters:   tblType - string value representing the group of tables to check
'                         either "db" -> backend data tables, links & app defaults
'                         or    "app" -> release version, bugs, user roles & logins
' Returns:      True if all tables exist, false if any do not
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, July 31, 2014 for NCPN WQ Utilities tool
' Adapted:      -
' Revisions:    BLC, 7/31/2014 - initial version
'               BLC, 4/22/2015 - shifted default arrays of sys db & app tables to globals
'                                DB_SYS_TABLES & APP_SYS_TABLES to accommodate & expose settings for
'                                multiple apps (NCPN Invasives Reporting tool), some that do not
'                                contain same/all tables
'               BLC, 10/17/2017 - moved from mod_Initialize_App to mod_Db
' ---------------------------------
Public Function SysTablesExist(tblType As String) As Boolean
On Error GoTo Err_Handler:
Dim sysTables As Variant
Dim i As Integer
Dim missingTable As String

    Select Case tblType
            
        Case "db"
            ' Confirm certain system tables exist --> if not, close the application
            '-----------------------------------------------------------------------
            '   tsys_App_Defaults -> default application settings
            '   tsys_BE_Updates   -> updates to post to remote back-end copies
            '   tsys_Link_Dbs     -> info about linked back-end dbs
            '   tsys_Link_Tables  -> info about linked tables
            '-----------------------------------------------------------------------
            sysTables = Split(DB_SYS_TABLES, ",")

        Case "app"
            ' Confirm certain backend system tables exist --> if not, set connected to false
            '-----------------------------------------------------------------------
            '   tsys_App_Releases -> list of application releases
            '   tsys_Bug_Reports  -> tracking for known issues
            '   tsys_Logins       -> system use monitoring
            '   tsys_User_Roles   -> assign user access priviledges
            '-----------------------------------------------------------------------
            sysTables = Split(APP_SYS_TABLES, ",")
        Case ""
    End Select
        
    For i = 0 To UBound(sysTables)
        If TableExists("tsys_" & Trim(sysTables(i))) = False Then
            missingTable = sysTables(i)
            GoTo Missing_Table:
        End If
    Next
    
    'return result
    SysTablesExist = True
    
Exit_Handler:
    Exit Function
    
Missing_Table:
    Dim strMsg As String
    strMsg = "Unable to find the system table: " & vbCrLf & vbCrLf & vbTab & _
                sysTables(i) & vbCrLf & vbCrLf

    Select Case missingTable
        Case "App_Defaults", "BE_Updates", "Link_Dbs", "Link_Tables", "Link_Files"
            strMsg = strMsg & "Notify the database administrator."
            DoCmd.SetWarnings True
            DoCmd.Quit acQuitSaveNone
        Case "App_Releases", "Bug_Reports", "Logins", "User_Roles"
            ' Close the application if missing one or more systems tables
            strMsg = strMsg & _
                "Either link to the correct back-end or quit and notify the" & vbCrLf & _
                "database administrator before using this application."
            TempVars.Item("Connected") = False
        Case ""
    End Select
    
    'display messages
    MsgBox strMsg, vbCritical, "Application error - Missing system table"
    
    'return result
    SysTablesExist = False

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SysTablesExist[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     QueryExists
' Description:  Determine if a query exists in a database
' Parameters:   strQueryName - query name(string)
' Returns:      true - if found (boolean); false - if not found
' Throws:       -
' References:   -
' Source/date:  SOS, 3/20/2010
'               http://www.access-programmers.co.uk/forums/showthread.php?t=190747
' Adapted:      Bonnie Campbell, May 1, 2015
' Revisions:    BLC, 5/1/2015 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Function QueryExists(strQueryName As String) As Boolean
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.QueryDef
    
    Set db = CurrDb
    Set tdf = db.QueryDefs(strQueryName)
    
    QueryExists = True

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3265
        QueryExists = False
        Resume Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - QueryExists[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          qryExists
' Description:  Checks if query exists in database as a permanent query(QueryDefs)
'               Function retained for backward compatibility
' Parameters:   strQueryName - query name as a string
' Returns:      true - if found (boolean); false - if not found
' Throws:       -
' References:   -
' Source/date:  Nick Vans, January 31, 2008
'               http://bytes.com/topic/access/answers/765384-determine-if-query-x-exists
' Adapted:      Bonnie Campbell, June 17, 2014
' Revisions:    6/17/2014 - BLC - initial version
'               BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                 multiple open connections
' ---------------------------------
Public Function qryExists(strQueryName As String) As Boolean

    Dim qdf As DAO.QueryDef
    
    'default
    qryExists = False
  
    For Each qdf In CurrDb.QueryDefs
'        Debug.Print qdf.Name
        If qdf.Name = strQueryName Then
            qryExists = True
            Exit For
        End If
    Next

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - qryExists[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsRecordset
' Description:  Determines if the object is a recordset or not
' Assumptions:
'               Error handling is ignored since rs.Recordcount would produce
'               an error if rs is not a recordset. In that case isRS remains
'               false and that is returned through the Exit_Handler
' Parameters:   rs - recordset object (object)
' Returns:      isRS - if object was determined to be a recordset (boolean)
'                      true = is a recordset object, false = is not a recordset object
' References:   -
' Source/date:  Bonnie Campbell, October 11 2016
' Revisions:    BLC, 10/11/2016 - initial version
' ---------------------------------
Public Function IsRecordset(rs As Object) As Boolean
On Error GoTo Err_Handler

    Dim isRS As Boolean
    
    isRS = False
    
    If Not rs Is Nothing Then
            
'        If Not IsError(IsNumeric(rs.RecordCount)) Then isRS = True
        If IsNumeric(rs.RecordCount) Then isRS = True
    
    End If

Exit_Handler:
    IsRecordset = isRS
    Exit Function
Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - IsRecordset[mod_Db])"
'    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     HasRecords
' Description:  Returns whether the specified table has records or not
' Parameters:   strName - string for the name of the table or query to check
' Returns:      True if the specified table/query has records, or False if not
' Throws:       none
' References:   Fionnuala, Oct 22, 2010
'               http://stackoverflow.com/questions/3994956/meaning-of-msysobjects-values-32758-32757-and-3-microsoft-access
' Source/date:  Bonnie Campbell, May 26, 2015
' Revisions:    BLC, 5/26/2015 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Function HasRecords(ByVal strName As String) As Boolean
    On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    Dim blnHasRecords As Boolean
    
    blnHasRecords = False
    
    ' check for table/query - 1(table), 4(Linked ODBC), 6(Linked Access), 5(query)
    If DCount("*", "MSysObjects", "(([Type] In (1,4,6,5)) AND ([Name]=""" & _
        strName & """))") > 0 Then
            Set rs = CurrDb.OpenRecordset("SELECT * FROM " & strName & ";")
            
            'check if empty (BOF & EOF = true)
            If Not (rs.BOF And rs.EOF) Then
                blnHasRecords = True
            End If
    End If

    HasRecords = blnHasRecords

Exit_Handler:
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HasRecords[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Db Property Methods/Functions
' ---------------------------------

' ---------------------------------
' SUB:          SetQueryProperty
' Description:  Sets query properties
' Assumptions:  -
' Parameters:   qdf - query to modify (DAO.QueryDef)
'               prop - name of property to add (string)
'               val - value of new property (variant)
' Returns:      -
' Throws:       none
' References:
'   LPurvis, September 13, 2008
'   http://www.utteraccess.com/forum/Set-query-property-VBA-t1713084.html
' Source/date:  -
' Adapted:      Bonnie Campbell, March 30, 2017 - for NCPN tools
' Revisions:
'   BLC - 3/30/2017 - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Sub SetQueryProperty(qdf As DAO.QueryDef, prop As String, val As Variant) 'qry As String, prop As String, val As Variant)
On Error Resume Next
'    Dim db As Database
'    Dim qdf As QueryDef
    Dim prp As DAO.Property
    
'    Set db = CurrDb
'    Set qdf = db.QueryDefs(qry)
    
    With qdf
        Set prp = qdf.Properties(prop)
        If Err Then
            Set prp = .CreateProperty(prop, dbText, val)
            .Properties.Append prp
        Else
            prp.Value = val
        End If
    End With
    
Exit_Handler:
    'cleanup
'    Set prp = Nothing
'    Set qdf = Nothing
'    Set db = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetQueryProperty[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetDescription
' Description:  retrieves object description property
' Assumptions:  object has a description property
' Parameters:   obj - item to check (object)
' Returns:      description text for object (string)
' Throws:       none
' References:
'   Allen Browne, April, 2010
'   http://allenbrowne.com/func-06.html
' Source/date:  Bonnie Campbell, September 2016 for NCPN tools
' Revisions:    BLC, 9/16/2016 - initial version
' ---------------------------------
Public Function GetDescription(obj As Object) As String
On Error GoTo Err_Handler

    GetDescription = obj.Properties("Description")

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescription[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     FieldCount
' Description:  Determines the number of fields in a table/query
' Assumptions:  -
' Parameters:   TableName - name of table/query (variant)
' Returns:      FieldCount - number of fields (variant)
' References:
'   Sinndho, May 8, 2012
'   http://www.dbforums.com/showthread.php?1678970-Count-the-number-of-columns-(fields)-in-a-table
' Source/date:  Bonnie Campbell, October 11 2016
' Revisions:    BLC, 10/11/2016 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                               multiple open connections
' ---------------------------------
Public Function FieldCount(ByVal TableName As String) As Long
'Public Function FIeldCount(ByVal TableName As Variant) As Variant <<-- if including in query
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset

    Set rs = CurrDb.OpenRecordset(TableName, dbOpenSnapshot)
    
    FieldCount = rs.Fields.Count

Exit_Handler:
    'cleanup
    rs.Close
    Set rs = Nothing
    
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FieldCount[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     MaxDbFieldCount
' Description:  Determines the maximum number of fields in Db tables/queries
' Assumptions:  -
' Parameters:   -
' Returns:      MaxDbFieldCount - maximum number of fields (long)
' References:
'   Sinndho, May 8, 2012
'   http://www.dbforums.com/showthread.php?1678970-Count-the-number-of-columns-(fields)-in-a-table
' Source/date:  Bonnie Campbell, October 11 2016
' Revisions:    BLC, 10/11/2016 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Function MaxDbFieldCount() As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim max As Long
    Dim qtName As String
    
    Set db = CurrDb
    
    'default
    max = 0
    
    For Each tdf In db.TableDefs
        
        If tdf.Fields.Count > max Then
            max = tdf.Fields.Count
            qtName = tdf.Name
        End If

    Next
    
    For Each qdf In db.QueryDefs
    
        If qdf.Fields.Count > max Then
            max = qdf.Fields.Count
            qtName = qdf.Name
        End If

    Next
    
    Debug.Print qtName
    
    MaxDbFieldCount = max

Exit_Handler:
    'cleanup
    Set tdf = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MaxDbFieldCount[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsLinked
' Description:  Determines if a table is linked
' Assumptions:  -
' Parameters:   tblName - name of table to evaluate (string)
' Returns:      IsLinked - whether table is linked (boolean)
'                          returns true for types 4 (ODBC linked), 6 (other linked)
'                                  false for type 1 (non-linked tables)
' References:
'   Douglas J. Steele, February 20, 2009
'   http://www.pcreview.co.uk/threads/check-if-a-table-is-linked.3748757/
' Source/date:  Bonnie Campbell, October 20, 2016
' Revisions:    BLC, 10/20/2016 - initial version
' ---------------------------------
Public Function IsLinked(tblName As String) As Boolean
On Error GoTo Err_Handler
    
    IsLinked = Nz(DLookup("Type", "MSysObjects", "Name='" & tblName & "'"), 0) <> 1

Exit_Handler:
    'cleanup
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsLinked[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          RetrieveTableColumnData
' Description:  Retrieves table column names & attributes
' Assumptions:  -
' Parameters:   tbl - table name (string)
' Returns:      array of column data (variant, 2-element array)
'                 0 - column data recordset (rs)
'                 1 - table column data as a comma separated string (string)
' Throws:       none
' References:   none
' Source/date:  -
' Adapted:      Bonnie Campbell, January 19, 2017 - for NCPN tools
' Revisions:
'   BLC - 1/19/2017 - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function RetrieveTableColumnData(tbl As String) As Variant
On Error GoTo Err_Handler

    'retrieve field info
    Dim aryFieldInfo() As Variant 'string
    
    aryFieldInfo = FetchDbTableFieldInfo(tbl)
    
    'clear table
    ClearTable "usys_temp_rs"

    'populate w/ table data
    Dim rs As DAO.Recordset
    Dim aryRecord() As String
    Dim i As Integer
    Dim strTableColumns As String
    
    'default
    strTableColumns = ""
    
    Set rs = CurrDb.OpenRecordset("usys_temp_rs", dbOpenDynaset)
    
    For i = 0 To UBound(aryFieldInfo)
        
        'create new record
        rs.AddNew
        
        aryRecord = Split(aryFieldInfo(i), "|")
        
        rs!Column = aryRecord(0)
        rs!ColType = aryRecord(5)
        rs!IsReqd = IIf(aryRecord(3) = False, 0, 1)
        rs!Length = aryRecord(2)
        rs!AllowZLS = IIf(aryRecord(4) = False, 0, 1)
    
        'add the new record
        rs.Update
        
        'prepare table columns list
        strTableColumns = strTableColumns & aryRecord(0) & ", "
        
    Next
    
    Dim ary() As Variant
    ary = Array(rs, strTableColumns)
    
    RetrieveTableColumnData = ary
    
Exit_Handler:
    'cleanup
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RetrieveTableColumnData[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' https://support.office.com/en-us/article/VarType-Function-1e08636c-1892-40c2-aff3-2b894389e82d?ui=en-US&rs=en-US&ad=US&fromAR=1
'   vbEmpty         0   Empty (uninitialized)
'   vbNull          1   Null (no valid data)
'   vbInteger       2   Integer
'   vbLong          3   Long integer
'   vbSingle        4   Single-precision floating-point number
'   vbDouble        5   Double-precision floating-point number
'   vbCurrency      6   Currency value
'   vbDate          7   Date Value
'   vbString        8   String
'   vbObject        9   Object
'   vbError         10  Error Value
'   vbBoolean       11  Boolean value
'   vbVariant       12  Variant (used only with arrays of variants)
'   vbDataObject    13  A data access object
'   vbDecimal       14  Decimal value
'   vbByte          17  Byte value
'   vbUserDefinedType 36    Variants that contain user-defined types
'   vbArray         8192    Array

' ---------------------------------
' FUNCTION:     FieldTypeName
' Description:  retrieves field type property name from the numeric field type
' Assumptions:  -
' Parameters:   fld - field to retrieve type for (DAO.field)
' Returns:      name for the field type (string)
' Throws:       none
' References:
'   Allen Browne, April, 2010
'   http://allenbrowne.com/func-06.html
'   TofuBug     May 28, 2015
'   http://stackoverflow.com/questions/30511987/why-does-vartype-always-return-8204-for-arrays
' Source/date:  Bonnie Campbell, September 2016 for NCPN tools
' Revisions:    BLC, 9/16/2016 - initial version
' ---------------------------------
Public Function FieldTypeName(fld As DAO.field) As String
On Error GoTo Err_Handler

    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) ' fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText

'        'Arrays
'        Case vbArray:
'            strReturn = "Array"                     '8192
'
'        Case Is > 8192
'            Select Case (fld.Type - 8192)
'                Case vbString                       '8 --> Overall 8200 = 8192+8
'                    strReturn = "String Array"
'                Case vbVariant                      '12 --> Overall 8204 = 8192+12
'                    strReturn = "Variant Array"
'                Case Else
'                    strReturn = "Field type " & fld.Type & " unknown"
'            End Select
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FieldTypeName[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     VarTypeName
' Description:  retrieves var type name from the numeric variable type
' Assumptions:  -
' Parameters:   vType - numeric type (integer)
' Returns:      name for the type (string)
' Throws:       none
' References:
'   Allen Browne, April, 2010
'   http://allenbrowne.com/func-06.html
'   TofuBug     May 28, 2015
'   http://stackoverflow.com/questions/30511987/why-does-vartype-always-return-8204-for-arrays
' Source/date:  Bonnie Campbell, September 2016 for NCPN tools
' Revisions:    BLC, 9/16/2016 - initial version
' ---------------------------------
Public Function VarTypeName(vType As Integer) As String
On Error GoTo Err_Handler
    
    Dim strReturn As String    'Name to return

    Select Case CLng(vType) ' vType is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
'            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
'            Else
'                strReturn = "AutoNumber"
'            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
'            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
'            Else
'                strReturn = "Text (fixed width)"        '(no interface)
'            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
'            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
'            Else
'                strReturn = "Hyperlink"
'            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        
        'Arrays
        Case vbArray:
            strReturn = "Array"                     '8192

        Case Is > 8192
            Select Case (vType - 8192)
                Case vbString                       '8 --> Overall 8200 = 8192+8
                    strReturn = "String Array"
                Case vbVariant                      '12 --> Overall 8204 = 8192+12
                    strReturn = "Variant Array"
                Case Else
                    strReturn = "Field type " & vType & " unknown"
            End Select
        Case Else: strReturn = "Field type " & vType & " unknown"
    End Select

    VarTypeName = strReturn

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VarTypeName[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     FetchDbTableFieldInfo
' Description:  retrieves field information from a database table
' Parameters:   tbl - name of database table (string)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  used for retrieving table field data for mapping fields, etc.
' Throws:       none
' References:
'   David W. Fenton, July 27, 2010
'   http://stackoverflow.com/questions/3343922/get-column-names
'   HansUp, October 18, 2013
'   http://stackoverflow.com/questions/19452952/how-to-count-number-of-fields-in-a-table
'   Stuart McCall, Sep 24, 2010
'   https://www.pcreview.co.uk/threads/stop-all-vba-code-running.4025375/
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
'               BLC, 1/19/2017 - added error handling for non-table inputs
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function FetchDbTableFieldInfo(tbl As String) As Variant 'DAO.Recordset
On Error GoTo Err_Handler

    Dim blnNoTable As Boolean
    'default
    blnNoTable = False
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset ', rsFields As ADODB.Recordset
    Dim fld As DAO.field
    Dim aryFieldInfo() As Variant
    Dim icols As Integer, iCol As Integer
    Dim strTypeName As String
    
    Set db = CurrDb()
    
    'determine if table is in database
    If Not DbTableExists(tbl) Then
        blnNoTable = True
        GoTo Err_Handler
    End If
    
    Set rs = db.OpenRecordset(tbl)
    
'    Set rsFields = New ADODB.Recordset
    
    'get count
    icols = rs.Fields.Count
    iCol = 0
    
    ReDim Preserve aryFieldInfo(0 To icols - 1)
    
    'iterate through fields
    For Each fld In rs.Fields
'        Debug.Print fld.Name
'        Debug.Print fld.Attributes
'        Debug.Print fld.Size
'        Debug.Print fld.Properties
'        Debug.Print fld.Required
'        Debug.Print fld.Type
'        Debug.Print fld.ValidationRule
'        With rsFields
'            .Append
'        End With
        
'        Debug.Print (fld)
        
        'fetch name for type
'        GetFieldTypeName fld
        strTypeName = VarTypeName(fld.Type)

        With fld
                
            'prepare array of info
            aryFieldInfo(iCol) = .Name & "|" & _
                            .Type & "|" & _
                            .Size & "|" & _
                            .Required & "|" & _
                            .AllowZeroLength & "|" & _
                            strTypeName

        End With
        
        iCol = iCol + 1
    Next

    FetchDbTableFieldInfo = aryFieldInfo

    'cleanup
    'Set fld = Nothing
'    Set rs = Nothing
'    Set db = Nothing
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 0
        If blnNoTable Then _
        MsgBox "The table name (" & tbl & ") provided does not exist or was typed " _
            & "incorrectly." & vbCrLf & vbCrLf _
            & "Please check it and try again or contact your data manager.", vbCritical, _
            "Error: Table Doesn't Exist ( FetchDbTableFieldInfo[mod_Db] )"
        'quit the process! (otherwise additional errors will occur w/in calling subs
        End
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FetchDbTableFieldInfo[mod_Db])"
    End Select
    Resume Exit_Handler
End Function
    
' ---------------------------------
' FUNCTION:     ListTables
' Description:  List database tables
' Assumptions:  -
' Parameters:   ShowMSysTables - whether or not to show msys_ tables (boolean)
'               ShowTsysTables - whether or not to show tsys_ tables (boolean)
'               ShowUsysTables - whether or not to show usys_ tables (boolean)
'               ShowLinkedTables - whether or not to show linked tables (boolean)
' Returns:      tables - delimited string of tables (string)
' References:   -
'   Daniel Pineault, June 10, 2010
'   http://www.devhut.net/2010/06/10/ms-access-vba-list-the-tables-in-a-database/
'   HansUp, December 17, 2013
'   http://stackoverflow.com/questions/20643263/how-can-one-search-tabledefs-for-linked-tables
' Source/date:  Bonnie Campbell, October 6 2016
' Revisions:    BLC, 10/6/2016 - initial version
'               BLC, 10/20/2016 - revised to include linked tables, added Tsys, Usys parameters
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
'               BLC, 1/12/2018 - update to eliminate ~TMP... tables, refresh TableDefs
' ---------------------------------
Public Function ListTables(ShowMSysTables As Boolean, _
                            ShowTSysTables As Boolean, _
                            ShowUSysTables As Boolean, _
                            ShowLinkedTables As Boolean) As String
On Error GoTo Err_Handler

    Dim tdf As DAO.TableDef
    Dim tbls As String
    
    'default
    tbls = ""
    
    'refresh tabledefs
    CurrDb.TableDefs.Refresh
    
    'fetch tables
    For Each tdf In CurrDb.TableDefs
'Debug.Print tdf.Name

        'handle MSys tables
        If Len(tdf.Name) > Len(Replace(tdf.Name, "MSys", "")) And ShowMSysTables = False Then GoTo Continue
        
        'handle tsys tables
        If Len(tdf.Name) > Len(Replace(tdf.Name, "tsys", "")) And ShowMSysTables = False Then GoTo Continue
                
        'handle usys tables
        If Len(tdf.Name) > Len(Replace(tdf.Name, "usys", "")) And ShowMSysTables = False Then GoTo Continue
        
        'handle linked tables
        If Len(tdf.connect) > 0 And ShowLinkedTables = False Then GoTo Continue
        
        'handle temp tables beginning w/ ~  (e.g. ~TMPCLP535461)
        If Left(tdf.Name, 1) = "~" Then GoTo Continue
        
        tbls = tbls & "|" & tdf.Name
        
Continue:
    Next
    
    'trim starting delimiter
    tbls = Right(tbls, Len(tbls) - 1)
'    Debug.Print tbls
    
Exit_Handler:
    ListTables = tbls
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ListTables[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Object Level Methods/Functions
' ---------------------------------

' ---------------------------------
' SUB:          CloseObject
' Description:  Checks if object exists, closes it if it does
' Assumptions:  Object does not require saving (acSaveNo)
' Parameters:   obj - object to close (variant)
'               oType - object type (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  -
' Adapted:      Bonnie Campbell, March 28, 2017 - for NCPN tools
' Revisions:
'   BLC - 3/28/2017 - initial version
' ---------------------------------
Public Sub CloseObject(obj As Variant, oType As String)
On Error GoTo Err_Handler

    Dim oGrp As AcObjectType
    
    Select Case LCase(oType)
        Case "qry"
            oGrp = acQuery
        Case "tbl"
            oGrp = acTable
        Case "frm"
            oGrp = acForm
        Case "rpt"
            oGrp = acReport
    End Select

    DoCmd.Close oGrp, obj, acSaveNo
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CloseObject[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     ConvertObjectToPointer
' Description:  Retrieve the pointer to the referenced object
' Parameters:   obj - reference to object to check
' Returns:      pointer to referenced object (long)
' Throws:       none
' References:
' Terry Harris, August 31, 2011
' http://www.utteraccess.com/forum/Pass-Object-Openargs-t1967468.html
' Source/date:  Bonnie Campbell, January 2, 2018
' Revisions:    BLC, 1/2/2018 - initial version
' ---------------------------------
Public Function ConvertObjectToPointer(ByRef obj As Object) As Long
On Error GoTo Err_Handler
    
    Dim objPointer As Long

    'RtlMoveMemory objPointer, obj, POINTERSIZE
    CopyMem objPointer, obj, POINTERSIZE

    ConvertObjectToPointer = objPointer

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConvertObjectToPointer[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ConvertPointerToObject
' Description:  Retrieve the referenced object from its pointer (long)
' Parameters:   ObjPointer - pointer to object (long)
' Returns:      referenced object (object)
' Throws:       none
' References:
' Terry Harris, August 31, 2011
' http://www.utteraccess.com/forum/Pass-Object-Openargs-t1967468.html
' Source/date:  Bonnie Campbell, January 2, 2018
' Revisions:    BLC, 1/2/2018 - initial version
' ---------------------------------
Public Function ConvertPointerToObject(ByVal objPointer As Long) As Object
On Error GoTo Err_Handler
    
    Dim obj As Object

    'RtlMoveMemory obj, objPointer, POINTERSIZE
    CopyMem obj, objPointer, POINTERSIZE

 'Set ConvertPointerToObject = obj
    Set ConvertPointerToObject = obj

'RtlMoveMemory obj, ZEROPOINTER, POINTERSIZE
    CopyMem obj, ZEROPOINTER, POINTERSIZE

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ConvertPointerToObject[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Template Related Methods/Functions
' ---------------------------------

' ---------------------------------
' SUB:     GetTemplates
' Description:  loads Templates into memory as a global dictionary object (dictTemplates)
'               makes current Templates available without querying the db tsys_SQL_Templates table
' Parameters:   strSyntax - specifies syntax of the Template to retrieve (T-SQL, JET, etc.)
'               strParams - specifies the parameters & their datatypes for the Template
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   tsys_Db_Templates, Microsoft Scripting Runtime (dictionary object)
'   HansUp, June 27, 2013
'   http://stackoverflow.com/questions/17328092/how-to-display-access-query-results-without-having-to-create-Temporary-query
' Source/date:  Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - initial version
'               BLC, 5/13/2016 - shifted from mod_Db_Templates to mod_Db & adjusted to match tsys_Db_Templates
'               BLC, 5/19/2016 - revised documentation & renamed GetTemplates() vs. GetSQLTemplates() since tsys_Db_Templates
'                                can accommodate more than SQL
'               BLC, 6/5/2016  - revised to set strSyntax to "T-SQL" to avoid error due to multiple items of same name in dict
'               BLC, 6/6/2016  - added error handling for duplicate Templates, renamed global to g_AppTemplates
'               BLC, 2/1/2017  - added error handling for improper parameter syntax (param name:param type)
'               BLC, 2/2/2017  - added clearing of global g_AppTemplates to allow re-definition
'                                without restarting db
'               BLC, 3/23/2017 - revised to use "SQL" for default syntax (most should be "SQL" i.e. usable in SQL server & Access)
'               BLC, 3/30/2017 - added ID, Dependencies, FieldCheck, FieldOK values to dictionary object
'                                to capture these properties of a Template
'               BLC, 4/18/2017 - revised Dim to set dictTemplates as Scripting.Dictionary vs. Dictionary (latter
'                                produces compile error on .Add - Method or Data Member not found)
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Sub GetTemplates(Optional strSyntax As String = "", Optional Params As String = "")

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, Key As String
    Dim Value As Variant
    
    'handle default
    strSQLWhere = " WHERE IsSupported > 0"
    
    If Len(strSyntax) = 0 Then
        strSyntax = "SQL"
    End If
    
    strSQLWhere = strSQLWhere & " AND LCase(Syntax) = LCase('" & strSyntax & "')"
    
    'sql -> ID, Version, IsSupported, Context, Syntax, TemplateName, Params, Template, Remarks,
    '       EffectiveDate, RetireDate, CreateDate, CreatedBy_ID, LastModified, LastModifiedBy_ID
    strSQL = "SELECT * FROM tsys_Db_Templates" & strSQLWhere & ";"
    
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)
    
    'handle no records
    If rs.EOF Then
        MsgBox "Sorry, no Templates were found for this database version.", vbExclamation, _
            "Linked Database Templates Not Found"
        DoCmd.CancelEvent
        GoTo Exit_Handler
    End If
    
    'prepare dictionary
    Dim dict As New Scripting.Dictionary, dictParam As New Scripting.Dictionary
    Dim ary(1 To 11) As String, ary2() As String, param() As String
    Dim i As Integer, j As Integer
    
    'prepare the dictionary key array
    ary(1) = "Context"
    ary(2) = "TemplateName"
    ary(3) = "Template" 'Template
    ary(4) = "Params"
    ary(5) = "Syntax"
    ary(6) = "ID"
    ary(7) = "FieldCheck"
    ary(8) = "FieldOK"
    ary(9) = "Dependencies"
    ary(10) = "DataScope"
    ary(11) = "Version"
    
    'prepare array of dictionaries
    Dim dictTemplates As Scripting.Dictionary
    Set dictTemplates = New Scripting.Dictionary
    
    rs.MoveLast
    rs.MoveFirst
    Do Until rs.EOF
        'create new dictionary object
        Set dict = New Scripting.Dictionary
        
        'populate the dictionary
        For i = 1 To UBound(ary)
            
            Key = ary(i)
            
            If Key = "Params" Then
                'create new dictionary for param name & data type
                Set dictParam = New Scripting.Dictionary

'Debug.Print rs.Fields(ary(i))

                'separate parameters
                ary2 = Split(Nz(rs.Fields(ary(i)), ":"), "|")
 
                'prepare sets of param name & data type --> split(ary2(i), ":") yields name & data type
                For j = 0 To UBound(ary2)
                
                    'split the param into name & data type
                    param = Split(ary2(j), ":")
                                        
                    If Not dictParam.Exists(param(0)) And Len(param(0)) <> 0 Then
                        
                        'catch parameters not in paramname:type format
                        If UBound(param) <> 1 Then
                            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                            "msg" & PARAM_SEPARATOR & "Parameter format must be name:type.  " _
                            & "Please contact a data manager to resolve this issue.  " _
                            & "Db says: ""I can't work this way, so I'm closing now.""" _
                            & "|Type" & PARAM_SEPARATOR & "caution" _
                            & "|Caption" & PARAM_SEPARATOR & "Invalid SQL Template Parameters for the '" & dict("TemplateName") & "' Template"
                            
                            'exit database since application won't function w/o valid Templates
                            DoCmd.CloseDatabase
                        End If
                            
                        dictParam.Add param(0), param(1)

                    End If
                
                Next
                
                Set Value = dictParam

            Else
                Value = Nz(rs.Fields(ary(i)), "")
            End If
            
            'add key if it isn't already there
            If Not dict.Exists(Key) Then
                If IsNull(Value) Then MsgBox Key, vbOKCancel, "is NULL"
                'Debug.Print Nz(Value, key & "-NULL")
                dict.Add Key, Value
            End If
        
        Next
        
        'add Template dictionary to dictionary of Templates
        dictTemplates.Add dict("TemplateName"), dict
        
'        Debug.Print dict("TemplateName") & " " & dict.Item("ID")
        rs.MoveNext
    Loop
    
    'load global AppTemplates As Scripting.Dictionary of Templates
    Set g_AppTemplates = Nothing    'clear first
    
    Set g_AppTemplates = dictTemplates
    
Exit_Handler:
    'cleanup
    Set rs = Nothing
    Set dict = Nothing
    Set dictTemplates = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 457  'Duplicate Template -- tsys_Db_Templates finds more than one w/ same name
        MsgBox "A duplicate Template was found." & vbCrLf & vbCrLf & _
            "When you click 'OK' a query will run to identify the problem Template." & vbCrLf & vbCrLf & _
            "You can close the query after it runs (save it if you like)." & vbCrLf & vbCrLf & _
            "Please contact your data manager to resolve this issue." & vbCrLf & vbCrLf & _
            "Error #" & Err.Number & " - GetTemplates[mod_Db]:" & vbCrLf & _
            Err.Description, vbExclamation, "Duplicate Db Template Found! [tsys_Db_Templates]"

            Dim strErrorSQL As String
            strErrorSQL = "SELECT TemplateName, Count(TemplateName) AS NumberOfDupes " & _
                    "FROM tsys_Db_Templates " & _
                    "GROUP By TemplateName " & _
                    "HAVING Count(TemplateName) > 1;"

            Dim qdf As DAO.QueryDef
            
            If Not QueryExists("UsysTempQuery") Then
                Set qdf = CurrDb.CreateQueryDef("UsysTempQuery")
            Else
                Set qdf = CurrDb.QueryDefs("UsysTempQuery")
            End If
            
            qdf.SQL = strErrorSQL
            
            DoCmd.OpenQuery "USysTempQuery", acViewNormal

            '********** FATAL ERROR ****************
            'terminate *ALL* VBA code to prevent other popups
            'Exit Sub
            Stop
            
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetTemplates[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetTemplateIDs
' Description:  retrieves Template numeric IDs from Templates global Template dictionary (AppTemplates)
'               returns them as a dictionary object with ID:Template name as key:value pair
' Parameters:   -
' Returns:      Template IDs - Template IDs (dictionary object, key=ID, value=Template name)
' Assumptions:  tsys_Db_Templates contains all desired Templates
'               g_AppTemplates contains Template info searchable by TemplateName which is its key
'               each of the values in that pair is itself a dictionary object containing
'               the various Template properties (ID, SQL, Version, etc.)
' Throws:       none
' References:   tsys_Db_Templates, Microsoft Scripting Runtime (dictionary object)
'   Craig Hatmaker, December 5, 2012
'   http://stackoverflow.com/questions/11296522/looping-through-a-scripting-dictionary-using-index-item-number
' Source/date:  Bonnie Campbell, March 30, 2017
' Revisions:    BLC, 3/30/2017 - initial version
'               BLC, 3/31/2017 - fix issue which caused g_AppTemplateIDs to report wrong ID for a Template
' ---------------------------------
Public Function GetTemplateIDs() As Scripting.Dictionary
On Error GoTo Err_Handler

    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates
    
    Dim d As Scripting.Dictionary, tIDs As Scripting.Dictionary
    Dim i As Integer
    Dim X As Variant
    
    Set d = g_AppTemplates
    
    Set tIDs = CreateObject("Scripting.Dictionary")
    
    'iterate through the global Template dictionary
    'For i = 0 To d.Count - 1
    For Each X In d
    
        'add @ Template to the dictionary
        '--------------------------------------------------------------
        ' Note: @ of the global Template dictionary's items is itself a dictionary
        '       so reference them via d.Items()(i).Item("keyname")
        '--------------------------------------------------------------
        tIDs.Add d.Item(X).Item("ID"), d.Item(X).Item("TemplateName")
        'tIDs.Add d.Items()(x).Item("ID"), d.Items()(x).Item("TemplateName")
        'tIDs.Add d.Items()(i).Item("ID"), d.Items()(i).Item("TemplateName")
        'tIDs.Add d.Items()(i).Item("ID"), d.Keys()(i)
        ' Debug.Print tIDs.Keys()(i) & " - " & tIDs.Items()(i)
    Next 'i
    
    'set global
    Set g_AppTemplateIDs = tIDs
    
    Set GetTemplateIDs = tIDs
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetTemplateIDs[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetTemplate
' Description:  retrieves Template from Templates global Template dictionary (AppTemplates)
' Parameters:   strTemplate - name of Template to fetch (string)
'               params - pipe (|) separated parameter listing w/ parameter name:value pairs (: separated) (string)
' Returns:      Template - value of the Template (string)
'               most Templates are SQL strings, so the SQL string (Template) field of the given
'               Template name is retrieved
' Assumptions:  tsys_Db_Templates correctly list parameter:parameter type values & AppTemplates contain them
'               params do not include PARAM_SEPARATOR w/in them as this is considered a separator
' Throws:       none
' References:   tsys_Db_Templates, Microsoft Scripting Runtime (dictionary object)
'   HansUp, June 27, 2013
'   http://stackoverflow.com/questions/17328092/how-to-display-access-query-results-without-having-to-create-Temporary-query
' Source/date:  Bonnie Campbell, May 2016
' Revisions:    BLC, 5/19/2016 - initial version
'               BLC, 6/6/2016  - added error handling for duplicate Templates, renamed global to g_AppTemplates
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Function GetTemplate(strTemplate As String, Optional Params As String = "") As String
On Error GoTo Err_Handler

    Dim aryParams() As Variant
    Dim ary() As String, ary2() As String
    Dim i As Integer
    Dim template As String, swap As String, param As String

Debug.Print strTemplate

    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates

    template = g_AppTemplates(strTemplate).Item("Template")
    
    If Len(Params) > 0 Then
    
        'prepare passed in param array --> array contains param:value pairs
        'ary = Split(params, "|")
        If InStr(Params, "|") Then
            ary = Split(Params, "|")
        Else
            ReDim Preserve ary(0) 'avoid Error #9 subscript out of range
            ary(0) = Params
            'ary = Split(params, PARAM_SEPARATOR)
        End If
        
        'prepare array of Template parameters w/ their data type
        'aryParams = Split(AppTemplates(strTemplate).item("Params"), "|")
        'AppTemplates("s_tagline").Item("Params").Item("SourceID") --> integer
    
        'iterate through params
        For i = 0 To UBound(ary)
            
            'split name:value pair --> ary2(0) = name, ary2(1) = value
            'If InStr(ary(1), PARAM_SEPARATOR) Then
            If InStr(ary(i), PARAM_SEPARATOR) Then
                ary2 = Split(ary(i), PARAM_SEPARATOR)
            Else
                ary2 = Split(ary(i), ":")
            End If
            'compare datatype to aryParams value
            If IsTypeMatch(ary2(1), g_AppTemplates(strTemplate).Item("Params").Item(ary2(0))) Then
                
                'prepare replaced value
                swap = "[" & ary2(0) & "]"
                
                'SQL-ize parameter values to avoid SQL syntax errors
                param = SQLencode(ary2(1))
'Debug.Print param
                'swap out the placeholder in the Template
                template = Replace(template, swap, ary2(1))
                
            End If
            
        Next
    
    End If
    
'Debug.Print Template
    
    GetTemplate = template
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 457  'Duplicate Template -- tsys_Db_Templates finds more than one w/ same name
        MsgBox "A duplicate Template was found." & vbCrLf & vbCrLf & _
            "When you click 'OK' a query will run to identify the problem Template." & vbCrLf & vbCrLf & _
            "You can close the query after it runs (save it if you like)." & vbCrLf & vbCrLf & _
            "Please contact your data manager to resolve this issue." & vbCrLf & vbCrLf & _
            "Error #" & Err.Number & " - GetTemplate[mod_Db]:" & vbCrLf & _
            Err.Description, vbExclamation, "Duplicate Db Template Found! [tsys_Db_Templates]"

            Dim strErrorSQL As String
            strErrorSQL = "SELECT TemplateName, Count(TemplateName) AS NumberOfDupes " & _
                    "FROM tsys_Db_Templates " & _
                    "GROUP By TemplateName " & _
                    "HAVING Count(TemplateName) > 1;"

            Dim qdf As DAO.QueryDef
            
            If Not QueryExists("UsysTempQuery") Then
                Set qdf = CurrDb.CreateQueryDef("UsysTempQuery")
            Else
                Set qdf = CurrDb.QueryDefs("UsysTempQuery")
            End If
            
            qdf.SQL = strErrorSQL
            
            DoCmd.OpenQuery "USysTempQuery", acViewNormal

            '********** FATAL ERROR ****************
            'terminate *ALL* VBA code to prevent other popups
            End
        
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetTemplate[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     RefreshTemplates
' Description:  Refreshes global template dictionary
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 13, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/13/2017 - initial version
' ---------------------------------
Public Function RefreshTemplates() As Boolean
On Error GoTo Err_Handler

    'clear existing templates
    Set g_AppTemplates = Nothing
    
    'run template generation to refresh
    GetTemplates
    
    'assume template dictionary is refreshed
    RefreshTemplates = True
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshTemplates[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     VirtualDAORecordset
' Description:  prepares a virtual -in memory only- DAO recordset
' Parameters:   strTable - name of virtual table (string)
'               iCount - number of records (integer)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  the virtual recordset is used in limited instances when
'               a recordset is needed but doesn't exist
' Throws:       none
' References:
'   Tom van Stiphout, July 17, 2006
'   https://bytes.com/topic/access/answers/512790-dao-connectionless-recordset
' Source/date:  Bonnie Campbell, May 2016
' Revisions:    BLC, 5/26/2016 - initial version
' ---------------------------------
Public Function VirtualDAORecordset(iCount As Integer, Optional strTable As String = "Temp") As Recordset
On Error GoTo Err_Handler

    Dim Counter As Long
    Dim rs As DAO.Recordset
    Dim i As Integer

    With DBEngine
        .BeginTrans
        With .Workspaces(0)(0)

            .Execute "CREATE TABLE " & strTable _
                    & "(RecCount INT CONSTRAINT RecCount UNIQUE);"

            Set rs = .OpenRecordset(strTable, dbOpenTable)
            With rs
                For i = 1 To iCount
                    .AddNew
                    .Fields("RecCount") = i
                    .Update
                Next
                
                .index = "RecCount"
                '.Close
            End With
        End With
    End With

    Set VirtualDAORecordset = rs

Exit_Handler:
    'Set rs = Nothing
    'DBEngine.Rollback
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 3010
        Counter = Counter + 1
        strTable = "Temp" & CStr(Counter)
        Resume Next
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - VirtualDAORecordset[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     RemoveVirtualDAORecordset
' Description:  removes a virtual -in memory only- DAO recordset
' Parameters:   strTemplate - name of virtual table (string)
'               iCount - number of records (integer)
' Returns:      rs - recordset containing # of records = iCount (DAO.recordset)
' Assumptions:  the virtual recordset is used in limited instances when
'               a recordset is needed but doesn't exist
' Throws:       none
' References:
'   Tom van Stiphout, July 17, 2006
'   https://bytes.com/topic/access/answers/512790-dao-connectionless-recordset
' Source/date:  Bonnie Campbell, May 2016
' Revisions:    BLC, 5/26/2016 - initial version
' ---------------------------------
Public Sub RemoveVirtualDAORecordset()
On Error GoTo Err_Handler

    DBEngine.Rollback

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveVirtualDAORecordset[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     CreateVirtualADORecordset
' Description:  creates a virtual -in memory only- ADO recordset
' Parameters:   strTemplate - name of virtual table (string)
'               iCount - number of records (integer)
' Returns:      rs - recordset containing # of records = iCount (ADO.recordset)
' Assumptions:  the virtual recordset is used in limited instances when
'               a recordset is needed but doesn't exist
' Throws:       none
' References:
'   Danny Lesandrini, November 2, 2009
'   http://www.databasejournal.com/features/msaccess/article.php/3846361/Create-In-Memory-ADO-Recordsets.htm
' Source/date:  Bonnie Campbell, June 2016
' Revisions:    BLC, 6/8/2016 - initial version
' ---------------------------------
Public Sub CreateVirtualADORecordset(iCount As Integer)
On Error GoTo Err_Handler

'    Dim rsADO As ADODB.Recordset
'    Dim fld As ADODB.Field
'
'    'create rs
'    Set rsADO = New ADODB.Recordset
'    With rsADO
'        .Fields.Append "Number", adInteger, , adFldMayBeNull
'
'        .CursorType = adOpenKeyset
'        .CursorLocation = adUseClient
'        .LockType = adLockPessimistic
'        .Open
'    End With
'
'    'populate rs
'    For i = 0 To iCount - 1
'        rsADO.AddNew
'    Next


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateVirtualADORecordset[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddRecords
' Description:  adds records to a recordset & table
' Assumptions:  -
' Parameters:   rs - (DAO.Recordset)
'               aryCols - field/column names (string array)
'               aryData - data for each record (variant array)
'               delimiter - separator(string)
' Returns:      -
' References:
'   simoco, February 9, 2014
'   http://stackoverflow.com/questions/21885101/can-you-use-a-variable-for-the-field-name-when-using-addnew-to-a-record-set
' Source/date:  Bonnie Campbell, September 20 2016
' Revisions:    BLC, 9/20/2016 - initial version
' ---------------------------------
Public Sub AddRecords(rs As DAO.Recordset, aryCols() As String, aryData() As Variant, _
                            delimiter As String)
On Error GoTo Err_Handler

    Dim aryRecord As String
    Dim i As Integer, j As Integer
    Dim strColName As String

    With rs
        
        'add new record
        .AddNew
        
        
        'iterate through data records
        For i = 0 To UBound(aryData)
        
            'get record array
            aryRecord = Split(aryData(i), delimiter)
            
            'iterate through columns
            For j = 0 To UBound(aryCols)
            
                strColName = aryCols(j)
                
                .Fields(strColName) = aryRecord
            
            Next
        
        Next
    
    
    End With
    
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddRecords[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          DeleteRecord
' Description:  Delete a specific record from a table
' Assumptions:  Assumes tbl name is properly capitalized & matches db table name
' Parameters:   tbl - table name (string)
'               ID - record ID (long)
'               ShowMsg - whether message should be displayed (boolean, default = true)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 6/2/2016 - moved from forms (TaglineList, EventsList) to mod_App_UI
'   BLC - 6/27/2016- revised to match
'   BLC - 3/30/2017- added displayMsg to enable silent deletes
'   BLC - 11/24/2017 - revised to ShowMsg vs displayMsg, updated to use DisplayMsg()
' ---------------------------------
Public Sub DeleteRecord(tbl As String, ID As Long, Optional ShowMsg As Boolean = True)
On Error GoTo Err_Handler
    Dim strSQL As String

    'find the form & populate its controls from the ID
    strSQL = GetTemplate("d_form_record", "tbl" & PARAM_SEPARATOR & tbl & "|id" & PARAM_SEPARATOR & ID)
    
    If IsNull(strSQL) Or Len(strSQL) = 0 Then GoTo Exit_Handler
Debug.Print strSQL
    DoCmd.SetWarnings False
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    
    If ShowMsg Then
        'show deleted record message & clear
        Dim msg As String
        msg = "Record # " & ID & " from " & tbl & " deleted."
        
        DisplayMsg "", msg, "info", "Record Deleted"
        
    End If
        
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DeleteRecord[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Delete_All_Records
' Description:  Cleanup function deletes all records in a basic table set
'               Tables included:
'                   tbl_Db_Revisions        tlu_Contacts
'                   tbl_Db_Meta             tbl_Events
'                   tbl_Event_Details       tbl_Event_Group
'                   tbl_Field_Data          tbl_Locations
'                   tbl_Data_Locations      tbl_Sites
'                   xref_Event_Contacts
'
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  NCPN, unknown
' Adapted:      Bonnie Campbell, April 28, 2017 for NCPN tools
' Revisions:    NCPN, unknown - initial version
'               BLC, 4/28/2017 - moved to mod_Db
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                multiple open connections
' ---------------------------------
Public Sub Delete_All_Records()
On Error GoTo Err_Handler:

    Dim strSQL As String
    Dim strTables(11) As String
    Const cstrSQL As String = "DELETE * FROM "
    Dim i As Integer
    
    strTables(0) = "tbl_Db_Revisions"
    strTables(1) = "tbl_Db_Meta"
    strTables(2) = "tbl_Event_Details"
    strTables(3) = "tbl_Field_Data"
    strTables(4) = "tbl_Data_Locations"
    strTables(5) = "xref_Event_Contacts"
    strTables(6) = "tlu_Contacts"
    strTables(7) = "tbl_Events"
    strTables(8) = "tbl_Event_Group"
    strTables(9) = "tbl_Locations"
    strTables(10) = "tbl_Sites"
    
    For i = 0 To UBound(strTables) - 1
        strSQL = cstrSQL & strTables(i)
        CurrDb.Execute strSQL
    Next i

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Delete_All_Records[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          HandleDependentQueries
' Description:  Runs or closes Template dependent queries
' Assumptions:  Recursion is supported -> if a query is run & has dependencies this
'                                         routine will recurse & open the
'                                         dependent queries dependent queries
' Parameters:   deps - dependent queries (string, comma delimited)
'               action - run or close & delete (string, run or remove)
' Returns:      -
' Throws:       none
' References:   -
'   Parfait, December 31, 2015
'   http://stackoverflow.com/questions/34512307/how-to-refresh-navigation-pane-using-vba
'   Rey Obrero (Capricorn1) July 29, 2010
'   https://www.experts-exchange.com/questions/26365999/Hide-and-Unhide-tables-queries-forms-code-etc.html
' Source/date:  -
' Adapted:      Bonnie Campbell, March 30, 2017 - for NCPN tools
' Revisions:
'   BLC - 3/30/2017 - initial version
'   BLC - 3/31/2017 - adjust to include check for g_AppTemplateIDs
'   BLC - 4/3/2017 - code cleanup
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Sub HandleDependentQueries(deps As String, action As String)
On Error GoTo Err_Handler
    
    Dim ary() As String
    Dim i As Integer
    Dim ids As Scripting.Dictionary
If deps = "41" Then
    Debug.Print deps
End If

    'list -> array
    ary = Split(deps, ",")
    
    If IsArray(ary) Then
    
        Dim db As DAO.Database
        Dim qdf As DAO.QueryDef
        Dim rs As DAO.Recordset
        Dim strTemplate As String, deps2 As String, strSQL As String
        Dim tID As Integer, iTemplate
        
        Set db = CurrDb
        
        'retrieve Template ID dictionary
        'initialize AppTemplates if not populated
        If g_AppTemplateIDs Is Nothing Then GetTemplateIDs
        Set ids = g_AppTemplateIDs
    
        'iterate through queries
        For i = LBound(ary) To UBound(ary)
                    
            tID = CInt(ary(i))
            
            iTemplate = ids(tID)
            
            strTemplate = g_AppTemplates(iTemplate).Item("TemplateName")
            
            Select Case LCase(action)
                Case "run"  'generates dep queries
                    
                    'handle dependencies of dependencies first
                    deps2 = g_AppTemplates(strTemplate).Item("Dependencies")

                    If Len(deps2) > 0 Then HandleDependentQueries deps2, "run"
                
                    'retrieve Template SQL
                    strSQL = g_AppTemplates(strTemplate).Item("Template")
                
                    'create & run query
                    'qdf.Name = strTemplate 'rs("TemplateName")
                    'qdf.SQL = strSQL 'rs("Template")
                    
                    'check if query exists
                    If Not QueryExists(strTemplate) Then 'db.QueryDefs(strTemplate) Is Nothing Then 'qdf.Name) Is Nothing Then
                        'create query
                        db.CreateQueryDef strTemplate, strSQL 'qdf.Name, qdf.SQL
                        
                        'hide new query
'                        db.QueryDefs (strTemplate)
                        Application.SetHiddenAttribute acQuery, strTemplate, True

                        'refresh UI
                        db.QueryDefs.Refresh
                        
                        'add query to open queries list (later to be closed)
                        g_OpenQueries = g_OpenQueries & "," & tID
                        
                    End If
                    
                    'run query
                Case "remove"   'remove dep queries
                    
                    If QueryExists(strTemplate) Then
                        'close & remove
                        DoCmd.Close acQuery, strTemplate, acSaveNo 'qdf.Name, acSaveNo
                        DoCmd.DeleteObject acQuery, strTemplate 'qdf.Name
                        
                        'remove from open query list
                        g_OpenQueries = Replace(Replace(g_OpenQueries, tID, ""), ",,", ",")
                    End If
                    
            End Select
    
        Next
    End If
        
    'clean up g_OpenQueries (remove opening comma)
    If Left(g_OpenQueries, 1) = "," Then
        g_OpenQueries = Right(g_OpenQueries, Len(g_OpenQueries) - 1)
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - HandleDependentQueries[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RemoveTemplateQueries
' Description:  Removes queries created from Templates
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, March 30, 2017 - for NCPN tools
' Revisions:
'   BLC - 3/30/2017 - initial version
' ---------------------------------
Public Sub RemoveTemplateQueries()
On Error GoTo Err_Handler
    
    Dim i As Integer
    
    'initialize AppTemplates if not populated (necessary for testing)
    If g_AppTemplates Is Nothing Then GetTemplates
    
    For i = 0 To g_AppTemplates.Count - 1
    
        Dim strTemplate As String
        Dim iTemplate As Integer
         
        strTemplate = g_AppTemplates.Items()(i).Item("TemplateName")
        iTemplate = g_AppTemplates.Items()(i).Item("ID")
                    
        If QueryExists(strTemplate) Then
            'close & remove
            DoCmd.Close acQuery, strTemplate, acSaveNo 'qdf.Name, acSaveNo
            DoCmd.DeleteObject acQuery, strTemplate 'qdf.Name
            
            'remove from open query list
            g_OpenQueries = Replace(Replace(g_OpenQueries, iTemplate, ""), ",,", ",")
        End If
    
    Next
        
    'clean up g_OpenQueries (remove opening comma)
    If Left(g_OpenQueries, 1) = "," Then
        g_OpenQueries = Right(g_OpenQueries, Len(g_OpenQueries) - 1)
    End If
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveTemplateQueries[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          RefreshTempTable
' Description:  Refreshes Temp table data
' Assumptions:  Temp table is generated by a query
'               w/ name Create_* where * = Temp table name
' Parameters:   tbl - Temp table name (string)
'               nav - nav group to put table into (string, default "Queries - Application", optional)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/18/2017 - initial version
' ---------------------------------
Public Sub RefreshTempTable(tbl As String, _
            Optional nav As String = "Queries - Application")
On Error GoTo Err_Handler

    Dim CreateQuery As String
    
    CreateQuery = "Create_" & tbl

    're-generate the Temp table source
    DoCmd.SetWarnings False
    If TableExists(tbl) Then
        DoCmd.DeleteObject acTable, tbl
    End If
    
    DoCmd.OpenQuery CreateQuery
    
    If Not Len(nav) = 0 Then
        'move tables to nav group
        SetNavGroup nav, tbl, "table"
    End If
    
    DoCmd.SetWarnings True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RefreshTempTable[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          CombineTableSQL
' Description:  Combines two tables into a resulting table
' Assumptions:  -
' Parameters:   Table1 - first table being combined (string)
'               Table2 - second table being combined (string)
'               Destination - name of destination database (string)
' Returns:      -
' Throws:       none
' References:
'   Laurence, September 13, 2013
'   https://stackoverflow.com/questions/18795263/how-to-merge-two-database-tables-when-only-some-fields-are-common
' Source/date:  -
' Adapted:      Bonnie Campbell, June 22, 2017 - for NCPN tools
' Revisions:
'   BLC - 6/22/2017 - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Function CombineTableSQL(Table1 As String, _
                         Table2 As String, _
                         Destination As String) As String
    Dim lDb As Database
    Dim lTd1 As TableDef, lTd2 As TableDef
    Dim lField As field, lF2 As field
    Dim lS1 As String, lS2 As String, lSep As String

    CombineTableSQL = "Select "
    lS1 = "Select "
    lS2 = "Select "

    Set lDb = CurrDb
    Set lTd1 = lDb.TableDefs(Table1)
    Set lTd2 = lDb.TableDefs(Table2)

    For Each lField In lTd1.Fields
        CombineTableSQL = CombineTableSQL & lSep & "x.[" & lField.Name & "]"
        lS1 = lS1 & lSep & "a.[" & lField.Name & "]"
        Set lF2 = Nothing
        On Error Resume Next
        Set lF2 = lTd2.Fields(lField.Name)
        On Error GoTo 0
        If lF2 Is Nothing Then
            lS2 = lS2 & lSep & "Null"
        Else
            lS2 = lS2 & lSep & "b.[" & lField.Name & "]"
        End If
        lSep = ", "
    Next

    For Each lField In lTd2.Fields
        Set lF2 = Nothing
        On Error Resume Next
        Set lF2 = lTd1.Fields(lField.Name)
        On Error GoTo 0
        If lF2 Is Nothing Then
            CombineTableSQL = CombineTableSQL & lSep & "x.[" & lField.Name & "]"
            lS1 = lS1 & lSep & "Null as [" & lField.Name & "]"
            lS2 = lS2 & lSep & "b.[" & lField.Name & "]"
        End If
        lSep = ", "
    Next

    CombineTableSQL = CombineTableSQL & " Into [" & Destination & "] From ( " & lS1 & " From [" & Table1 & "] a Union All " & lS2 & " From [" & Table2 & "] b ) x"

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CombineTableSQL[mod_Db])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          SetColumnOrdinalPosition
' Description:  Sets the position for a column in a table
' Assumptions:  -
' Parameters:   tdf - table being modified (TableDef)
'               MoveCol - name of column to position (move) (string)
'               AfterCol - name of column to position after (string)
' Returns:      True or False depending on whether move occurred
' Throws:       none
' References:
'   Paul Shapiro, Jan 6,2010
'   https://www.pcreview.co.uk/threads/how-can-i-change-the-column-order-using-vba.3948127/
' Source/date:  -
' Adapted:      Bonnie Campbell, June 22, 2017 - for NCPN tools
' Revisions:
'   BLC - 6/22/2017 - initial version
'   BLC - 10/5/2017 - changed from Function to Subroutine
' ---------------------------------
Public Sub SetColumnOrdinalPosition( _
    tdf As DAO.TableDef, MoveCol As String, AfterCol As String) 'As Boolean
    
    'move MoveCol from tbl to the position immediately following AfterCol
    Dim fldNew As DAO.field
    Dim MoveTo As Long
    
    'Get the ordinal position desired for the field
    Set fldNew = tdf.Fields(AfterCol)
    MoveTo = fldNew.OrdinalPosition + 1
    Set fldNew = Nothing
    
    'Increment ordinal positions and make space for the newly-assigned field
    For Each fldNew In tdf.Fields
        If fldNew.Name = MoveCol Then
            fldNew.OrdinalPosition = MoveTo
        ElseIf fldNew.OrdinalPosition >= MoveTo Then
            fldNew.OrdinalPosition = fldNew.OrdinalPosition + 1
        End If
    Next
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetColumnOrdinalPosition[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub