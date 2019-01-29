Attribute VB_Name = "mod_App_Data"
Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Data
' Level:        Application module
' Version:      1.62
' Description:  data functions & procedures specific to this application
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015  - 1.00 - initial version
'               BLC - 2/18/2015 - 1.01 - included subforms in fillList
'               BLC - 5/1/2015  - 1.02 - integerated into Invasives Reporting tool
'               BLC - 5/22/2015 - 1.03 - added PopulateList()
'               BLC - 6/3/2015  - 1.04 - added IsUsedTargetArea()
'               BLC - 5/5/2016  - 1.05 - added GetRiverSegments(), GetProtocolVersion()
'                                        changed to Exit_Handler vs. Exit_Function
'               BLC - 6/28/2016 - 1.06 - added ToggleIsActive(), revised getParkState() to GetParkState()
'               BLC - 7/26/2016 - 1.07 - added SetRecord(), GetRecords()
'               BLC - 7/28/2016 - 1.08 - added UpsertRecord()
'               BLC - 7/30/2016 - 1.09 - added ToggleSensitive()
'               BLC - 8/8/2016  - 1.10 - updated UpsertRecord() for additional forms
'               BLC - 9/1/2016  - 1.11 - added UploadSurveyFile(), updated UpsertRecord()
'               BLC - 9/13/2016 - 1.12 - added FetchAddlData()
'               BLC - 9/21/2016 - 1.13 - updated SetRecord() i_login parameters
'               BLC - 9/22/2016 - 1.14 - added templates
'               BLC - 10/16/2016 - 1.15 - fixed PopulateCombobox() to properly set recordset
'               BLC - 10/19/2016 - 1.16 - renamed UploadSurveyFile() to UploadCSVFile() to genericize
'               BLC - 10/24/2016 - 1.17 - updated SetRecord(), ToggleIsActive()
'               BLC - 10/28/2016 - 1.18 - updated i_task, TempVars("ContactID") -> TempVars("AppUserID")
'               BLC - 1/9/2017   - 1.19 - revised UpsertRecord from ContactID to ID,
'                                         added GetRecords templates
'               BLC - 1/24/2017  - 1.20 - added IsNPS flag for SetRecord() contacts
'               BLC - 2/1/2017   - 1.21 - updated UpsertRecord() to handle form upserts
'                                         for forms w/o lists/msg & msg icons
'               BLC - 2/3/2017   - 1.22 - location adjustments for UpsertRecord() & SetRecord()
'               BLC - 2/7/2017   - 1.23 - added template - s_location_with_loctypeID_sensitivity
' --------------------------------------------------------------------
'               BLC, 3/22/2017          added updated version to Upland db
' --------------------------------------------------------------------
'               BLC, 3/22/2017  - 1.24 - removed big rivers only components
'                                        revised for uplands
'               BLC, 3/29/2017  - 1.25 - added FieldCheck, FieldOK, Dependencies for templates
'               BLC, 3/30/2017  - 1.26 - added non-parameterized query option for GetRecords()
'               BLC, 4/3/2017   - 1.27 - added qc_species_by_plot_visit
' --------------------------------------------------------------------
'               BLC, 4/17/2017          added updated version to Invasives db
' --------------------------------------------------------------------
'               BLC, 4/17/2017  - 1.28 - revised for Invasives
'               BLC, 7/5/2017   - 1.29 - SetRecords() added inserts for transect quadrats &
'                                        surface cover records, GetRecords() added
'                                        surface microhabitat & quadrat IDs templates
'               BLC, 7/14/2017  - 1.30 - add transect update template
'               BLC, 7/16/2017  - 1.31 - revise u_transect_data to exclude NULLable start time,
'                                        Add u_transect_start_time
'               BLC, 7/17/2017  - 1.32 - add u_quadrat_flags
'               BLC, 7/18/2017  - 1.33 - add species cover templates
'               BLC, 7/24/2017  - 1.34 - added get surface ID from col name template
'               BLC, 7/26/2017  - 1.35 - added u_surfacecover_by_ID template
' --------------------------------------------------------------------
'               BLC, 9/7/2017  - 1.36 - merged common code for framework from Upland, Invasives, Big Rivers dbs
' --------------------------------------------------------------------
'                   BLC - 6/3/2015  - 1.04 - added IsUsedTargetArea
'                   BLC - 12/1/2015 - 1.05 - "extra" vs target area renaming (IsUsedTargetArea > IsUsedExtraArea)
'                   BLC - 6/14/2017 - 1.06 - add SetRecord(), GetRecords()
'                   ------------
'                   BLC, 8/14/2017  - 1.28 - add error handling to address error 3048 on SetPlotCheckResult(),
'                                        GetRecords()
' --------------------------------------------------------------------
'               BLC, 9/28/2017 - 1.37 - update ToggleSensitive, SetRecord for sensitive locations/species
'               BLC, 9/29/2017 - 1.38 - update UpsertRecord for location, add i_location site ID (SetRecord)
'               BLC, 10/4/2017 - 1.39 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC, 10/6/2017 - 1.40 - moved UploadCSVFile() to mod_CSV
'               BLC, 10/17/2017 - 1.41 - added error 3022 - duplicate record handling,
'                                        updated SetRecord u_site parameters
'               BLC, 10/18/2017 - 1.42 - added tsys_Datasheet_defaults parameters
'               BLC, 10/30/2017 - 1.43 - change location CollectionSourceID to use text vs. index (UpsertRecord)
'               BLC, 10/31/2017 - 1.44 - added ReplicatePlot, CalibrationPlot flags, updated VegPlot case
'               BLC, 11/2/2017  - 1.45 - added s_plot_numbers, s_transect_numbers (GetRecords)
'               BLC, 11/3/2017  - 1.46 - added s_modal_sediment_size (GetRecords)
'               BLC, 11/6/2017  - 1.47 - use GetRecords (GetSiteID), add i_site_vegtransect (SetRecord),
'                                        added s_site_id_by_code (GetRecords)
'               BLC, 11/8/2017  - 1.48 - update UpsertRecord, SetRecord for % MSS
'               BLC, 11/9/2017  - 1.49 - added s_location_by_site, s_location_by_feature (GetRecords),
'                                        added Comment case (UpsertRecord)
'               BLC, 11/10/2017 - 1.50 - added Transducer distances (UpsertRecord)
'               BLC, 11/11/2017 - 1.51 - updated UpsertRecord, SetRecord for vegplot cases
'               BLC, 11/12/2017 - 1.52 - added unknown cases (SetRecord)
'               BLC, 11/24/2017 - 1.53 - fix i_feature parameters (SetRecord)
'               BLC, 11/26/2017 - 1.54 - added VegSpecies, updated VegWalk, VegPlot (UpsertRecord),
'                                        updated VegPlot for social trails pct
'               BLC, 12/5/2017  - 1.55 - added VegPlot BeaverBrowse (SetRecord, UpsertRecord)
'               BLC, 12/8/2017  - 1.56 - update VegPlot (UpsertRecord)
'               BLC, 12/13/2017 - 1.57 - added s_access_lvl case (GetRecords)
'               BLC, 1/10/2018  - 1.58 - added s_comments, s_tasks (GetRecords),
'                                        SetReportRecordset & rpt_recordset global variable,
'                                        add PhotoBatchUpdate, Logger cases (UpsertRecord)
'                                        update i_photo case (SetRecord), add GetPhotoTypes
'               BLC, 1/16/2018  - 1.59 - add i_event_photo case (SetRecord), update i_photo (UpsertRecord) case
'               BLC, 1/17/2018  - 1.60 - updated i_photo_event, i_photo, u_photo cases
'               BLC, 1/24/2018  - 1.61 - added s_photo_year_by_site (GetRecords)
'               BLC, 1/xx/2019  - 1.62 - added FOREST VEG section and s_annual_data_export (GetRecords)
' =================================

' =================================
'   Declarations
' =================================
Public rpt_recordset As DAO.Recordset   'used in SetReportRecordset()
'   AppPhotoTypes global dictionary --> defined in app photo types (Big Rivers AppEnum) [mod_App_Data]
Public g_AppPhotoTypes As Scripting.Dictionary

' =================================
'   Dictionary Methods
' =================================
' ---------------------------------
' SUB:     GetPhotoTypes
' Description:  loads PhotoTypes into memory as a global dictionary object (dictTemplates)
'               makes current Templates available without querying the db AppEnum table
' Parameters:   PType - specifies phototype to retrieve (F,T,O,R,OO, etc.)
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   tsys_Db_Templates, Microsoft Scripting Runtime (dictionary object)
'   HansUp, June 27, 2013
'   http://stackoverflow.com/questions/17328092/how-to-display-access-query-results-without-having-to-create-Temporary-query
' Source/date:  Bonnie Campbell, June 2014
' Revisions:    BLC, 6/16/2014 - initial version
'               BLC, 5/13/2016 - shifted from mod_Db_Templates to mod_Db & adjusted to match tsys_Db_Templates
'               BLC, 5/19/2016 - revised documentation & renamed GetPhotoTypes() vs. GetSQLTemplates() since tsys_Db_Templates
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
Public Sub GetPhotoTypes(Optional PType As String = "")

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strSQLWhere As String, Key As String
    Dim Value As Variant
    
    'sql -> ID, Version, IsSupported, Context, Syntax, TemplateName, Params, Template, Remarks,
    '       EffectiveDate, RetireDate, CreateDate, CreatedBy_ID, LastModified, LastModifiedBy_ID
    strSQL = "SELECT * FROM AppEnum WHERE EnumType = 'PhotoType';"
    
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)
    
    'handle no records
    If rs.EOF Then
        MsgBox "Sorry, no photo types were found in this database version.", vbExclamation, _
            "Photo Types Not Found"
        DoCmd.CancelEvent
        GoTo Exit_Handler
    End If
    
    'prepare dictionary
    Dim dict As New Scripting.Dictionary
    Dim ary(1 To 4) As String
    Dim i As Integer
    
    'prepare the dictionary key array
'    ary(1) = "Context"
'    ary(2) = "TemplateName"
'    ary(3) = "Template" 'Template
'    ary(4) = "Params"
'    ary(5) = "Syntax"
'    ary(6) = "ID"
'    ary(7) = "FieldCheck"
'    ary(8) = "FieldOK"
'    ary(9) = "Dependencies"
'    ary(10) = "DataScope"
'    ary(11) = "Version"
    ary(1) = "ID"
    ary(2) = "Label"
    ary(3) = "Summary"
    
    'prepare array of dictionaries
    Dim dictPhotoTypes As Scripting.Dictionary
    Set dictPhotoTypes = New Scripting.Dictionary
    
    rs.MoveLast
    rs.MoveFirst
    Do Until rs.EOF
        'create new dictionary object
        Set dict = New Scripting.Dictionary
        
        'populate the dictionary
        For i = 1 To UBound(ary)
            
            Key = ary(i)
            
            Value = Nz(rs.Fields(ary(i)), "")
            
            'add key if it isn't already there
            If Not dict.Exists(Key) Then
                If IsNull(Value) Then MsgBox Key, vbOKCancel, "is NULL"
                'Debug.Print Nz(Value, key & "-NULL")
                dict.Add Key, Value
            End If
        
        Next
        
        'add Template dictionary to dictionary of Templates
        dictPhotoTypes.Add dict("PhotoType"), dict
        
'        Debug.Print dict("TemplateName") & " " & dict.Item("ID")
        rs.MoveNext
    Loop
    
    'load global AppTemplates As Scripting.Dictionary of Templates
    Set g_AppPhotoTypes = Nothing    'clear first
    
    Set g_AppPhotoTypes = dictPhotoTypes
    
Exit_Handler:
    'cleanup
    Set rs = Nothing
    Set dict = Nothing
    Set dictPhotoTypes = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 457  'Duplicate PhotoType -- AppEnum finds more than one w/ same name (improbable)
        MsgBox "A duplicate PhotoType was found." & vbCrLf & vbCrLf & _
            "When you click 'OK' a query will run to identify the problem PhotoType." & vbCrLf & vbCrLf & _
            "You can close the query after it runs (save it if you like)." & vbCrLf & vbCrLf & _
            "Please contact your data manager to resolve this issue." & vbCrLf & vbCrLf & _
            "Error #" & Err.Number & " - GetPhotoTypes[mod_Db]:" & vbCrLf & _
            Err.Description, vbExclamation, "Duplicate Db PhotoType Found! [AppEnum]"

            Dim strErrorSQL As String
            strErrorSQL = "SELECT Lable, Count(Label) AS NumberOfDupes " & _
                    "FROM AppEnum " & _
                    "GROUP By Label " & _
                    "HAVING Count(Label) > 1;"

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
            "Error encountered (#" & Err.Number & " - GetPhotoTypes[mod_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:     GetDictionary
' Description:  Retrieves a dictionary object based on a template results
' Parameters:   template - SQL template to retrieve data (string)
'               keys - array of keys to retrieve (array)
'               dictname - name of resulting dictionary object (string)
' Returns:      -
' Assumptions:  -
' Throws:       none
' References:   tsys_Db_Templates, Microsoft Scripting Runtime (dictionary object)
' Source/date:  Bonnie Campbell, January 2018
' Revisions:    BLC, 1/24/2018 - initial version
' ---------------------------------
Public Function GetDictionary(template As String, keys As Variant, _
                                dictname As String) As Scripting.Dictionary

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim Key As String
    Dim Value As Variant
    Dim i As Integer
    
    If IsArray(keys) Then
        i = keys.Count
    Else
        GoTo Exit_Handler
    End If
    
    Set db = CurrDb
    Set rs = GetRecords(template)
    
    'handle no records
    If rs.EOF Then
        MsgBox "Sorry, no records were found in this database version.", vbExclamation, _
            "No Records Found"
        DoCmd.CancelEvent
        GoTo Exit_Handler
    End If
    
    'prepare dictionary
'    Dim dictNew As New Scripting.Dictionary
'    Dim ary(1 To X) As String
'    Dim i As Integer
'
'    'prepare the dictionary key array
''    ary(1) = "Context"
''    ary(2) = "TemplateName"
''    ary(3) = "Template" 'Template
''    ary(4) = "Params"
''    ary(5) = "Syntax"
''    ary(6) = "ID"
''    ary(7) = "FieldCheck"
''    ary(8) = "FieldOK"
''    ary(9) = "Dependencies"
''    ary(10) = "DataScope"
''    ary(11) = "Version"
'    ary(1) = "ID"
'    ary(2) = "Label"
'    ary(3) = "Summary"
    
    'prepare array of dictionaries
    Dim dictNew As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dictNew = New Scripting.Dictionary
    
    rs.MoveLast
    rs.MoveFirst
    
    Do Until rs.EOF
    
        'create new dictionary object
        Set dict = New Scripting.Dictionary
        
        'populate the dictionary
        For i = 1 To UBound(keys)
            
            Key = keys(i)
            
            Value = Nz(rs.Fields(keys(i)), "")
            
            'add key if it isn't already there
            If Not dict.Exists(Key) Then
                If IsNull(Value) Then MsgBox Key, vbOKCancel, "is NULL"
                'Debug.Print Nz(Value, key & "-NULL")
                dict.Add Key, Value
            End If
        
        Next
        
        'add item dictionary to overall dictionary of items
        dictNew.Add dict(dictname), dict
        
'        Debug.Print dict("TemplateName") & " " & dict.Item("ID")
        rs.MoveNext
    Loop
    
    'load global AppTemplates As Scripting.Dictionary of Templates
    'Set g_AppPhotoTypes = Nothing    'clear first
    
    Set GetDictionary = dictNew
    
Exit_Handler:
    'cleanup
    Set rs = Nothing
    Set dict = Nothing
    Set dictNew = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 457  'Duplicate Item -- finds more than one w/ same name (improbable)
        MsgBox "A duplicate item was found." & vbCrLf & vbCrLf & _
            "When you click 'OK' a query will run to identify the problem PhotoType." & vbCrLf & vbCrLf & _
            "You can close the query after it runs (save it if you like)." & vbCrLf & vbCrLf & _
            "Please contact your data manager to resolve this issue." & vbCrLf & vbCrLf & _
            "Error #" & Err.Number & " - GetDictionary[mod_Db]:" & vbCrLf & _
            Err.Description, vbExclamation, "Duplicate Item Found! [AppEnum]"

            Dim strErrorSQL As String
            strErrorSQL = "SELECT Label, Count(Label) AS NumberOfDupes " & _
                    "FROM AppEnum " & _
                    "GROUP By Label " & _
                    "HAVING Count(Label) > 1;"

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
            "Error encountered (#" & Err.Number & " - GetDictionary[mod_Db])"
    End Select
    Resume Exit_Handler
End Function


' =================================
'   List Methods
' =================================
' ---------------------------------
' SUB:          fillList
' Description:  Fill a list (or listbox like subform) from specific queries for datasheets, species or other items
' Assumptions:  Either a listbox or subform control is being populated
' Parameters:   frm - main form object
'               ctrl - either:
'                      lbx - main form listbox object (for filling a listbox control)
'                      sfrm - subform object (for populating a subform control)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/6/2015  - initial version
'   BLC, 2/18/2015 - adapted to include subform as well as listbox controls
'   BLC, 5/1/2015  - integrated into Invasives Reporting tool
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'                   - un-comment out
' --------------------------------------------------------------------
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Sub fillList(frm As Form, ctrlSource As Control, Optional ctrlDest As Control)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String

    'output to form or listbox control?

    'determine data source
    Select Case ctrlSource.Name

        Case "lbxDataSheets", "sfrmDatasheets" 'Datasheets
            strQuery = "qry_Active_Datasheets"
            strSQL = CurrDb.QueryDefs(strQuery).SQL

        Case "lbxSpecies", "lbxTgtSpecies", "fsub_Species_Listbox" 'Species
            strQuery = "qry_Plant_Species"
            strSQL = CurrDb.QueryDefs(strQuery).SQL

    End Select

    'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    If Not ctrlDest Is Nothing Then
        'populate list & headers
        PopulateList ctrlSource, rs, ctrlDest
    Else
        'populate only ctrlSource headers
        PopulateListHeaders ctrlSource, rs
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fillList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          PopulateList
' Description:  Populate listbox and similar controls from recordset
' Assumptions:  -
' Parameters:   ctrlSource - source control (listbox/listview)
'               rs - recordset used to populate control (recordset object)
'               ctrlDest - destination control (listbox/listview)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' krish KM, Aug. 27, 2014
' http://stackoverflow.com/questions/25526904/populate-listbox-using-ado-recordset
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/20/2015 - changed from tbxMasterCode to tbxLUCode
'   BLC - 5/22/2015 - moved to mod_App_Data from mod_List
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'                   - un-comment out
'               BLC - 12/1/2015 - "extra" vs. target area renaming (tbxTgtAreaID > tbxExtraAreaID, Target_Area_ID > Extra_Area_ID)
' --------------------------------------------------------------------
' ---------------------------------
Public Sub PopulateList(ctrlSource As Control, rs As Recordset, ctrlDest As Control)
On Error GoTo Err_Handler

    Dim frm As Form
    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer, iZeroes As Integer
    Dim strItem As String, strColHeads As String, aryColWidths() As String

    Set frm = ctrlSource.Parent

    rows = rs.RecordCount
    cols = rs.Fields.Count

    'address no records
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Handler
    End If

    'handle sfrm controls (acSubform = 112)
    If ctrlSource.ControlType = acSubform Then
        Set ctrlSource.Form.Recordset = rs

        ctrlSource.Form.Controls("tbxCode").ControlSource = "Code"
        ctrlSource.Form.Controls("tbxSpecies").ControlSource = "Species"
        'ctrlSource.Form.Controls("tbxMasterCode").ControlSource = "Master_PLANT_Code"
        ctrlSource.Form.Controls("tbxLUCode").ControlSource = "LUCode"
        ctrlSource.Form.Controls("tbxTransectOnly").ControlSource = "Transect_Only"
        ctrlSource.Form.Controls("tbxExtraAreaID").ControlSource = "Target_Area_ID"

        'set the initial record count (MoveLast to get full count, MoveFirst to set display to first)
        rs.MoveLast
        ctrlSource.Parent.Form.Controls("lblSfrmSpeciesCount").Caption = rs.RecordCount & " species"
        rs.MoveFirst

        GoTo Exit_Handler
    End If

    'fetch column widths array
    aryColWidths = Split(ctrlSource.ColumnWidths, ";")

    'count number of 0 width elements
    iZeroes = CountArrayValues(aryColWidths, "0")

    'clear out existing values
    ClearList ctrlSource

    'populate column names (if desired)
    If ctrlSource.ColumnHeads = True Then
        PopulateListHeaders ctrlSource, rs

        'populate second listbox headers if present
        If ctrlDest.ColumnHeads = True Then
            ClearList ctrlDest
            PopulateListHeaders ctrlDest, rs
        End If
    End If

    'populate data
    Select Case ctrlSource.RowSourceType
        Case "Table/Query"
            Set ctrlSource.Recordset = rs
        Case "Value List"

            'initialize
            i = 0

            Do Until rs.EOF

                'initialize item
                strItem = ""

                'generate item
                For j = 0 To cols - 1
                    'check if column is displayed width > 0
                    If CInt(aryColWidths(j)) > 0 Then

                        strItem = strItem & rs.Fields(j).Value & ";"

                        'determine how many separators there are (";") --> should equal # cols
                        matches = (Len(strItem) - Len(Replace$(strItem, ";", ""))) / Len(";")

                        'add item if not already in list --> # of ; should equal cols - 1
                        'but # in list should only be # of non-zero columns --> cols - iZeroes
                        If matches = cols - iZeroes Then
                            ctrlSource.AddItem strItem
                            'reset the string
                            strItem = ""
                        End If

                    End If

                Next

                i = i + 1

                rs.MoveNext
            Loop
        Case "Field List"
    End Select

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddListToTable
' Description:  Populate table from listbox
' Assumptions:  -
' Parameters:   lbx - listbox control
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, June 3, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/3/2015 - initial version
'   BLC - 12/1/2015 - "extra" vs. target area renaming (iTgtAreaID > iExtraAreaID, Target_Area_ID > Extra_Area_ID)
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'                   - un-comment out
' --------------------------------------------------------------------
' ---------------------------------
Public Sub AddListToTable(lbx As ListBox)
On Error GoTo Err_Handler

Dim aryFields() As String
Dim aryFieldTypes() As Variant
Dim strCode As String, strSpecies As String, strLUCode As String
Dim iRow As Integer, iTransectOnly As Integer, iExtraAreaID As Integer

    iRow = lbx.ListCount - 1 'Forms("frm_Tgt_Species").Controls("lbxTgtSpecies").ListCount - 1

    ReDim Preserve aryFields(0 To iRow)

    'header row (iRow = 0)
    aryFields(0) = "Code;Species;LUCode;Transect_Only;Extra_Area_ID"   'iRow = 0
    aryFieldTypes = Array(dbText, dbText, dbText, dbInteger, dbInteger)

    'data rows (iRow > 0)
    For iRow = 1 To lbx.ListCount - 1

        ' ---------------------------------------------------
        '  NOTE: listbox column MUST have a non-zero width to retrieve its value
        ' ---------------------------------------------------
         strCode = lbx.Column(0, iRow) 'column 0 = Master_PLANT_Code (Code)
         strSpecies = lbx.Column(1, iRow) 'column 1 = Species name (Species)
         strLUCode = lbx.Column(2, iRow) 'column 2 = LU_Code (LUCode)
         iTransectOnly = Nz(lbx.Column(3, iRow), 0) 'column 3 = Transect_Only (TransectOnly)
         iExtraAreaID = Nz(lbx.Column(4, iRow), 0) 'column 4 = Extra_Area_ID (ExtraAreaID)

        aryFields(iRow) = strCode & ";" & strSpecies & ";" & strLUCode & ";" & iTransectOnly & ";" & iExtraAreaID

    Next

    'save the existing records to temp_Listbox_Recordset & replace any existing records
    SetListRecordset lbx, True, aryFields, aryFieldTypes, "temp_Listbox_Recordset", True

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddListToTable[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     getListLastModifiedDate
' Description:  Retrieve the last modified date with a park (via tbl_Target_List)
' Assumptions:  -
' Parameters:   tgtYear - 4 digit year of list (integer)
'               parkCode - 4 character park designator (string)
' Returns:      date - last modified date (mmm-d-yyyy H:nn AMPM format) for the specified target list (string)
'                      if NULL (no last modified date) returns empty string
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/10/2015  - initial version
' ---------------------------------
Public Function getListLastModifiedDate(TgtYear As Integer, ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim strCriteria As String

    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Or TgtYear < 2000 Then
        GoTo Exit_Handler
    End If
    
    'set lookup criteria
    strCriteria = "Park_Code LIKE '" & ParkCode & "' AND CInt(Target_Year) = " & CInt(TgtYear)
    
    'Debug.Print strCriteria
        
    'lookup last modified date & return value
    getListLastModifiedDate = Nz(Format(DLookup("Last_Modified", "tbl_Target_List", strCriteria), "mmm-d-yyyy H:nn AMPM"), "")
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getListLastModifiedDate[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IsUsedExtraArea
' Description:  Determine if the extra/target area is in use by a target list
' Parameters:   ExtraAreaID - extra/target area idenifier (integer)
' Returns:      boolean - true if extra/target area is in use, false if not
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
'   BLC - 12/1/2015 - "extra" vs target area renaming (IsUsedTargetArea > IsUsedExtraArea)
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function IsUsedExtraArea(ExtraAreaID As Integer) As Boolean
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    'default
    IsUsedExtraArea = False
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 Target_Area_ID FROM tbl_Target_Species WHERE Target_Area_ID = " & ExtraAreaID & ";"
            
    'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        IsUsedExtraArea = True
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsUsedExtraArea[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' =================================
'   Combobox Methods
' =================================
' ---------------------------------
' SUB:          PopulateCombobox
' Description:  Populate priority/status comboboxes
' Parameters:   cbx - combobox control to populate (ComboBox)
'               BoxType - type of combobox, priority or status (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
'  https://msdn.microsoft.com/en-us/library/office/ff845773.aspx
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
'   BLC - 10/12/2016 - fixed to set combobox recordset
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Sub PopulateCombobox(cbx As ComboBox, BoxType As String)
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Select Case BoxType
        Case ""
        Case "priority"
            strSQL = "SELECT ID, Priority FROM Priority ORDER BY Sequence ASC;"
        Case "status"
            strSQL = "SELECT ID, Status FROM Status ORDER BY Sequence ASC;"
    End Select
 
     'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)
 
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        Set cbx.Recordset = rs
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateCombobox[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   Tree Methods
' =================================
' ---------------------------------
' SUB:     PopulateTree
' Description:  Populate the treeview control
' Parameters:   TreeType - treeview type (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Sub PopulateTree(TreeType As String)
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    Select Case TreeType
        Case "ParkSiteFeatureTransectPlot"
            strSQL = "SELECT * FROM qry_Park_Site_Feature_Transect_Plot"
        Case "Photo"
    End Select
                   
    'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        
    Else
        GoTo Exit_Handler
    End If
       
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateTree[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   Toggle Methods
' =================================
' ---------------------------------
' Sub:          ToggleIsActive
' Description:  Toggle IsActive button click actions
' Assumptions:  -
' Parameters:   Context - form context for the action (string)
'               ID - id of record to toggle (long)
'               IsActive - state to change IsActiveFlag to (Byte), 0 - active, 1 - inactive
'                          optional for ModWentworth scale retire date
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 20, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/20/2016 - initial version
'   BLC - 6/28/2016 - shifted from ContactList form to mod_App_Data
'   BLC - 10/20/2016 - added ModWentworth retire date toggle
'   BLC - 10/24/2016 - revised to use SetRecord()
' ---------------------------------
Public Sub ToggleIsActive(Context As String, ID As Long, Optional IsActive As Byte)
On Error GoTo Err_Handler
    
'    Dim strSQL As String
'
'    Select Case Context
'        Case "Contact"
'            strSQL = GetTemplate("u_contact_isactive_flag", _
'                      "IsActiveFlag" & PARAM_SEPARATOR & IsActive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'        Case "Site"
'            strSQL = GetTemplate("u_site_isactive_flag", _
'                      "IsActiveFlag" & PARAM_SEPARATOR & IsActive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'        Case "ModWentworthScale"
'            strSQL = GetTemplate("u_mod_wentworth_retireyear", _
'                      "RetireDate" & PARAM_SEPARATOR & Date & "|ID" & _
'                      PARAM_SEPARATOR & ID)
'    End Select
'
'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (strSQL)
'    DoCmd.SetWarnings True
    
    Dim template As String
    
    Select Case Context
        Case "Contact"
            template = "u_contact_isactive_flag"
        Case "Site"
            template = "u_site_isactive_flag"
        Case "ModWentworthScale"
            template = "u_mod_wentworth_retireyear"
            
    End Select
    
    Dim Params(0 To 3) As Variant
    
    Params(0) = template
    Params(1) = ID
    Params(2) = IIf(InStr(template, "wentworth") > 0, Year(Date), IsActive)
        
    SetRecord template, Params
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleIsActive[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ToggleSensitive
' Description:  Toggle Sensitive button click actions
' Assumptions:  -
' Parameters:   Context - form context for the action (string)
'               ID - id of record to toggle (long)
'               Sensitive - state to change SensitiveFlag to (Byte), 0 - active, 1 - inactive
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 30, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/30/2016 - initial version
'   BLC - 9/28/2017 - revised template to lower case
' ---------------------------------
Public Sub ToggleSensitive(Context As String, ID As Long, Sensitive As Byte)
On Error GoTo Err_Handler
    
    Dim template As String
    
    template = IIf(Sensitive = 1, "i_", "d_")
    
    template = LCase(template & "sensitive_" & Context)
    
    If Right(template, 1) <> "s" Then template = template & "s"
    
'    Select Case Context
'        Case "Locations"
''            strSQL = GetTemplate("u_location_sensitive_flag", _
''                      "SensitiveFlag" & PARAM_SEPARATOR & Sensitive & _
''                      "|ID" & PARAM_SEPARATOR & ID)
'            strToggle = strToggle & "Sensitive" & Context & "s"
'        Case "Species"
'            strSQL = GetTemplate("u_species_sensitive_flag", _
'                      "SensitiveFlag" & PARAM_SEPARATOR & Sensitive & _
'                      "|ID" & PARAM_SEPARATOR & ID)
'    End Select

'    DoCmd.SetWarnings False
'    DoCmd.RunSQL (strSQL)
'    DoCmd.SetWarnings True
    
    Dim Params(0 To 3) As Variant
    
    Params(0) = template
    Params(1) = ID
    Params(2) = Sensitive
        
    SetRecord template, Params
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleSensitive[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   Hierarchy Methods
' =================================
' ---------------------------------
' Sub:          GetHierarchyLevel
' Description:  Determine the hierarchy level set
' Assumptions:  -
' Parameters:   -
' Returns:      lvl - maximum level set in the application (string)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, January 1, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/11/2017 - initial version
' ---------------------------------
Public Function GetHierarchyLevel() As String
On Error GoTo Err_Handler
    
    Dim lvl As String
    
    'default
    lvl = ""
    
    If Not TempVars("ParkCode") Is Nothing Then
        lvl = "park"
        If Not TempVars("River") Is Nothing Then
            lvl = "river"
            If Not TempVars("SiteCode") Is Nothing Then
                lvl = "site"
                If Not TempVars("Feature") Is Nothing Then
                    lvl = "feature"
                End If
            End If
        End If
    End If

    GetHierarchyLevel = lvl

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetHierarchyLevel[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' =================================
'   Recordset Methods
' =================================
' ---------------------------------
' Sub:          SetReportRecordset
' Description:  Sets the global variable rpt_recordset for setting
'               report record source
' Assumptions:  1) rpt_recordset is a global public variable
'                   defined w/in framework but accessible to referencing dbs
'               2) rpt_recordset is set before the report is opened
'                  (e.g. in ReportList's report name click action for Big Rivers)
'                       rpt_recordset = GetRecords("s_tasks")
'               3) Report recordsource is set in the Report_Open method via
'                       Me.RecordSource = rpt_recordset.Name
' Parameters:   -
' Returns:      recordset for reports
' Throws:       none
' References:
'   Andy Baron, unknown
'   http://access.mvps.org/access/reports/rpt0014.htm
' Source/date:  Bonnie Campbell, January 10, 2018 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/10/2018 - initial version
' ---------------------------------
Public Sub SetReportRecordset(Context As String)
On Error GoTo Err_Handler
    
    Set rpt_recordset = GetRecords(Context)

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetReportRecordset[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
'   Record Methods
' =================================
' ---------------------------------
' Sub:          GetRecords
' Description:  Retrieve records based on template
' Assumptions:  -
' Parameters:   Template - SQL template name (string)
'               params - additional parameters to pass into routine (variant, optional)
' Returns:      rs - data retrieved (recordset)
' Throws:       none
' References:
'   user1938742, October 17, 2014
'   http://stackoverflow.com/questions/26422970/run-query-with-parameters-and-display-in-listbox-ms-access-2013
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/22/2016 - added templates
'   BLC - 1/9/2017 - added templates
'   BLC - 2/7/2017 - added template - s_location_with_loctypeID_sensitivity
'   BLC - 3/28/2017 - added upland templates, removed big rivers templates
'   BLC - 3/30/2017 - added option for non-parameterized queries (Else)
'   BLC - 4/3/2017 - added qc_species_by_plot_visit
' --------------------------------------------------------------------
'   BLC - 4/18/2017 - added updated version to Invasives db
' --------------------------------------------------------------------
'   BLC - 4/18/2017 - adjusted for invasives templates
'   BLC - 4/24/2017 - added microhabitat surface & species templates
'   BLC - 7/5/2017  - added surface microhabitat & quadrat IDs templates
'   BLC - 7/24/2017 - added get surface ID from col name template
'   BLC - 7/26/2017 - added get route transects template
'   BLC - 7/27/2017 - added "s_speciescover_dupes" template, optional params parameter
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'       BLC - 8/14/2017 - redo error handling to address error 3048
' --------------------------------------------------------------------
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
'   BLC - 10/18/2017 - added tsys_Datasheet_defaults parameters
'   BLC - 11/2/2017 - added s_plot_numbers, s_transect_numbers
'   BLC - 11/3/2017 - added s_modal_sediment_size
'   BLC - 11/6/2017 - added s_site_id_by_code
'   BLC - 11/9/2017 - added s_location_by_site, s_location_by_feature
'   BLC - 12/5/2017 - updated s_feature_id, s_record_action_by_refID
'   BLC - 12/12/2017 - add s_usys_temp_photo_list
'   BLC - 12/13/2017 - add s_access_lvl
'   BLC - 1/10/2018  - added s_comments, s_tasks
'   BLC - 1/24/2018  - added s_photo_year_by_site
'   BLC - 1/xx/2019  - added FOREST VEG section, s_annual_data_export
' ---------------------------------
Public Function GetRecords(template As String, _
                            Optional Params As Variant) As DAO.Recordset
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set db = CurrDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .SQL = GetTemplate(template)
        
            Select Case template
                                        
        '-----------------------
        '  QC
        '-----------------------
                Case "qc_ndc_notrecorded_all_methods_by_plot_visit", _
                    "qc_photos_missing_by_plot_visit", _
                    "qc_species_by_plot_visit"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("pid") = TempVars("plotID")
                    .Parameters("vdate") = TempVars("SampleDate")
        
        '-----------------------
        '  SELECTS
        '-----------------------
        
            '-------------------
            ' --- BIG RIVERS ---
            '-------------------
                Case "s_access_level"
                    '-- required parameters --
                    .Parameters("lvl") = TempVars("TempLvl")

                Case "s_app_enum_list"
                    '-- required parameters --
                    .Parameters("etype") = TempVars("EnumType")
                
                Case "s_comments"
                     '-- required parameters --
                
                Case "s_contact_list"
                    '-- required parameters --
                    'N/A
                                
                Case "s_datasheet_defaults_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                
                Case "s_datasheet_defaults_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                                
                Case "s_events_by_feature"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                    .Parameters("feat") = TempVars("Feature")
                                
                Case "s_event_by_park_river_w_location"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                    
                Case "s_events_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_events_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")

                Case "s_events_list_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_feature_by_park_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_feature_id"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    '.Parameters("scode") = TempVars("SiteCode")
                    .Parameters("feat") = TempVars("Feature")
                
                Case "s_feature_list"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                                        
                Case "s_feature_list_by_site", _
                     "s_feature_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_location_by_feature"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                    .Parameters("feat") = TempVars("Feature")
                
                Case "s_location_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_location_by_park_river_segment"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("seg") = TempVars("River")
                
                Case "s_location_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_location_list_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_location_with_loctypeID_sensitivity"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_mod_wentworth_for_eventyr"
                    '-- required parameters --
                    'default event year to current year if not passed in
                    .Parameters("eventyr") = Nz(TempVars("EventYear"), Year(Now))
                
                Case "s_modal_sediment_size"    'from AppEnum
                    '-- required parameters --
                
                Case "s_photo_year_by_site"
                    '-- required parameters --
                    .Parameters("sid") = TempVars("SiteID")
                
                Case "s_plot_numbers"
                    '-- required parameters --
                    .Parameters("maxnum") = TempVars("MaxPlotNumber")
                
                Case "s_record_action_by_refID"
                    '-- required parameters --
                    .Parameters("reftype") = Params(0)  'RefType
                    .Parameters("refID") = Params(1)    'RefID
                
                Case "s_river_segment_id"
                    '-- required parameters --
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_river_list"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                
                Case "s_site_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_site_by_park_river_segment"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("seg") = TempVars("River")
                
                Case "s_site_id_by_code"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_site_list_by_park_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_site_list_by_park_river_segment"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                
                Case "s_site_list_active"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("seg") = TempVars("River")
            
                Case "s_species_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                
                Case "s_tasks"
                     '-- required parameters --
                
                Case "s_top_rooted_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .SQL = Replace(Replace(.SQL, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_rooted_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .SQL = Replace(Replace(.SQL, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_understory_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .SQL = Replace(Replace(.SQL, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_understory_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .SQL = Replace(Replace(.SQL, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_woody_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .SQL = Replace(Replace(.SQL, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_top_woody_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
    
                    'revise TOP X --> 99 is replaced by # species to return (from datasheet defaults)
                    '             -->  8 is replaced by # blanks to return (")
                    .SQL = Replace(Replace(.SQL, 99, TempVars("TopSpecies")), 8, TempVars("TopBlanks"))
                    
                Case "s_transect_numbers"
                    '-- required parameters --
                    .Parameters("maxnum") = TempVars("MaxTransectNumber")
                
                Case "s_usys_temp_photo_list"
                    '-- required parameters --
                                
                Case "s_veg_walk_species_last_year_by_park"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    
                    'revise TOP X --> 8 is replaced by # blanks to return (from # rows remaining)
                    .SQL = Replace(.SQL, 8, TempVars("Blanks"))
                
                Case "s_veg_walk_species_last_year_by_river"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("waterway") = TempVars("River")
                    
                    'revise TOP X --> 8 is replaced by # blanks to return (from # rows remaining)
                    .SQL = Replace(.SQL, 8, TempVars("Blanks"))
                    
                    '-- optional parameters --
        
                Case "s_vegplot_number_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_vegtransect_by_feature"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                    .Parameters("feat") = TempVars("Feature")
                
'                Case "s_vegtransect_by_park_site"
'                    '-- required parameters --
'                    .Parameters("pkcode") = TempVars("ParkCode")
'                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_vegtransect_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
                
                Case "s_vegtransect_number_by_site"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("scode") = TempVars("SiteCode")
            
            '-------------------
            ' --- UPLAND ---
            '-------------------
                Case "s_template_num_records"
                    '-- required parameters --
                
                Case "s_surface"
                    '-- required parameters --
                
                Case "s_surface_by_ID"
                    '-- required parameters --
                    .Parameters("sid") = TempVars("SurfaceID")
                
                Case "s_speciescover_by_transect"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("eid") = TempVars("Event_ID")
                    .Parameters("tid") = TempVars("Transect_ID")
                
                Case "s_surfacecover_by_transect"
                    '-- required parameters --
                    '.Parameters("pkcode") = TempVars("ParkCode")
                    '.Parameters("eid") = TempVars("Event_ID")
                    .Parameters("tid") = TempVars("Transect_ID")

            '-------------------
            ' --- INVASIVES --
            '-------------------
                Case "s_access_level"
                    '-- required parameters --
                    .Parameters("lvl") = TempVars("tempLvl")
                    
                    'clear the tempvar
                    TempVars.Remove "tempLvl"
                                                                                                               
                Case "s_get_parks"
                    '-- required parameters --
                                                                    
                Case "s_park_id"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                                
                Case "s_route_transects"
                    '-- required parameters --
                    .Parameters("eid") = TempVars("EventID")
                
                Case "s_surface"
                    '-- required parameters --
                
                Case "s_surface_by_colname"
                    '-- required parameters --
                    .Parameters("cname") = TempVars("SurfaceColName")
                
                Case "s_surface_by_ID"
                    '-- required parameters --
                    .Parameters("sid") = TempVars("SurfaceID")
                
                Case "s_surface_IDs"
                    '-- required parameters --
                                
                Case "s_speciescover_by_transect"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("eid") = TempVars("Event_ID")
                    .Parameters("tid") = TempVars("Transect_ID")
                
                Case "s_surfacecover_by_transect"
                    '-- required parameters --
                    '.Parameters("pkcode") = TempVars("ParkCode")
                    '.Parameters("eid") = TempVars("Event_ID")
                    .Parameters("tid") = TempVars("Transect_ID")
                
                Case "s_speciescover_dupes"
                    '-- required parameters --
                    .Parameters("eid") = Params(1)
                    .Parameters("tid") = Params(2)
                    .Parameters("pcode") = Params(3)
                    .Parameters("dead") = Params(4)
                    
                Case "s_template_num_records"
                    '-- required parameters --

                Case "s_transect_quadrat_IDs"
                    '-- required parameters --
                    .Parameters("tid") = TempVars("TransectQuadratID")
                
                Case "s_tsys_datasheet_defaults"
                    '-- required parameters --
                   .Parameters("parkID") = TempVars("ParkID")
                
            '-------------------
            ' --- FOREST VEG --
            '-------------------
                Case "s_annual_data_export"
                    '-- required parameters --
                                                                                                               
                Case Else
                    'handle other non-parameterized queries
                    
            End Select
            
            Set rs = .OpenRecordset(dbOpenDynaset)
            
        End With
        
    End With
    
    Set GetRecords = rs
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case 3048 'Cannot open any more databases
        Debug.Print "Error 3048: " & Err.Description & " " & Err.Source
        'close & re-open forms (frm_Data_Entry & PlotCheck)
'         DoCmd.Close acForm, "PlotCheck", acSaveNo
'         DoCmd.SelectObject acForm, "frm_Data_Entry"
'         DoCmd.Close acForm, "frm_Data_Entry", acSaveYes
'         DoCmd.OpenForm "frm_Data_Entry", , , TempVars("CriteriaLoc") & " AND " & TempVars("CriteriaEvent"), , , TempVars("CriteriaEvent")
'         DoCmd.Minimize
'         DoCmd.OpenForm "PlotCheck", acNormal, , , , acWindowNormal
'        DoCmd.SelectObject acForm, "PlotCheck"
'        DoCmd.Close acForm, "frm_Visit_Date"
'        DoCmd.SelectObject acForm, "frm_Data_Entry"
      
'        MsgBox "Sorry, I'm overtaxed right now..." & vbCrLf & vbCrLf & _
'                "...I can't seem to get the " & Template & _
'                " query to run.", vbOKOnly, "Oops!"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRecords[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     SetRecord
' Description:  Insert/update/delete record based on template
' Assumptions:  -
' Parameters:   template - SQL template name (string)
'               params - array of parameters for template (variant)
' Returns:      id - ID of record inserted, updated, deleted (long integer)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/21/2016 - updated i_login parameters
'   BLC - 10/24/2016 - added flag templates (contact, site, mod wentworth)
'   BLC - 10/28/2016 - updated TempVars("ContactID") -> TempVars("AppUserID"), updated i_task
'   BLC - 1/24/2017 - added IsNPS flag parameter for contacts
'   BLC - 3/24/2017 - set SkipRecordAction = False for uplands, removed unused big rivers cases,
'                     added uplands cases, delete cases
'   BLC - 3/29/2017 - added FieldOK, FieldCheck, Dependencies parameters for templates
'   BLC - 4/24/2017 - add surface/species cover, set SkipRecordAction = false (invasives, uplands)
'   BLC - 7/14/2017 - add u_transect_data
'   BLC - 7/16/2017 - revise u_transect_data to exclude NULLable start time, add u_transect_start_time
'   BLC - 7/17/2017 - add u_quadrat_flags, u_event_(startdate,observer,comments)
'   BLC - 7/18/2017 - add species cover templates (u_speciescover, d_speciescover, i_speciescover)
'   BLC - 7/26/2017 - add u_surfacecover_by_ID template
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
' --------------------------------------------------------------------
'   BLC - 9/28/2017 - update i_sensitive_locations to pull parkID from TempVars("ParkID")
'   BLC - 9/29/2017 - add SiteID parameter for i_location
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
'   BLC - 10/17/2017 - added error 3022 - duplicate record handling, update u_site, i_feature, u_feature
'                      parameters
'   BLC - 10/31/2017 - added ReplicatePlot, CalibrationPlot flags, updated VegPlot case
'   BLC - 11/6/2017 - add i_site_vegtransect
'   BLC - 11/11/2017 - updated i_vegplot, u_vegplot
'   BLC - 11/12/2017 - added i_unknown, u_unknown_identify, u_unknown
'   BLC - 11/24/2017 - fix i_feature parameters
'   BLC - 11/26/2017 - revise i_vegplot parameters (social trails)
'   BLC - 12/5/2017 - add vegplot BeaverBrowse, update i_vegplot pctfa parameter
'   BLC - 1/10/2018 - update i_photo case
'   BLC - 1/16/2018 - add i_photo_event case
'   BLC - 1/17/2018 - updated i_photo_event, i_photo, u_photo cases
' ---------------------------------
Public Function SetRecord(template As String, Params As Variant) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim SkipRecordAction As Boolean
    Dim ID As Long
    
    'exit w/o values
    If Not IsArray(Params) Then GoTo Exit_Handler
    
    'default <-- upland/invasives donot have RecordAction table implemented so skip!
    SkipRecordAction = True 'False
            
    'default ID (if not set as param)
    ID = 0
    
    Set db = CurrDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .SQL = GetTemplate(template)
            
            '-------------------
            ' set SQL parameters --> .Parameters("") = params()
            '-------------------
            
            '-------------------------------------------------------------------------
            ' NOTE:
            '   param(0) --> reserved for record action RefTable (ReferenceType)
            '   last param(x) --> used as record ID for updates
            '-------------------------------------------------------------------------
            Select Case template
            
        '-----------------------
        '  INSERTS
        '-----------------------
                
            '-------------------
            ' --- BIG RIVERS ---
            '-------------------
                Case "i_comment"
                    '-- required parameters --
                    .Parameters("comtype") = Params(1)             'CommentType -> table
                    .Parameters("ctid") = Params(2)                 'TypeID
                    .Parameters("cmt") = Params(3)                  'Comment
                    .Parameters("CID") = Params(4)                  'CommentorID
                    
'                    .Parameters("CreateDate") = Now()
'                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID
'                    .Parameters("LastModified") = Now()
                    .Parameters("LMID") = TempVars("AppUserID")     'LastModifiedByID -> ContactID
        
                Case "i_contact", "i_contact_new"
                    '-- required parameters --
                    .Parameters("First") = Params(1)
                    .Parameters("Last") = Params(2)
                    .Parameters("EmailAddress") = Params(3)
                    .Parameters("Login") = Params(4)
                    .Parameters("Org") = Params(5)
                    .Parameters("MI") = Params(6)
                    .Parameters("Position") = Params(7)
                    .Parameters("Phone") = Params(8)
                    .Parameters("Ext") = Params(9)
                    .Parameters("IsActiveFlag") = Params(10)
                    .Parameters("IsNPSFlag") = Params(11)
                    
                Case "i_contact_access"
                    '-- required parameters --
                    .Parameters("ContactID") = Params(1)
                    .Parameters("AccessID") = Params(2)
                
                    'don't record the action or return ID
                    SkipRecordAction = True
                
                Case "i_cover_species"
                    'set the table name in the template --> handles WCC, URC, ARC species
                    .SQL = Replace(.SQL, "INTO tbl ", "INTO " & Params(0) & " ")
                                    
                    '-- required parameters --
                    .Parameters("VegPlotID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("PctCover") = Params(3)
                        
'        params(0) = "WoodyCanopySpecies"
'        params(1) = .VegPlotID
'        params(2) = .MasterPlantCode
'        params(3) = .PercentCover

'        params(0) = "RootedSpecies"
'        params(1) = .VegPlotID
'        params(2) = .MasterPlantCode
'        params(3) = .PercentCover
                
                Case "i_event"
                    '-- required parameters --
                    .Parameters("SID") = Params(1)
                    .Parameters("LID") = Params(2)
                    .Parameters("PID") = Params(3)
                    .Parameters("Start") = Params(4)
     
                Case "i_event_photo"
                    '-- required parameters --
                    .Parameters("eid") = Params(0)
                    .Parameters("pid") = Params(1)
     
                Case "i_feature"
                    '-- required parameters --
                    .Parameters("LocID") = Params(1)            'FeatureID
                    .Parameters("feat") = Params(2)             'FeatureName
                    .Parameters("Descr") = Params(3)            'Description
                    .Parameters("Dirs") = Params(4)             'Directions
                
                Case "i_imported_data"
                    '-- required parameters --
                    .Parameters("idate") = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                    .Parameters("sfile") = Params(1)
                    .Parameters("dtbl") = Params(2)
                    .Parameters("nrec") = Params(3)
                    .Parameters("srec") = Params(4)
                    .Parameters("erec") = Params(5)
                    
                Case "i_location"
                    '-- required parameters --
                    .Parameters("csn") = Params(1)           'CollectionSourceName
                    .Parameters("ltype") = Params(2)         'LocationType
                    .Parameters("lname") = Params(3)         'LocationName
                    .Parameters("dist") = Params(4)          'HeadtoOrientDistance
                    .Parameters("brg") = Params(5)           'HeadtoOrientBearing
                    .Parameters("lnotes") = Params(6)        'Notes
                    .Parameters("sid") = TempVars("SiteID")  'Site
                    
                    '.Parameters("CreateDate") = Now()
                    .Parameters("CID") = TempVars("AppUserID")  'CreatedByID
                    '.Parameters("LastModified") = Now()
                    .Parameters("LMID") = TempVars("AppUserID") 'LastModifiedByID
                                                        
                Case "i_login"
                    '-- required parameters --
                    .Parameters("uname") = Params(1) 'username
                    .Parameters("activity") = Params(2) 'activity
                    .Parameters("version") = TempVars("AppVersion")
                    .Parameters("accesslvl") = TempVars("UserAccessLevelID")

Debug.Print "uname: " & Params(1) & " activity: " & Params(2) & _
            " version: " & TempVars("AppVersion") & " accesslvl: " & TempVars("UserAccessLevelID")
                    
                    SkipRecordAction = True
                
                Case "i_park"
                    '-- required parameters --
                    .Parameters("ParkCode") = Params(1)
                    .Parameters("ParkName") = Params(2)
                    .Parameters("ParkState") = Params(3)
                    .Parameters("IsActiveForProtocol") = Params(4)
                                                        
                Case "i_photo"
                    '-- required parameters --
                    .Parameters("pdate") = Params(1)            'PhotoDate
                    .Parameters("ptype") = Params(2)            'PhotoType
                    .Parameters("photogid") = Params(3)         'PhotographerID
                    .Parameters("digfname") = Params(4)         'FileName
                    .Parameters("NID") = Params(5)              'NCPNImageID
                    .Parameters("pfacing") = Params(6)          'DirectionFacing
                    .Parameters("ploc") = Params(7)             'PhotogLocation
                    .Parameters("plocdesc") = Params(13)        'PhotogLocationDesc
                    .Parameters("porient") = Params(14)         'PhotogOrientation
                    .Parameters("sptid") = Params(15)           'SurveyPointID
                    .Parameters("sloc") = Params(16)            'SubjectLocation
                    .Parameters("closeup") = Params(8)          'IsCloseup
                    .Parameters("inact") = Params(9)            'IsInActive
                    .Parameters("skip") = Params(10)            'IsSkipped
                    .Parameters("replacemt") = Params(11)       'IsReplacement
                    .Parameters("lpupdate") = Params(12)        'LastPhotoUpdate
                    .Parameters("pfname") = Params(17)          'PhotoFilename
                    .Parameters("path") = Params(18)            'PhotoPath
                    
'                    .Parameters("") = Now()CreateDate
                    .Parameters("CID") = TempVars("AppUserID")  'CreatedByID ContactID")
'                    .Parameters("") = Now() 'LastModified
                    .Parameters("LMID") = TempVars("AppUserID") 'LastModifiedByID ContactID")
                
                Case "i_record_action"
                    '-- required parameters --
                    .Parameters("RefTable") = Params(0)
                    .Parameters("RefID") = Params(1)
                    .Parameters("ID") = Params(2)
                    .Parameters("Activity") = Params(3)
                    .Parameters("ActionDate") = Params(4)
                    
                    SkipRecordAction = True
                
                Case "i_sensitive_locations"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkID")
                    .Parameters("lid") = Params(1)
                    .Parameters("CID") = TempVars("AppUserID")
                    .Parameters("LMID") = TempVars("AppUserID")
                    
                Case "i_sensitive_species"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkID")
                    .Parameters("sp") = Params(1)
                    .Parameters("CID") = TempVars("AppUserID")
                    .Parameters("LMID") = TempVars("AppUserID")
                    
                Case "i_site"
                    '-- required parameters --
                    .Parameters("parkID") = Params(1)
                    .Parameters("riverID") = Params(2)
                    .Parameters("code") = Params(3)         'SiteCode
                    .Parameters("sname") = Params(4)        'SiteName
                    'use |flag| to force 1/0 values vs. Access False (0) & True (-1)
                    .Parameters("flag") = Abs(Params(5))    'IsActiveForProtocol
                    
                    '-- optional parameters --
                    'NOTE: parameters are limited to 255 char
                    '      dir may be truncated via parameter since it's a MEMO field
                    .Parameters("dir") = Params(6)          'Directions
                    .Parameters("descr") = Params(7)        'Description
                
                Case "i_site_vegtransect"           'populate site_vegtransect linking table
                    '-- required parameters --
                    .Parameters("sid") = Params(1) 'Site ID
                    .Parameters("tid") = Params(2) 'VegTransect ID
                    
                    SkipRecordAction = True
                
                Case "i_tagline"
                    '-- required parameters --
                    .Parameters("LineDistSource") = Params(1)
                    .Parameters("LineDistSourceID") = Params(2)
                    .Parameters("LineDistType") = Params(3)
                    .Parameters("LineDistance") = Params(4)
                    .Parameters("HeightType") = Params(5)
                    .Parameters("Height") = Params(6)
                
                Case "i_task"
                    '-- required parameters --
                    .Parameters("descr") = Params(1)         'Task
                    .Parameters("stat") = Params(2)         'Status
                    .Parameters("prio") = Params(3)         'Priority
                    .Parameters("ttype") = Params(4)        'TaskType
                    .Parameters("typeident") = Params(5)    'TaskTypeID
                    .Parameters("RID") = Params(6)          'RequestedByID
                    .Parameters("reqdate") = Params(7)      'RequestDate
                    .Parameters("CID") = Params(8)          'CompletedByID
                    .Parameters("compldate") = Params(9)    'CompleteDate
                
                    '.Parameters("CreateDate") = Now()                  'CreateDate
                    '.Parameters("CreatedByID") = TempVars("ContactID") 'CreatedByID
                    '.Parameters("LastModified") = Now()                'LastModified
                    .Parameters("LMID") = TempVars("AppUserID") 'ContactID")  'lastmodifiedID
                
                Case "i_template"
                    '-- required parameters --
                    .Parameters("tname") = Params(1)        'TemplateName
                    .Parameters("contxt") = Params(2)       'Context
                    '.Parameters("tmpl").Type = dbMemo       'set it to a memo field
                    'Limit template SQL to 255 characters to avoid
                    'error 3271 SetRecord mod_App_Data Invalid property value.
                    'templates > 255 characters must be edited directly in the table
                    .Parameters("tmpl") = Left(Params(3), 255) 'TemplateSQL
                    .Parameters("rmks") = Params(4)         'Remarks
                    .Parameters("effdate") = Params(5)      'EffectiveDate
                    .Parameters("cid") = Params(6)          'CreatedBy_ID (contactID)
                    .Parameters("prms") = Params(7)         'Params
                    .Parameters("syntx") = Params(8)        'Syntax
                    .Parameters("vers") = Params(9)         'Version
                    .Parameters("sflag") = Params(10)       'IsSupported
                    .Parameters("lmid") = TempVars("AppUserID") 'lastmodifiedID
                
                Case "i_transducer"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("TransducerType") = Params(2)
                    .Parameters("TransducerNumber") = Params(3)
                    .Parameters("SerialNumber") = Params(4)
                    .Parameters("IsSurveyed") = Params(5)
                    .Parameters("Timing") = Params(6)
                    .Parameters("ActionDate") = Params(7)
                    .Parameters("ActionTime") = Params(8)
                
                Case "i_understory_species"
                    '-- required parameters --
                    .Parameters("VegPlotID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("PercentCover") = Params(3)
                    .Parameters("IsSeedling") = Params(4)
                
                Case "i_unknown"
                    '-- required parameters --
                    .Parameters("ucode") = Params(1)
                    .Parameters("ptype") = Params(2)
                    .Parameters("pdescr") = Params(3)
                    .Parameters("sfeat") = Params(4)
                    .Parameters("ltype") = Params(5)
                    .Parameters("lmarg") = Params(6)
                    .Parameters("lch") = Params(7)
                    .Parameters("sch") = Params(8)
                    .Parameters("fch") = Params(9)
                    .Parameters("gch") = Params(10)
                    .Parameters("forb") = Params(11)
                    .Parameters("pere") = Params(12)
                    .Parameters("bguess") = Params(13)
                    .Parameters("pix") = Params(14)
                    .Parameters("coll") = Params(15)
                    .Parameters("collmeth") = Params(16)
                    .Parameters("lid") = Params(17)
                    .Parameters("cid") = Params(18)
                     
                Case "i_vegplot"
                    '-- required parameters --
                    .Parameters("EID") = Params(1)              'Event ID
                    .Parameters("SID") = Params(2)              'Site ID
                    .Parameters("FID") = Params(3)              'Feature ID
                    .Parameters("TID") = Params(4)              'VegTransect ID
                    .Parameters("pnum") = Params(5)             'PlotNumber
                    .Parameters("pdist") = Params(6)            'PlotDistance
                    .Parameters("MSSID") = Params(7)            'ModalSedimentSize ID
                    .Parameters("pctmss") = Params(13)          'PctModalSedimentSize
                    .Parameters("pctfines") = Params(8)         'PercentFines
                    .Parameters("pctwater") = Params(9)         'PercentWater
                    .Parameters("pcturc") = Params(10)          'UnderstoryRootedPctCover
                    .Parameters("pctwcc") = Params(11)          'WoodyCanopyPctCover
                    .Parameters("pctarc") = Params(12)          'AllRootedPctCover
                    .Parameters("pd") = Params(18)              'PlotDensity
                    .Parameters("nocanopy") = Params(19)        'NoCanopyVeg
                    .Parameters("norooted") = Params(20)        'NoRootedVeg
                    '.Parameters("hastrails") = Params(21)       'HasSocialTrails
                    .Parameters("pctfa") = Params(14)           'PctFilamentousAlgae
                    .Parameters("pctlitter") = Params(15)       'PctLitter
                    .Parameters("pctwd") = Params(16)           'PctWoodyDebris
                    .Parameters("pctsd") = Params(17)           'PctStandingDead
                    .Parameters("noindsp") = Params(22)         'NoIndicatorSpecies
                    .Parameters("cplot") = Params(23)           'CalibrationPlot
                    .Parameters("rplot") = Params(24)           'ReplicatePlot
                    .Parameters("pctst") = Params(25)           'SocialTrailsPct
                    .Parameters("bb") = Params(26)              'BeaverBrowse
                    
                Case "i_vegtransect"
                    '-- required parameters --
                    .Parameters("LocationID") = Params(1)
                    .Parameters("EventID") = Params(2)
                    .Parameters("TransectNumber") = Params(3)
                    .Parameters("SampleDate") = Params(4)
        
                Case "i_vegwalk"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("CollectionPlaceID") = Params(2)
                    .Parameters("CollectionType") = Params(3)
                    .Parameters("StartDate") = Params(4)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                    
                Case "i_vegwalk_species"
                    '-- required parameters --
                    .Parameters("VegWalkID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("IsSeedling") = Params(3)
                    
                Case "i_waterway"
                    '-- required parameters --
                    .Parameters("ParkID") = Params(1)
                    .Parameters("Name") = Params(2)
                    .Parameters("Segment") = Params(3)
                    
                Case "i_usys_temp_photo"
                    '-- required parameters --
                    .Parameters("ppath") = Params(1)
                    .Parameters("pfile") = Params(2)
                    .Parameters("pdate") = Params(3)
                    .Parameters("ptype") = Params(4)
            
            '-------------------
            ' --- UPLAND & INVASIVES ---
            '-------------------
                Case "i_new_transect_quadrat"
                    '-- required parameters --
                    .Parameters("tid") = Params(1)  'record ID
                    .Parameters("qnum") = Params(2) 'quadrat number (1-3)
                
                Case "i_new_transect_quadrat_sfccover"
                    '-- required parameters --
                    .Parameters("qid") = Params(1)  'record ID
                    .Parameters("sid") = Params(2)  'surface microhabitat ID
                
                Case "i_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)  'record ID
                    .Parameters("num") = Params(2)  'number of records
                    .Parameters("fok") = Params(3)  'field ok? (QC pass/fail)
                    
                Case "i_speciescover"
                    '-- required parameters --
                    .Parameters("qid") = Params(1)      'quadrat ID
                    .Parameters("plant") = Params(2)    'plant lookup code
                    .Parameters("dead") = Params(3)     'is dead flag
                    .Parameters("pct") = Params(4)      'percent cover
                    
                Case "i_template"
                    '-- required parameters --
                    .Parameters("tname") = Params(1)        'TemplateName
                    .Parameters("contxt") = Params(2)       'Context
                    '.Parameters("tmpl").Type = dbMemo       'set it to a memo field
                    'Limit template SQL to 255 characters to avoid
                    'error 3271 SetRecord mod_App_Data Invalid property value.
                    'templates > 255 characters must be edited directly in the table
                    .Parameters("tmpl") = Left(Params(3), 255) 'TemplateSQL
                    .Parameters("rmks") = Params(4)         'Remarks
                    .Parameters("effdate") = Params(5)      'EffectiveDate
                    .Parameters("cid") = Params(6)          'CreatedBy_ID (contactID)
                    .Parameters("prms") = Params(7)         'Params
                    .Parameters("syntx") = Params(8)        'Syntax
                    .Parameters("vers") = Params(9)         'Version
                    .Parameters("sflag") = Params(10)       'IsSupported
                    .Parameters("lmid") = TempVars("AppUserID") 'lastmodifiedID
                    .Parameters("fqc") = Params(11)         'FieldCheck
                    .Parameters("fok") = Params(12)         'FieldOK
                    .Parameters("dep") = Params(13)         'Dependencies
                
                Case "i_surface_cover"
                    '-- required parameters --
                    .Parameters("qid") = Params(1)
                    .Parameters("sid") = Params(2)
                    .Parameters("pct") = Params(3)
                                    
        '-----------------------
        '  UPDATES
        '-----------------------
                
            '-------------------
            ' --- BIG RIVERS ---
            '-------------------
                Case "u_comment"
                    '-- required parameters --
                    .Parameters("CommentType") = Params(1)
                    .Parameters("TypeID") = Params(2)
                    .Parameters("Comment") = Params(3)
                    .Parameters("CommentorID") = Params(4)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                    
                Case "u_contact"
                    '-- required parameters --
                    .Parameters("First") = Params(1)
                    .Parameters("Last") = Params(2)
                    .Parameters("EmailAddress") = Params(3)
                    .Parameters("Login") = Params(4)
                    .Parameters("Org") = Params(5)
                    .Parameters("MI") = Params(6)
                    .Parameters("Position") = Params(7)
                    .Parameters("Phone") = Params(8)
                    .Parameters("Ext") = Params(9)
                    .Parameters("IsActiveFlag") = Params(10)
                    .Parameters("IsNPSFlag") = Params(11)
                    .Parameters("ContactID") = Params(12)
                    ID = Params(12)
                
                Case "u_contact_access"
                    '-- required parameters --
                    .Parameters("ContactID") = Params(1)
                    .Parameters("AccessID") = Params(2)
                    ID = Params(1)
                
                Case "u_contact_isactive_flag"
                    '-- required parameters --
                    .Parameters("cid") = Params(1)
                    .Parameters("flag") = Params(2)
                
                Case "u_cover_species"
                    'set the table name in the template --> handles WCC, URC, ARC species
                    .SQL = Replace(.SQL, " tbl ", " " & Params(0) & " ")
                                    
                    '-- required parameters --
                    .Parameters("VegPlot_ID") = Params(1)
                    .Parameters("Master_PLANT_Code") = Params(2)
                    .Parameters("PctCover") = Params(3)
                
                Case "u_event"
                    '-- required parameters --
                    .Parameters("SID") = Params(1)
                    .Parameters("LID") = Params(2)
                    .Parameters("PID") = Params(3)
                    .Parameters("Start") = Params(4)
                    .Parameters("EID") = Params(5)
                    ID = Params(5)
                    
                Case "u_feature"
                    '-- required parameters --
                    .Parameters("LocID") = Params(1)        'LocationID
                    .Parameters("feat") = Params(2)         'Feature
                    .Parameters("Descr") = Params(3)        'FeatureDescription
                    .Parameters("Dirs") = Params(4)         'FeatureDirections
                    .Parameters("FID") = Params(5)
                    ID = Params(5)
                    
                Case "u_location"
                    '-- required parameters --
                    .Parameters("CollectionSourceName") = Params(1)
                    .Parameters("LocationType") = Params(2)
                    .Parameters("LocationName") = Params(3)
                    .Parameters("HeadtoOrientDistance") = Params(4)
                    .Parameters("HeadtoOrientBearing") = Params(5)
                    
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") 'ContactID")
                
                Case "u_mod_wentworth_retireyear"
                    '-- required parameters --
                    .Parameters("mwsid") = Params(1)
                    .Parameters("yr") = Params(2)
                
                Case "u_park"
                    '-- required parameters --
                    .Parameters("ParkCode") = Params(1)
                    .Parameters("ParkName") = Params(2)
                    .Parameters("ParkState") = Params(3)
                    .Parameters("IsActiveForProtocol") = Params(4)
                        
                Case "u_photo"
                    '-- required parameters --
                    .Parameters("pdate") = Params(1)            'PhotoDate
                    .Parameters("ptype") = Params(2)            'PhotoType
                    .Parameters("photogid") = Params(3)         'PhotographerID
                    .Parameters("digfname") = Params(4)         'FileName
                    .Parameters("NID") = Params(5)              'NCPNImageID
                    .Parameters("pfacing") = Params(6)          'DirectionFacing
                    .Parameters("ploc") = Params(7)             'PhotogLocation
                    .Parameters("plocdesc") = Params(13)        'PhotogLocationDesc
                    .Parameters("porient") = Params(14)         'PhotogOrientation
                    .Parameters("sptid") = Params(15)           'SurveyPointID
                    .Parameters("sloc") = Params(16)            'SubjectLocation
                    .Parameters("closeup") = Params(8)          'IsCloseup
                    .Parameters("inact") = Params(9)            'IsInActive
                    .Parameters("skip") = Params(10)            'IsSkipped
                    .Parameters("replacemt") = Params(11)       'IsReplacement
                    .Parameters("lpupdate") = Params(12)        'LastPhotoUpdate
                    .Parameters("pfname") = Params(17)          'PhotoFilename
                    .Parameters("path") = Params(18)            'PhotoPath
                    
                    .Parameters("CID") = TempVars("AppUserID")  'CreatedByID ContactID")
                    .Parameters("LMID") = TempVars("AppUserID") 'LastModifiedByID ContactID")
                
                Case "u_site"
                    '-- required parameters --
                    '.Parameters("ParkID") = Params(1)
                    '.Parameters("RiverID") = Params(2)
                    .Parameters("scode") = Params(3)
                    .Parameters("sname") = Params(4)
                    .Parameters("flag") = Params(5)         'IsActiveForProtocol
                    
                    '-- optional parameters --
                    .Parameters("sdir") = Params(6)             'Directions
                    .Parameters("sdesc") = Params(7)             'Description
                    .Parameters("sid") = Params(8)
                
                Case "u_site_isactive_flag"
                    '-- required parameters --
                    .Parameters("sid") = Params(1)
                    .Parameters("flag") = Params(2)
                
                Case "u_tagline"
                    '-- required parameters --
                    .Parameters("LineDistSource") = Params(1)
                    .Parameters("LineDistSourceID") = Params(2)
                    .Parameters("LineDistType") = Params(3)
                    .Parameters("LineDistance") = Params(4)
                    .Parameters("HeightType") = Params(5)
                    .Parameters("Height") = Params(6)
                
                Case "u_task"
                    '-- required parameters --
                    .Parameters("tid") = Params(14)         'task ID
                    .Parameters("descr") = Params(1)        'task
                    .Parameters("stat") = Params(2)         'status
                    .Parameters("prio") = Params(3)         'priority
                    .Parameters("ttype") = Params(4)        'task type
                    .Parameters("typeident") = Params(5)    'task type ID
                    .Parameters("RID") = Params(3)          'requested by ID
                    .Parameters("reqdate") = Params(7)      'request date
                    .Parameters("CID") = Params(5)          'completed by ID
                    .Parameters("compldate") = Params(9)    'complete date
                
                    .Parameters("LMID") = TempVars("AppUserID") 'last modified by ID
                
                Case "u_transducer"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("TransducerType") = Params(2)
                    .Parameters("TransducerNumber") = Params(3)
                    .Parameters("SerialNumber") = Params(4)
                    .Parameters("IsSurveyed") = Params(5)
                    .Parameters("Timing") = Params(6)
                    .Parameters("ActionDate") = Params(7)
                    .Parameters("ActionTime") = Params(8)
                
                Case "u_template"
                    '-- required parameters --
                    .Parameters("id") = Params(1)
                
                Case "u_tsys_datasheet_defaults"
                    '-- required parameters --
                    .Parameters("id") = Params(1)
                    .Parameters("pid") = Params(2)
                    .Parameters("rid") = Params(3)
                    .Parameters("cover") = Params(4)
                    .Parameters("species") = Params(5)
                    .Parameters("blanks") = Params(6)
                    
                    '-- optional parameters --
                
                Case "u_usys_temp_photo"
                    '-- required parameters --
                    .Parameters("iid") = Params(1)
                    .Parameters("ptype") = Params(4)
                
                Case "u_vegplot"
                    '-- required parameters --
                    .Parameters("vid") = Params(23)             'VegPlot ID
                    .Parameters("EID") = Params(1)              'Event ID
                    .Parameters("SID") = Params(2)              'Site ID
                    .Parameters("FID") = Params(3)              'Feature ID
                    .Parameters("TID") = Params(4)              'VegTransect ID
                    .Parameters("pnum") = Params(5)             'PlotNumber
                    .Parameters("pdist") = Params(6)            'PlotDistance
                    .Parameters("MSSID") = Params(7)            'ModalSedimentSize ID
                    .Parameters("pctmss") = Params(13)          'PctModalSedimentSize
                    .Parameters("pctfines") = Params(8)         'PctFines
                    .Parameters("pctwater") = Params(9)         'PctWater
                    .Parameters("pcturc") = Params(10)          'UnderstoryRootedPctCover
                    .Parameters("pctwcc") = Params(11)          'WoodyCanopyPctCover
                    .Parameters("pctarc") = Params(12)          'AllRootedPctCover
                    .Parameters("pd") = Params(18)              'PlotDensity
                    .Parameters("nocanopy") = Params(19)        'NoCanopyVeg
                    .Parameters("norooted") = Params(20)        'NoRootedVeg
                    '.Parameters("hastrails") = Params(21)       'HasSocialTrails
                    .Parameters("fa") = Params(14)              'FilamentousAlgae
                    .Parameters("pctlitter") = Params(15)       'PctLitter
                    .Parameters("pctwd") = Params(16)           'PctWoodyDebris
                    .Parameters("pctsd") = Params(17)           'PctStandingDead
                    .Parameters("noindsp") = Params(22)         'NoIndicatorSpecies
                    .Parameters("cplot") = Params(23)           'CalibrationPlot
                    .Parameters("rplot") = Params(24)           'ReplicatePlot
                    .Parameters("pctst") = Params(25)           'PctSocialTrails
                    .Parameters("bb") = Params(26)              'BeaverBrowse
                
                Case "u_vegtransect"
                    '-- required parameters --
                    .Parameters("LocationID") = Params(1)
                    .Parameters("EventID") = Params(2)
                    .Parameters("TransectNumber") = Params(3)
                    .Parameters("SampleDate") = Params(4)
                
                Case "u_vegwalk"
                    '-- required parameters --
                    .Parameters("EventID") = Params(1)
                    .Parameters("CollectionPlaceID") = Params(2)
                    .Parameters("CollectionType") = Params(3)
                    .Parameters("StartDate") = Params(4)
                    
                    .Parameters("CreateDate") = Now()
                    .Parameters("CreatedByID") = TempVars("AppUserID") 'ContactID")
                    .Parameters("LastModified") = Now()
                    .Parameters("LastModifiedByID") = TempVars("AppUserID") '"ContactID")
                
                Case "u_vegwalk_species"
                    '-- required parameters --
                    .Parameters("VegWalkID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("IsSeedling") = Params(3)
                    
                Case "u_understory_species"
                    '-- required parameters --
                    .Parameters("VegPlotID") = Params(1)
                    .Parameters("MasterPlantCode") = Params(2)
                    .Parameters("PercentCover") = Params(3)
                    .Parameters("IsSeedling") = Params(4)
                
                Case "u_unknown"
                    '-- required parameters --
                    .Parameters("uid") = Params(25)
                    .Parameters("ucode") = Params(1)
                    .Parameters("ptype") = Params(2)
                    .Parameters("pdescr") = Params(3)
                    .Parameters("sfeat") = Params(4)
                    .Parameters("ltype") = Params(5)
                    .Parameters("lmarg") = Params(6)
                    .Parameters("lch") = Params(7)
                    .Parameters("sch") = Params(8)
                    .Parameters("fch") = Params(9)
                    .Parameters("gch") = Params(10)
                    .Parameters("forb") = Params(11)
                    .Parameters("pere") = Params(12)
                    .Parameters("bguess") = Params(13)
                    .Parameters("pix") = Params(14)
                    .Parameters("coll") = Params(15)
                    .Parameters("collmeth") = Params(16)
                    .Parameters("lid") = Params(17)
                    .Parameters("cid") = Params(18)
                
                Case "u_unknown_identify"
                    '-- required parameters --
                    .Parameters("uid") = Params(1)
                    .Parameters("ccode") = Params(2)
                    .Parameters("idate") = Params(3)
                    .Parameters("cid") = Params(4)
                
                Case "u_waterway"
                    '-- required parameters --
                    .Parameters("ParkID") = Params(1)
                    .Parameters("Name") = Params(2)
                    .Parameters("Segment") = Params(3)
            
            '-------------------
            ' --- UPLAND & INVASIVES ---
            '-------------------
                Case "u_event_comments"
                    '-- required parameters --
                    .Parameters("eid") = Params(1)
                    .Parameters("cmt") = Params(2)
                
                Case "u_event_observer"
                    '-- required parameters --
                    .Parameters("eid") = Params(1)
                    .Parameters("oid") = Params(2)
                
                Case "u_event_startdate"
                    '-- required parameters --
                    .Parameters("eid") = Params(1)
                    .Parameters("start") = Params(2)
                
                Case "u_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)
                    .Parameters("num") = Params(2)
                    .Parameters("fok") = Params(3)
                    
                Case "u_quadrat_flags"
                    '-- required parameters --
                    .Parameters("qid") = Params(1)
                    .Parameters("is") = Params(2)
                    .Parameters("ne") = Params(3)
                
                Case "u_speciescover"
                    '-- required parameters --
                    .Parameters("scid") = Params(1)     'species cover record ID
                    .Parameters("qid") = Params(2)      'quadrat ID
                    .Parameters("plant") = Params(3)    'plant lookup code
                    .Parameters("dead") = Params(4)     'is dead flag
                    .Parameters("pct") = Params(5)      'percent cover
                
                Case "u_surface_cover"
                    '-- required parameters --
                    .Parameters("sfcid") = Params(1)
                    .Parameters("qid") = Params(2)
                    .Parameters("sid") = Params(3)
                    .Parameters("pct") = Params(4)
                    
                Case "u_surfacecover_by_id"
                    '-- required parameters --
                    .Parameters("sfcid") = Params(1)
                    .Parameters("pct") = Params(2)
                
                Case "u_template"
                    '-- required parameters --
                    .Parameters("id") = Params(1)
                
                Case "u_transect_data"
                    '-- required parameters --
                    .Parameters("oid") = Params(1)      'observer
                    .Parameters("cmt") = Params(2)      'comments
                    .Parameters("tid") = Params(3)      'transect quadrat ID
                    
                Case "u_transect_comments"
                    '-- required parameters --
                    .Parameters("cmt") = Params(1)      'comments
                    .Parameters("tid") = Params(2)      'transect quadrat ID
                
                Case "u_transect_observer"
                    '-- required parameters --
                    .Parameters("oid") = Params(1)      'observer
                    .Parameters("tid") = Params(2)      'transect quadrat ID
                    
                Case "u_transect_start_time"
                    '-- required parameters --
                    .Parameters("start") = Params(1)    'start time
                    .Parameters("tid") = Params(2)      'transect quadrat ID
                    
        '-----------------------
        '  DELETES
        '-----------------------
                Case "d_num_records_all"
                    '-- required parameters --
                
                Case "d_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)
            
                Case "d_speciescover"
                    '-- required parameters --
                    .Parameters("scid") = Params(1)
'                    .Parameters("qid") = params(2)
'                    .Parameters("plant") = params(3)
'                    .Parameters("dead") = params(4)
            
                Case "d_sensitive_locations"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkID")
                    .Parameters("lid") = Params(1)
                
                Case "d_sensitive_species"
                    '-- required parameters --
                    .Parameters("pkid") = TempVars("ParkID")
                    .Parameters("sid") = Params(1)
                
            End Select
'Debug.Print .sql
            .Execute dbFailOnError
                
            If ID = 0 Then
                'retrieve identity
                ID = db.OpenRecordset("SELECT @@IDENTITY;")(0)
            End If
    
    ' -------------------
    '  Record Action
    ' -------------------
            'handle unrecorded actions & those which don't generate an ID
            If SkipRecordAction Then GoTo Exit_Handler
            
            'set record action
            .SQL = GetTemplate("i_record_action")
                                            
            '-- required parameters --
            .Parameters("RefTable") = Params(0)
            .Parameters("RefID") = ID
            .Parameters("ID") = TempVars("AppUserID") 'TempVars("ContactID")
            .Parameters("Activity") = "DE"
            .Parameters("ActionDate") = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                                
            .Execute dbFailOnError
            
            'cleanup
            .Close
        
        End With

    End With
                
Exit_Handler:
    'add temp var to capture cases when object is no longer available (e.g. photo batch update)
    SetTempVar "NewRecordID", ID
    SetRecord = ID
    
    'cleanup
    Set qdf = Nothing
    Set db = Nothing

    Exit Function
Err_Handler:
    Select Case Err.Number
          Case 3022
            'show added record message & clear
            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                        "msg" & PARAM_SEPARATOR & "Please check your " & _
                        "values && retry if necessary. " & _
                        vbCrLf & "[" & template & " - SetRecord]" & _
                        "|Type" & PARAM_SEPARATOR & "caution" & _
                        "|Title" & PARAM_SEPARATOR & "Duplicate Record!"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetRecord[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     UpsertRecord
' Description:  Handle insert/update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
'   gecko_1, February 10, 2005
'   http://www.access-programmers.co.uk/forums/showthread.php?t=81221
'   Khinsu, August 19, 2013
'   http://stackoverflow.com/questions/18317059/how-to-test-if-item-exists-in-recordset
'   HansUp, April 4, 2013
'   http://stackoverflow.com/questions/15823687/findfirst-vba-access2010-unbound-form-runtime-error
' Source/date:  Bonnie Campbell, July 28, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/28/2016 - initial version
'   BLC - 9/1/2016  - added vegwalk, photo
'   BLC - 10/4/2016 - added template, adjusted for form w/o list
'   BLC - 10/14/2016 - updated to accommodate non-users for contacts
'   BLC - 1/9/2017 - revised retrieve ID from ContactID to ID, revised i_event to use TempVar("SiteID")
'   BLC - 2/1/2017 - handle form upserts for forms w/o lists/msg & msg icons
'   BLC - 2/3/2017 - location adjustments
'   BLC - 3/27/2017 - removed big rivers cases, replaced w/ upland cases
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'   BLC - 9/29/2017 - update location ContactID
'   BLC - 10/30/2017 - change location CollectionSourceID to use text vs. index
'   BLC - 10/31/2017 - added ReplicatePlot, CalibrationPlot flags (VegPlot)
'   BLC - 11/9/2017 - add Comment case
'   BLC - 11/10/2017 - added Transducer distances
'   BLC - 11/11/2017 - updated VegPlot
'   BLC - 11/12/2017 - add UnknownSpecies case
'   BLC - 11/26/2017 - added VegSpecies, updated VegWalk, VegPlot
'   BLC - 12/5/2017 - added VegPlot BeaverBrowse, PctSocialTrails
'   BLC - 12/8/2017 - update VegPlot EventID, VegTransectID
'   BLC - 12/12/2017 - add TempPhoto
'   BLC - 1/10/2018  - add PhotoBatchUpdate, Logger cases
'   BLC - 1/16/2018  - update PhotoBatchUpdate case
' ---------------------------------
Public Sub UpsertRecord(ByRef frm As Form)
On Error GoTo Err_Handler
    
' ----------------------------------------------------------------------------------
'    1) Click to edit
'       a) populates form fields
'       b) tbxID is set
'
'       c) change values --> i) compare against existing values
'                           ii) no existing values match ==> update
'                           iii) existing values match ==> message no change
'
'   2) Enter new values
'       a) enables save button
'       b) click save -->   i) compare against existing values
'                           ii) no existing values match ==> insert
'                           iii) existing values match ==> message no change
' ----------------------------------------------------------------------------------
    
    Dim DoAction As String, strCriteria As String, strTable As String
    Dim NoList As Boolean
    Dim obj As Object
    
    'use generic object to handle multiple obj types
    With obj
    
        'default
        NoList = False
        strTable = frm.Name
    
        Select Case frm.Name
            
            '-------------------
            ' --- BIG RIVERS ---
            '-------------------
            Case "Comment"
                Dim cmt As New AppComment
                With cmt
                    'form values
                    .CommentType = frm!tbx
                    .TypeID = Trim(Split(frm!Context, "-")(1))
                    .Comment = frm!tbxComment
                    .CommentorID = TempVars("")
                
                End With
            
            Case "Contact"
                Dim p As New person
    
                With p
                    'values passed into form
                            
                    'form values
                    .LastName = frm!tbxLast.Value
                    .FirstName = frm!tbxFirst.Value
                    If Not IsNull(frm!tbxMI.Value) Then p.MiddleInitial = frm!tbxMI.Value  'FIX EMPTY STRING
                    .Email = frm!tbxEmail.Value
                    '.Username = frm!tbxUsername.Value
                    If Not IsNull(frm!tbxUsername.Value) Then p.UserName = frm!tbxUsername.Value
                    If Not IsNull(frm!tbxOrganization.Value) Then p.Organization = frm!tbxOrganization.Value
                    If Not IsNull(frm!tbxPosition.Value) Then .PosTitle = frm!tbxPosition.Value
                    If Not IsNull(frm!tbxPhone.Value) And Len(frm!tbxPhone.Value) > 0 Then
                        .WorkPhone = RemoveChars(frm!tbxPhone.Value, True) 'remove non-numerics
                    Else
                        .WorkPhone = Null
                    End If
                    If Not IsNull(frm!tbxExtension.Value) And Len(frm!tbxExtension.Value) > 0 Then
                        .WorkExtension = RemoveChars(frm!tbxExtension.Value, True) 'remove non-numerics
                    Else
                        .WorkExtension = Null
                    End If
                    If Not IsNull(frm!cbxUserRole.Column(1)) Then .AccessRole = frm!cbxUserRole.Column(1)
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    strCriteria = "[FirstName] = '" & .FirstName _
                                    & "' AND [LastName] = '" & .LastName _
                                    & "' AND [MiddleInitial] = '" & .MiddleInitial _
                                    & "' AND [Email] = '" & .Email & "'"
                    
                    'set the generic object --> Contact
                    Set obj = p
                    
                    'cleanup
                    Set p = Nothing
                End With

            Case "Events"
                Dim ev As New EventVisit
                strTable = "Event"
                
                With ev
                    'values passed into form
                    
                    'form values
                    .LocationID = frm!cbxLocation.Column(0)
                    .ProtocolID = 1 ' assumes this is for big rivers protocol
                    .SiteID = TempVars("SiteID") 'frm!cbxSite.Column(0)
                    
                    .StartDate = frm!tbxStartDate.Value
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                    
                    strCriteria = "[Site_ID] = " & .SiteID & " AND [Location_ID] = " & .LocationID & " AND [StartDate] = " & Format(.StartDate, "YYYY-mm-dd")
                    
                    'set the generic object --> EventVisit
                    Set obj = ev
                    
                    'cleanup
                    Set ev = Nothing
                End With
            
            Case "Feature"
                Dim f As New Feature

                With f
                    'values passed into form
                            
                    'form values
                    .LocationID = frm!cbxLocation.Column(0)
                    .Name = frm!tbxFeature.Value
                    
                    If Not IsNull(frm!tbxFeatureDirections.Value) Then f.Directions = frm!tbxFeatureDirections.Value
                    If Not IsNull(frm!tbxDescription.Value) Then .Directions = frm!tbxDescription.Value
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    strCriteria = "[Location_ID] = " & .LocationID & " AND [Feature] = '" & .Name & "'"
                    
                    'set the generic object --> Feature
                    Set obj = f
                
                    'cleanup
                    Set f = Nothing
                End With

            Case "Location"
                Dim loc As New Location
                
                With loc
                    'form values
                    
                    'location types: F- feature, T- transect, P - plot
                    .LocationType = frm.LocationType 'cbxLocationType.SelText
                    
                    'CollectionSourceName is the identifier for which
                    'feature/transect/plot the location is located on
                    'collection feature ID (A, B, C...) or Transect number (1-8)
                    '.CollectionSourceName = frm.cbxCollectionSourceID
                    'capture text vs. index
                    .CollectionSourceName = frm.CollectionSourceID
                                                                    
                    .LocationName = frm!tbxName.Value
            
                    .HeadtoOrientDistance = frm!tbxDistance.Value
                    .HeadtoOrientBearing = frm!tbxBearing.Value
                    
                    .LocationNotes = frm!tbxNotes.Value
                    
                    '.CreateDate = ""
                    '.CreatedByID = 0
                    .LastModified = Now()
                    .LastModifiedByID = TempVars("AppUserID") '0
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0

                    'ignore location notes in criteria
                    strCriteria = "[LocationName] = '" & .LocationName _
                                & "' AND [LocationType] = '" & .LocationType _
                                & "' AND [CollectionSourceName] = '" & .CollectionSourceName _
                                & "' AND [HeadtoOrientDistance_m] = " & .HeadtoOrientDistance _
                                & " AND [HeadtoOrientBearing] = " & .HeadtoOrientBearing '_
'                                    & " AND [LastModified] = " & .LastModified _
'                                    & " AND [LastModifiedBy_ID] = " & .LastModifiedByID
                
                    'set the generic object --> Location
                    Set obj = loc
                    
                    'cleanup
                    Set loc = Nothing
                End With
                                        
            Case "Logger"
                'Dim lg As New Transducer
                'With lg
                
                
                'End With
                
                                                        
            Case "Photo"
                Dim ph As New Photo
                
                With ph
                    'values passed into form
                
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                    .PhotographerID = frm!cbxPhotographer
                    .PhotoType = frm!cbxPhotoType
                    .FileName = frm!tbxFilename
                    .PhotoNumber = frm!tbxPhotoNumber
                    .SubjectLocation = frm!tbxSubjectLocation
                    .PhotogLocation = frm!tbxPhotogLocation
                    .DirectionFacing = frm!optDirFacing
                    .PhotoDate = frm!PhotoDate
                    .IsCloseup = frm!chkIsCloseup
                    .IsReplacement = frm!chkIsReplacement
                    .IsSkipped = frm!chkIsSkipped
                    .NCPNImageID = frm!tbxNCPNImageID
                    .PhotoFilename = frm!tbxFilename
                    .PhotoPath = frm!lblPath.Caption
                    
                    'set the generic object --> Photo
                    Set obj = ph
                    
                    'cleanup
                    Set ph = Nothing
                End With
            
            Case "PhotoBatchUpdate"
                Dim pbu As New Photo
                
                With pbu
                    Dim FilePath As String
                    Dim aryFileInfo() As Variant
                    
                    .PhotoType = frm!cbxPhotoType
                    Select Case .PhotoType
                        Case "U" 'unclassified
                        Case "F" 'feature
                        Case "T" 'transect
                        Case "O" 'overview
                        Case "R" 'reference
                        Case "OO" 'other
                    End Select
                    
                    .PhotographerID = frm!cbxPhotographer
                    .FileName = frm!lbxPhotos.Column(3)
                    .PhotoDate = frm!lbxPhotos.Column(4)
                    .PhotoFilename = frm!lbxPhotos.Column(3)
                    .PhotoPath = frm!lbxPhotos.Column(2)
                    
                    .ID = frm!lbxPhotos.Column(0)
                    
                    'set the generic object  -->  Photo
                    Set obj = pbu
                    
                    'cleanup
                    Set pbu = Nothing

                End With
                
                'inserts only, no ID?
                NoList = True
            
            Case "PhotoOtherDetails"
                Dim pho As New Photo
                
                With pho
                    'Dim FilePath As String < already defined (PhotoBatchUpdate)
                    'Dim aryFileInfo() As Variant < already defined (PhotoBatchUpdate)
                    Dim nodeinfo() As String
                    '0 - M, 1- C, 2-full file path, 3-file name w/o extension
                    nodeinfo = Split(frm.Parent!tvwTree.Object.SelectedItem.Tag, "|")
                    FilePath = nodeinfo(2)
                    'filepath = frm.Parent!tvwTree.Object.SelectedItem.Tag 'frm!tvw.SelectedNode.Tag
                    'aryFileInfo = GetFileEXIFInfo()
                    'values passed into form
'        Params(0) = "Photo"
'        Params(1) = .PhotoDate
'        Params(2) = .PhotoType
'        Params(3) = .PhotographerID
'        Params(4) = .FileName
'        Params(5) = .NCPNImageID
'        Params(6) = .DirectionFacing
'        Params(7) = .PhotogLocation
'        Params(8) = .IsCloseup
'        Params(9) = .IsInActive
'        Params(10) = .IsSkipped
'        Params(11) = .IsReplacement
'        Params(12) = .LastPhotoUpdate
                    .PhotoType = frm!lblPhotoType
                    Select Case .PhotoType
                        Case "U" 'unclassified
                        Case "F" 'feature
                        Case "T" 'transect
                        Case "O" 'overview
                        Case "R" 'reference
                        Case "OO" 'other
                    End Select
                    .PhotographerID = frm.fsub.Form.Controls("cbxPhotog")
                    .FileName = "" 'lblPhotoFilename 'aryFileInfo(0)
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                                
                    'set the generic object --> Location
                    Set obj = pho
                    
                    'cleanup
                    Set pho = Nothing
                End With
                                                                                
            Case "SetObserverRecorder"
                Dim ra As New RecordAction
                
                With ra
                    'values passed into form
                    .RefTable = frm.RefTable
                    .RefID = frm.RefID
                    .ContactID = frm.RAContactID
                    .RefAction = frm.RAAction
                    '.ActionType = frm.RAAction
                    .ActionDate = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                                
                    strCriteria = "[Contact_ID] = " & .ContactID _
                                & " AND [Activity] = '" & .RefAction _
                                & "'"

                    'set the generic object --> Location
                    Set obj = ra
                    
                    'cleanup
                    Set ra = Nothing
                End With
            
            Case "Site"
                Dim s As New Site
                
                With s
                    'values passed into form
                    .Park = TempVars("ParkCode")
                    .River = TempVars("River")
                    
                    'form values
                    .Code = frm!tbxSiteCode.Value
                    .Name = frm!tbxSiteName.Value
                    .Directions = Nz(frm!tbxSiteDirections.Value, "")
                    .Description = Nz(frm!tbxDescription.Value, "")
                    
                    'assumed
                    .IsActiveForProtocol = 1 'all sites assumed active when added
        
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    strCriteria = "[SiteCode] = '" & .Code & "' AND [SiteName] = '" & .Name & "'"
                
                    'set the generic object --> Site
                    Set obj = s
                    
                    'cleanup
                    Set s = Nothing
                End With
                
            Case "SurveyFile"
            
            
            Case "TempPhoto"
                Dim tp As New TempPhoto
                With tp
                
                    .ID = frm!tbxID.Value
                    .EventID = frm!cbxEvent
                    .PhotographerID = frm!cbxPhotographer
                    .FileName = frm!tbxFilename
                    .Path = frm!tbxPhotoPath
                    .PhotoDate = frm!tbxPhotoDate
                    .PhotoType = frm!cbxPhotoType
                    
                    'set the generic object --> TempPhoto
                    Set obj = tp
                    
                    'cleanup
                    Set tp = Nothing
                
                End With
            Case "Task"
                Dim tk As New Task
                
                With tk
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                    .RequestDate = frm!tbxRequestDate.Value
                    .RequestedByID = frm!cbxRequestedBy.Column(0)
                    .Status = frm!cbxStatus.Column(0)
                    .Priority = frm!cbxPriority.Column(0)
                    .Task = frm!tbxTask.Value
                    .TaskType = frm.ContextType
                    
                    strCriteria = "[TaskType] = '" & .TaskType _
                                & "' AND [Task] = '" & .Task _
                                & "'"
                
                    'set the generic object --> Task
                    Set obj = tk
                    
                    'cleanup
                    Set tk = Nothing
                End With
            
            Case "Transducer"
                Dim t As New Transducer
        
                With t
                    'values passed into form
                    .EventID = 1
                            
                    'form values
                    .TransducerType = ""
                    .TransducerNumber = frm!cbxTransducer.SelText
                    .SerialNumber = frm!tbxSerialNo.Value
                    .IsSurveyed = frm!chkSurveyed.Value
                    .Timing = frm!cbxTiming.SelText
                    .ActionDate = Format(frm!tbxSampleDate.Value, "YYYY-mm-dd")
                    .ActionTime = Format(frm!tbxSampleTime.Value, "hh:mm.ss")
                    .RefToEyebolt = frm!tbxRefToEyebolt
                    .RefToWaterline = frm!tbxRefToWaterline
                    .EyeboltToWaterline = frm!tbxEyeboltToWaterline
                    .EyeboltToScribeline = frm!tbxEyeboltToScribeline
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    strCriteria = "[TransducerNumber] = " & .TransducerNumber _
                                & " AND [Timing] = '" & .Timing _
                                & "' AND [SerialNumber] = '" & .SerialNumber _
                                & "' AND [ActionDate] = " & .ActionDate
                
                    'set the generic object --> Transducer
                    Set obj = t
                    
                    'cleanup
                    Set t = Nothing
                End With
            
            Case "Transect"
                Dim vt As New VegTransect
                strTable = "VegTransect"
                
                With vt
                    'values passed into form
                    .Park = TempVars("ParkCode")
                    .LocationID = 1
                    .EventID = 1
                            
                    'form values
                    .TransectNumber = frm!tbxNumber.Value
                    .SampleDate = Format(frm!tbxSampleDate.Value, "YYYY-mm-dd")
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                    
                    strCriteria = "[TransectNumber] = " & .TransectNumber _
                                & "' AND [SampleDate] = " & .SampleDate
                
                    'set the generic object --> VegTransect
                    Set obj = vt
                    
                    'cleanup
                    Set vt = Nothing
                End With
            
            Case "UserRole"
                Dim u As New person
                    
                With u
                    'values passed into form
            '        .EventID = 1
                            
                    'form values
            '        .UserRoleType = ""
            '        .UserRoleNumber = cbxUserRole.SelText
            '        .SerialNumber = tbxSerialNo.value
            '        .IsSurveyed = chkSurveyed.value
            '        .Timing = cbxTiming.SelText
            '        .ActionDate = Format(tbxSampleDate.value, "YYYY-mm-dd")
            '        .ActionTime = Format(tbxSampleTime.value, "hh:mm.ss")
                    
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
                
                    'strCriteria = "[UserRoleNumber] = " & .UserRoleNumber
                
                    'set the generic object --> Location
                    Set obj = u
                    
                    'cleanup
                    Set u = Nothing
                End With

            Case "UnknownSpecies"
                Dim unk As New UnknownSpecies
                
                With unk
                
                    .ID = frm!tbxID.Value '0 if new, edit if > 0
            
                    .UnknownCode = frm!tbxUnknownCode
                    .PlantType = frm!optgPlantType
                    .PlantDescription = frm!tbxDescription
                    .SalientFeature = frm!tbxFeature
                    .LeafType = frm!tbxLeafType
                    .LeafMargin = frm!tbxLeafMargin
                    .LeafCharacter = frm!tbxLeafCharacter
                    .StemCharacter = frm!tbxStemCharacter
                    .FlowerCharacter = frm!tbxFlowerCharacter
                    .GeneralCharacter = frm!tbxGeneralCharacter
                    .ForbGrassType = frm!optgForbGrassType
                    .PerennialGrassType = frm!optgPerennialGrassType
                    .BestGuess = frm!tbxBestGuess
                    .HasPhotos = IIf(frm!chkHasPhotos = True, 1, 0)
                    .Collected = IIf(frm!chkCollected = True, 1, 0)
                    .CollectionMethod = frm!tbxCollectionMethod
                    .LocationID = frm!tbxLocationID
                    .CollectedByID = frm!cbxCollectedById
                    
                    'set the generic object --> unk
                    Set obj = unk
                    
                    'cleanup
                    Set unk = Nothing
                End With
            
            Case "VegPlot"
                Dim vp As New VegPlot
                With vp
                    'values passed into form
                    .SiteID = TempVars("SiteID")
                    .FeatureID = TempVars("FeatureID")
                    
                    'form values
                    .EventID = frm!cbxEvent
                    .VegTransectID = Ne(Nz(frm!cbxTransect, 0), 0)
                    .PlotNumber = Ne(frm!tbxPlotNumber, 0)
                    .PlotDistance = Ne(frm!tbxPlotDistance, 0)
                    .ModalSedimentSizeID = frm!cbxModalSedSize
                    
                    .PlotDensity = Ne(Nz(frm!tbxPlotDensity, 0), 0)
                    
                    'pct values
                    .WoodyCanopyPctCover = SetTrace(Ne(frm!tbxPctWCC, 0), 0.5)
                    .UnderstoryRootedPctCover = SetTrace(Ne(frm!tbxPctURC, 0), 0.5)
                    .AllRootedPctCover = SetTrace(Ne(frm!tbxPctARC, 0), 0.5)
                    .PctFines = SetTrace(frm!tbxPctFines, 0.5)
                    .PctWater = SetTrace(frm!tbxPctWater, 0.5)
                    .PctSocialTrails = SetTrace(frm!tbxPctSocialTrails, 0.5)
                    .PctLitter = SetTrace(frm!tbxPctLitter, 0.5)
                    .PctWoodyDebris = SetTrace(frm!tbxPctWoodyDebris, 0.5)
                    .PctStandingDead = SetTrace(frm!tbxPctStandingDead, 0.5)
                    .PctFilamentousAlgae = SetTrace(frm!tbxPctFA, 0.5)
                    
                    .PctModalSedimentSize = SetTrace(frm!tbxPctMSS, 0.5)
                    
                    'chk/tgl values -> change Access -1 (true) to clearer 1 for use in SQL
                    '                  so value of 1 = has no canopy veg, 0 = has canopy veg etc.
                    
                    '.HasSocialTrails = IIf(frm!tglHasSocialTrails = True, 1, 0)
                    .NoCanopyVeg = IIf(frm!tglNoCanopyVeg = True, 1, 0)
                    .NoRootedVeg = IIf(frm!tglNoRootedVeg = True, 1, 0)
                    '.NoIndicatorSpecies = IIf(frm!tglNoIndicatorSpecies = True, 1, 0) 'replaced w/ PctSocialTrails
                    .BeaverBrowse = IIf(frm!tglBeaverBrowse = True, 1, 0)
                    
                    .CalibrationPlot = IIf(frm!chkCalibrationPlot = True, 1, 0)
                    .ReplicatePlot = IIf(frm!chkReplicatePlot = True, 1, 0)
                    
                    'set the generic object --> VegPlot
                    Set obj = vp
                    
                    'cleanup
                    Set vp = Nothing
                End With
            
            Case "VegSpecies"
                Select Case frm.FormContext
                    Case "AllRootedSpecies"
                        Dim ars As New RootedSpecies
                        
                        With ars
                            'values passed into form
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                            
                            .VegPlotID = frm!tbxVegPlotID
                            .MasterPlantCode = frm!cbxSpecies
                            .PercentCover = frm!tbxPctCover
                            
                            'set the generic object --> Woody Canopy Species
                            Set obj = ars
                            
                            'cleanup
                            Set ars = Nothing
                        End With
                    
                    Case "UnderstoryRootedSpecies"
                        Dim ucs As New UnderstoryCoverSpecies
                        
                        With ucs
                            'values passed into form
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                            
                            .VegPlotID = frm!tbxVegPlotID
                            .MasterPlantCode = frm!cbxSpecies
                            .PercentCover = frm!tbxPctCover
                            .IsSeedling = IIf(frm!chkIsSeedling = True, 1, 0)
                            
                            'set the generic object --> Woody Canopy Species
                            Set obj = ucs
                            
                            'cleanup
                            Set ucs = Nothing
                        End With

                    Case "VegWalkSpecies"
                        Dim vws As New VegWalkSpecies
                        
                        With vws
                            'values passed into form
                    
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                            
                            .VegWalkID = ""
                            .MasterPlantCode = frm!cbxSpecies
                            .IsSeedling = IIf(frm!chkIsSeedling = True, 1, 0)
                            
                            'set the generic object --> Location
                            Set obj = vws
                            
                            'cleanup
                            Set vws = Nothing
                        End With
                    
                    Case "WoodyCanopySpecies"
                        Dim wcs As New WoodyCanopySpecies
                        
                        With wcs
                            'values passed into form
                            .ID = frm!tbxID.Value '0 if new, edit if > 0
                            
                            .VegPlotID = frm!tbxVegPlotID
                            .MasterPlantCode = frm!cbxSpecies
                            .PercentCover = frm!tbxPctCover
                            
                            'set the generic object --> Woody Canopy Species
                            Set obj = wcs
                            
                            'cleanup
                            Set wcs = Nothing
                        End With
                
                End Select

                Case "VegWalk"
                    Dim vw As New VegWalk
                    
                    With vw
                        'values passed into form
                    
                        .ID = frm!tbxID.Value '0 if new, edit if > 0
                        .EventID = frm!cbxEventID
                        .CollectionPlaceID = frm.CollectionPlaceID
                        .CollectionType = frm.CollectionType
                        .StartDate = frm!tbxWalkStartDate
                        'set the generic object --> Location
                        Set obj = vw
                        
                        'cleanup
                        Set vw = Nothing
                    End With
                    
            
            '-------------------
            ' --- UPLAND & INVASIVES ---
            '-------------------
            Case "Template"
                'Dim tpl As New Template
                Dim tpl As template
                
                With tpl
                    .IsSupported = 1
                    .Context = ""
                    .EffectiveDate = Date
                    .Remarks = ""
                    .TemplateName = ""
                    .Version = ""
                    .TemplateSQL = ""
                    .Syntax = ""
    
                End With
                
                'set the generic object --> Template
                Set obj = tpl
                
                'cleanup
                Set tpl = Nothing
                           
            Case "TemplateAdd"
                'Dim tpl As New Template
                
                With tpl
                    .TemplateName = frm!tbxTemplate
                    .Context = .TemplateName
                    .IsSupported = 1 '.IsSupported default = 1 (i.e. yes)
                    .Version = frm!tbxVersion
                    .Syntax = frm!cbxSyntax
                    .TemplateSQL = frm!tbxTemplateSQL
                    .EffectiveDate = frm!tbxEffectiveDate
                    '.Params handled when .TemplateSQL set
                    '.Params = GetParamsFromSQL(.TemplateSQL)
                    .Remarks = frm!tbxRemarks
                    .ContactID = TempVars("AppUserID")
                    
                    'set the generic object --> Transducer
                    Set obj = tpl
                    
                    'cleanup
                    Set tpl = Nothing
                End With
                
                'inserts only, no ID?
                NoList = True

            Case Else
                GoTo Exit_Handler
        End Select
                
        'set insert/update based on whether its an edit or new entry
        DoAction = IIf(frm!tbxID.Value > 0, "u", "i")
        
        If NoList Then
                    
            'form doesn't contain list subform or message/icon fields
            'so cut to the chase -> do nothing here
            
        Else
        
            'check if the record already exists by checking event list form records
            'event list form pulls active records for park, river segment
            Dim rs As DAO.Recordset
            
            Set rs = frm!list.Form.RecordsetClone
            rs.FindFirst strCriteria
            
            If rs.NoMatch Then
                ' --- INSERT ---
                frm!lblMsg.forecolor = lngLime
                frm!lblMsgIcon.forecolor = lngLime
                frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                frm!lblMsg.Caption = IIf(DoAction = "i", "Inserting new record...", "Updating record...")
            Else
                ' --- UPDATE ---
                'record already exists & ID > 0
                
                'retrieve ID
                If frm!tbxID.Value = rs("ID") Then 'rs("Contact.ID") Then
                    'IDs are equivalent, just change the data
                    frm!lblMsg.forecolor = lngLime
                    frm!lblMsgIcon.forecolor = lngLime
                    frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                    frm!lblMsg.Caption = "Updating record..."
                Else
                    'prevent duplicate record entries
                    frm!lblMsg.forecolor = lngYellow
                    frm!lblMsgIcon.forecolor = lngYellow
                    frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
                    frm!lblMsg.Caption = "Oops, record already exists."
                    GoTo Exit_Handler
                End If
                
            End If
        End If
        
        'T/F refers to whether the record is an update (T) or insert (F)
        obj.SaveToDb IIf(DoAction = "i", False, True)
        
        'add the action record --> DONE via SaveToDb (thru SetRecord)
        
        'set the tbxID.value ==> tbxID is a bound control, can't set it this way
        'tbxID = .ID
        'frm!tbxID.Value = obj.ID
        'frm.Controls("tbxID").Value = obj.ID
    End With
    
    'clear values & refresh display
    frm.ReadyForSave 'Application defined error? --> ensure ReadyForSave is Public Sub
    'Forms!frm.ReadyForSave
    
    'handle situations where Access is saving same record
    
    'save record changes from form first to avoid "Write Conflict" errors
    'where form & SQL are attempting to save record
    'frm.Dirty = False
    
'    If frm.Dirty Then
    If frm.Dirty And Not NoList Then
        Debug.Print "UpsertRecord " & frm.Name & " DIRTY"
        'frm.Dirty = False
        
        frm!lblMsg.forecolor = lngYellow
        frm!lblMsgIcon.forecolor = lngYellow
        frm!lblMsgIcon.Caption = StringFromCodepoint(uDoubleTriangleBlkR)
        frm!lblMsg.Caption = "** DIRTY **" 'UNSAVED CHANGES! **"
        
    Else
        Debug.Print "UpsertRecord " & frm.Name & " CLEAN"
    End If
        
' CHECK IF POPULATING FORM IS THE ISSUE...
'    PopulateForm frm, frm!tbxID.Value
    
'    'refresh list
'    frm!list.Requery
    
    frm.Requery
    
    'handle list forms - update messages, icon & refresh
    If Not NoList Then
        'clear messages & icon
        frm!lblMsgIcon.Caption = ""
        frm!lblMsg.Caption = ""
        
        'refresh list
        frm!list.Requery
    End If
    
    'exit
    GoTo Exit_Handler
    
Form_Without_List:
    DoAction = "i"
    Resume Next

Exit_Handler:
    'cleanup
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpsertRecord[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          SetObserverRecorder
' Description:  Sets data observer & recorder
' Assumptions:  -
' Parameters:   obj - object to set observer/recorder on (object)
'               tbl - name of table being modified (string)
'               activity - observer (O) or recorder (R) (string)
'               person - ID of individual (long)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 9, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 8/9/2016 - initial version
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'                   - un-comment out
' --------------------------------------------------------------------
'   BLC - 12/5/2017 - fix to RefID vs ID, handle either observer or recorder
' ---------------------------------
Public Sub SetObserverRecorder(obj As Object, tbl As String, _
                                activity As String, person As Long)
On Error GoTo Err_Handler

    'handle record actions
    Dim act As New RecordAction
    With act

'    'Recorder
'        .RefAction = "R"
'        .ContactID = obj.RecorderID
'        .RefID = obj.ID
'        .RefTable = tbl
'        .SaveToDb
'
'    'Observer
'        .RefAction = "O"
'        .ContactID = obj.ObserverID
'        .RefID = obj.ID
'        .RefTable = tbl
'        .SaveToDb

        'Observer or Recorder
        .RefAction = activity
        .ContactID = person
        .RefID = obj.RefID
        .RefTable = tbl
        .SaveToDb

    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetObserverRecorder[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          CollapseRows
' Description:  Collapses TCount, PctCover, SE for one species/IsDead into one row
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
'   R.Hicks, Sept 15, 2002
'   http://www.utteraccess.com/forum/copy-table-structure-vb-t117555.html
' Source/date:  Bonnie Campbell, June 22 2017
' Adapted:      -
' Revisions:    BLC - 6/22/2017 - initial version
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'                   - shifted from frm_Species_Cover_by_Route
' --------------------------------------------------------------------
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Sub CollapseRows(tbl As String) 'tdf As DAO.TableDef)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim col As field
    Dim rs As DAO.Recordset
    Dim rsPctCover As DAO.Recordset
    Dim rsSE As DAO.Recordset
    Dim strTableNew As String
    Dim strCol As String
    Dim Park As String, VisitYear As String, Species As String, CommonName As String, _
        IsDead As String, Route As String
    Dim PrevPark As String, PrevVisitYear As String, PrevSpecies As String, _
        PrevCommonName As String, PrevIsDead As String, PrevRoute As String
    Dim TCount As Integer
    Dim PctCover As String
    Dim SE As Double
    Dim Concat As String, PrevConcat As String
    
    strTableNew = tbl & "_NEW"
    
    PrevPark = ""
    PrevVisitYear = ""
    PrevSpecies = ""
    PrevCommonName = ""
    PrevIsDead = ""
    
    Set db = CurrDb
    Set tdf = db.TableDefs(tbl)
    
    Set rs = db.OpenRecordset(tdf.Name)
    
    'remove table if it exists
    If TableExists(strTableNew) Then DoCmd.DeleteObject acTable, strTableNew
    
    'create empty table w/ same columns to fill into
    DoCmd.TransferDatabase acExport, "Microsoft Access", db.Name, acTable, tdf.Name, strTableNew, True
    
    With tdf
        'iterate through ALL
        Do Until rs.BOF And rs.EOF
        
            'iterate result columns (fields)
            For Each col In tdf.Fields
            
                'get park, visit year, species, common name & isdead
                If col.OrdinalPosition = 1 Then Park = rs!Unit_Code
                If col.OrdinalPosition = 2 Then VisitYear = rs!Visit_Year
                If col.OrdinalPosition = 3 Then Species = rs!Species
                If col.OrdinalPosition = 4 Then CommonName = rs!Master_Common_Name
                If col.OrdinalPosition = 5 Then IsDead = rs!IsDead
            
                'ignore 1-5 (static Park, Year, Species, Master Common Name, IsDead)
                If col.OrdinalPosition > 5 Then
                
                    'get column & route name
                    strCol = col.Name
                    Route = Left(strCol, InStr(col.Name, ") ") + 1)
                                     
                    Select Case Replace(strCol, Route, "")
                        Case "TCount"
                            TCount = col.Value
                        Case "AvgCover"
                            PctCover = col.Value
                        Case "SE"
                            SE = col.Value
                    End Select

                    'concatenate for comparison
                    Concat = Park & VisitYear & Species & CommonName & IsDead & Route
                    
                    If Concat <> PrevConcat Then
                
                    'add these to the new table if they don't already exist
                    End If
    
                    If PrevSpecies = rs!Species And PrevIsDead = rs!IsDead Then
                    
                    End If
                End If
                
                'capture the previous values
                PrevConcat = Concat
                PrevPark = Park
                PrevVisitYear = VisitYear
                PrevSpecies = Species
                PrevCommonName = CommonName
                PrevIsDead = IsDead
                
           Next
           
        Loop
    
    End With

Exit_Procedure:
    Set tdf = Nothing
    Set rs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CollapseRows[mod_App_Data])"
    End Select
    Resume Exit_Procedure
End Sub

' =================================
'   SOP Methods
' =================================
' ---------------------------------
' FUNCTION:     GetProtocolVersion
' Description:  Retrieve protocol version, effective & retire dates
' Assumptions:  Assumes only one version of the protocol is active at once
' Parameters:   blnAllVersions - indicator if all versions should be retrieved (boolean)
' Returns:      Protocol name, version, effective & retire dates, last modified date
' Note:         To retrieve values, data must be retrieved from the array:
'                   ary(0,0) = ProtocolName
'                   ary(1,0) = Version
'                   ary(2,0) = EffectiveDate
'                   ary(3,0) = RetireDate
'                   ary(4,0) = LastModified
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function GetProtocolVersion(Optional blnAllVersions As Boolean = False) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String, strWHERE As String
    Dim Count As Integer
    Dim metadata() As Variant
   
    'handle only appropriate park codes
    If blnAllVersions Then
        strWHERE = ""
    Else
        strWHERE = "WHERE RetireDate IS NULL"
    End If
    
    'generate SQL
'    strSQL = "SELECT ProtocolName, Version, EffectiveDate, RetireDate, LastModified FROM Protocol " _
'                & strWHERE & ";"
    strSQL = GetTemplate("s_protocol_info", "strWHERE" & PARAM_SEPARATOR & strWHERE)
    
    'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)
        
    If rs.BOF And rs.EOF Then GoTo Exit_Handler
        
    With rs
        .MoveLast
        .MoveFirst
        Count = .RecordCount
    
        metadata = rs.GetRows(Count)
 
        .Close
    End With
    
    'return value
    GetProtocolVersion = metadata
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetProtocolVersion[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetSOPMetadata
' Description:  Retrieve SOP metadata (abbreviation code, #, version, effective date)
' Assumptions:  Assumes only one active/effective SOP # for a given area
' Parameters:   area - area covered by the SOP (string)
' Returns:      SOP metadata - Code, SOP #, Version, EffectiveDate
' Note:         To retrieve value, data must be retrieved from the array:
'                   ary(0,0) = SOP #
'               Assuming there is only one matching SOP for each area
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
'   BLC - 5/11/2016 - revised to getting full SOP metadata vs. number only
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function GetSOPMetadata(area As String) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
        
    'generate SQL
    '---------------------------------------------------------------------
    ' NOTE: use * vs % for the LIKE wildcard
    '       if it is not used strSQL will work in a query directly,
    '       but will fail to return records via a VBA recordset
    '       So    "...LIKE '" & LCase(area) & "*';"   works
    '       But   "...LIKE '" & LCase(area) & "%';"   does not (except in direct Query SQL)
    '
    ' c.f.  Hans Up, May 17, 2011 & discussion
    '       http://stackoverflow.com/questions/6037290/use-of-like-works-in-ms-access-but-not-vba
    '---------------------------------------------------------------------
    strSQL = GetTemplate("s_sop_metadata", "area" & PARAM_SEPARATOR & LCase(area))
    
    'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL, dbOpenDynaset)
        
    'return value
    Set GetSOPMetadata = rs
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSOPNum[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' =================================
'   Discrete Data Methods
' =================================
' ---------------------------------
' FUNCTION:     GetParkID
' Description:  Retrieve the ID associated with a park
' Assumptions:  -
' Parameters:   ParkCode - 4 character park designator (string)
' Returns:      ID - unique park identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 1/12/2017  - revised to use GetRecords() vs. GetTemplate()
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function GetParkID(ParkCode As String) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
'    strSQL = GetTemplate("s_park_id", "ParkCode" & PARAM_SEPARATOR & ParkCode)
            
    'fetch data
'    Set db = CurrDb
    Set rs = GetRecords("s_park_id") 'db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    
    If Not (rs.BOF And rs.EOF) Then
        ID = rs.Fields("ID")
    End If
    
    rs.Close
    
    'return value
    GetParkID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParkID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetParkState
' Description:  Retrieve the state associated with a park (via tlu_Parks)
' Assumptions:  Park state is properly identified in tlu_Parks
' Parameters:   parkCode - 4 character park designator
' Returns:      ParkState - 2 character state abbreviation
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
'   BLC - 6/28/2016  - revised to uppercase GetParkState vs getParkState
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function getParkState(ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim state As String, strSQL As String
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 ParkState FROM tlu_Parks WHERE ParkCode LIKE '" & ParkCode & "';"
            
    'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        state = rs.Fields("ParkState").Value
    End If
   
    'return value
    getParkState = state
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParkState[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetRiverSegments
' Description:  Retrieve the river segments associated with a park
' Assumptions:  River segments are properly associate w/ park
' Parameters:   ParkCode - 4 character park designator
' Returns:      segments - river segments (Green, CAC, GBC, Yampa, CBC, GBC, etc.)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 5, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/5/2016  - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function GetRiverSegments(ParkCode As String) As Variant
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim Count As Integer
    Dim segments() As Variant
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
    strSQL = GetTemplate("s_get_river_segments", "ParkCode" & PARAM_SEPARATOR & ParkCode)

            
    'fetch data
    Set db = CurrDb
    Set rs = db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    Count = rs.RecordCount
    
    'retrieve 2D array of records
    'segments(intField, intRecord) --> segments(0,1) = 2nd record, 1st field
    segments = rs.GetRows(Count)
 
    rs.Close
    
    'return value
    GetRiverSegments = segments
    
Exit_Handler:
    Set rs = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRiverSegments[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetRiverSegmentID
' Description:  Retrieve the ID associated with a River
' Assumptions:  -
' Parameters:   segment - river segment designator (string)
' Returns:      ID - unique river identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 1/17/2017  - revise to use GetRecords() vs. GetTemplate()
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function GetRiverSegmentID(segment As String) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate River codes
    If Len(segment) < 1 Then
        GoTo Exit_Handler
    End If
    
    'generate SQL
'    strSQL = GetTemplate("s_river_segment_id", "waterway" & PARAM_SEPARATOR & segment)
            
    'fetch data
'    Set db = CurrDb
    Set rs = GetRecords("s_river_segment_id") 'db.OpenRecordset(strSQL)

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    
    If Not (rs.BOF And rs.EOF) Then
        ID = rs.Fields("ID")
    End If
    
    rs.Close
    
    'return value
    GetRiverSegmentID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRiverSegmentID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetSiteID
' Description:  Retrieve the ID associated with a site
' Assumptions:  -
' Parameters:   ParkCode - park designator (4-character string)
'               SiteCode - site designator (2-character string)
' Returns:      ID - unique site identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
'   BLC - 11/6/2017 - use GetRecords() vs direct SQL
' ---------------------------------
Public Function GetSiteID(ParkCode As String, SiteCode As String) As Long
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    Dim ID As Long
   
    'handle only appropriate River codes
    If Len(ParkCode) <> 4 Or Len(SiteCode) <> 2 Then
        GoTo Exit_Handler
    End If
    
    'fetch data
    Set rs = GetRecords("s_site_id_by_code")

    If rs.BOF And rs.EOF Then GoTo Exit_Handler

    rs.MoveLast
    rs.MoveFirst
    
    If Not (rs.BOF And rs.EOF) Then
        ID = rs.Fields("ID")
    End If
    
    rs.Close
    
    'return value
    GetSiteID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSiteID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetFeatureID
' Description:  Retrieve the ID associated with a feature
' Assumptions:  -
' Parameters:   ParkCode - park designator (4-character string)
'               Feature - feature designator (2-character string)
' Returns:      ID - unique feature identifier (long)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016  - initial version
'   BLC - 10/4/2016  - update to use parameter query
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
'   BLC - 12/5/2017 - revise to use GetRecords()
' ---------------------------------
Public Function GetFeatureID(ParkCode As String, Feature As String) As Long
On Error GoTo Err_Handler
    
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim strSQL As String
    Dim ID As Long
   
    'handle only appropriate River codes
    If Len(ParkCode) <> 4 Or Len(Feature) < 1 Then
        GoTo Exit_Handler
    End If
    
    Dim Params(0 To 2) As Variant
    Params(0) = ParkCode
    Params(1) = Feature
    
    ID = GetRecords("s_feature_id", Params)(0)
    
'    'generate SQL
'    strSQL = GetTemplate("s_feature_id", _
'            "ParkCode" & PARAM_SEPARATOR & ParkCode & _
'            "|feature" & PARAM_SEPARATOR & Feature)
'
'    'fetch data
'    Set db = CurrDb
'    Set rs = db.OpenRecordset(strSQL)
'
'    If rs.BOF And rs.EOF Then GoTo Exit_Handler
'
'    rs.MoveLast
'    rs.MoveFirst
'
'    If Not rs.BOF And rs.EOF Then
'        ID = rs.GetRows(1)
'    End If
'
'    rs.Close
    
    'return value
    GetFeatureID = ID
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetFeatureID[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          GetSurfaceIDs
' Description:  Sets a collection of surface IDs where IDs can be retrieved
'               from column names
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/24/2016 - initial version
' ---------------------------------
Public Function GetSurfaceIDs() As Scripting.Dictionary
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim strKey As String, strItem As String

    'prepare dictionary
    Dim dict As Scripting.Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    'retrieve surfaces & surface IDs
    Set rs = GetRecords("s_surface")
    
    If Not (rs.BOF And rs.EOF) Then
        Do Until rs.EOF
            
            With dict
            
                strKey = rs("ColName")
                strItem = rs("ID")
                
                If Not .Exists(strKey) Then
                    'add the ColName (key) & ID (value)
                    '--------------------------------------
                    ' NOTE:
                    '   Cannot use notation w/ rs("fieldname")
                    '   notation because as soon as you leave
                    '   the Do Until the dictionary forgets
                    '   the values since rs("fieldname") is
                    '   out of scope.
                    '   -> Error 3420: Object invalid or no longer set
                    '   Use:
                    '       .Add strKey, strItem
                    '   Not:
                    '       .Add strKey, rs("ID")
                    '--------------------------------------
                    .Add strKey, strItem
                End If
            
            End With
            
            'Debug.Print strKey & ": " & strItem & " = " & dict(strKey)
            
            rs.MoveNext
        Loop
    End If
    
    'set global dictionary
    Set g_AppSurfaces = dict
    
    'return dictionary
    Set GetSurfaceIDs = dict
    
Exit_Handler:
    'Set dict = Nothing
    Set rs = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetSurfaceIDs[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          GetModSedimentSizeIDs
' Description:  Sets a collection of modal sediment size IDs where IDs can be retrieved
'               from class names based on the year passed in
' Assumptions:  Assumes the year is the current sampling year
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 3, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 11/3/2017 - initial version
' ---------------------------------
Public Function GetModSedimentSizeIDs() As Scripting.Dictionary
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset
    Dim strKey As String, strItem As String

    'prepare dictionary
    Dim dict As Scripting.Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    'retrieve surfaces & surface IDs
    Set rs = GetRecords("s_modal_sediment_size")
    
    If Not (rs.BOF And rs.EOF) Then
        Do Until rs.EOF
            
            With dict
            
                strKey = rs("Label")
                strItem = rs("ID")
                
                If Not .Exists(strKey) Then
                    'add the ColName (key) & ID (value)
                    '--------------------------------------
                    ' NOTE:
                    '   Cannot use notation w/ rs("fieldname")
                    '   notation because as soon as you leave
                    '   the Do Until the dictionary forgets
                    '   the values since rs("fieldname") is
                    '   out of scope.
                    '   -> Error 3420: Object invalid or no longer set
                    '   Use:
                    '       .Add strKey, strItem
                    '   Not:
                    '       .Add strKey, rs("ID")
                    '--------------------------------------
                    .Add strKey, strItem
                End If
            
            End With
            
            'Debug.Print strKey & ": " & strItem & " = " & dict(strKey)
            
            rs.MoveNext
        Loop
    End If
    
    'set global dictionary
    Set g_ModalSedimentSizeIDs = dict
    
    'return dictionary
    Set GetModSedimentSizeIDs = dict
    
Exit_Handler:
    'Set dict = Nothing
    Set rs = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetModSedimentSizeIDs[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          GetQuadratPositions
' Description:  Sets a collection of quadrat positions where positions can be retrieved
'               from quadrat control names
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/24/2016 - initial version
' ---------------------------------
Public Function GetQuadratPositions() As Scripting.Dictionary
On Error GoTo Err_Handler

    Dim ctrl As Variant 'control name
    Dim strKey As String, strItem As String

    'prepare dictionary
    Dim dict As Scripting.Dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim aryControls() As String
    
    'prepare positions
    aryControls = Split("Q1,Q2,Q3,Q1_3m,Q2_8m,Q3_13m,Q1_hm,Q2_5m,Q3_10m", ",")
    
    For Each ctrl In aryControls
        
        With dict
        
            strKey = ctrl
            
            Select Case ctrl
                Case "Q1", "Q2", "Q3"
                    strItem = vbNullString 'position NULL
                Case "Q1_3m"
                    strItem = 3
                Case "Q1_hm"
                    strItem = 0
                Case "Q2_8m", "Q2_5m"
                    strItem = Replace(Right(ctrl, 2), "m", "")
                Case "Q3_13m", "Q3_10m"
                    strItem = Replace(Right(ctrl, 3), "m", "")
            End Select
            
            If Not .Exists(strKey) Then
                    'add the ctrl name (key) & position (value)
                    .Add strKey, strItem
            End If
            
            Debug.Print strKey & ": " & strItem & " = " & dict(strKey)
            
        End With
    
    Next
    
    'set global dictionary
    Set g_AppQuadratPositions = dict
    
    'return dictionary
    Set GetQuadratPositions = dict
    
Exit_Handler:
    'Set dict = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetQuadratPositions[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' =================================
'   PlotCheck Methods
' =================================
' ---------------------------------
' Sub:          RunPlotCheck
' Description:  Run plot check queries
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 27, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/27/2017 - initial version
'   BLC - 3/29/2017 - adjusted to accommodate FieldOK (pass/fail/unknown) values
'   BLC - 3/30/2017 - handle dependencies (queries dependent on queries)
' ---------------------------------
Public Function RunPlotCheck()
On Error GoTo Err_Handler

    Dim strTemplate As String
    Dim X As Variant

    'clear num records
    ClearTable "NumRecords"
    
    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates
        
    'use g_AppTemplates scripting dictionary vs. recordset to avoid missing dependencies
    'iterate through queries
 '   For i = 0 To g_AppTemplates.Count - 2
    For Each X In g_AppTemplates
    
        With g_AppTemplates.Item(X) 'g_AppTemplates.Items()(i)
            strTemplate = .Item("TemplateName")
            
            Debug.Print strTemplate
            
            If Len(.Item("FieldOK")) > 0 And .Item("FieldCheck") Then _
                SetPlotCheckResult strTemplate, "insert"
'            iTemplate = .Item("ID")
'            strDeps = .Item("Dependencies")
'            strFieldOK = .Item("FieldOK")
'            blnFieldCheck = .Item("FieldCheck")
        End With
        
    Next
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RunPlotCheck[mod_App_Data form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          SetPlotCheckResult
' Description:  Run plot check queries & determine results
' Assumptions:  -
' Parameters:   strTemplate - template name (string)
'               action - insert or update (string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 30, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/30/2017 - initial version
'   BLC - 3/29/2017 - adjusted to accommodate FieldOK (pass/fail/unknown) values
'   BLC - 3/30/2017 - handle dependencies (queries dependent on queries)
'                     only queries used for field checks are checked
' --------------------------------------------------------------------
'   BLC - 9/7/2017  - merge uplands, invasives, big rivers dbs modifications
' --------------------------------------------------------------------
'       BLC - 8/14/2017 - add error handling to address error 3048
' --------------------------------------------------------------------
' ---------------------------------
Public Function SetPlotCheckResult(strTemplate As String, action As String)
On Error GoTo Err_Handler

    Dim rs As DAO.Recordset, rs2 As DAO.Recordset
    Dim strDeps As String, strFieldOK As String, _
        strOperator As String, strField As String, CompareTo As String
    Dim iTemplate As Long
    Dim i As Integer, iOK As Integer
    Dim blnFieldCheck As Boolean, isOK As Boolean
    
    'initialize AppTemplates if not populated
    If g_AppTemplates Is Nothing Then GetTemplates
        
'Debug.Print "SetPlotCheckResult"
'Debug.Print strTemplate
        
    With g_AppTemplates(strTemplate)
        iTemplate = .Item("ID")
        strDeps = .Item("Dependencies")
        strFieldOK = .Item("FieldOK")
        blnFieldCheck = .Item("FieldCheck")
    End With
        
    'handle dependencies first
    'Dependencies = comma separated list of queries template is dependent on
    If Len(strDeps) > 0 Then _
        HandleDependentQueries strDeps, "run"
    
    'run query & retrieve record #s
    Set rs = GetRecords(strTemplate)
        
    'catch missing recordsets due to Error 3048: Cannot open any more databases.
    If rs Is Nothing Then GoTo Exit_Handler
        
    'identify proper count
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveLast
        rs.MoveFirst
        Debug.Print "# records = " & rs.RecordCount
    End If
        
    'default
    isOK = 0
        
    'add values to numrecords
    Dim Params(0 To 3) As Variant
    
    Params(0) = LCase(Left(action, 1)) & "_num_records"
    Params(1) = iTemplate
    Params(2) = rs.RecordCount
    
    If Len(strFieldOK) > 0 Then
        'assess if field check is fulfilled
        
        'determine comparitor
        iOK = CInt(Right(strFieldOK, 1))
        
        'fetch the operator
        strOperator = Left(Right(strFieldOK, Len(strFieldOK) - InStr(strFieldOK, "]")), 1)
        
        'fetch the field/item to check
        strField = Replace(Left(strFieldOK, InStr(strFieldOK, "]") - 1), "[", "")
        
        Select Case strField
            Case "NumRecords"
                CompareTo = rs.RecordCount
            Case Else
                CompareTo = strField
        End Select
    
        Select Case strOperator
            Case "="
                isOK = IIf(CompareTo = iOK, 1, 0)
            Case "<"
                isOK = IIf(CompareTo < iOK, 1, 0)
            Case ">"
                isOK = IIf(CompareTo > iOK, 1, 0)
        End Select
    
    End If
    
    Params(3) = IIf(isOK = True, 1, 0) 'convert to 1/0 as true/false instead of 0/-1
    
    'clear original value
    DeleteRecord "NumRecords", iTemplate, False
    
    SetRecord "i_num_records", Params
    
    Debug.Print Params(1) & " " & strTemplate & " " & Params(2) & " " & Params(3)
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetPlotCheckResult[mod_App_Data form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          UpdateNumRecords
' Description:  Update NumRecords # of records
' Assumptions:  -
' Parameters:   iRecord - template ID (integer)
'               numRecords - # of records (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 30, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 3/30/2017 - initial version
' ---------------------------------
Public Function UpdateNumRecords(iRecord As Integer, NumRecords As Integer)
On Error GoTo Err_Handler

    'add values to numrecords
    Dim Params(0 To 3) As Variant

Debug.Print "UpdateNumRecords"
    
    Params(0) = "u_num_records"
    Params(1) = iRecord
    Params(2) = NumRecords
            
    SetRecord "u_num_records", Params
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateNumRecords[mod_App_Data form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          PrepareSpeciesQuery
' Description:  Craft the species query
' Assumptions:  -
' Parameters:   Park - park code (string)
'               SampleDate - sampling visit year (integer)
'               PlotID - plot # (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 3, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/3/2017 - initial version (consider for future use)
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function PrepareSpeciesQuery(Park As String, SampleYear As Integer, PlotID As Integer)
On Error GoTo Err_Handler
    'set statusbar notice
    SysCmd acSysCmdSetStatus, "Running report ..."
    
    Screen.MousePointer = 11 'Hour Glass

    Dim strFilter As String, strWHERE As String, strParkWhere As String, strPlotWhere As String, strYrWhere As String, strSpeciesYear As String
    Dim stDocName As String

    'defaults
    strFilter = ""
    strWHERE = ""
    strParkWhere = ""
    strPlotWhere = ""
    strYrWhere = ""

    stDocName = "rpt_Species_by_Park"
    
    ' Set where condition if needed
    If (IsNull(Park) + IsNull(SampleYear) + IsNull(PlotID)) > -3 Then
      
      'park
      If Not IsNull(Park) Then
        strParkWhere = "Unit_Code = '" & Park & "'"
        strFilter = Park
      End If
      
      'plot --> NOTE: assumes UI will not allow plot selection w/o park
      If Not IsNull(PlotID) Then
        strPlotWhere = "Plot_ID = " & PlotID
        strFilter = strFilter & "- plot #" & PlotID
      End If
      
      'year
      If Not IsNull(SampleYear) Then
'        strYrWhere = "Len(SpeciesYear) > Len(Replace(SpeciesYear, CStr(" & Me!Visit_Date & "), ''))"
        '(qry_Sp_Rpt_All.Utah_Species+"-"+CStr(qry_Sp_Rpt_All.Year)) AS SpeciesYear
        strSpeciesYear = "(qry_Sp_Rpt_All.Utah_Species+' - '+CStr(qry_Sp_Rpt_All.Year))"
        strYrWhere = "Len(" & strSpeciesYear & ") > Len(Replace(" & strSpeciesYear & ", CStr(" & SampleYear & "), ''))"
        
        'set filter display
        Select Case Len(strFilter)
            Case 0 'year only
                strFilter = CStr(SampleYear)
            Case 4 'park only
                strFilter = strFilter & "-" & CStr(SampleYear)
            Case Is > 4 'park & plot
                strFilter = Replace(strFilter, "-", "-" & CStr(SampleYear) & " ")
        End Select
      
      Else
        'clear extra "-" for park & plot filter
        strFilter = Replace(strFilter, "-", "")
      End If
      
      'prepare where using string array & PrepareWhereClause
      Dim ary() As String
      ary = Split(strParkWhere & ";" & strPlotWhere & ";" & strYrWhere, ";")
      strWHERE = PrepareWhereClause(ary)
      
'      If Not IsNull(Me!Park_Code) Then
'        strWhere = "Unit_Code = '" & Me!Park_Code & "'"
'        If Not IsNull(Me!Plot) Then
'          strWhere = strWhere & " And Plot_ID = " & Me!Plot
'        End If
'        If Not IsNull(Me!Visit_Date) Then
'          'strWhere = strWhere & " AND Visit_Year = " & Me!Visit_Date
'          'strWhere = strWhere & " AND " & Me!Visit_Date & " IN (replace(SpeciesYears, '|', ','))"
'          'strWhere = strWhere & " AND " & Me!Visit_Date & " LIKE SpeciesYears"
'          strWhere = strWhere & " AND Len(SpeciesYear) > Len(Replace(SpeciesYear, CStr(" & Me!Visit_Date & "), ''))"
'        End If
'      Else
'        'strWhere = "Visit_Year = " & Me!Visit_Date
'        'strWhere = Me!Visit_Date & " IN (replace(SpeciesYears, '|', ','))"
'        'WHERE Len(SpeciesYears) > Len(Replace(SpeciesYears, CStr(2014), ''));
'        strWhere = "Len(SpeciesYear) > Len(Replace(SpeciesYear, CStr(" & Me!Visit_Date & "), ''))"
'      End If
    End If
    
    'retrieve querydef
    Dim qdf As QueryDef
    Dim strSQL As String
    
    Set qdf = CurrDb.QueryDefs("qry_Sp_Rpt_by_Park_Complete_Create_Table")
    strSQL = qdf.SQL

'SELECT DISTINCT
'qry_Sp_Rpt_All.Unit_Code,
'qry_Sp_Rpt_All.Year,
'qry_Sp_Rpt_All.Plot_ID,
'qry_Sp_Rpt_All.Master_Family,
'qry_Sp_Rpt_All.Utah_Species,
'(qry_Sp_Rpt_All.Utah_Species+"-"+CStr(qry_Sp_Rpt_All.Year)) AS SpeciesYear,
'(qry_Sp_Rpt_All.Unit_Code+"-"+CStr(qry_Sp_Rpt_All.Plot_ID)+"-"+CStr(qry_Sp_Rpt_All.Utah_Species)) AS ParkPlotSpecies,
'(qry_Sp_Rpt_All.Unit_Code+"-"+CStr(qry_Sp_Rpt_All.Utah_Species)) AS ParkSpecies,
'(qry_Sp_Rpt_All.Unit_Code+"-"+CStr(qry_Sp_Rpt_All.Plot_ID)) AS ParkPlot INTO temp_Sp_Rpt_by_Park_Complete
'FROM qry_Sp_Rpt_All
'WHERE Len(SpeciesYears) > Len(Replace(SpeciesYears, CStr(2014), ''))
'ORDER BY qry_Sp_Rpt_All.Unit_Code, qry_Sp_Rpt_All.Plot_ID, qry_Sp_Rpt_All.Master_Family, qry_Sp_Rpt_All.Utah_Species;

    'update the SQL if parameters exist
    If Len(strWHERE) > 0 Then
        Dim iOrderBy As Integer
        Dim strSQLNew As String
        
        'replace ORDER with WHERE clause + ORDER
        strSQLNew = Replace(strSQL, "ORDER", " WHERE " & strWHERE & " ORDER")
        qdf.SQL = strSQLNew 'was strSQL
    End If
    
    'update underlying table (temp_Sp_Rpt_by_Park_Complete is used in report's underlying table temp_Sp_Rpt_by_Park_Rollup)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qry_Sp_Rpt_by_Park_Complete_Create_Table", acViewNormal
    
    'update status bar
    SysCmd acSysCmdSetStatus, "Generating complete results..."
    'DoEvents
    'Application.Echo False, "Generating complete results..."
    'Application.Echo True, ""
    
    'add an index to improve report performance
    Dim strIdxSQL As String
    
    strIdxSQL = "CREATE INDEX idxParkPlotSpeciesYear ON temp_Sp_Rpt_by_Park_Complete (ParkPlotSpecies, Year)"
    CurrDb.Execute strIdxSQL
    
    DoCmd.SetWarnings True
    
    'reset qdf SQL
    qdf.SQL = strSQL
    
    'update underlying table (temp_Sp_Rpt_by_Park_Rollup)
    DoCmd.SetWarnings False
    DoCmd.OpenQuery "qry_Sp_Rpt_by_Park_Rollup_Create_Table", acViewNormal
    
    'update status bar
    SysCmd acSysCmdSetStatus, "Generating rollup..."
    'DoEvents
    'Application.Echo False, "Generating rollup..."
    'Application.Echo True, ""
    
    'add an index to improve report performance
    strIdxSQL = "CREATE INDEX idxParkPlotSpeciesYears ON temp_Sp_Rpt_by_Park_Rollup (ParkPlotSpecies, SpeciesYears)"
    CurrDb.Execute strIdxSQL
    
Debug.Print strSQL

    DoCmd.SetWarnings True
    
    'update status bar
    SysCmd acSysCmdSetStatus, "Preparing report..."
    'DoEvents
    'Application.Echo False, "Preparing report..."
    'Application.Echo True, ""
    
    'translate SQL Where for rollup --> SpeciesYear = SpeciesYears, ,qry_Sp_Rpt_All.Year = SpeciesYears, qry_Sp_Rpt_All.Utah_species = "Utah.species"
    Dim aryText() As String
    aryText = Split("SpeciesYear|SpeciesYears||qry_Sp_Rpt_All.Year|SpeciesYears||qry_Sp_Rpt_All.Utah_species|Utah_species", "||")
    strWHERE = ReplaceMulti(strWHERE, aryText)
    'strWhere = Replace(strWhere, Replace(strSpeciesYear, "SpeciesYear", "SpeciesYears"), "SpeciesYears")
    
    'open report --> strWhere = WHERE clause filter, strFilter = display for filter if present
    DoCmd.OpenReport stDocName, acViewPreview, , strWHERE, acWindowNormal, strFilter
    
    SysCmd acSysCmdSetStatus, "Report complete."
    
    Screen.MousePointer = 1 'Standard Cursor
    'clear status bar
    SysCmd acSysCmdSetStatus, " "

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PrepareSpeciesQuery[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' =================================
'   Update Methods
' =================================
' ---------------------------------
' Function:     UpdateTransect
' Description:  Updates transect record values (Invasives)
'               and returns the submitted value (single)
' Assumptions:  Controls of transect visit form trigger the
'               function using:
'                   =UpdateTransect()
'               in their change event properties
' Parameters:   -
' Returns:      if successful - submitted cover value (single)
'               or 0 if not
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 13, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/13/2016 - initial version
' ---------------------------------
Public Function UpdateTransect() As Single
On Error GoTo Err_Handler

    Dim transectID As String
    Dim ObserverID As String
    Dim Comments As String
    Dim StartTime As Variant
    
    Dim vt As New VegTransect
    
    With Forms("frm_Data_Entry").Controls("frm_Quadrat_Transect").Form
        
        'set transect values
        transectID = .Controls("tbxTransectID")     ' Quadrat-Transect ID
        ObserverID = .Controls("cbxObserver")
        Comments = Nz(.Controls("tbxComments"), "")
        'If Not IsNull(.Controls("tbxStartTime")) Then StartTime = .Controls("tbxStartTime")
        StartTime = .Controls("tbxStartTime")
        
        With vt
            
            .TransectQuadratID = transectID
            .Observer = ObserverID
            .Comments = Comments
            
            '.UpdateTransectData
            
            'update start time if it is set
            If Not IsNull(StartTime) Then
                
                .StartTime = StartTime
                
                .UpdateStartTime
                
            End If
            
        End With
        
    End With
       
    'skip if NULL
'    If IsNull(TempVars("Transect_ID")) Then GoTo Exit_Handler
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateTransect[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          UpdateMicrohabitat
' Description:  Updates microhabitat surface cover values (Invasives)
'               and returns the submitted value (single)
' Assumptions:  Comboboxes of surface microhabitats set the
'               calling control using:
'                   =UpdateMicrohabitat([Screen].[ActiveControl])
'               in their change event properties
' Parameters:   caller - calling control (control)
' Returns:      if successful - submitted cover value (single)
'               or 0 if not
' Throws:       none
' References:
'   Douglas J Steele, Dec 5, 2005
'   https://www.pcreview.co.uk/threads/get-control-name-in-click-event-procedure.2274312/
' Source/date:  Bonnie Campbell, April 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/24/2016 - initial version
' ---------------------------------
Public Function UpdateMicrohabitat(caller As Control) As Single
', transectID As String) As Single
', sfcID As Long, pctCover As Single)
On Error GoTo Err_Handler

'    Dim caller As Control
    Dim strSurface As String, strControl As String
    Dim sfcID As Integer
    Dim PctCover As Single
    Dim rs As DAO.Recordset
    
    'set surface ID (pull from global dictionary using control name - _Q#)
    strSurface = Left(caller.Name, Len(caller.Name) - 3)
    
    'if global dictionary not available, set it
    If IsNothing(g_AppSurfaces) Then GetSurfaceIDs
    sfcID = g_AppSurfaces(strSurface)
    
    'retrieve values
    PctCover = Nz(caller.Value, 0)
    
    'skip if NULL
    If IsNull(TempVars("Transect_ID")) Then GoTo Exit_Handler
    
    Dim sfc As New SurfaceCover
    
    With sfc
        '.QuadratID = CInt(Right(CStr(caller.Name), 1))
        .PercentCover = PctCover
        .SurfaceID = sfcID
        
        'fetch the appropriate QuadratID
        strControl = "tbxQ" & Right(CStr(caller.Name), 1)
        .QuadratID = Forms("frm_Data_Entry").Controls("frm_Quadrat_Transect").Form.Controls(strControl)
        
        'update values
        .SaveToDb True
    End With
    
    'SetRecord "u_surfacecover", params
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateMicrohabitat[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          UpdateCoverSpecies
' Description:  Updates microhabitat surface cover values (Invasives)
'               and returns the submitted value (single)
' Assumptions:  Comboboxes of surface microhabitats set the
'               calling control using:
'                   =UpdateCoverSpecies([Screen].[ActiveControl])
'               in their change event properties
' Parameters:   caller - calling control (control)
' Returns:      if successful - submitted cover value (single)
'               or 0 if not
' Throws:       none
' References:
'   Douglas J Steele, Dec 5, 2005
'   https://www.pcreview.co.uk/threads/get-control-name-in-click-event-procedure.2274312/
' Source/date:  Bonnie Campbell, April 24, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/24/2016 - initial version
' ---------------------------------
Public Function UpdateCoverSpecies(caller As Control) As Single
On Error GoTo Err_Handler

'    Dim caller As Control
    Dim strQuadrat As String, strControl As String, strPosition As String
    Dim sfcID As Integer
    Dim PctCover As Single
    Dim rs As DAO.Recordset

    'retrieve calling control
    
    
    'set quadrat # (pull from global dictionary using control name - _Q#)
    strQuadrat = Replace(Left(caller.Name, 2), "Q", "")
    
    'retrieve values
    PctCover = Nz(caller.Value, 0)
    
    'skip if NULL
    If IsNull(TempVars("Transect_ID")) Then GoTo Exit_Handler
    
    Dim sp As New InvasiveCoverSpecies
    
    With sp
        '.QuadratID = CInt(Right(CStr(caller.Name), 1))
        .PctCover = PctCover
        '.IsDead = cbxIsDead
        '.Position =
        
        'fetch the appropriate QuadratID
        strControl = "tbxQ" & strQuadrat
        .QuadratID = Forms("frm_Data_Entry").Controls("frm_Quadrat_Transect").Form.Controls(strControl)
        
        'determine quadrat position (pull from global dictionary using control name)
        strPosition = g_AppQuadratPositions(caller.Name)
        
        'update values
        .SaveToDb True
    End With
    
    'SetRecord "u_surfacecover", params
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateCoverSpecies[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     FetchAddlData
' Description:  Retrieves additional data field(s)
' Assumptions:
'               fields are delimited w/ a pipe (|)
' Parameters:   tbl - name of table to retrieve from (string)
'               field(s) - name of field to retrieve (string)
'               id - record to retrieve's ID (long)
' Returns:      field value(s) for record (DAO.Recordset)
' Throws:       none
' References:
'   Steven Thomas, November 28, 2011
'   https://blogs.office.com/2011/11/28/display-real-time-information-with-the-controltip-property/
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function FetchAddlData(tbl As String, Fields As String, ID As Long) As DAO.Recordset
On Error GoTo Err_Handler
    
    'values are required --> exit if not
    If Len(tbl) = 0 Or Len(Fields) = 0 Or Not (ID > 0) Then GoTo Exit_Handler
    
    'begin retrieval
    Dim field As String
    Dim strFields As String
    Dim strSQL As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = CurrDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
            
            'check for multiple fields
            If InStr(Fields, "|") > 0 Then
                Dim aryFlds() As String
                Dim i As Integer
                
                aryFlds = Split(Fields, "|")
                
                For i = 0 To UBound(aryFlds)
                    strFields = aryFlds(i) & ","
                Next
                
                'remove extra comma
                strFields = IIf(Right(strFields, 1) = ",", RTrim(strFields), strFields)
            
            Else
                
                strFields = Fields
            End If
            
            'base
            strSQL = "SELECT " & strFields & " FROM " & tbl & " WHERE ID = " & ID & ";"
            
            'update the query SQL
            .SQL = strSQL
            
            Dim rs As DAO.Recordset

            Set rs = .OpenRecordset
                        
            'send results
            Set FetchAddlData = rs
            
            'cleanup
            Set rs = Nothing
            Set qdf = Nothing
            Set db = Nothing

        End With
    End With
    

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FetchAddlData[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function


