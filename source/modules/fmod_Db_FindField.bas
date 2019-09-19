Option Compare Database
Option Explicit

'Purpose:   Search your database (tables, queries, forms, reports)
'           to find where a particular field name is used.
'Release:   April 2007 (a work in progress.)
'Documentation: http://allenbrowne.com/ser-73.html
'Author:    Allen Browne    (allen@allenbrowne.com)
'Versions:  Requires Access 2000 and later.
'           For Access 2000, you will need to remove this from the end of several lines:
'                   , WindowMode:=acHidden

'Usage examples
'==============
' To find where InvoiceID is used in Report1:
'           ? FindField("InvoiceID", "Report1")
' To find where ClientID is used in all forms and reports:
'           ? FindField("ClientID",,ffoForm + ffoReport)
' To find anywhere EventDate is used:
'           ? FindField("EntryDate")

'Summary
'=======
' Tables    Searches the Name and Caption of the fields, and the Filter and OrderBy of the table.
' Queries:  Searches the Name, SourceName, and Caption of fields; Filter and OrderBy of query.
' Forms:    Searches Name, ControlSource, Caption of controls,
'               and LinkMasterFields/LinkChildFields of subform controls
' Reports:  Searches Name, ControlSoruce, Caption of controls, Control Source of Group Levels,
'               and LinkMasterFields/LinkChildFields of subreport controls

'Notes
'=====
' When you type a SQL statement into the RecordSource of a form/report, or the RowSource
'           of a combo/listbox, Access creates a saved query with a name prefixed with ~sq_.
' With reports, click Ok if notified the report was set up for another printer.
'Does not search RecordSource of form/report, nor RowSource of combo/list box.
'Does not handle renamed fields that might be under the control of Name AutoCorrect.
'Does not handle query parameters (which are not fields.)

'Bitfield constants: their sum indicates which types of object to search.
Public Enum FindFieldObject
    ffoTable = 1        'Search table fields.
    ffoQuery = 2        'Search query fields.
    ffoForm = 4         'Search form controls and properties.
    ffoReport = 8       'Search report controls, properties, and group levels.
    ffoAll = 15         'Search all (tables, queries, forms, and reports.)
End Enum

Public Function FindField(strFieldName As String, _
    Optional strObjectName As String, _
    Optional iObjectTypes As FindFieldObject = ffoAll, _
    Optional bExactMatchOnly As Boolean) As Long
On Error GoTo Err_Handler
    'Purpose:   Search the current database for where a field name is used. MAIN FUNCTION.
    'Arguments: strFieldName:    the field name to find (or part of field name.)
    '           strObjectName:   Leave blank to search all objects. Only named object if entered.
    '           iObjectTypes:  determines what objects to search for. Sum of the types you want.
    '           bExactMatchOnly: not matched with wildcards if this is True.
    'Return:    Number of matches found.
    '           List of items in the Immediate Window (Ctrl+G.)
    'Usage:     To search tables and queries for a field named Inactive:
    '               Call FindField("Inactive", ffoTable + ffoQuery)
    Dim db As DAO.Database          'This database
    Dim tdf As DAO.TableDef         'Each table
    Dim qdf As DAO.QueryDef         'Each query
    Dim accObj As AccessObject      'Each form/report.
    Dim strDoc As String            'Name of form/report.
    Dim strText2Match As String     'strFieldName with wildcards.
    Dim bLeaveOpen As Boolean       'Flag to leave the form/report open if it was already open.
    Dim lngKt As Long               'Count of matches.
    
    'Initialize
    Set db = CurrentDb()
    If bExactMatchOnly Then
        strText2Match = strFieldName
    Else
        strText2Match = "*" & strFieldName & "*"
    End If
    
    'Search Tables
    If (iObjectTypes And ffoTable) <> 0 Then
        If strObjectName <> vbNullString Then
            'Just one table (if it exists)
            If ObjectExists(db.TableDefs, strObjectName) Then
                Set tdf = db.TableDefs(strObjectName)
                lngKt = lngKt + FindInTableQuery(tdf, strText2Match)
            End If
        Else
            'All tables
            For Each tdf In db.TableDefs
                lngKt = lngKt + FindInTableQuery(tdf, strText2Match)
            Next
        End If
    End If
    
    'Search Queries
    If (iObjectTypes And ffoQuery) <> 0 Then
        If strObjectName <> vbNullString Then
            'Just one query (if it exists)
            If ObjectExists(db.QueryDefs, strObjectName) Then
                Set qdf = db.QueryDefs(strObjectName)
                lngKt = lngKt + FindInTableQuery(qdf, strText2Match)
            End If
        Else
            'All queries
            For Each qdf In db.QueryDefs
                lngKt = lngKt + FindInTableQuery(qdf, strText2Match)
            Next
        End If
    End If

    'Search Forms.
    If (iObjectTypes And ffoForm) <> 0 Then
        If strObjectName <> vbNullString Then
            'Just one form (if it exists)
            If ObjectExists(CurrentProject.AllForms, strObjectName) Then
                Set accObj = CurrentProject.AllForms(strObjectName)
                strDoc = accObj.Name
                bLeaveOpen = accObj.IsLoaded
                DoCmd.OpenForm strDoc, acDesign, WindowMode:=acHidden
                'Search
                lngKt = lngKt + FindInFormReport(Forms(strDoc), strText2Match)
                'Close unless already open.
                If Not bLeaveOpen Then
                    DoCmd.Close acForm, strDoc, acSaveNo
                End If
            End If
        Else
            'All forms
            For Each accObj In CurrentProject.AllForms
                strDoc = accObj.Name
                bLeaveOpen = accObj.IsLoaded
                DoCmd.OpenForm strDoc, acDesign, WindowMode:=acHidden
                'Search
                lngKt = lngKt + FindInFormReport(Forms(strDoc), strText2Match)
                'Close unless already open.
                If Not bLeaveOpen Then
                    DoCmd.Close acForm, strDoc, acSaveNo
                End If
            Next
        End If
    End If
    
    'Search Reports.
    If (iObjectTypes And ffoReport) <> 0 Then
        If strObjectName <> vbNullString Then
            'Just one report (if it exists)
            If ObjectExists(CurrentProject.AllReports, strObjectName) Then
                Set accObj = CurrentProject.AllReports(strObjectName)
                strDoc = accObj.Name
                bLeaveOpen = accObj.IsLoaded
                DoCmd.OpenReport strDoc, acDesign, WindowMode:=acHidden
                'Search
                lngKt = lngKt + FindInFormReport(Reports(strDoc), strText2Match)
                'Check the Group Levels as well.
                lngKt = lngKt + FindInGroupLevel(Reports(strDoc), strText2Match)
                'Close unless already open.
                If Not bLeaveOpen Then
                    DoCmd.Close acReport, strDoc, acSaveNo
                End If
            End If
        Else
            'All reports
            For Each accObj In CurrentProject.AllReports
                strDoc = accObj.Name
                bLeaveOpen = accObj.IsLoaded
                DoCmd.OpenReport strDoc, acDesign, WindowMode:=acHidden
                'Search
                lngKt = lngKt + FindInFormReport(Reports(strDoc), strText2Match)
                'Check the Group Levels as well.
                lngKt = lngKt + FindInGroupLevel(Reports(strDoc), strText2Match)
                'Close unless already open.
                If Not bLeaveOpen Then
                    DoCmd.Close acReport, strDoc, acSaveNo
                End If
            Next
        End If
    End If

Exit_Handler:
    FindField = lngKt
    'Clean up
    Set accObj = Nothing
    Set qdf = Nothing
    Set tdf = Nothing
    Set db = Nothing
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "FindField"
    Resume Exit_Handler
End Function

Private Function FindInTableQuery(obj As Object, strText2Match As String) As Long
On Error GoTo Err_Handler
    'Purpose:   Find fields where the Name, SourceField, or Caption matches the string.
    'Arguments: obj = the TableDef or QueryDef to search.
    '           strText2Match is the text to search for, including any wildcards.
    'Return:    Count of matches listed.
    Dim fld As DAO.Field
    Dim lngKt As Long
    
    For Each fld In obj.Fields
        'Search the name
        If fld.Name Like strText2Match Then
            Debug.Print obj.Name & "." & fld.Name
            lngKt = lngKt + 1&
        'Search the SourceField (for aliased query fields.)
        ElseIf fld.SourceField Like strText2Match Then
            Debug.Print obj.Name & "." & fld.Name & ".SourceField: " & fld.SourceField
            lngKt = lngKt + 1&
        'Search the Caption.
        ElseIf HasProperty(fld, "Caption") Then
            If fld.Properties("Caption") Like strText2Match Then
                Debug.Print obj.Name & "." & fld.Name
                lngKt = lngKt + 1&
            End If
        End If
    Next
    Set fld = Nothing
    
    'Search the Filter and OrderBy properties too.
    lngKt = lngKt + FindInProperty(obj, "Filter", strText2Match)
    lngKt = lngKt + FindInProperty(obj, "OrderBy", strText2Match)

Exit_Handler:
    FindInTableQuery = lngKt
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "FindInTableQuery"
    Resume Exit_Handler
End Function

Private Function FindInFormReport(obj As Object, strText2Match As String) As Long
On Error GoTo Err_Handler
    'Purpose:   Search for controls where the Name, Control Source, or Caption matches the string.
    'Arguments: obj = a reference to the form or report.
    '           strText2Match is the text to search for, including any wildcards.
    'Return:    Count of matches listed.
    Dim ctl As Control      'Each control on the form/report.
    Dim lngKt As Long       'Count of matches.

    For Each ctl In obj.Controls
        'Search the name
        If ctl.Name Like strText2Match Then
            Debug.Print obj.Name & "." & ctl.Name & " (" & ControlTypeName(ctl.ControlType) & ")"
            lngKt = lngKt + 1&
        'LinkMasterFields/LinkChildFields for subform/subreport.
        ElseIf ctl.ControlType = acSubform Then
            If ctl.LinkMasterFields Like strText2Match Then
                Debug.Print obj.Name & "." & ctl.Name & ".LinkMasterFields: " & ctl.LinkMasterFields
                lngKt = lngKt + 1&
            End If
            If ctl.LinkChildFields Like strText2Match Then
                Debug.Print obj.Name & "." & ctl.Name & ".LinkChildFields: " & ctl.LinkChildFields
                lngKt = lngKt + 1&
            End If
        'Search the Control Source
        ElseIf HasProperty(ctl, "ControlSource") Then
            If ctl.Properties("ControlSource") Like strText2Match Then
                Debug.Print obj.Name & "." & ctl.Name & ".ControlSource: " & ctl.ControlSource
                lngKt = lngKt + 1&
            End If
        'Search the caption (less any hotkey.)
        ElseIf HasProperty(ctl, "Caption") Then
            If ctl.Caption Like Replace(strText2Match, "&", vbNullString) Then
                Debug.Print obj.Name & "." & ctl.Name & ".Caption: " & ctl.Caption
                lngKt = lngKt + 1&
            End If
        End If
    Next
    
    'Search the Filter and OrderBy properties too.
    lngKt = lngKt + FindInProperty(obj, "Filter", strText2Match)
    lngKt = lngKt + FindInProperty(obj, "OrderBy", strText2Match)
    
Exit_Handler:
    FindInFormReport = lngKt
    Set ctl = Nothing
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "FindInFormReport"
    Resume Exit_Handler
End Function

Private Function FindInProperty(obj As Object, strPropName As String, strText2Match As String) As Long
On Error GoTo Err_Handler
    'Purpose:   Search the Filter an OrderBy properties of the object for the string.
    'Arguments: obj           = a reference to TableDef, QueryDef, Form, or Report.
    '           strPropName   = name of property to search, e.g. "Filter" or "OrderBy".
    '           strText2Match = the text to search for, including any wildcards.
    'Return:    1 if found; 0 if not.
    
    If obj.Properties(strPropName) Like strText2Match Then
        Debug.Print obj.Name & "." & strPropName & ": " & obj.Properties(strPropName)
        FindInProperty = 1&
    End If
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
    Case 438&, 3270&                'Property doesn't apply; Property not found.
        'do nothing
    Case Else
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, ".FindInProperty"
    End Select
    Resume Exit_Handler
End Function

Private Function FindInGroupLevel(rpt As Report, strText2Match As String) As Long
On Error GoTo Err_Handler
    'Purpose:   Search the Control Source of each Group Level of a report.
    'Arguments: rpt = a reference to the report.
    '           strText2Match is the text to search for, including any wildcards.
    'Return:    Count of matches listed.
    'Note:      Assumes the report is open.
    Dim i As Integer        'Loop controller (group levels.)
    Dim lngKt As Long       'Count of matches
    
    Do      'Loop will terminate by error when there are no more group levels.
        If rpt.GroupLevel(i).ControlSource Like strText2Match Then
            Debug.Print rpt.Name & ".GroupLevel(" & i & "): " & rpt.GroupLevel(i).ControlSource
            lngKt = lngKt + 1&
        End If
        i = i + 1
    Loop
    
Exit_Handler:
    FindInGroupLevel = lngKt
    Exit Function

Err_Handler:
    If Err.Number <> 2464& Then     'No more group levels.
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "FindInGroupLevel()"
    End If
    Resume Exit_Handler
End Function

Public Function ObjectExists(obj As Object, strObjectName As String) As Boolean
    'Purpose:   Return True if the named object exists.
    'Examples:  If ObjectExists(CurrentDb.TableDefs, "Table1") Then
    '           If ObjectExists(CurrentProject.AllForms, "Form1") Then
    Dim varDummy As Variant
    On Error Resume Next
    varDummy = obj.Item(strObjectName).Name
    ObjectExists = (Err.Number = 0&)
End Function

Public Function ControlTypeName(lngControlType As AcControlType) As String
On Error GoTo Err_Handler
    'Purpose:   Return the name of the ControlType.
    'Argument:  A Long Integer that is one of the acControlType constants.
    'Return:    A string describing the type of control.
    'Note:      The ControlType is a Byte, but the constants are Long.
    Dim strReturn As String

    Select Case lngControlType
        Case acBoundObjectFrame: strReturn = "Bound Object Frame"
        Case acCheckBox: strReturn = "Check Box"
        Case acComboBox: strReturn = "Combo Box"
        Case acCommandButton: strReturn = "Command Button"
        Case acCustomControl: strReturn = "Custom Control"
        Case acImage: strReturn = "Image"
        Case acLabel: strReturn = "Label"
        Case acLine: strReturn = "Line"
        Case acListBox: strReturn = "List Box"
        Case acObjectFrame: strReturn = "Object Frame"
        Case acOptionButton: strReturn = "Object Button"
        Case acOptionGroup: strReturn = "Option Group"
        Case acPage: strReturn = "Page (of Tab)"
        Case acPageBreak: strReturn = "Page Break"
        Case acRectangle: strReturn = "Rectangle"
        Case acSubform: strReturn = "Subform/Subrport"
        Case acTabCtl: strReturn = "Tab Control"
        Case acTextBox: strReturn = "Text Box"
        Case acToggleButton: strReturn = "Toggle Button"
        Case Else: strReturn = "Unknown: type" & lngControlType
    End Select
    
    ControlTypeName = strReturn

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ControlTypeName()"
    Resume Exit_Handler
End Function

Public Function HasProperty(obj As Object, strPropName As String) As Boolean
    'Purpose:   Return true if the object has the property.
    Dim varDummy As Variant
    
    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function