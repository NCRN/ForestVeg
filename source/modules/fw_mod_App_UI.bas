Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_App_UI
' Level:        Application module
' Version:      1.43

' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               ----------- invasives reports -----------------
'               BLC, 5/26/2015 - 1.01 - added PopulateSpeciesPriorities function from mod_Species
'               BLC, 6/1/2015  - 1.02 - changed View to Search tab
'               BLC, 6/12/2015 - 1.03 - added EnableTargetTool button
'               ----------- big rivers ------------------------
'               BLC, 6/30/2015 - 1.04 - added ClearFields()
'               BLC, 7/27/2015 - 1.05 - added SetHints()
'               ----------- uplands ---------------------------
'               BLC, 8/21/2015 - 1.06 - added CaptureEscapeKey
'               BLC, 2/3/2016  - 1.07 - added SetNoDataCheckbox()
'               BLC, 2/9/2016  - 1.08 - added public dictionary for NoData checkboxes
'                                       dictionary is used within subforms to identify if checkboxes
'                                       should be checked, GetNoDataCollected(), SetNoDataCollected()
'               BLC, 2/9/2016 - 1.09 - added constants, functions & subroutine supporting transect overlays
'                                       (LWA_ALPHA, GWL_EXSTYLE, WS_EX_LAYERED, GetWindowLong(),
'                                       SetWindowLong(), SetLayeredWindowAttributes(), SetFormOpacity())
'               BLC, 3/17/2016 -1.10 - added SetControlBackcolor(), CTRL_DEFAULT_BACKCOLOR, Check1000hrFuels
'               BLC, 3/29/2016 -1.11 - added SetControlHighlight()
'               BLC, 4/1/2016 - 1.12 - added AddTallyValue()
'               BLC, 3/22/2017 - 1.13 - added SortListForm() from big rivers,
'                                       moved to mod_Forms (6/1/2016 big rivers dev):
'                                       CaptureEscapeKey(), SetFormOpacity()
'               BLC, 3/23/2017 - 1.14 - added PopulateForm(), DeleteRecord() from big rivers
'               BLC, 3/30/2017 - 1.15 - moved DeleteRecord() to mod_Db
' --------------------------------------------------------------------
'               BLC, 9/18/2017 - 1.22 - merged prior work:
'
'                   ----------- big rivers -------------
'                   BLC, 11/19/2015 - 1.02 - added CreateEnums call to initApp
'                   BLC, 4/26/2016  - 1.03 - added ClickAction() for handling various app actions
'                   BLC, 6/24/2016 - 1.04 - replaced Exit_Function > Exit_Handler
'                   BLC, 7/5/2016  - 1.05 - added ClearFields() to support Species Search
'                   BLC, 8/8/2016 - 1.06 - revised to use default table name in PopulateForm()
'                   BLC, 8/29/2016 - 1.07 - revised to use usys_temp_qdf & Contact_ID in PopulateForm()
'                                           for Contact form
'                   BLC, 8/30/2016 - 1.08 - added Batch Upload Photos to ClickAction()
'                   BLC, 9/13/2016 - 1.09 - added SortList()
'                   BLC, 10/14/2016 - 1.10 - added SetContext()
'                   BLC, 10/19/2016 - 1.11 - revised to use UploadCSVFile() vs. UploadSurveyFile()
'                   BLC, 10/24/2016 - 1.12 - added modwentworth form
'                   BLC, 10/25/2016 - 1.13 - added originForm TempVar for species seach
'                   BLC, 12/9/2016 -  1.14 - added PopulateCSVFields()
'                   BLC, 12/13/2016 - 1.15 - added SetCurrentPseudoRecord()
'                   BLC, 1/9/2017   - 1.16 - revised ClickAction() to use SetTempVar()
'                   BLC, 1/12/2017 - 1.17 - revised to VegTransect vs. Transect form
'                   BLC, 1/31/2017 - 1.18 - adjusted SortListForm() to accommodate template list form
'                   BLC, 2/2/2017  - 1.19 - commented CreateEnums call in Initialize(),
'                                           most/not all enums handled through calls to AppEnum table
'                   BLC, 2/14/2017 - 1.20 - added Task form to PopulateForm()
'                   BLC, 2/21/2017 - 1.21 - adjusted SortListForm() to accommodate Contact list form,
'                                           revised to use Photo vs. Tree form
'                   ----------- inavsive reports -------
'                   BLC, 9/21/2015 - 1.05 - added park species list, park summary report
' --------------------------------------------------------------------
'               BLC, 9/29/2017 - 1.23 - added logger case
'               BLC, 10/4/2017 - 1.24 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC, 10/16/2017 - 1.25 - adjusted Contact to include IsNPS (PopulateForm())
'               BLC, 10/18/2017 - 1.26 - added ClickAction() AppSettings case, SortListForm() Comment cases
'               BLC, 10/19/2017 - 1.27 - adjusted Comment and Location cases (PopulateForm())
'               BLC, 10/30/2017 - 1.28 - add Location cbxCollectionSourceID setting (PopulateForm())
'               BLC, 10/31/2017 - 1.29 - added ReplicatePlot, CalibrationPlot (VegPlot)
'               BLC, 11/3/2017  - 1.30 - update Location case (PopulateForm())
'               BLC, 11/9/2017  - 1.31 - update VegPlot case, checkboxes & toggles;
'                                        Transducer case distances (PopulateForm())
'               BLC, 11/11/2017 - 1.32 - update VegPlot case (PopulateForm())
'               BLC, 12/5/2017 - 1.33 - add VegPlot BeaverBrowse (PopulateForm())
'               BLC, 12/7/2017 - 1.34 - add VegPlot, Event SortListForm cases
'               BLC, 12/8/2017 - 1.35 - added obs-photos ClickAction() case, VegPlot PopulateForm()
'               BLC, 12/14/2017 - 1.36 - updated Loggers ClickAction() case
'               BLC, 12/27/2017 - 1.37 - update PopulateForm VegPlot case to set combobox values
'               BLC, 1/10/2018  - 1.38 - added DisplayFormats()
'               BLC, 1/11/2018  - 1.39 - added ClearMsgIcon()
'               BLC, 1/17/2018  - 1.40 - added Task list case (SortListForm)
'               BLC, 1/19/2018  - 1.41 - added SetMsgIcon()
'               BLC, 5/16/2019  - 1.42 - added fw_ module prefix
'               BLC, 3/9/2020   - 1.43 - 64-bit OS update
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
' -- Constants --
Private Const LWA_ALPHA     As Long = &H2
Private Const GWL_EXSTYLE   As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000

Public Const CTRL_DEFAULT_BACKCOLOR  As Long = 65535  'RGB(255, 255, 0) highlight yellow

' -- Values --
Public NoData As Scripting.Dictionary

' -- Functions --
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long
 
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal crKey As Long, _
   ByVal bAlpha As Byte, _
   ByVal dwFlags As Long) As Long


' ---------------------------------
'  Methods
' ---------------------------------

' *********************************
'    Common
' *********************************
' =================================
' SUB:     PopulateInsetTitle
' Description:  Sets inset title on form
' Assumptions:
' Parameters:   ctrl - control to update (Control)
'               strContext - context for title (string)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - initial version
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                                Converted QAQC to Create, Logs to View
'               BLC, 5/26/2015 - Added error handling
'               BLC, 6/4/2015 - Changed View to Search tab, added "or modify" for create tab
'               BLC, 9/21/2015 - added park species list, park summary report
' =================================
Public Sub PopulateInsetTitle(ctrl As Control, strContext As String)
On Error GoTo Err_Handler
    
    Dim strTitle As String
    
    Select Case strContext
        Case "Create" ' Create main
            strTitle = "Choose what you'd like to create"
        Case "CreateTgtLists" ' Create species target lists
            strTitle = "Create > Species Target Lists"
        Case "AddTgtArea" ' Add target areas
            strTitle = "Create > Add Target Area"
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates"  ' QA/QC > Outliers etc.
            strContext = Replace(Replace(strContext, "Suspect", "Suspect "), "Missing", "Missing ")
            strTitle = "Data Validation > " & strContext
        Case "Data Validation" ' QA/QC analysis project selection
            strTitle = "Data Validation > Field > Duplicates (NFV)" '<<<<< Make this so it ties back to the selected analysis
        Case "View" ' View main
            strTitle = "View"
        Case "Search" ' Search main
            strTitle = "Species Search"
        Case "Reports" ' Reports main
            strTitle = "Reports"
        Case "CrewSpeciesList" ' Reports > Field Crew Species List
            strTitle = "Reports > Field Crew Species List"
        Case "ParkSpeciesList" ' Reports > Park Personnel Species List
            strTitle = "Reports > Park Personnel Species List"
        Case "SpeciesListByPark" ' Reports > Species List By Park
            strTitle = "Reports > Species List By Park"
        Case "TgtListAnnualSummary" ' Reports > Annual Species List Summary
            strTitle = "Reports > Annual Species List Summary"
        Case "TgtListParkSummary" ' Reports > Park Species List Summary
            strTitle = "Reports > Park Species List Summary"
        Case "CrewVegWalk" ' Reports > Field Crew Species List
            strTitle = "Reports > Field Crew Species List"
        Case "VegWalkByPark" ' Reports > Species List By Park
            strTitle = "Reports > Species List By Park"
        Case "TgtListAnnualSummary" ' Reports > Annual Species List Summary
            strTitle = "Reports > Annual Species List Summary"
        Case "Precision", "Effectiveness", "Bias", "Stage", "Flow" ' Reports > Precision etc.
            strTitle = "Reports > " & strContext
        Case "Export" ' Export main
            strTitle = "Export"
        Case "UtahLab" ' Exports > Utah Lab etc.
            strContext = Replace(strContext, "Lab", " Lab")
            strTitle = "Exports > " & strContext
        Case "DB Admin" ' DB Admin main
            strTitle = ""
    End Select
    
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strTitle
        If strContext <> "DbAdmin" Then
            ctrl.visible = True
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateInsetTitle[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' SUB:     PopulateInstructions
' Description:  Sets form instruction strings
' Assumptions:  -
' Parameters:   strTab - tab for instruction string
' Returns:      aryCrumbs - array of breadcrumb values
' Throws:       none
' References:   none
' Source/date:
'               Created 06/12/2014 blc; Last modified 06/12/2014 blc.
' Revisions:    Bonnie Campbell, June 12, 2014 - initial version
'               --------------------------------------------------------------------------------------
'               BLC, 4/21/2015 - Adapted for NCPN Invasives Reports - Species Target List tool
'                                Converted QAQC to Create, Logs to View
'               BLC, 5/26/2015 - Added error handling
'               BLC, 6/4/2015  - Changed View to Search
'               BLC, 9/21/2015 - added park species list, park summary report
' =================================
Public Sub PopulateInstructions(ctrl As Control, strContext As String)
On Error GoTo Err_Handler
    Dim strInstructions As String
    
    'MsgBox strContext
    
    Select Case strContext
        Case "Create" ' Create main
            strInstructions = "Choose what you would like to create."
        Case "CreateTgtLists" ' Create > Species Target Lists
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your list." & vbCrLf & vbCrLf & _
                    "Only existing lists for the current or future years may be modified." & vbCrLf & vbCrLf & _
                    "Please contact the project lead or data management if a prior year list must be modified."
        Case "AddTgtArea" ' Create > Add Target Area
            strInstructions = "" '"Choose the park and year for your target area. Click 'Continue' to create your area."
        Case "Outliers", "MissingData", "SuspectValues", "SuspectDO", "SuspectpH", "SuspectSC", "SuspectWT", "Duplicates" ' QA/QC main
            strInstructions = "Complete the fields to define the data set or subset you are validating. " _
                    & "Leave the fields blank if you are validating all data. Click 'Run' to validate."
        Case "View" ' View main
            strInstructions = "The view menu is currently not in use for this application."
            'strInstructions = "Log your modifications to data within the edit log. " _
            '        & "Be as complete as possible to aid others in tracing data changes."
        Case "Search" ' Search main
            strInstructions = "Search for species family, name, codes. " & _
                    "Latin, common, and state specific (UT, CO, WY) genus species names " & _
                    "and lookup (6-letter) and ITIS codes are included." & vbCrLf & vbCrLf & _
                    "Searches can be made across all or only a few species names/codes."
            'strInstructions = "Log your modifications to data within the edit log. " _
            '        & "Be as complete as possible to aid others in tracing data changes."
        Case "Reports" ' Reports main
            strInstructions = "Choose the report you would like to run."
        Case "CrewVegWalk" ' Reports > Field Crew Species List
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "VegWalkByPark" ' Reports > Species List By Park
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "CrewSpeciesList" ' Reports > Field Crew Species List
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "ParkSpeciesList" ' Reports > Park Personnel Species List
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "SpeciesListByPark" ' Reports > Species List By Park
            strInstructions = "Choose the park and year for your list. Click 'Continue' to prepare your report."
        Case "TgtListAnnualSummary"
            strInstructions = "Choose the year(s) for your list. Click 'Continue' to prepare your report." & vbCrLf & vbCrLf & _
                            "This report may take a minute to create and display. " & vbCrLf & _
                            "Calculated summary values will display once the report has finished rendering. " & vbCrLf & vbCrLf & _
                            "Your patience is appreciated."
        Case "Precision", "Effectiveness", "Bias", "Stage", "Flow" ' Reports > Precision etc.
            strInstructions = "Complete the fields to define the data set or subset you are reporting. " _
                    & "Leave the fields blank if you are reporting on all data. Click 'Run' to validate."
        Case "Export" ' Export main
            strInstructions = "After opening a report from the report tab, use the Export menu above in the application menu to export reports to your desired format."
        Case "UtahLab" ' Exports > Utah Lab etc.
            strInstructions = "Choose the export you would like to run."
        Case "DbAdmin" ' DB Admin main
            strInstructions = "The database administration tab is currently not in use for this application."
            'strInstructions = ""
    End Select
    
    'populate caption & display instructions
    If ctrl.ControlType = acLabel Then
        ctrl.Caption = strInstructions
        If strContext <> "DbAdmin" Then
            ctrl.visible = True
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateInstructions[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     PopulateSpeciesPriorities
' Description:  Populate species priority values from species priority concatenation
' Assumptions:  Park priority textboxes are named tbxPARKPriority (e.g. tbxZIONPriority)
' Parameters:   parkCode - 4 character park code (string)
'               priorities - species priority string concatenation for all parks (e.g. "BLCA-1|COLM-Transect|FOBU-1")
'               TargetYear - year for target species (integer)
' Returns:      Priority - value for park species priority (string)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, April 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 4/9/2015 - initial version
'   BLC - 5/26/2015 - moved from mod_Species to mod_App_UI
'   BLC - 9/30/2015 - added optional TargetYear for park summary report
' ---------------------------------
Public Function PopulateSpeciesPriorities(ParkCode As String, priorities As String, _
                                            Optional TargetYear As Integer = -1) As String
On Error GoTo Err_Handler

Dim ParkPriorities As Variant
Dim i As Integer, z As Integer

    'check if parkCode is in priorities string
    If Len(priorities) > Len(Replace(priorities, ParkCode, "")) Then
    
        'prepare the Park Priority values
        ParkPriorities = Split(priorities, "|")
        
        'set park priority values
        For i = 0 To UBound(ParkPriorities)
        
            'does Park have a priority value?
            If ParkCode = Left(ParkPriorities(i), 4) Then
            
                'park summary report check
                If TargetYear > 0 Then
                
                    If TargetYear = CInt(mid(ParkPriorities(i), 6, 4)) Then
                        'priority is for park & target year
                        PopulateSpeciesPriorities = Replace(ParkPriorities(i), ParkCode + "-" + CStr(TargetYear) + "-", "")
                        Exit For
                    Else
                        'not listed
                        PopulateSpeciesPriorities = "X"
                    End If
                    
                Else
                    'annual summary report
                    PopulateSpeciesPriorities = Replace(ParkPriorities(i), ParkCode + "-", "")
                End If
                
            End If
        
        Next
        
    Else
        'not listed
        PopulateSpeciesPriorities = "X"
    
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateSpeciesPriorities[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          Initialize
' Description:  initialize application values
' Assumptions:  -
' Parameters:   N/A
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015  - initial version
'   BLC - 2/19/2015 - added dynamic getParkState() & standard error handling
'   BLC - 3/4/2015  - shifted colors to mod_Color, removed setting of park, state, tgtYear TempVars
'   BLC - 5/13/2015 - stub only
'   BLC - 11/19/2015 - added CreateEnums call to create application specific Enums,
'                      updated documentation to reflect mod_App_UI vs. mod_Init
'   BLC - 2/2/2017  - comment: most enums handled through calls to AppEnum table
'                     however some require CreateEnums()
'   BLC - 10/18/2017 - add setting for turning ON/OFF enum creation
' ---------------------------------
Public Sub Initialize()
On Error GoTo Err_Handler

    'create the enums specific to this application from the Enums table &
    'mod_App_Enum stub module
    'CreateEnums requires BOTH mod_Enum & mod_App_Enum files to be re-imported to
    'the database
    If CREATE_ENUMS = True Then CreateEnums

    'set application UI display
'     SetStartupOptions "AppTitle", dbText, "NCPN Big Rivers"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Initialize[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

'----------------------------------------------
' RETIRED - 7/1/2020 - covered in mod_App_UI
'----------------------------------------------
'' ---------------------------------
'' Sub:          SortListForm
'' Description:  form label sort on click actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:
''   pere_de_chipstic, August 5, 2012
''   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
''   Allen Browne, June 28, 2006
''   https://bytes.com/topic/access/answers/506322-using-orderby-multiple-fields
'' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 1/19/2017 - initial version
''   BLC - 1/31/2017 - adjusted to accommodate templates list
''   BLC - 2/21/2017 - adjusted to accommodate Contact list
''   BLC - 10/18/2017 - added cases for Comment list
''   BLC - 12/7/2017 - added cases for VegPlot, Event lists
''   BLC - 1/17/2018 - added cases for Task list
'' ---------------------------------
'Public Sub SortListForm(frm As Form, ctrl As Control)
'On Error GoTo Err_Handler
'
'    Dim strSort As String
'
'    'default
'    strSort = ""
'
'    'set sort field
'    Select Case Replace(ctrl.Name, "lbl", "")
'        Case "Comment"
'            strSort = "Comment"
'        Case "CommentType"
'            strSort = "CommentType"
'        Case "CommentTypeID"
'            strSort = "CommentType_ID"
'        Case "Email"
'            strSort = "Email"
'        Case "HdrID"
'            strSort = "ID"
'            Select Case frm.Name
'                Case "ContactList"
'                    strSort = "c.ID"
'                Case "TaskList"
'                    strSort = "t.ID"
'            End Select
'        Case "Location"
'            strSort = "Location"
'        Case "ModalSedSize"
'            strSort = "ModalSedimentSize_ID"
'        Case "Name"
'            strSort = "LastName"
'        Case "PctMSS"
'            strSort = "PctModalSedimentSize"
'        Case "PlotNumDist"
'            strSort = IIf(TempVars("ParkCode") = "DINO", "PlotNumber", "PlotDistance_m")
'        Case "Priority"
'            strSort = "Priority"
'        Case "Site"
'            strSort = "Site"
'        Case "SOP"
'            strSort = "FullName"
'        Case "SOPNum"
'            strSort = "SOPNumber"
'        Case "StartDate"
'            strSort = "StartDate"
'        Case "Syntax"
'            strSort = "Syntax"
'        Case "Task"
'            strSort = "Task"
'        Case "TaskType"
'            strSort = "TaskType"
'        Case "Template"
'            strSort = "TemplateName"
'        Case "Status"
'            strSort = "Status"
'        Case "Version"
'            strSort = "Version"
'        Case "LastModifiedDate"
'            strSort = "LastModified"
'        Case "EffectiveDate"
'            strSort = "EffectiveDate"
'        Case ""
'    End Select
'
'    'set the sort
'    If InStr(frm.OrderBy, strSort) = 0 Then
'        frm.OrderBy = strSort
'    ElseIf Right(frm.OrderBy, 4) = "Desc" Then
'        frm.OrderBy = strSort
'    Else
'        frm.OrderBy = strSort & " Desc"
'    End If
'
'    frm.OrderByOn = True
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - SortListForm[fw_mod_App_UI form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' SUB:          ClearFields
' Description:  initialize application values
' Assumptions:  -
' Parameters:   frm - Form whose fields should be cleared
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015  - initial version
'   BLC - 5/18/2015  - fixed error documentation ClearFields vs. ITIS_Click, mod_Forms vs. frm_SpeciesSearch
'   BLC - 6/30/2015  - moved to mod_App_UI
'   BLC - 7/5/2016   - added from Invasives Reporting mod_App_UI to support Species Search
' ---------------------------------
Public Sub ClearFields(frm As Form)
On Error GoTo Err_Handler

    Select Case frm.Name
    
        Case "frm_Species_Search"
            frm.Controls("cbxCO").DefaultValue = False
            frm.Controls("cbxUT").DefaultValue = False
            frm.Controls("cbxWY").DefaultValue = False
            frm.Controls("cbxITIS").DefaultValue = False
            frm.Controls("cbxCommon").DefaultValue = False
            frm.Controls("tbxSearchFor").Value = ""
    End Select
    
Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearFields[fw_mod_App_UI])"
    End Select
    Resume Exit_Sub
End Sub

' *********************************
'    Big Rivers
' *********************************
' ---------------------------------
' SUB:          SetHints
' Description:  Sets hints for form actions
' Assumptions:  -
' Parameters:   frm - form object to set hints on (form)
'               strForm - form name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, July 27, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/27/2015  - initial version
' ---------------------------------
Public Sub SetHints(frm As Form, strForm As String)
On Error GoTo Err_Handler

' Forms!Mainform!Subform1.Form!
 
    With frm!fsub.Form
 
        Select Case strForm
 
            Case "fsub_Photo_FTOR_Details"
 
                !lblCloseupHint.Caption = "Is the photo a closeup?"
                !lblReplacementHint.Caption = "Does photo replace another?"
                !lblCommentHint.Caption = ""
 
                Select Case TempVars("phototype")
                    Case "R" 'reference
                        !lblPhotogLocHint.Caption = "from river, 10m upstream, etc."
                        !lblSubjectLocHint.Caption = "CP1, RM2, etc."
                    Case "O" 'overview
                        !lblPhotogLocHint.Caption = ""
                        !lblSubjectLocHint.Caption = "O1, O2, etc."
                    Case "T" 'transect
                        !lblPhotogLocHint.Caption = "T + transect# - order# (T2-1)"
                        !lblSubjectLocHint.Caption = ""
                    Case "F" 'feature
                        !lblPhotogLocHint.Caption = "F + transect# - order# " & vbCrLf & "(F3/4-2)"
                        !lblSubjectLocHint.Caption = ""
                End Select
 
            Case "fsub_Photo_Other_Details"
                !lblDescriptionHint.Caption = ""
            Case Else
 
        End Select
 
        !lblPhotoNumHint.Caption = "P + Month" & vbCrLf & "(Jan-Sep=0-9,Oct-Dec=A-C) + day(01-31) + " & vbCrLf & "4-digit camera seq# " & vbCrLf & "(PA010300 = Jan 1, #300)"
                  
      End With
      
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetHints[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     SetCurrentPseudoRecord
' Description:  sets a pseudo current record # based on the combobox w/ current focus
' Assumptions:  -
' Parameters:   ctrl - control with focus (control)
' Returns:      current # for combobox (integer)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, December 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/13/2016 - initial version
' ---------------------------------
Public Function SetCurrentPseudoRecord(ctrl As Control) As Integer
On Error GoTo Err_Handler
              
    'set psuedo current record <- set the # of the cbx
'    'MsgBox ActiveControl.Name
'    MsgBox Screen.ActiveForm.Name & " is the active form."
'    MsgBox Screen.ActiveControl.Name & "is the active control."
'    If InStr(Me.ActiveControl, "cbxColumnName") Then
'        Me.Parent.Controls("tbxCSVRecord").Value = Replace(Me.ActiveControl.Name, "cbxColumnName", "")
'    End If

    If InStr(ctrl.Name, "cbxColumnName") Then
        ctrl.Parent.Form.Parent.Form.Controls("tbxCSVRecord").Value = Replace(ctrl.Name, "cbxColumnName", "")
        ChangeBackColor ctrl, lngYelLime
        Call Forms("ImportMap").tbxCSVRecord_Change
    End If
              
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetCurrentPseudoRecord[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

'----------------------------------------------
' RETIRED - 7/1/2020 - compile issues
'----------------------------------------------
'' ---------------------------------
'' SUB:          ClickAction
'' Description:  Handles click events for various form links
'' Assumptions:  Link caption and tag text matches action text values.
''               If a link caption &/or tag changes, the corresponding action must change
''               here too.
'' Parameters:   action - concatenated link label caption & tag (string)
'' Returns:      N/A
'' Throws:       none
'' References:   none
'' Source/date:
'' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
'' Revisions:
''   BLC - 4/26/2016  - initial version
''   BLC - 8/30/2016  - added Batch Upload Photos
''   BLC - 10/19/2016 - revised to use UploadCSVFile() vs. UploadSurveyFile()
''   BLC - 10/25/2016 - revised species search to add originForm TempVar, callingform oArg
''   BLC - 1/9/2017   - revised to use SetTempVar()
''   BLC - 2/21/2017  - revised to use Photo vs. Tree form
''   BLC - 9/29/2017  - added Logger case
''   BLC - 10/18/2017 - added AppSettings case
''   BLC - 12/8/2017  - add photos
''   BLC - 12/14/2017 - updated Loggers case
'' ---------------------------------
'Public Sub ClickAction(action As String)
'On Error GoTo Err_Handler
'
'    Dim fName As String, rName As String, oArgs As String
'    Dim StartFolder As String, strPics As String, strPath As String
'
'    action = LCase(Nz(Trim(action), ""))
'
'    'defaults
'    fName = ""
'    rName = ""
'    oArgs = ""
'
'    Select Case Trim(action)
'        'Where?
'        Case "site"
'            fName = "Site"
'        Case "feature"
'            fName = "Feature"
'        Case "transect"
'            fName = "VegTransect"
'            oArgs = ""
'        Case "plot"
'            fName = "VegPlot"
'        Case "location"
'            fName = "Location"
'        'Sampling
'        Case "event"
'            fName = "Events"
'            oArgs = "" 'site & protocol IDs
'        Case "vegplots"
'            fName = "VegPlot"
'            oArgs = "" 'site & protocol IDs
'        Case "locations"
'            fName = "Location"
'            oArgs = "" 'collection source name - feature (A-G), transect #(1-8) &
'        Case "logger"
'            fName = "Logger"
'            oArgs = ""
'        Case "people"
'            fName = "Contact"
'            oArgs = "Main"
'        'Vegetation
'        Case "veg plots"
'            fName = "VegPlot"
'        Case "woody canopy cover"
'            fName = "VegWalk" '"WoodyCanopyCover"
'            oArgs = "" '"1|2016|WCC"
'        Case "understory cover"
'        Case "vegetation walk"
'            fName = "VegWalk"
'        Case "species"
'            fName = "Species"
'        Case "unknowns"
'            fName = "Unknown"
'        Case "species search"
'            fName = "SpeciesSearch"
'            oArgs = "Main"
'            'disable double click events
'            SetTempVar "originForm", "DisableDoubleClick"
''            If Not IsNull(TempVars("originForm")) Then
''                TempVars("originForm") = "DisableDoubleClick"
''            Else
''                TempVars.Add "originForm", "DisableDoubleClick"
''            End If
'        'Observations
'        Case "photos", "obs-photos"
'            fName = "Photo" '"Tree"
'        Case "transducers"
'            fName = "Transducer"
'            oArgs = ""
'        Case "Survey Files"
'            fName = "SurveyFile"
'            oArgs = ""
'        Case "Upload Survey File"
'            fName = ""
'            oArgs = ""
'
'            'handle upload
'            StartFolder = GetSpecialFolderPath("FOLDERID_Recent")
'
'            strPath = BrowseFolder("Select survey file to upload", "Confirm File", _
'                                    StartFolder, , msoFileDialogFilePicker, "Survey files-CSV")
'
'            If Len(strPath) > 0 Then
'                'open data form before upload
'                DoCmd.OpenForm "SurveyFile", acNormal, , , , , strPath
'
'                'upload survey file
''                UploadCSVFile strPath
'            End If
'
'            'restore Main
''            ToggleForm "Main", 0
'
'        Case "batch upload photos"
'            fName = ""
'            oArgs = ""
'
'            'handle upload
'
'            StartFolder = GetSpecialFolderPath("FOLDERID_Recent")
'
'            strPath = BrowseFolder("Select directory with photos to upload", "Confirm Directory", _
'                                    StartFolder)
'
'            If Len(strPath) > 0 Then
'                'ingest photos as "U" - unclassified
'                IngestPhotos strPath, "U"
''            Else
''                MsgBox "Oops. Missed the directory the photos are in. " _
''                        & "Please re-select it.", vbOKOnly, "Missing Directory"
'            End If
'
'            'restore Main
'            ToggleForm "Main", 0
'        'Trip Prep
'        Case "vegplot"
'            rName = "VegPlot"
'            oArgs = ""
'        Case "vegwalk"
'            rName = "VegWalk"
'            oArgs = ""
'        Case "photo"
'            rName = "Photo"
'            oArgs = ""
'        Case "transducer"
'            rName = "Transducer"
'        Case "loggers"
'            fName = "Logger"
'            oArgs = ""
'        Case "tasks"
'            fName = "Task"
'        Case "application settings"
'            fName = "AppSettings"
'        Case "sediment class settings"
'            fName = "ModWentworth"
'        Case "sheet settings"
'            fName = "SetDatasheetDefaults"
'        'Reports
'        Case "# Plots"
'            rName = "NumPlots"
'        Case "VegPlot - Species #s"
'            rName = "NumSpecies"
'        Case "VegPlot - Species"
'            rName = "SpeciesCommon"
'        Case "VegWalk - Species #s"
'            rName = "NumSpeciesCommon"
'        Case "VegWalk - Species"
'            rName = "SpeciesUnique"
'        Case "More..."
'            fName = "AppReport"
'    End Select
'
'    If Len(fName) > 0 Then
'        Forms("Main").visible = False
'        DoCmd.OpenForm fName, acNormal, OpenArgs:=oArgs
'    ElseIf Len(rName) > 0 Then
'        'print preview mode - acViewPreview
'        DoCmd.OpenReport rName, acViewPreview
'    End If
'
'Exit_Handler:
'    Exit Sub
'
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - ClickAction[fw_mod_App_UI])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' SUB:          GetParks
' Description:  Retrieves list of parks from database
' Assumptions:  -
' Parameters:   active - flag if park is currently being sampled, 1-active, 0-inactive (boolean)
' Returns:      parks - list of park codes separated by "|" (string)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, May 18, 2016 - for NCPN tools
' Revisions:
'   BLC - 5/18/2016  - initial version
' ---------------------------------
Public Function GetParks() As String
On Error GoTo Err_Handler

    'defaults
        

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetParks[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          DisplayFormats
' Description:  Converts formats field to string of icons
' Assumptions:  -
' Parameters:   formats - document formats available (string) - uPDF, uDOC, etc.
' Returns:      display - icons displayed via unicode (string)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, January 10, 2018 - for NCPN tools
' Revisions:
'   BLC - 1/10/2018  - initial version
' ---------------------------------
Public Function DisplayFormats(formats As String) As String
On Error GoTo Err_Handler

    Dim ary As Variant
    Dim i As Integer
    Dim display As String
    
    ary = Split(formats, "|")
    
    For i = 0 To UBound(ary)
    
        If i > 0 Then display = display & " | "
    
        display = display & StringFromCodepoint(ary(i))
    
    Next
    
    DisplayFormats = display

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisplayFormats[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ClearMsgIcon
' Description:  Clears the captions of msg and msgIcon labels on a form
' Assumptions:  -
' Parameters:   frm - form whose msg & msgIcon captions are to be cleared (form)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, January 11, 2018 - for NCPN tools
' Revisions:
'   BLC - 1/11/2018  - initial version
' ---------------------------------
Public Sub ClearMsgIcon(frm As Form)
On Error GoTo Err_Handler

    With frm
        'clear msg & icon
        frm.Controls("lblMsg").forecolor = lngRobinEgg
        frm.Controls("lblMsgIcon").forecolor = lngRobinEgg
        frm.Controls("lblMsg").Caption = ""
        frm.Controls("lblMsgIcon").Caption = ""
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearMsgIcon[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetMsgIcon
' Description:  Sets the captions of msg and msgIcon labels on a form
' Assumptions:  Assumes both icon & message are the same forecolor
'               Default optionals = uDoubleTriangleBlk icon & yellow color
' Parameters:   frm - form whose msg & msgIcon captions are to be set (form)
'               msg - message to set (string)
'               icon - icon to set (string)
'               color - fore color (long)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, January 19, 2018 - for NCPN tools
' Revisions:
'   BLC - 1/19/2018  - initial version
' ---------------------------------
Public Sub SetMsgIcon(frm As Form, msg As String, _
            Optional icon As String = "default", _
            Optional color As Long = lngYellow)
On Error GoTo Err_Handler

    If icon = "default" Then _
        icon = StringFromCodepoint(uDoubleTriangleBlkR)

    With frm
        'set msg & icon
        frm.Controls("lblMsg").forecolor = color
        frm.Controls("lblMsgIcon").forecolor = color
        frm.Controls("lblMsg").Caption = msg
        frm.Controls("lblMsgIcon").Caption = icon
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetMsgIcon[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' *********************************
'    Invasives
' *********************************

' *********************************
'    Invasives Reports
' *********************************
' ---------------------------------
' SUB:          EnableTargetTool
' Description:  enable the target tool button
' Assumptions:  -
' Parameters:   ctrl - button to enable (control)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, June 4, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/4/2015  - initial version
'   BLC - 6/12/2015 - replaced TempVars.item("... with TempVars("...
' ---------------------------------
Public Sub EnableTargetTool(ctrl As Control)
On Error GoTo Err_Handler
    
    'enable button if connected
    If TempVars("Connected") Then
        ctrl.Enabled = True
    Else
        ctrl.Enabled = False
    End If

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - EnableTargetTool[fw_mod_Init])"
    End Select
    Resume Exit_Sub
End Sub


' *********************************
'    Uplands
' *********************************

' =================================
' SUB:          RollupReportbyPark
' Description:  Prepares concatenated report data
'               Looks for the number of records (years) for each ParkPlotSpecies (species found on a given park plot)
'               and concatenates the years (e.g. 2008|2009|2013 ) so that a species only takes up a single
'               row for a specific park plot in the report. This reduces report length by 50% or more.
' Assumptions:  Assumes that tlu_NCPN_Plants contains Utah_Species names for all species
'               identified in the plots. Also assumes temp_Sp_Rpt_by_Park_Complete has been run prior to
'               running this so the report is updated with the most recent data.
' Note:         -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, August 27, 2015 - for NCPN tools
' Revisions:    BLC, 8/27/2015 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                 multiple open connections
' =================================
Public Sub RollupReportbyPark()
On Error GoTo Err_Handler

    Dim strParkPlotSpecies As String, strSpeciesYears As String
    Dim strPark As String, strFamily As String, strUtah_Species As String, strParkPlot As String
    Dim intPlotID As Integer, i As Integer, iCount As Integer
    Dim rs As DAO.Recordset, rsTemp As DAO.Recordset, rsCount As DAO.Recordset
    'Dim blnAdd As Boolean
    'Dim strSpeciesYr As String
    Dim strSQL As String
    
    Dim strPrevParkPlotSpecies As String
    
    Set rs = CurrDb.OpenRecordset("temp_Sp_Rpt_by_Park_Complete")

    'remove existing table
    If DCount("[Name]", "MSysObjects", "[Name] = 'temp_Sp_Rpt_by_Park_Rollup'") = 1 Then _
            CurrDb.TableDefs.Delete "temp_Sp_Rpt_by_Park_Rollup"
    
    'create empty table
    CreateRollupTable
    Set rsTemp = CurrDb.OpenRecordset("temp_Sp_Rpt_by_Park_Rollup")
    
    'default
    strParkPlotSpecies = ""
    strSpeciesYears = ""
    'blnAdd = False
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do Until rs.EOF
            
            'set the current record's values
            strPark = rs("Unit_Code")
            intPlotID = rs("Plot_ID")
            strFamily = rs("Master_Family")
            strUtah_Species = rs("Utah_Species")
            strParkPlotSpecies = rs("ParkPlotSpecies")
            strParkPlot = rs("ParkPlot")
            'strSpeciesYr = rs("Year")
            
            If Not iCount > 0 Then
              'determine how many have the same ParkPlotSpecies
              strSQL = "SELECT COUNT(Year) AS NumRecords FROM temp_Sp_Rpt_by_Park_Complete WHERE ParkPlotSpecies = '" & strParkPlotSpecies & "';"
              Set rsCount = CurrDb.OpenRecordset(strSQL, dbOpenSnapshot)
              iCount = rsCount!NumRecords
            End If
          
            For i = 1 To iCount
              'add year if it's a new year
              If Len(strSpeciesYears) = Len(Replace(strSpeciesYears, CStr(rs("Year")), "")) Then
                  strSpeciesYears = IIf(Len(strSpeciesYears) > 0, strSpeciesYears & "|" & rs("Year"), rs("Year"))
              End If
              rs.MoveNext
            Next
            
            ' add new record
            With rsTemp
                .AddNew
                !Unit_Code = strPark
                !Plot_ID = intPlotID
                !Master_Family = strFamily
                !Utah_Species = strUtah_Species
                !SpeciesYears = IIf(Len(strSpeciesYears) > 0, strSpeciesYears, rs!Year)
                !PlotParkSpecies = strParkPlotSpecies
                !ParkPlot = strParkPlot
                'update when rs!ParkPlotSpecies <> strParkPlotSpecies
                .Update
            End With
            'reset values
            strSpeciesYears = ""
            iCount = 0
        Loop
    End If
    
Exit_Sub:
    Set rs = Nothing
    Set rsTemp = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RollupReportbyPark[fw_mod_App_UI])"
    End Select
    Resume Exit_Sub
End Sub

' =================================
' SUB:          CreateRollupTable
' Description:  Prepares rollup temporary table
' Assumptions:
' Note:         -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, August 27, 2015 - for NCPN tools
' Revisions:    BLC, 8/27/2015 - initial version
'               BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                                 multiple open connections
' =================================
Public Sub CreateRollupTable()
On Error GoTo Err_Handler

    Dim tdf As DAO.TableDef
    
    Set tdf = CurrDb.CreateTableDef("temp_Sp_Rpt_by_Park_Rollup")
    
    'add the new record
    With tdf
        .Fields.Append .CreateField("Unit_Code", dbText)
        .Fields.Append .CreateField("Plot_ID", dbInteger)
        .Fields.Append .CreateField("Master_Family", dbText)
        .Fields.Append .CreateField("Utah_Species", dbText)
        .Fields.Append .CreateField("SpeciesYears", dbText)
        .Fields.Append .CreateField("PlotParkSpecies", dbText)
        .Fields.Append .CreateField("ParkPlot", dbText)
    End With

    CurrDb.TableDefs.Append tdf
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateRollupTable[fw_mod_App_UI])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          GetNoDataCollected
' Description:  Gets no data collected information from NoDataCollected table for event ID
' Assumptions:  -
' Parameters:   levelID - ID for event or event|transect as appropriate
'               level - event or transect (E = event, T = transect)
' Returns:      Dictionary of no data collected information (scripting.dictionary)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
'   BLC, 2/11/2016 - added level to accommodate both event & transect level identifiers
'   BLC, 3/18/2016 - added 1000hr fuel A-D to handle no fuels reported in comments for transects
'   BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
Public Function GetNoDataCollected(levelID As String, level As String) As Scripting.Dictionary
On Error GoTo Err_Handler

    Dim strSQL As String, strItem As String
    Dim rs As DAO.Recordset
    
    Set NoData = New Scripting.Dictionary 'publicly set
    
    'prepare default dictionary
    With NoData
        .Add "1mBelt-Shrub", 0
        .Add "1mBelt-TreeSeedling", 0
'        .Add "1mBelt-ExoticPerennial", 0
        .Add "1mBelt-Exotics", 0
        .Add "OverstoryTree-Sapling", 0
        .Add "OverstoryTree-Census", 0
        .Add "Fuel-1000hr", 0
        .Add "Fuel-1000hr-A", 0
        .Add "Fuel-1000hr-B", 0
        .Add "Fuel-1000hr-C", 0
        .Add "Fuel-1000hr-D", 0
        .Add "SiteImpact-Disturbance", 0
        .Add "SiteImpact-Exotic", 0
    End With
    
    strSQL = "SELECT SampleType FROM NoDataCollected WHERE ID = '" & levelID & "' AND SampleLevel = '" & level & "';"
    
    Set rs = CurrDb.OpenRecordset(strSQL)
    
    'rs.MoveFirst
    
    If Not (rs.EOF And rs.BOF) Then
    
        Do Until rs.EOF
    
            strItem = rs("SampleType") 'cannot use directly in NoData.item(rs("SampleType")) -> adds new item
            NoData.Item(strItem) = 1
            
            rs.MoveNext
            
        Loop
        
    End If
    
    Set GetNoDataCollected = NoData
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetNoDataCollected[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SetNoDataCollected
' Description:  Sets no data checkbox
' Assumptions:  Absolute value of Access/VBA checkbox is sent to drive 1 = true, 0 = false
'               SampleLevel is used vs. level in SQL (Access restricted word)
' Parameters:   levelID - ID for event/transect
'               level - sampling level identifier (E-event, T-transect)
'               SampleType - sub-protocol w/o data "1mBelt-Shrub", "OverstoryTree-Sapling", etc.
'               cbxValue - the value (1 or 0) to add or remove the flag
' Returns:      No data collected dictionary (scripting.dictionary)
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
'   BLC, 2/11/2016 - added level to accommodate both event & transect level identifiers
' ---------------------------------
Public Function SetNoDataCollected(levelID As String, level As String, SampleType As String, cbxValue As Integer) As Scripting.Dictionary
On Error GoTo Err_Handler
    
    Dim strSQL As String, strItem As String
    Dim rs As DAO.Recordset
    
    Set NoData = New Scripting.Dictionary 'publicly set
    Set NoData = GetNoDataCollected(levelID, level)
    
    NoData.Item(SampleType) = cbxValue
    
    'update the table appropriately
    If cbxValue = 1 Then
        strSQL = "INSERT INTO NoDataCollected(ID, SampleLevel, SampleType) VALUES ('" & levelID & "', '" & level & "', '" & SampleType & "');"
    ElseIf cbxValue = 0 Then
        strSQL = "DELETE * FROM NoDataCollected WHERE ID = '" & levelID & "' AND SampleLevel = '" & level & _
                    "' AND SampleType = '" & SampleType & "';"
    End If
    
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings (True)
    
    'return current dictionary object
    Set SetNoDataCollected = NoData
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetNoDataCollected[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          SetControlBackcolor
' Description:  sets controls backcolor based on control value
' Parameters:   ctrl - textbox control (textbox)
'               threshold - value to compare against (variant)
'               compareType - type of comparison (string)
'               color - numeric value for color (long) - result of RGB(r,g,b)
'               checkNULL - check if the control's value is NULL (boolean)
'               checkEmpty - check if the control's value is an empty string (boolean)
' Returns:      -
' Assumptions:  Assumes CTRL_DEFAULT_BACKCOLOR is set for the application
'               and that this is the typical backcolor for the controls
'               using SetControlBackcolor.
'               Assumes threshold value is numeric.
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/17/2016 - initial version
' ---------------------------------
Public Sub SetControlBackcolor(ctrl As TextBox, color As Long, checkNULL As Boolean, _
                        checkEmpty As Boolean, Optional threshold As Variant, Optional compareType As String)
On Error GoTo Err_Handler
    
    Dim resetcolor As Boolean
    
    'default
    resetcolor = False
    
    'change the backcolor --> revert to default only if the conditions aren't met
    ctrl.backcolor = color
    
    'null
    If checkNULL Then
        'reset backcolor if null
        If IsNull(Trim(ctrl.text)) Then
            resetcolor = True
            GoTo Exit_Handler
        End If
    End If
    
    'empty
    If checkEmpty Then
        'reset backcolor if empty
        If Len(Trim(ctrl.text)) = 0 Then
            resetcolor = True
            GoTo Exit_Handler
        End If
    End If
    
    'threshold
    If Not IsNull(threshold) And IsNumeric(ctrl.text) Then
        'set value base on compareType & threshold
        Select Case compareType
            Case "gt"
                If Not CDbl(ctrl.text) > threshold Then
                    resetcolor = True
                End If
            Case "gteq"
                If Not CDbl(ctrl.text) >= threshold Then
                    resetcolor = True
                End If
            Case "lt"
                If Not CDbl(ctrl.text) < threshold Then
                    resetcolor = True
                End If
            Case "lteq"
                If Not CDbl(ctrl.text) <= threshold Then
                    resetcolor = True
                End If
            Case "eq"
                If Not CDbl(ctrl.text) = threshold Then
                    resetcolor = True
                End If
        End Select
    End If
    
Exit_Handler:
    'reset to default backcolor
    If resetcolor Then
        ctrl.backcolor = CTRL_DEFAULT_BACKCOLOR
    End If
    
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetControlBackcolor[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Check1000hrFuels
' Description:  Handles 1000hr fuel check actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 3/18/2016  - initial version
'   BLC, 3/23/2016  - remove setting values when no records found
' ---------------------------------
Public Sub Check1000hrFuels()
On Error GoTo Err_Handler

    Dim frm As Form
    Set frm = Forms!frm_Data_Entry!fsub_Fuels_1000.Form

    '-----------------------------------
    ' update the NoDataCollected info IF no records now exist
    '-----------------------------------
    If frm.RecordsetClone.RecordCount = 0 Then
    
'        Dim NoData As Scripting.Dictionary
'
'        With frm.Parent.Form
'            'add the no data collected record
'            Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr", 1)
'
'            'update checkbox/rectangle --> No1000hr is not set here (leave commented out)
''            .Controls("cbxNo1000hr") = 1
''            .Controls("cbxNo1000hr").Enabled = True
''            .Controls("rctNo1000hr").Visible = True
'
'            'update A, B, C, D transect 1000hr fuels as well
'            .Controls("cbxNo1000hrA") = 1
'            .Controls("cbxNo1000hrA").Enabled = True
'            .Controls("rctNo1000hrA").Visible = True
'
'            .Controls("cbxNo1000hrB") = 1
'            .Controls("cbxNo1000hrB").Enabled = True
'            .Controls("rctNo1000hrB").Visible = True
'
'            .Controls("cbxNo1000hrC") = 1
'            .Controls("cbxNo1000hrC").Enabled = True
'            .Controls("rctNo1000hrC").Visible = True
'
'            .Controls("cbxNo1000hrD") = 1
'            .Controls("cbxNo1000hrD").Enabled = True
'            .Controls("rctNo1000hrD").Visible = True
'
'            'add the database records for A-D
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-A", 1
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-B", 1
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-C", 1
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-D", 1
'        End With
        
    Else
    
        'default values
        With frm.Parent.Form
            'update checkbox/rectangle (leave 1000hr commented here)
'            .Controls("cbxNo1000hr") = 0
'            .Controls("cbxNo1000hr").Enabled = True
'            .Controls("rctNo1000hr").Visible = True
            
            'update A, B, C, D transect 1000hr fuels as well
            .Controls("cbxNo1000hrA") = 0
            .Controls("cbxNo1000hrA").Enabled = True
            .Controls("rctNo1000hrA").visible = True
            
            .Controls("cbxNo1000hrB") = 0
            .Controls("cbxNo1000hrB").Enabled = True
            .Controls("rctNo1000hrB").visible = True
            
            .Controls("cbxNo1000hrC") = 0
            .Controls("cbxNo1000hrC").Enabled = True
            .Controls("rctNo1000hrC").visible = True
        
            .Controls("cbxNo1000hrD") = 0
            .Controls("cbxNo1000hrD").Enabled = True
            .Controls("rctNo1000hrD").visible = True
        End With
    
        'check for A, B, C, D transect 1000hr fuels
        Dim rs As DAO.Recordset
        
        Set rs = frm.RecordsetClone
        With rs
            .MoveFirst
            Do While Not .EOF
            Select Case .Fields("Transect")
            
                Case "A"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-A", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrA") = 0
                        .Controls("cbxNo1000hrA").Enabled = False
                        .Controls("rctNo1000hrA").visible = False
                    End With
                    
                Case "B"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-B", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrB") = 0
                        .Controls("cbxNo1000hrB").Enabled = False
                        .Controls("rctNo1000hrB").visible = False
                    End With
                    
                Case "C"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-C", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrC") = 0
                        .Controls("cbxNo1000hrC").Enabled = False
                        .Controls("rctNo1000hrC").visible = False
                    End With
                    
                Case "D"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-D", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrD") = 0
                        .Controls("cbxNo1000hrD").Enabled = False
                        .Controls("rctNo1000hrD").visible = False
                    End With
            End Select
            .MoveNext
            Loop
        End With
        
        'set checkboxes based on NoDataCollected (catch unchanged checkboxes)
        Dim dNoDataEvent As Scripting.Dictionary
        Set dNoDataEvent = GetNoDataCollected(frm.Parent.Form.Controls("Event_ID"), "E")
        
        With dNoDataEvent
            frm.Parent.Form.Controls("cbxNo1000hrA") = .Item("Fuel-1000hr-A")
            frm.Parent.Form.Controls("cbxNo1000hrB") = .Item("Fuel-1000hr-B")
            frm.Parent.Form.Controls("cbxNo1000hrC") = .Item("Fuel-1000hr-C")
            frm.Parent.Form.Controls("cbxNo1000hrD") = .Item("Fuel-1000hr-D")
        End With
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Check1000hrFuels[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetControlHighlight
' Description:  handles control highlight actions
' Parameters:   ctrl - textbox control (textbox)
'               -- optional --
'               threshold - value to compare control value to (double, default = 0)
'               compareType - how control value should be compared to threshold (string, default = "gteq")
' Returns:      -
' Assumptions:  highlighting will be consistent across all textboxes
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/29/2016 - initial version
' ---------------------------------
Public Sub SetControlHighlight(ctrl As TextBox, Optional threshold As Double, Optional compareType As String)
On Error GoTo Err_Handler

    'set defaults if optional values aren't set
    If Not IsNumeric(threshold) Then threshold = 0
    If Len(compareType) > 0 Then compareType = "gteq"

    'set the backcolor to white when the value reaches a threshold >= 0, checking for NULL and empty values
    SetControlBackcolor ctrl, RGB(255, 255, 255), True, True, threshold, compareType
   
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetControlHighlight[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddTallyValue
' Description:  Adds tally amount to control
' Assumptions:  -
' Parameters:   ctrl - control being changed (textbox)
'               tallyAmount - amount to add (integer - positive or negative)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 4/1/2016  - initial version
' ---------------------------------
Public Sub AddTallyValue(ctrl As TextBox, tallyAmount As Integer)
On Error GoTo Err_Handler
  
  'handle when the user keeps cursor in field & tallyAmount would drive the value to < 0 (negative)
  If (ctrl.Value + tallyAmount < 0) Or (IsNull(ctrl.Value) And tallyAmount < 0) Then GoTo Exit_Handler
  
  If tallyAmount = 0 Then ctrl.Value = 0
  
  Select Case ctrl.Name
    Case "SeedTotal"
        ctrl.Value = Nz(ctrl.Value, 0) + tallyAmount
  End Select
  
  'return focus
  ctrl.SetFocus
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddTallyValue[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          DisableTallyButtons
' Description:  Disable tally buttons on control
' Assumptions:  -
' Parameters:   frm - form where tally buttons are being changed (form)
'               lookFor - common part of tally button name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 4/1/2016  - initial version
' ---------------------------------
Public Sub DisableTallyButtons(frm As Form, lookFor As String)
On Error GoTo Err_Handler
  
  Dim ctrl As Control
  
  For Each ctrl In frm.Controls
  
    If Len(ctrl.Name) > Len(Replace(ctrl.Name, lookFor, "")) Then
    
            ctrl.Enabled = False
    
    End If
  
  Next
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableTallyButtons[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateForm
' Description:  Populate a form using a specific record for edits
' Assumptions:  -
' Parameters:   frm - form to populate (form)
'               ID - identifier for record to populate from (long)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 6/2/2016 - moved from forms (EventsList, TaglineList)
'   BLC - 8/8/2016 - revised to use default table name
'   BLC - 8/29/2016 - adjusted for Contact form (requires both Contact, Contact_Access data)
'                     using usys_temp_qdf & adjusting ID to Contact_ID in final SQL
'   BLC - 10/24/2016 - added ModWentworth form
'   BLC - 1/12/2017 - code cleanup
'   BLC - 2/14/2017 - added Task form
' --------------------------------------------------------------------
'   BLC - 3/23/2017 - adapted version for Upland db
' --------------------------------------------------------------------
'   BLC - 9/18/2017 - added back in from big rivers: Location, ModWentworth, SetDatasheetDefaults,
'                     Site, Tagline, Task, Transducer, VegPlot, VegTransect
'   BLC - 9/29/2017 - added Logger case
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
'   BLC - 10/16/2017 - adjusted Contact to include IsNPS
'   BLC - 10/18/2017 - added Comment case
'   BLC - 10/19/2017 - added Location toggle
'   BLC - 10/30/2017 - add Location cbxCollectionSourceID setting
'   BLC - 10/31/2017 - added ReplicatePlot, CalibrationPlot (VegPlot)
'   BLC - 11/3/2017 - update Location case
'   BLC - 11/9/2017 - update VegPlot case, checkboxes & toggles; Transducer case distances
'   BLC - 11/11/2017 - update VegPlot case
'   BLC - 12/5/2017 - add VegPlot BeaverBrowse
'   BLC - 12/8/2017 - update VegPlot case
'   BLC - 12/27/2017 - update VegPlot to set combobox values
' ---------------------------------
Public Sub PopulateForm(frm As Form, ID As Long)
On Error GoTo Err_Handler
    Dim strSQL As String, strTable As String

    With frm
        'default
        strTable = .Name
        
        'find the form & populate its controls from the ID
        Select Case .Name
            Case "Comment"
                strTable = "AppComment"
                .Controls("tbxComment").ControlSource = "Comment"
                .Controls("tbxID").ControlSource = "ID"
            Case "Contact"
                'requires Contact & Contact_Access data
                Dim qdf As DAO.QueryDef
                CurrDb.QueryDefs("usys_temp_qdf").SQL = GetTemplate("s_contact_access")
                
                strTable = "usys_temp_qdf"
                'set form fields to record fields as datasource
                'contact data
                .Controls("tbxID").ControlSource = "c.ID"
                .Controls("tbxFirst").ControlSource = "FirstName"
                .Controls("tbxMI").ControlSource = "MiddleInitial"
                .Controls("tbxLast").ControlSource = "LastName"
                .Controls("tbxEmail").ControlSource = "Email"
                .Controls("tbxUsername").ControlSource = "Username"
                .Controls("tbxOrganization").ControlSource = "Organization"
                .Controls("tbxPhone").ControlSource = "WorkPhone"
                .Controls("tbxPosition").ControlSource = "PositionTitle"
                .Controls("tbxExtension").ControlSource = "WorkExtension"
                .Controls("tglIsNPS").ControlSource = "IsNPS" 'IIf("IsNPS" = 1, True, False)
                'contact_access data
                .Controls("cbxUserRole").ControlSource = "Access_ID"
            Case "Events"
                strTable = "Event"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxSite").ControlSource = "Site_ID"
                .Controls("cbxLocation").ControlSource = "Location_ID"
                .Controls("tbxStartDate").ControlSource = "StartDate"
            Case "Feature"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxFeature").ControlSource = "Feature"
                '.Controls("cbxLocation").ControlSource = ""
            Case "Location"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxName").ControlSource = "CollectionSourceName"
                .Controls("tbxDistance").ControlSource = "HeadToOrientDistance_m"
                .Controls("tbxBearing").ControlSource = "HeadToOrientBearing"
                .Controls("tbxNotes").ControlSource = "LocationNotes"
                .Controls("optgLocationType").ControlSource = "LocationType"
                
                With .Controls("cbxCollectionSourceID")

                    Select Case frm.LocationType
                        Case "F" 'Feature
                            Set .Recordset = GetRecords("s_feature_by_site")
                            .BoundColumn = 2
                            .ColumnCount = 2
                            .ColumnWidths = "0;1in"
                        Case "T" 'Transect
                            Set .Recordset = GetRecords("s_transect_numbers")
                                .BoundColumn = 1
                                .ColumnCount = 2
                                .ColumnWidths = "0;1in"
                        Case "P" 'Plot
                            Set .Recordset = GetRecords("s_plot_numbers")
                                .BoundColumn = 1
                                .ColumnCount = 1
                                .ColumnWidths = "1in"
                        Case ""  'default
                    End Select
                                
                    'select the value
'                   .Controls("cbxCollectionSourceID") = "LocationType"
'                    .Controls("cbxCollectionSourceID").SelText = frm.Controls("list").Form.Controls("tbxLocTypeID")
                    .Controls("cbxCollectionSourceID").SelText = frm.Controls("list").Controls("tbxLocTypeID") 'Form.Controls("tbxLocTypeID")
                End With
                
                'unhide fields
                .Controls("cbxCollectionSourceID").visible = True
                .Controls("lblCollectionSourceID").visible = True
                
            Case "Logger"
                'set form fields to record fields as datasource
                .Controls("cbxSite").ControlSource = "Site_ID"
                .Controls("cbxLoggerType").ControlSource = "SensorType"
                .Controls("tbxAbbreviation").ControlSource = "SensorNumber"
                .Controls("tbxSampleOrder").ControlSource = "SamplingOrder"
            Case "ModWentworth"
                strTable = "ModWentworthScale"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxClass").ControlSource = "Label"
                .Controls("tbxCode").ControlSource = "Code"
                .Controls("tbxSizeRange").ControlSource = "DiameterRange_mm"
                .Controls("tbxEffectiveDate").ControlSource = "ActiveYear"
                .Controls("tbxRetireDate").ControlSource = "RetireYear"
            Case "SetDatasheetDefaults"
                strTable = "tsys_Datasheet_Defaults"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxSpecies").ControlSource = "SpeciesRows"
                .Controls("tbxBlanks").ControlSource = "BlankRows"
            Case "Site"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxSiteCode").ControlSource = "SiteCode"
                .Controls("tbxSiteName").ControlSource = "SiteName"
                .Controls("tbxDescription").ControlSource = "SiteDescription"
                .Controls("tbxSiteDirections").ControlSource = "SiteDirections"
            Case "Tagline"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxCause").ControlSource = "HeightType"
                .Controls("tbxDistance").ControlSource = "LineDistance_m"
                .Controls("tbxHeight").ControlSource = "Height_cm"
            Case "Task"
                'set form fields to record fields as datasource
                .Controls("tbxType").ControlSource = "TaskType"
                .Controls("tbxTypeID").ControlSource = "TaskType_ID"
                '.Controls("lblTaskContext").Caption = .Controls("tbxType") & " (" & .Controls("tbxTypeID") & ")"
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxStatus").ControlSource = "Status_ID"
                .Controls("cbxPriority").ControlSource = "Priority_ID"
                .Controls("tbxTask").ControlSource = "Task"
                .Controls("cbxRequestedBy").ControlSource = "RequestedBy_ID"
                .Controls("tbxRequestDate").ControlSource = "RequestDate"
            Case "Transducer"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxTiming").ControlSource = "Timing"
                .Controls("cbxTransducer").ControlSource = "TransducerNumber"
                .Controls("tbxSerialNo").ControlSource = "SerialNumber"
                .Controls("tbxSampleDate").ControlSource = "ActionDate"
                .Controls("tbxSampleTime").ControlSource = "ActionTime"
                .Controls("chkSurveyed").ControlSource = IIf("IsSurveyed" = 1, True, False)
                .Controls("tbxRefToWaterline").ControlSources = "RefToWaterline"
                .Controls("tbxRefToEyebolt").ControlSources = "RefToEyebolt"
                .Controls("tbxEyeboltToWaterline").ControlSources = "EyeboltToWaterline"
                .Controls("tbxEyeboltToScribeline").ControlSources = "EyeboltToScribeline"
            Case "Unknown"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxUnknownCode").ControlSource = "UnknownCode"
                .Controls("optgPlantType").ControlSource = "PlantType"
                .Controls("tbxDescription").ControlSource = "Description"
                .Controls("tbxFeature").ControlSource = "SalientFeature"
                .Controls("tbxLeafType").ControlSource = "LeafType"
                .Controls("tbxLeafMargin").ControlSource = "LeafMargin"
                .Controls("tbxLeafCharacter").ControlSource = "LeafCharacter"
                .Controls("tbxStemCharacter").ControlSource = "StemCharacter"
                .Controls("tbxFlowerCharacter").ControlSource = "FlowerCharcter"
                .Controls("tbxGeneralCharacter").ControlSource = "GeneralCharacter"
                .Controls("optgForbGrassType").ControlSource = "ForbGrassType"
                .Controls("optgPerennialGrassType").ControlSource = "PerennialGrassType"
                .Controls("tbxBestGuess").ControlSource = "BestGuess"
                .Controls("chkHasPhotos").ControlSource = "HasPhotos"
                .Controls("chkCollected").ControlSource = "Collected"
                .Controls("tbxCollectionMethod").ControlSource = "CollectionMethod"
                .Controls("cbxLocationID").ControlSource = "Location_ID"
                .Controls("cbxCollectedByID").ControlSource = "CollectedBy_ID"
            Case "VegPlot"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                '.Controls("cbxEvent").ControlSource = "Event_ID"
                .Controls("tbxEventID").ControlSource = "Event_ID"
                '.Controls("cbxTransect").ControlSource = "VegTransect_ID"
                .Controls("tbxTransectID").ControlSource = "VegTransect_ID"
                .Controls("tbxMSSID").ControlSource = "ModalSedimentSize_ID"
                '.Controls("cbxModalSedSize").ControlSource = "ModalSedSize"
                Dim i As Integer
                With .Controls("cbxModalSedSize")
                    For i = 0 To .ListCount - 1
                        If .Column(0, i) = frm.Controls("tbxMSSID") Then
                            .Value = .ItemData(i)
                            Exit For
                        End If
                    Next
                End With
                .Controls("tbxNumber").ControlSource = "PlotNumber"
                .Controls("tbxDistance").ControlSource = "PlotDistance_m"
                .Controls("tbxPctWCC").ControlSource = "WoodyCanopyPctCover"
                .Controls("tbxPctURC").ControlSource = "UnderstoryRootedPctCover"
                .Controls("tbxPctARC").ControlSource = "AllRootedPctCover"
                .Controls("tbxPctFines").ControlSource = "PctFines"
                .Controls("tbxPctWater").ControlSource = "PctWater"
                .Controls("tbxPctLitter").ControlSource = "PctLitter"
                .Controls("tbxPctWoodyDebris").ControlSource = "PctWoodyDebris"
                .Controls("tbxPctStandingDead").ControlSource = "PctStandingDead"
                .Controls("tbxPctMSS").ControlSource = "PctModalSedimentSize"
                .Controls("tbxPctFA").ControlSource = "PctFilamentousAlgae"
                .Controls("tbxPctSocialTrails").ControlSource = "PctSocialTrails"
                .Controls("tbxPlotDensity").ControlSource = "PlotDensity"
                .Controls("tglNoCanopyVeg").ControlSource = "NoCanopyVeg" 'IIf("NoCanopyVeg" = 1, True, False)
                .Controls("tglNoRootedVeg").ControlSource = "NoRootedVeg" 'IIf("NoRootedVeg" = 1, True, False)
                .Controls("tglNoIndicatorSpecies").ControlSource = "NoIndicatorSpecies" 'IIf("NoIndicatorSpecies" = 1, True, False)
                .Controls("tglBeaverBrowse").ControlSource = "BeaverBrowse"
                '.Controls("tglHasSocialTrails").ControlSource = "HasSocialTrails" 'IIf("HasSocialTrails" = 1, True, False)
                .Controls("chkCalibrationPlot").ControlSource = "CalibrationPlot" 'IIf("CalibrationPlot" = 1, True, False)
                .Controls("chkReplicatePlot").ControlSource = "ReplicatePlot" 'IIf("ReplicatePlot" = 1, True, False)
            Case "VegTransect"

            Case "VegWalk"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxWalkStartDate").ControlSource = "StartDate"

        End Select

        'clear msg/msg icon if present
        If ControlExists("lblMsgIcon", frm) Then .Controls("lblMsgIcon").Caption = ""
        If ControlExists("lblMsg", frm) Then .Controls("lblMsg").Caption = ""
        
        strSQL = GetTemplate("s_form_edit", "tbl" & PARAM_SEPARATOR & strTable & "|id" & PARAM_SEPARATOR & ID)
        
        'alter to retrieve proper ID
        Select Case .Name
            Case "Contact"
                strSQL = Replace(strSQL, " ID = ", " c.ID = ")
        End Select
        
        .RecordSource = strSQL
        
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateForm[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateCSVFields
' Description:  CSV field combobox populating actions
' Assumptions:  Control OnChange event = PopulateCSVFields([Screen].[ActiveControl])
'               where Screen.ActiveControl passes in the proper combobox
' Parameters:   ctrl - control to populate (control)
' Returns:      -
' Throws:       none
' References:
'   Jeremy Cook, September 13, 2013
'   http://stackoverflow.com/questions/8787979/how-do-i-reference-the-current-form-in-an-expression-in-microsoft-access
' Source/date:  Bonnie Campbell, October 6, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/8/2016 - initial version
'   BLC - 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
' ---------------------------------
'Public Sub PopulateCombobox(ctrl As ComboBox)
Public Function PopulateCSVFields(ctrl As Control) 'frm As Form) 'strName As String) 'ByRef ctrl As ComboBox)
On Error GoTo Err_Handler
    
    'Dim ctrl As ComboBox
    
    'Set ctrl = Forms("ImportMap").Controls("listCSV").Form.Controls("cbxColumnName2")
    
'    Set ctrl = Me.ActiveControl
    
'    'set displayed title
'    lblTitle.Caption = "CSV fields"
    
    'retrieve field info
    Dim aryFieldInfo() As Variant 'string
    
    aryFieldInfo = FetchDbTableFieldInfo("usys_temp_csv")
    
    'clear table
    ClearTable "usys_temp_rs2"
    
    'populate w/ table data
    Dim rs2 As DAO.Recordset
    Dim aryRecord() As String
    Dim i As Integer
    
    Set rs2 = CurrDb.OpenRecordset("usys_temp_rs2", dbOpenDynaset)
    
    'add the "None" value
    rs2.AddNew
    rs2.Fields(0) = "None"
    rs2.Update
    
    For i = 0 To UBound(aryFieldInfo)
    
        'create new record
        rs2.AddNew
        
        aryRecord = Split(aryFieldInfo(i), "|")
        
        'rs!Column = aryRecord(0)
        rs2.Fields(0) = aryRecord(0)
    
        'add the new record
        rs2.Update
        
    Next
    
    Set ctrl.Recordset = rs2 '<--ERROR #5302
    
    Debug.Print "mod_App_UI PopulateCSVFields rs2 count = " & rs2.RecordCount
    
'    Dim strControl As String
'
'Debug.Print Me.NumColumns
'
'    'expose & populate the proper # of dropdowns
'    For i = 1 To Me.NumColumns 'CInt(Me.Records.RecordCount)
'        strControl = "cbxColumnName" & i
'Debug.Print strControl
'
''FIX HERE!
'        If i = 30 Then
'            Debug.Print "30"
'        End If
'
'        Me.Controls(strControl).Visible = True
''        Set Me.Controls(strControl).Recordset = rs2 '<--ERROR #5302
'        'Me.Controls(strControl).AddItem item:="None", index:=0
'
'        'set "None" to red --> Conditional formmating = "None"
'
'        'requery to refresh displayed controls
'        Me.Controls(strControl).Requery
'Debug.Print Me.Controls(strControl).ListRows
'    Next
'
'    If Me.NumColumns > 0 Then
'        'set detail to proper height
'        Me.Detail.Height = Me.Controls(strControl).Height * Me.NumColumns 'Me.Records.RecordCount
'    End If
'
''    Set Me.Recordset = rs
'
''    Set cbxColumnName.Recordset = rs2
'
'    'set the # of repeats of the cbx
''    Set Me.Recordset = rs

Exit_Handler:
    'cleanup
    Set rs2 = Nothing
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateCSVFields[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          DisplayIcons
' Description:  Prepare icon set for reports
' Assumptions:  -
' Parameters:   icons - icons to display (delimited string)
'               delimiter - character splitting icons (string)
' Returns:      icon display translating icons field delimited string to display (string)
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, August 24, 2016 - for NCPN tools
' Revisions:
'   BLC - 8/24/2016  - initial version
' ---------------------------------
Public Function DisplayIcons(icons As String, delimiter As String)
On Error GoTo Err_Handler

    Dim dDocIcons As Dictionary
    Dim ary() As String
    Dim strIcons As String
    Dim i As Integer
    
    Set dDocIcons = CreateObject("scripting.dictionary")
    
    'setup dictionary
    dDocIcons.Add "uDocument", uDocument
    dDocIcons.Add "uPDF", uNotepad ' uPDF
    
    'default
    strIcons = ""
    
    ary = Split(icons, delimiter)
    
    For i = LBound(ary) To UBound(ary)
    
        strIcons = strIcons & StringFromCodepoint(dDocIcons(ary(i))) & Space(2)
    
    Next
    
    DisplayIcons = strIcons
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisplayIcons[fw_mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          FilterListForm
' Description:  form filter click actions
' Assumptions:  -
' Parameters:   frm - form to filter (form)
'               ctrl - control to filter by (control)
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Filter-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Public Sub FilterListForm(frm As Form, ctrl As Control)
On Error GoTo Err_Handler

    Dim strFilter As String
    
    'default
    strFilter = ""
    
    'set Filter field
    Select Case Replace(ctrl.Name, "lbl", "")
        Case "HdrID"
            strFilter = "ID"
        Case "Version"
            strFilter = "Version"
        Case "Template"
            strFilter = "Template"
        Case "Remarks"
            strFilter = "Remarks"
        Case "EffectiveDate"
            strFilter = "EffectiveDate"
        Case ""
    End Select

    'set the Filter
    If InStr(frm.OrderBy, strFilter) = 0 Then
        frm.OrderBy = strFilter
    ElseIf Right(frm.OrderBy, 4) = "Desc" Then
        frm.OrderBy = strFilter
    Else
        frm.OrderBy = strFilter & " Desc"
    End If
    
    frm.OrderByOn = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterListForm[fw_mod_App_UI form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     GetContext
' Description:  set the context based on tempvars
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, October 14, 2015 - for NCPN tools
' Revisions:
'   BLC - 10/14/2016  - initial version
' ---------------------------------
Public Function GetContext() As String
On Error GoTo Err_Handler

    Dim strContext
    
    strContext = Nz(TempVars("ParkCode"), "") & Space(2) & ">" & Space(2) & _
                 Nz(TempVars("River"), "-") & Space(2) & ">" & Space(2) & _
                 Nz(TempVars("SiteCode"), "-")

    Select Case Nz(TempVars("ParkCode"), "")
    
        Case "BLCA"
            'add feature
            strContext = strContext & Space(2) & ">" & Space(2) & _
                 Nz(TempVars("Feature"), "-")
        Case "CANY"
            'site level
        Case "DINO"
            'site level
    End Select
    
    GetContext = strContext

Exit_Sub:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetContext[fw_mod_App_UI])"
    End Select
    Resume Exit_Sub
End Function