Attribute VB_Name = "mod_Forms"
Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Forms
' Level:        Framework module
' Version:      1.15
' Description:  generic form functions & procedures
'
' Source/date:  Bonnie Campbell, 2/19/2015
' Revisions:    BLC - 2/19/2015 - 1.00 - initial version
'               BLC - 5/18/2015 - 1.01 - fixed ClearFields documentation
'               BLC - 6/9/2015  - 1.02 - added CloseFormsReports()
'               BLC - 6/30/2015 - 1.03 - shifted to mod_UI: ChangeBackColor
'                                        shifted from mod_UI: FormIsOpen, FormIsLoaded, SwitchboardIsOpen
'                                        shifted to mod_App_UI: ClearFields
'               BLC - 6/1/2016  - 1.04 - added SetFormOpacity(), CaptureEscapeKey(), constants & functions
'                                        from Uplands mod_App_UI
'               BLC - 6/24/2016 - 1.05 - added ToggleForm(), replaced Exit_Function > Exit_Handler
'               BLC - 7/1/2016  - 1.06 - added font weight constants
'               BLC - 7/28/2016 - 1.07 - added clearing lblMsg caption for ClearForm()
'               BLC - 2/22/2017 - 1.08 - added notes to ToggleForm()
' --------------------------------------------------------------------
'               BLC, 3/22/2017          added to Upland db
' --------------------------------------------------------------------
'               BLC, 9/14/2017  - 1.09 - added: notes re: IsLoaded function
'                                               from mod_Utilities: FormAssist()
'                                               documentation & error handling
'               BLC, 10/5/2017  - 1.10 - update documentation
'               BLC, 10/6/2017  - 1.11 - added from mod_UI:
'                                        SetWindowSize(), PopulateSubformControl()
'                                        RepaintParentForm(), ChangeBackColor()
'                                        ResetHeaders(), ShowControls(),
'                                        ControlExists(), AddFormControl()
'               BLC - 11/10/2017 - 1.12 - add control existance checks
'               BLC - 12/14/2017 - 1.13 - add checkbox and toggle button
'               BLC - 12/27/2017 - 1.14 - update to avoid black box inside checkboxes (ClearForm)
'               BLC - 1/10/2018  - 1.15 - added list control existance check (ClearForm)
' =================================

'=================================================================
'  References
'=================================================================
' ---------------------------------
'  Access Control Types
' ---------------------------------
' dbtech1, March 13, 2008
' http://www.utteraccess.com/forum/control-type-vba-t1609220.html
' 126 - acAttachment         119 - acCustomControl  114 - acObjectFrame    101 - acRectangle
' 108 - acBoundObjectFrame   103 - acImage          105 - acOptionButton   112 - acSubform
' 106 - acCheckBox           100 - acLabel          107 - acOptionGroup    123 - acTabCtl
' 111 - acComboBox           102 - acLine           124 - acPage           109 - acTextBox
' 104 - acCommandButton      110 - acListBox        118 - acPageBreak      122 - acToggleButton
' ---------------------------------

' ---------------------------------
'  Access Form Sections
' ---------------------------------
'   acDetail        0   (Default) Detail section    acGroupLevel1Footer 6   Group-level 1 footer (reports only)
'   acFooter        2   Form or report footer       acGroupLevel1Header 5   Group-level 1 header (reports only)
'   acHeader        1   Form or report header       acGroupLevel2Footer 8   Group-level 2 footer (reports only)
'   acPageFooter    4   Page footer                 acGroupLevel2Header 7   Group-level 2 header (reports only)
'   acPageHeader    3   Page header
' ---------------------------------

' ---------------------------------
'  Access Backstyle Property
' ---------------------------------
'  Transparent  0           Normal  1
' ---------------------------------

' ---------------------------------
'  Access FontWeight Property
' ---------------------------------
'   Thin    100         Extra Light         200
'   Light   300         (Default) Normal    400
'   Medium  500         Semi-Bold           600
'   Bold    700         Extra Bold          800
'   Heavy   900
' ---------------------------------

'=================================================================
'  Constants
'=================================================================

' -- font weight constants --
Public Const wtThin = 100
Public Const wtExtraLight = 200
Public Const wtLight = 300
Public Const wtNormal = 400
Public Const wtMedium = 500
Public Const wtSemiBold = 600
Public Const wtBold = 700
Public Const wtExtraBold = 800
Public Const wtHeavy = 900

'-- text align constants --
Public Const aGeneral = 0
Public Const aLeft = 1
Public Const aCenter = 2
Public Const aRight = 3
Public Const aDistribute = 4

'=================================================================
'  Declarations
'=================================================================
Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As _
     Integer
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As _
     Integer

' -- Constants --
Private Const LWA_ALPHA     As Long = &H2
Private Const GWL_EXSTYLE   As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000

Public Const CTRL_DEFAULT_BACKCOLOR  As Long = 65535  'RGB(255, 255, 0) highlight yellow

' -- Values --
Public NoData As Scripting.Dictionary

' -- Functions --
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal crKey As Long, _
   ByVal bAlpha As Byte, _
   ByVal dwFlags As Long) As Long

Public RefSub As String 'referring subroutine

'=================================================================
'  Properties
'=================================================================


'=================================================================
'  Subroutines & Functions
'=================================================================

' ---------------------------------
'  Open/Close/Loaded
' ---------------------------------

' ---------------------------------
' FUNCTION:     CloseFormsReports
' Description:  close forms, reports
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Susan Harkins, July 21, 2009
'   http://www.techrepublic.com/blog/microsoft-office/automatically-close-all-the-open-forms-and-reports-in-an-access-database/
' Adapted:      Bonnie Campbell, June 9, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/9/2015  - initial version
'   BLC - 10/5/2017 - change from function to subroutine
' ---------------------------------
Public Sub CloseFormsReports()
On Error GoTo Err_Handler

    'Close all open forms
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).Name
    Loop
    
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).Name
    Loop

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CloseFormsReports[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     FormIsOpen
' Description:  Indicates whether or not the specific form is open in form view
' Parameters:   none
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/5/2006 as fxnSwitchboardIsOpen
' Adapted:      Bonnie Campbell, 4/30/2015 for NCPN tools
' Revisions:    BLC, 4/30/2015 - initial version
' ---------------------------------
Public Function FormIsOpen(strFormName As String) As Boolean
    On Error GoTo Err_Handler

    Dim frm As Form

    FormIsOpen = False    ' Default in case of error
 
    'search for form in Forms collection (all open forms)
    For Each frm In Forms
      If frm.Name = strFormName Then
        'check form is in Form view: 0 - Design View, 1 - Form View, 2 - Datasheet View
        If frm.CurrentView = 1 Then
            FormIsOpen = True
            'Exit Function
        End If
      End If
    Next

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormIsOpen[mod_Forms])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SwitchboardIsOpen
' Description:  Indicates whether or not the switchboard form is open in form view
' Parameters:   none
' Returns:      True or False
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 5/5/2006
' Revisions:    JRB, 5/5/2006 - initial version
'               BLC, 4/30/2015  - moved to mod_Db framework module from mod_Custom_Functions
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 6/12/2016 - revised to use AppSettings SWITCHBOARD value
' ---------------------------------
Public Function SwitchboardIsOpen() As Boolean
    On Error GoTo Err_Handler

    SwitchboardIsOpen = False    ' Default in case of error

    'check for switchboard in all open forms ( AllForms.IsLoaded() )
    If CurrentProject.AllForms(SWITCHBOARD).IsLoaded = True Then
        If CurrentProject.AllForms(SWITCHBOARD).CurrentView = 1 Then
            SwitchboardIsOpen = True
        End If
    End If

Exit_Handler:
   Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SwitchboardIsOpen[mod_Forms])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     FormIsLoaded
' Description:  Returns whether the specified form is loaded in Form or Datasheet view
' Parameters:   strFormName - string for the name of the form to check
' Returns:      True if the specified form is open in Form view or Datasheet view
' Throws:       none
' References:   none
' Source/date:  From Northwind sample database, date unknown
' Revisions:    John R. Boetsch, 6/17/2009 - error trapping, documentation
'               BLC, 4/30/2015 - moved from mod_Utilities to mod_UI
'               BLC, 5/18/2015 - renamed, removed fxn prefix
'               BLC, 9/14/2017 - removed mod_Utilities IsLoaded() is the same function
' ---------------------------------
Public Function FormIsLoaded(ByVal strFormName As String) As Integer
    On Error GoTo Err_Handler
 
    ' These variables are used to test the return values of the SysCmd function
    '  and the CurrentView property of the requested form.
    Const cObjStateClosed = 0
    Const cDesignView = 0

    ' Use the SysCmd function to check the current state of the requested form.
    '  Possible states: not open or nonexistent, open, new, or changed but not saved
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> cObjStateClosed Then
        ' Checks for the current view of the requested form, assuming the previous statement
        '   found it to be open ... return True if open and not in design view
        If Forms(strFormName).CurrentView <> cDesignView Then
            FormIsLoaded = True
        End If
    End If
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormIsLoaded[mod_Forms])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
'  Form Help
' ---------------------------------
' ---------------------------------
' Sub:          FormAssist
' Description:  Responds to OnAction on custom menu or toolbar help command
'               Checks for an active form, then looks for a "FormHelp" handler
'               subroutine on that form
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' Usage:        -
' References:   -
' Source/date:  NCPN unknown
' Adapted:      Bonnie Campbell, September 14, 2017 - for NCPN tools
' Revisions:
'   Unknown - unknown - initial version
'   BLC - 9/14/2017 - moved from mod_Utilities to mod_Forms, error handling &
'                     documentation added
' ---------------------------------
Public Sub FormAssist()
On Error GoTo Err_Handler
    
    Dim frm As Form

    ' Try to locate a form that has the focus
    Set frm = Screen.ActiveForm
    If Err <> 0 Then
        ' Error means no active form,
        '  so open standard Office Assistant
        Application.Assistant.Help
        'Exit Function
        GoTo Exit_Handler
    End If
    
    ' No error, so try to call the FormHelp
    '   method of the active form
    frm.FormHelp
    If Err <> 0 Then
        ' Error means no FormHelp method for
        '  the current form,
        '  so open standard Office Assistant
        Application.Assistant.Help
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FormAssist[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Entire Form
' ---------------------------------
' ---------------------------------
' SUB:          ToggleForm
' Description:  Minimizes, maximizes, or restores form display
' Assumptions:  Form is not opened if it is not already opened
'               In part this is to avoid endless loops with forms
'               like PreSplash which call routines that shouldn't be re-called.
' Note:         -
' Parameters:   strForm - form to change (string)
'               Sizing - how to change display (integer) -1 = minimize, 0 = normal/restore, 1 = maximize
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 24, 2016  - for NCPN tools
' Revisions:    BLC, 6/24/2016 - initial version
'               BLC, 2/22/2017 - added documentation notes @ opening form
' ---------------------------------
Public Sub ToggleForm(strForm As String, Sizing As Integer)
On Error GoTo Err_Handler
    
    'ensure form is open, if not -> exit
    If Not FormIsOpen(strForm) Then GoTo Exit_Handler
    
    Forms(strForm).SetFocus
    
    Select Case Sizing
        Case -1 'minimize
            DoCmd.Minimize
        Case 0 'restore
            DoCmd.Restore
        Case 1 'maximize
            DoCmd.Maximize
    End Select
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleForm[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ClearForm
' Description:  Clear form fields
' Assumptions:  Form setup is similar to big rivers contact form w/ data entry
'               above and list below
' Parameters:   frm - form
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 23, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/23/2016 - initial version
'   BLC - 6/27/2016 - shifted to mod_Forms from big rivers forms
'   BLC - 7/28/2016 - added clearing lblMsg caption
'   BLC - 8/30/2016 - added RefSub to identify form subs called by ClearForm
'   BLC - 11/10/2017 - add control existance checks
'   BLC - 12/14/2017 - add checkbox and toggle button
'   BLC - 12/27/2017 - update to avoid black box inside checkboxes
'   BLC - 1/10/2018 - added list control existance check
' ---------------------------------
Public Sub ClearForm(ByRef frm As Form)
On Error GoTo Err_Handler
    
    'set global
    RefSub = "ClearForm"
    
    With frm
    
        'clear recordsource
        .RecordSource = ""
        
        'clear values so they no longer look for original control sources
        Dim ctrl As Control
        
        'clear the control sources to clear the textboxes
        For Each ctrl In frm.Controls
            Select Case ctrl.ControlType
                Case acTextBox
                    ctrl.ControlSource = ""
                    ctrl.Value = ""
                Case acComboBox
                    'ctrl.Value = "" '<< error: 2448 can't assign value to object
                    'ctrl.Value = Null '<< error: 2448 can't assign value to object
                    'ctrl.ItemData (0)
                    ' Johanness, October 12, 2012
                    ' http://stackoverflow.com/questions/12697427/vba-clear-selections-of-a-combobox
                Case acCheckBox
                    ctrl.Value = 0 'false vs "" to avoid black box inside checkbox
                Case acToggleButton
                    ToggleCaption ctrl, False
            End Select
        Next
        
        'set values if controls exist
        If ControlExists("tbxIcon", frm) Then
            .Controls("tbxIcon") = StringFromCodepoint(uBullet)
            .Controls("tbxIcon").ForeColor = lngRed
        End If
        
        If ControlExists("tbxID", frm) Then .Controls("tbxID") = 0
        
        'assume MsgIcon & Msg both are either on form or not
        If ControlExists("lblMsgIcon", frm) Then
            .Controls("lblMsgIcon").Caption = ""
            .Controls("lblMsg").Caption = ""
            .Controls("lblMsgIcon").ForeColor = lngRobinEgg
        End If
        
        If ControlExists("btnSave", frm) Then .Controls("btnSave").Enabled = False
                
        If ControlExists("list", frm) Then .list.Requery
        
        .Requery
    
    End With
    
Exit_Handler:
    RefSub = ""
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearForm[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetWindowSize
' Description:  sets form size (width & height)
' Assumptions:  -
' Note:         dimensions are in twips (1 inch = 1440 twips)
' Parameters:   frm - form to set size for (form)
'               lngHeight - desired height (long)
'               lngWidth - desired width (long)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Hasup, February 26,2014
'   http://stackoverflow.com/questions/22021802/resize-form-in-ms-access-by-changing-detail-height
' Adapted:      Bonnie Campbell, May 27, 2015 - for NCPN tools
' Revisions:    BLC, 5/27/2015 - initial version
'               BLC, 10/6/2017 - moved from mod_UI to mod_Form
' ---------------------------------
Public Sub SetWindowSize(ByRef frm As Form, ByRef lngHeight As Long, ByRef lngWidth As Long)
On Error GoTo Err_Handler

'    If Me.WindowHeight = 4044 Then
'        lngHeight = 8000
'    Else
'        lngHeight = 4044
'    End If
    frm.Move frm.WindowLeft, Height:=lngHeight, Width:=lngWidth
    
Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetWindowSize[mod_Forms])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          SetFormOpacity
' Description:  Sets form opacity
' Assumptions:  place in forms module mod_Form for protocols which utilize that module
' Parameters:   frm - form to prepare
'               sngOpacity - opacity of the form (single)
'               TColor - color for the form display (long)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Thenman, September 24, 2009
' http://www.access-programmers.co.uk/forums/showthread.php?t=154907
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
'   BLC, 6/1/2016  - moved to mod_Forms from mod_App_UI (uplands)
' ---------------------------------
Public Sub SetFormOpacity(frm As Form, sngOpacity As Single, TColor As Long)
On Error GoTo Err_Handler

    Dim lngStyle As Long
    
    ' get the current window style, then set transparency
    lngStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
    SetWindowLong frm.hwnd, GWL_EXSTYLE, lngStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes frm.hwnd, TColor, (sngOpacity * 255), LWA_ALPHA
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetFormOpacity[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RepaintParentForm
' Description:  Repaints the control's parent(or grandparent or great grandparent...) form
' Parameters:   ctl - control whose parent form you're looking to repaint
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell August, 2014 - NCPN tools
' Adapted:      -
' Revisions:    BLC, 8/20/2014 - initial version
'               BLC, 4/30/2015 - moved from mod_Common_UI to mod_UI
' ---------------------------------
Public Sub RepaintParentForm(ctl As Control)
On Error GoTo Err_Handler:
Dim parentControl As Object

    Set parentControl = ctl.Parent

    Do Until parentControl Is Nothing

        If TypeName(parentControl.Name) = "String" Then
            'form? -> refresh the display
            If getAccessObjectType(parentControl.Name) = -32768 Then
                parentControl.Repaint
                Exit Do
            End If
            Set parentControl = parentControl.Parent
        Else
            'form? -> refresh the display
            If CurrentProject.AllForms(parentControl.Name).IsLoaded Then
                parentControl.Repaint
                Exit Do
            End If
        End If
    Loop

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RepaintParentForm[mod_Forms])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
'  Form Controls
' ---------------------------------

' ---------------------------------
' SUB:          PopulateSubformControl
' Description:  Populate a subform control with a specific form
'               Allows swapping of subform with context
' Parameters:   ctrl - subform control to populate
'               strSubFormName - name of the subform to use in the control
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, 5/1/2015 for NCPN tools
' Revisions:    BLC, 5/1/2015 - initial version
'               BLC, 10/6/2017 - moved from mod_UI to mod_Forms
' ---------------------------------
Public Sub PopulateSubformControl(ctrl As SubForm, strSubFormName As String)
    On Error GoTo Err_Handler

    ctrl.SourceObject = strSubFormName 'Forms(strSubFormName)

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateSubformControl[mod_Forms])"
    End Select
    Resume Exit_Procedure
End Sub

' ---------------------------------
' FUNCTION:     ControlExists
' Description:  determines if a control exists in a form
' Parameters:   ctlName - control to check for (string)
'               frm - form to check on (form)
' Returns:      boolean - true if control exists, false if not
' Throws:       none
' References:   none
' Source/date:
'   VBslammer, March 22, 2005
'   http://www.tek-tips.com/viewthread.cfm?qid=1029435
'   Mike Lyons September 21, 2007
'   http://www.utteraccess.com/forum/Control-Exist-Form-t1505884.html
' Adapted:      Bonnie Campbell, May 15, 2015 - for NCPN tools
' Revisions:    BLC, 5/12/2015 - initial version
'               BLC, 9/1/2016  - added false path, updated documentation
'               BLC, 10/6/2017 - moved from mod_UI to mod_Forms
' ---------------------------------
Public Function ControlExists(ByRef ctlName As String, ByRef frm As Form) As Boolean
On Error GoTo Err_Handler
  Dim ctl As Control
  
  For Each ctl In frm.Controls
    If ctl.Name = ctlName Then
      ControlExists = True
      GoTo Exit_Handler
    End If
  Next ctl
  
  'doesn't exist
  ControlExists = False
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ControlExists[mod_Forms])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          AddControl
' Description:  add control to form
' Assumptions:  -
' Parameters:   frm - form (object)
'               ctrl - control (object)
'               ctrlName - name of control (string)
'               xPos - horizontal position (twips)
'               yPos - vertical position (twips)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' meloncolly, October 27, 2006
' http://forums.aspfree.com/microsoft-access-help-18/add-controls-form-dynamically-139627.html
' https://msdn.microsoft.com/en-us/library/bb237827(office.12).aspx
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
'   BLC - 9/14/2017  - updated documentation from form_frmSpeciesSearch to mod_Forms
' ---------------------------------
Public Sub AddControl(frm As Form, ctrl As Control, ctrlName As String, _
                        xPos As Integer, yPos As Integer)
On Error GoTo Err_Handler

    ' Create ctrl
    Set ctrl = CreateControl(frm.Name, ctrl.ControlType, , "", "", xPos, yPos)
    
    ' Restore form
    DoCmd.Restore

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddControl[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddFormControl
' Description:  Adds a control to a form
' Assumptions:  -
' Parameters:   frm - form to add controls to (form)
'               ctlType - type of control to add (control)
'               ctlName - name of control to add (string)
'               ctlData - data for control (optional, variant)
' Returns:      -
' Throws:       none
' References:
'   Chip Pearson, unknown
'   http://www.ozgrid.com/Excel/free-training/ExcelVBA2/excelvba2lesson21.htm
' Source/date:  Bonnie Campbell, October 11, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/11/2016 - initial version
'   BLC - 10/6/2017 - moved from mod_UI to mod_Forms
' ---------------------------------
Public Sub AddFormControl(frmName As String, ctlType As Long, ctlName As String, Optional ctlData As Variant, _
                        Optional w As Integer, Optional h As Integer, _
                        Optional xPos As Integer, Optional yPos As Integer)
On Error GoTo Err_Handler
    
    'Dim progID As String
    Dim c As Control
    
    'progID = "Forms." & ctlType & ".1"
    
'    Set c = frm.Controls.Add(progID, ctlName)
    Set c = CreateControl(frmName, ctlType, acDetail)

    c.Name = ctlName
    
    If Not ctlData Is Nothing Then
        Set c.Recordset = ctlData
    End If
    
    'set dimensions & location
    If IsNumeric(w) Then c.Width = w
    If IsNumeric(h) Then c.Height = h
    If IsNumeric(xPos) Then c.Left = xPos
    If IsNumeric(yPos) Then c.Top = yPos

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddFormControl[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     ChangeBackColor
' Description:  change background color of control
' Assumptions:  -
' Parameters:   ctrl- control to change color
'               lngColor = color (long)
' Returns:      N/A
' Throws:       none
' References:   none
' Note:         MUST be a function vs. sub to be called w/in form event ( =ChangeBackColor(Me,lngYelLime) )
' Source/date:  Bonnie Campbell, March 4, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/4/2015  - initial version
' ---------------------------------
Public Function ChangeBackColor(ctrl As Control, lngColor As Long)
On Error GoTo Err_Handler

    ctrl.backcolor = lngColor
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeBackColor[mod_Forms])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          ResetHeaders
' Description:  reset header fields to their original backcolor
' Assumptions:  if only a subset of form controls are to be reset, these controls should have the same Tag property value
' Parameters:   frm - form to reset headers on
'               allCtrls - if all form controls should be reset (boolean) (true = reset all controls,
'                           false = reset one control [requires oCtrl to be populated])
'               ctrlTag - control's tag string if resetting only a subset of forms controls (string)
'               fontBold - whether text should be bold (boolean) (true = make font bold, false not bold),  (optional)
'               backstyle - if back control back color is normal or transparent (integer) (1-normal 0-transparent) (optional)
'               forecolor - text color (long) (optional)
'               backcolor - backgound color of control (long) (optional)
'               oCtrl - control to change, if only one control is to be changed (optional)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Fionnuala January 20, 2013
' http://stackoverflow.com/questions/3344649/how-to-loop-through-all-controls-in-a-form-including-controls-in-a-subform-ac
' Adapted:      Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015  - initial version
'   BLC - 10/6/2017  - moved from mod_UI to mod_Forms
' ---------------------------------
Public Sub ResetHeaders(frm As Form, _
                        allCtrls As Boolean, _
                        ctrlTag As String, _
                        Optional FontBold As Boolean = True, _
                        Optional backstyle As Integer = 1, _
                        Optional ForeColor As Long, _
                        Optional backcolor As Long, _
                        Optional oCtrl As Control)
On Error GoTo Err_Handler

Dim ctrl As Control

    If allCtrls = True Then
    
        'iterate through all form controls
        For Each ctrl In frm
            
            'check control type
             If ctrl.ControlType = acTextBox Or _
                ctrl.ControlType = acComboBox Or _
                ctrl.ControlType = acListBox Or _
                ctrl.ControlType = acLabel _
             Then
             
                'check tag
                If ctrl.Tag = ctrlTag Then
                    If varType(FontBold) = vbBoolean Then ctrl.FontBold = FontBold
                    If varType(backstyle) = vbInteger Then ctrl.backstyle = backstyle
                    If varType(backcolor) = vbLong Then ctrl.backcolor = backcolor
                    If varType(ForeColor) = vbLong Then ctrl.ForeColor = ForeColor
                End If
                
          End If
          
        Next
    Else
        'reset only oCtrl

        'check tag
        If oCtrl.Tag = ctrlTag Then
        
            'check control type
            If oCtrl.ControlType = acTextBox Or _
                oCtrl.ControlType = acComboBox Or _
                oCtrl.ControlType = acListBox Or _
                oCtrl.ControlType = acLabel _
            Then
          
                If varType(FontBold) = vbBoolean Then oCtrl.FontBold = FontBold
                If varType(backstyle) = vbInteger Then oCtrl.backstyle = backstyle
                If varType(backcolor) = vbLong Then oCtrl.backcolor = backcolor
                If varType(ForeColor) = vbLong Then oCtrl.ForeColor = ForeColor
             
            End If
            
        End If

    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ResetHeaders[mod_Forms])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ShowControls
' Description:  toggle control visibility
' Assumptions:  if only a subset of form controls are to be reset, these controls should have the same Tag property value
' Parameters:   frm - form controls are on
'               allCtrls - if all form controls should be reset (boolean) (true = reset all controls,
'                           false = reset one control [requires oCtrl to be populated])
'               ctrlTag - control's tag string if resetting only a subset of forms controls (string)
'               visibility - whether control should be visible or not (boolean) (true = make font bold, false not bold),  (optional)
'               oCtrl - control to change, if only one control is to be changed (optional)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Fionnuala January 20, 2013
' http://stackoverflow.com/questions/3344649/how-to-loop-through-all-controls-in-a-form-including-controls-in-a-subform-ac
' Adapted:      Bonnie Campbell, February 20, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/20/2015 - initial version
'   BLC - 6/30/2015 - update documentation
'   BLC - 10/6/2017 - moved from mod_UI to mod_Forms
' ---------------------------------
Public Sub ShowControls(frm As Form, _
                        allCtrls As Boolean, _
                        ctrlTag As String, _
                        visibility As Boolean, _
                        Optional oCtrl As Control)
On Error GoTo Err_Handler

Dim ctrl As Control

    If allCtrls = True Then
    
        'iterate through all form controls
        For Each ctrl In frm

            'check tag
            If ctrl.Tag = ctrlTag Then
                ctrl.Visible = visibility
            End If

        Next
    Else
        'reset only oCtrl

        'check tag
        If oCtrl.Tag = ctrlTag Then
                oCtrl.Visible = visibility
        End If

    End If

Exit_Sub:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ShowControls[mod_Forms])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
'  Control Response
' ---------------------------------

' ---------------------------------
' SUB:          ContinuousUpDown
' Description:  Respond to Up/Down in a continuous form by moving to next record
' Assumptions:  Active control's EnterKeyBehavior is OFF
' Usage:        Call ContinuousUpDown(Me, KeyCode)
' Parameters:   frm - form for key behavior
'               KeyCode - code for key being pressed (integer)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Allen Browne via Jeanette Cunningham, Apr 13, 2010
' http://www.pcreview.co.uk/threads/need-to-get-the-up-down-arrow-keys-working.3995845/
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015  - initial version
' ---------------------------------
Public Sub ContinuousUpDown(frm As Form, KeyCode As Integer)
On Error GoTo Err_Handler

    Dim strForm As String
    
    strForm = frm.Name
    
    'determine key being used
    Select Case KeyCode
        Case vbKeyUp
            If ContinuousUpDownOk Then
                
                'Save any edits
                If frm.Dirty Then
                    RunCommand acCmdSaveRecord
                End If
                
                'Go previous: error if already there.
                    RunCommand acCmdRecordsGoToPrevious
                KeyCode = 0 'Destroy the keystroke
            End If
    
    Case vbKeyDown
        If ContinuousUpDownOk Then
            
            'Save any edits
            If frm.Dirty Then
                frm.Dirty = False
            End If
            
            'Go to the next record, unless at a new record.
            If Not frm.NewRecord Then
                RunCommand acCmdRecordsGoToNext
            End If
            KeyCode = 0 'Destroy the keystroke
        End If
    End Select

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 2046, 2101, 2113, 3022, 2465 'Already at first record, or save
            'failed, or The value you entered isn't valid for this field.
            KeyCode = 0
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - ContinuousUpDown[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     ContinuousUpDownOk
' Description:  Suppress moving up/down a record in a continuous form if:
'                - control is not in the Detail section
'                - multi-line text box (vertical scrollbar or EnterKeyBehavior true)
' Assumptions:  Active control's EnterKeyBehavior is OFF
' Usage:        Called by ContinuousUpDown SUB
' Parameters:   N/A
' Returns:      boolean - true if moving up/down a record in continuous form is ok, false if not
' Throws:       none
' References:   none
' Source/date:
' Allen Browne via Jeanette Cunningham, Apr 13, 2010
' http://www.pcreview.co.uk/threads/need-to-get-the-up-down-arrow-keys-working.3995845/
' Adapted:      Bonnie Campbell, March 5, 2015 - for NCPN tools
' Revisions:
'   BLC - 3/5/2015  - initial version
' ---------------------------------
Private Function ContinuousUpDownOk() As Boolean
On Error GoTo Err_Handler
    Dim blnDontDoIt As Boolean
    Dim ctl As Control
    
    Set ctl = Screen.ActiveControl
    If ctl.Section = acDetail Then
        If TypeOf ctl Is TextBox Then
            blnDontDoIt = ((ctl.EnterKeyBehavior) Or (ctl.ScrollBars > 1))
        End If
    Else
        blnDontDoIt = True
    End If

Exit_Handler:
    ContinuousUpDownOk = Not blnDontDoIt
    Set ctl = Nothing

Exit Function

Err_Handler:
    Select Case Err.Number
        Case 2474 'No active control
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - ContinuousUpDownOk[mod_Forms])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          CaptureEscapeKey
' Description:  Handles ESCAPE key actions for certain forms
' Assumptions:
' Note:         Handles ESC for the following modal forms:
'               fsub_Soil_Stability, fsub_Fuels_LD, frm_Locations, frm_Unknown_Species
' Parameters:   KeyCode - keycode detected (key down)
' Returns:      -
' Throws:       none
' References:
'  John Spencer, 3/11/2010
'  http://msgroups.net/microsoft.public.access/how-best-to-disable-esc-key-on-form/21881
' Source/date:  Bonnie Campbell, August 21, 2015 - for NCPN tools
' Revisions:    BLC, 8/21/2015 - initial version
'               BLC, 6/1/2016  - added to mod_Forms from mod_App_UI (uplands)
' ---------------------------------
Public Sub CaptureEscapeKey(KeyCode As Integer)
On Error GoTo Err_Handler

    If KeyCode = vbKeyEscape Then
        If MsgBox("Undo changes?" & vbCrLf & vbCrLf & _
            "If yes, this may undo all recent changes (not just for a single field)." & vbCrLf & vbCrLf & _
            "Note:" & vbCrLf & _
            "If your cursor was in a..." & vbCrLf & _
            "+ text field, dropdown listbox, or checkbox field >> ALL changes will be undone." & vbCrLf & _
            "+ text field changed immediately before you clicked ESCAPE >> only the text field changes will be undone." & vbCrLf & vbCrLf & _
            "Previously saved data will remain unchanged.", vbYesNo, "ESCAPE Pressed!") = vbNo Then
            KeyCode = 0
        End If
        'KeyCode = 0
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CaptureEscapeKey[mod_Forms])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' Sub:          LimitKeyPress
' Description:  Limit form fields to a set number of characters
' Assumptions:  Control passed in is a text or combo box
' Parameters:   ctrl - textbox/combobox (control)
'               iMaxLen - # of allowed characters (integer)
'               KeyAscii - character passed in (integer)
' Returns:      -
' Throws:       none
' Usage:        Call LimitKeyPress(Me.MyTextBox, 12, KeyAscii) in control's KeyPress event
' References:   LimitChange() required in control's Change event also
' Source/date:
'   Allen Browne, unknown
'   http://allenbrowne.com/ser-34.html
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Public Sub LimitKeyPress(ctrl As Control, iMaxLen As Integer, KeyAscii As Integer)
On Error GoTo Err_Handler
    
    With ctrl
        If Len(.text) - .SelLength >= iMaxLen Then
            If KeyAscii <> vbKeyBack Then
                KeyAscii = 0
                Beep
            End If
        End If
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LimitKeyPress[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          LimitChange
' Description:  Limit form fields to a set number of characters
' Assumptions:  Control passed in is a textbox
' Parameters:   ctrl - textbox cotnrol
'               iMaxLen - maximum # of characters (integer)
' Returns:      -
' Throws:       none
' Usage:        Call LimitChange(Me.MyTextBox, 12) in control's Change event
' References:   LimitKeyPress() required in controls KeyPress event also
' Source/date:
'   Allen Browne, unknown
'   http://allenbrowne.com/ser-34.html
' Adapted:      Bonnie Campbell, June 28, 2016 - for NCPN tools
' Revisions:
'   BLC - 6/28/2016 - initial version
' ---------------------------------
Public Sub LimitChange(ctrl As Control, iMaxLen As Integer)
On Error GoTo Err_Handler

    Dim msg As String
    
    With ctrl
        If Len(.text) > iMaxLen Then
            msg = "Oops! " & .Name & " field too long. Truncated to " & iMaxLen & " characters."
        
            DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                "msg" & PARAM_SEPARATOR & msg & _
                "|Type" & PARAM_SEPARATOR & "caution"
            
            .text = Left(.text, iMaxLen)
            .SelStart = iMaxLen
        End If
    End With
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LimitChange[mod_Forms])"
    End Select
    Resume Exit_Handler
End Sub
