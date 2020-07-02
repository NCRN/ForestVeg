Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_UI
' Level:        Application module
' Version:      1.03
'
' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2018
' Revisions:    BLC, 4/19/2018  - 1.00 - initial version
'               BLC, 5/21/2018  - 1.01 - accommodate NULL if user hasn't set value
'               BLC, 5/3/2109   - 1.02 - shifted GoToForm, WriteRecordCriteria from frm_Data_Gateway
'               BLC, 6/29/2020  - 1.03 - added in development message
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
' SUB:          DisableControls
' Description:  disables all form controls
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 19, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/4/2018 - initial version
' ---------------------------------
Public Sub DisableControls(frm As Form)
On Error GoTo Err_Handler
    
    
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableControls[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          DisplayMessage
' Description:  displays user message
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 29, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/29/2018 - initial version
' ---------------------------------
Public Sub DisplayMessage(topic As String)
On Error GoTo Err_Handler
    
    Dim msg As String
    Dim Title As String
    Dim msgType As Long
    
    Select Case topic
        Case "notready"
            Title = "Patience Required - Feature Not Yet Ready for Prime Time"
            msg = "Sorry, this feature is not quite ready." _
                    & vbCrLf & "Please check back in the next release." _
                    & vbCrLf & vbCrLf & "Thank you for your patience..." _
                    & vbCrLf & "...and for checking out new features!"
            msgType = vbInformation
    End Select
    
    MsgBox msg, msgType, Title

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisplayMessage[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Functions
' ---------------------------------
' ---------------------------------
' FUNCTION:     CheckboxToBit
' Description:  convert True/False (-1,0) to Byte (1,0) values
' Note:         Access sets checkbox values to True (-1) or False (0)
'               Any number other than 0 is treated as True
'               (because it's Not False)
' Assumptions:  -
' Parameters:   chkValue - checkbox value
' Returns:      -
' Throws:       none
' References:
'   David W. Fenton, September 29, 2010
'   https://stackoverflow.com/questions/3813760/determine-whether-a-access-checkbox-is-checked-or-not
' Source/date:  Bonnie Campbell, April 21, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/21/2018 - initial version
' ---------------------------------
Public Function CheckboxToBit(chkValue As Integer) As Byte
On Error GoTo Err_Handler
    
    'reject values |x|>1
    If Abs(chkValue) > 1 Then GoTo Exit_Handler
    
    'convert to viable value
    CheckboxToBit = Abs(chkValue)
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CheckboxToBit[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          LaunchKeypad
' Description:  keypad launch actions
' Requires:     Keypad Utils module
' Assumptions:  -
' Parameters:   frm - form to update (form)
'               keypad - name of keypad form (string)
'               ctlName - name of control to update (string)
' Returns:      -
' Throws:       none
' References:   Mark Lehman/Geoffrey Sanders, unknown
' Source/date:  Bonnie Campbell, April 22, 2018
' Adapted:      -
' Revisions:    BLC - 4/22/2018 - 1.00 - initial version
' ---------------------------------
Public Sub LaunchKeypad(frm As Form, keypad As String, ctlName As String)
On Error GoTo Err_Handler
    
    Call OpenKeypad(keypad, frm, ctlName)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - LaunchKeypad[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ValidPct
' Description:  percent validating actions
' Usage:        =ValidPct(ctrlValue, NullOK) in the LostFocus event of the control
'               for example:
'               =ValidPct([Screen].[ActiveControl],True)
'               used to trigger ValidationRule, ValidationText
' Assumptions:  -
' Parameters:   pct - value for the percent (double)
'               NullOK - whether NULL is an acceptable value (boolean, optional, default = False)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/22/2018 - initial version
'   BLC - 5/21/2018 - accommodate NULL if user hasn't set value
' ---------------------------------
Public Function ValidPct(Pct As Variant, Optional NullOK As Boolean = False) As Double
On Error GoTo Err_Handler
    
    Dim IsValid As Boolean
    
    'default
    ValidPct = 0
    IsValid = False
    
    'handle when NULLs are OK (i.e. when no value is yet set)
    If (NullOK = True) And (IsNull(Pct) = True) Then
        IsValid = True
        GoTo Exit_Handler
    End If
    
    Select Case Pct
'        Case Is = 0
'            ValidPct = pct
'            IsValid = True
        Case 0 To 100
            ValidPct = Pct
            IsValid = True
'        Case Is = 100
'            ValidPct = pct
'            IsValid = True
        Case Else
            'use default
'           ValidPct = 0
    End Select
    
    'set the control value?
    'Screen.ActiveControl = ValidPct
    If IsValid = False Then
        Screen.ActiveControl.backcolor = lngYellow
        Screen.ActiveControl.forecolor = lngRed
        MsgBox "Percent cover values range from 0 to 100 (inclusive). " _
                & vbCrLf & "Please check the highlighted value.", vbOKOnly, _
                "NCRN Vegetation Monitoring > Invalid Percent Value"
    Else
        Screen.ActiveControl.backcolor = lngWhite
        Screen.ActiveControl.forecolor = lngBlack
    End If
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ValidPct[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GoToForm
' Description:  open desired form
' Assumptions:
' Referenced Libraries: framework.DbObjectExists
' Parameters:   frm - name of form to open (string)
'               caller - name of calling form (optional, string)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 24, 2018
' Adapted:      -
' Revisions:
'   BLC - 5/24/2018 - initial version
'   BLC - 5/3/2019  - shifted from frm_Data_Gateway
' ---------------------------------
Public Function GoToForm(frm As String, Optional caller As String)
On Error GoTo Err_Handler
    
    'write record if on gateway
    'If caller = "frm_Data_Gateway" Then
    'Call Forms("frm_Data_Gateway").
    WriteRecordCriteria
    
    If DbObjectExists(frm, "frm") Then _
        DoCmd.OpenForm frm, acNormal
        
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnGoToTags_Click[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          WriteRecordCriteria
' Description:  Records Location & Event IDs of the current record so that it can be made the current record when coming
'               back to the form from another form (=bookmark).
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   GetCriteriaString
' Source/date:  Simon Kingston, 1/17/2007
'               Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   SK - 1/17/2007 - initial version
'   MEL/GS - unknown - initial NCRN version
'   BLC - 5/24/2018 - update documentation, error handling
'   BLC - 5/3/2019 - shift from frm_Data_Gateway & adapt for global
' ---------------------------------
Private Sub WriteRecordCriteria()
On Error GoTo Err_Handler

'    If Not IsNothing(Me!Location_ID) Then
'        strCurrentRecordCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
'        If IsNothing(Me!Event_ID) Then
'            strCurrentRecordCriteria = strCurrentRecordCriteria & " AND Event_ID Is Null"
'        Else
'            strCurrentRecordCriteria = strCurrentRecordCriteria & " AND " & GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
'        End If
'    End If
    
    Dim strCurrentRecordCriteria As String
    
    If Not IsNothing(TempVars("plot")) Then
        strCurrentRecordCriteria = "[Location_ID]='" & TempVars("plot") & "'"
        If IsNothing(TempVars("eventID")) Then
            strCurrentRecordCriteria = strCurrentRecordCriteria & " AND Event_ID Is Null"
        Else
            strCurrentRecordCriteria = strCurrentRecordCriteria & " AND [Event_ID]='" & TempVars("eventID") & "'"
        End If
    End If
    
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - WriteRecordCriteria[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetAppIcon
' Description:  Sets the application icon.
' Assumptions:  IconFile is actually an icon (*.ico) file
'               A check is made to see that it is a file & has the ico extension
'               however this doesn't guarantee 100% that the file is an icon.
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Vishal Monpara, December 15, 2009
'   https://www.vishalon.net/blog/change-ms-access-application-title-and-icon-using-vba
' Source/date:  Bonnie Campbell, June 4, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/4/2020 - initial version
' ---------------------------------
Public Sub SetAppIcon(IconFile As String)
On Error GoTo Err_Handler
  
    'set file path
    Dim IconFullPath As String
    IconFullPath = CurrentProject.Path & "\" & IconFile
  
    'is IconFile actually present?
    If FileExists(IconFullPath) And Right(LCase(IconFullPath), 3) = "ico" Then
  
        With CurrDb
            .Properties("AppIcon").Value = IconFullPath
            .Properties("AppTitle").Value = .Properties("Title")
            'if you want to extend icon to reports
            '.Properties("UseAppIconForFrmRpt").Value = True
        
            Application.RefreshTitleBar
        End With
        
    Else
        MsgBox "Sorry that is not a valid icon file.", vbInformation, "Invalid Icon File"
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - " & Application.VBE.ActiveCodePane & " [" & Application.VBE.a & "])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          SortListForm
' Description:  form label sort on click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
'   Allen Browne, June 28, 2006
'   https://bytes.com/topic/access/answers/506322-using-orderby-multiple-fields
' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/19/2017 - initial version
'   BLC - 1/31/2017 - adjusted to accommodate templates list
'   BLC - 2/21/2017 - adjusted to accommodate Contact list
'   BLC - 10/18/2017 - added cases for Comment list
'   BLC - 12/7/2017 - added cases for VegPlot, Event lists
'   BLC - 1/17/2018 - added cases for Task list
'   BLC - 6/11/2020 - repurposed for NCRN ForestVeg
' ---------------------------------
Public Sub SortListForm(frm As Form, ctrl As Control)
On Error GoTo Err_Handler

    Dim strSort As String
    
    'default
    strSort = ""
    
    'set sort field
    Select Case Replace(ctrl.Name, "lbl", "")
        Case "Citation"
            strSort = "LongCitation"
        Case "Comment"
            strSort = "Comment"
        Case "CommentType"
            strSort = "CommentType"
        Case "CommentTypeID"
            strSort = "CommentType_ID"
        Case "Email"
            strSort = "Email"
        Case "HdrID"
            strSort = "ID"
            Select Case frm.Name
                Case "ContactList"
                    strSort = "c.ID"
                Case "TaskList"
                    strSort = "t.ID"
            End Select
        Case "Location"
            strSort = "Location"
        Case "ModalSedSize"
            strSort = "ModalSedimentSize_ID"
        Case "Name"
            strSort = "LastName"
        Case "PctMSS"
            strSort = "PctModalSedimentSize"
        Case "PlotNumDist"
            strSort = IIf(TempVars("ParkCode") = "DINO", "PlotNumber", "PlotDistance_m")
        Case "Priority"
            strSort = "Priority"
        Case "Reference"
            strSort = "ShortCitation"
        Case "Site"
            strSort = "Site"
        Case "SOP"
            strSort = "FullName"
        Case "SOPNum"
            strSort = "SOPNumber"
        Case "StartDate"
            strSort = "StartDate"
        Case "Syntax"
            strSort = "Syntax"
        Case "Task"
            strSort = "Task"
        Case "TaskType"
            strSort = "TaskType"
        Case "Template"
            strSort = "TemplateName"
        Case "Status"
            strSort = "Status"
        Case "Version"
            strSort = "Version"
        Case "LastModifiedDate"
            strSort = "LastModified"
        Case "EffectiveDate"
            strSort = "EffectiveDate"
        Case ""
    End Select

    'set the sort
    If InStr(frm.OrderBy, strSort) = 0 Then
        frm.OrderBy = strSort
    ElseIf Right(frm.OrderBy, 4) = "Desc" Then
        frm.OrderBy = strSort
    Else
        frm.OrderBy = strSort & " Desc"
    End If
    
    frm.OrderByOn = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortListForm[mod_App_UI form])"
    End Select
    Resume Exit_Handler
End Sub