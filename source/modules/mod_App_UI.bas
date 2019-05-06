Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_UI
' Level:        Application module
' Version:      1.02
'
' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2018
' Revisions:    BLC, 4/19/2018  - 1.00 - initial version
'               BLC, 5/21/2018  - 1.01 - accommodate NULL if user hasn't set value
'               BLC, 5/3/2109   - 1.02 - shifted GoToForm, WriteRecordCriteria from frm_Data_Gateway
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
Public Function ValidPct(pct As Variant, Optional NullOK As Boolean = False) As Double
On Error GoTo Err_Handler
    
    Dim IsValid As Boolean
    
    'default
    ValidPct = 0
    IsValid = False
    
    'handle when NULLs are OK (i.e. when no value is yet set)
    If (NullOK = True) And (IsNull(pct) = True) Then
        IsValid = True
        GoTo Exit_Handler
    End If
    
    Select Case pct
'        Case Is = 0
'            ValidPct = pct
'            IsValid = True
        Case 0 To 100
            ValidPct = pct
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
        Screen.ActiveControl.BackColor = lngYellow
        Screen.ActiveControl.ForeColor = lngRed
        MsgBox "Percent cover values range from 0 to 100 (inclusive). " _
                & vbCrLf & "Please check the highlighted value.", vbOKOnly, _
                "NCRN Vegetation Monitoring > Invalid Percent Value"
    Else
        Screen.ActiveControl.BackColor = lngWhite
        Screen.ActiveControl.ForeColor = lngBlack
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