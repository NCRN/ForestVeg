Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_UI
' Level:        Application module
' Version:      1.00

' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2018
' Revisions:    BLC, 4/19/2018  - 1.00 - initial version
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