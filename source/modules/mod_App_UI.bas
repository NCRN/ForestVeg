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