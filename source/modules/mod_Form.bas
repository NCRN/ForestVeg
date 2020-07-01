Option Compare Database
Option Explicit


Public Function fxnFormCheck(frm As Form) As Boolean
' =================================
' FUNCTION:     fxnFormCheck
' Description:  Checks controls on a form that have "<data>" in the tag property
'               to see if they contain a value.
' Parameters:   frm = the form to be checked
' Returns:      True if any of the selected controls contain a value, false otherwise
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, ?
' Revisions:    Simon Kingston, Feb 28, 2007 - documentation and error trapping
' =================================

Dim ctl As Control
Dim booHasData As Boolean

On Error GoTo Error_Handler

For Each ctl In frm.Controls
    If InStr(ctl.Tag, "<data>") > 0 Then
        If Not IsNull(ctl.value) Then
            booHasData = True
            Exit For
        End If
    End If
Next
fxnFormCheck = booHasData

Exit_Handler:
    On Error Resume Next
    Set ctl = Nothing
    Exit Function

Error_Handler:
    MsgBox Err.Number & " - " & Err.Description
    Resume Exit_Handler

End Function


Public Sub fxnUpdateControl(varOpenArgs As Variant)
' =================================
' SUBROUTINE:   fxnUpdateControl
' Description:  Requeries a selected control on a selected form
' Parameters:   varOpenArgs = a string (usually the OpenArgs property from a form) that
'                             may contain an XML formatted reference to a form and control
' Returns:      nothing
' Throws:       none
' References:   XML_Read, IsLoaded
' Source/date:  Simon Kingston, ?
' Revisions:    Simon Kingston, Feb 28, 2007 - documentation
' =================================
Dim strFormName As String
Dim strControlName As String

On Error Resume Next

If Not IsNothing(varOpenArgs) Then
    strFormName = XML_Read("FormFrom", CStr(varOpenArgs))
    strControlName = XML_Read("ControlFrom", CStr(varOpenArgs))
    
    If Len(strFormName) > 0 And Len(strControlName) > 0 Then
        If IsLoaded(strFormName) Then
            Forms(strFormName)(strControlName).Requery
        End If
    End If
End If

End Sub