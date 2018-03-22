Option Compare Database

Public CtrlToUpdate As Control

Function GetFormType(frmIn As Form) As String
  ' Comments   : determines if the passed form is open as a form, subform or a sub-subform
  ' Parameters : frmIn - handle to the form
  ' Returns    : text description of form type
  ' Created    : Mark Lehman
  ' Modified   : 5/21/2007
  ' --------------------------------------------------------
  Dim varfrmcheck As Variant
  Dim varfrmcheck2 As Variant
  
    On Error Resume Next
    varFormCheck = frmIn.Parent
    If Err = 1 Or IsEmpty(varFormCheck) Then
      GetFormType = "CtrlOnMAIN"
    Else
        varFormCheck2 = frmIn.Parent.Parent
        If Err = 1 Or IsEmpty(varFormCheck2) Then
            GetFormType = "CtrlOnSUB"
        Else
            GetFormType = "CtrlOnSUBSUB"
      End If
    End If
  On Error GoTo 0
End Function

Sub OpenKeypad(strKeypadFormName As String, frmFormToUpdate As Form, strControlToUpdate As String)
  ' Comments   : Opens a keypad with a link established to a specific control
  ' Parameters : The name of the keypad form, the form where the control to update is located, the name of the control to update
  ' Returns    : Does not return a value
  ' Created    : Mark Lehman
  ' Modified   : 5/21/2007
  ' --------------------------------------------------------
    Dim strFormType As String
    Dim strFormParent As String
    Dim strFormParentUp1 As String
    Dim strFormParentUp2 As String
    Dim strLinkCriteria As String

  strFormType = GetFormType(frmFormToUpdate)
  strFormParent = frmFormToUpdate.Name
  Select Case strFormType
    Case "CtrlOnMain"
      Set CtrlToUpdate = Forms(strFormParent).Controls(strControlToUpdate)
      DoCmd.Close acForm, strKeypadFormName, acSavePrompt
      DoCmd.OpenForm strKeypadFormName, , , strLinkCriteria, , , strOpenArg
    Case "CtrlOnSUB"
      strFormParentUp1 = frmFormToUpdate.Parent.Name
      Set CtrlToUpdate = Forms(strFormParentUp1).Controls(strFormParent).Form.Controls(strControlToUpdate)
      DoCmd.Close acForm, strKeypadFormName, acSavePrompt
      DoCmd.OpenForm strKeypadFormName, , , strLinkCriteria, , , strOpenArg
    Case "CtrlOnSUBSUB"
      strFormParentUp1 = frmFormToUpdate.Parent.Name
      strFormParentUp2 = frmFormToUpdate.Parent.Parent.Name
      Set CtrlToUpdate = Forms(strFormParentUp2).Controls(strFormParentUp1).Form.Controls(strFormParent).Form.Controls(strControlToUpdate)
      DoCmd.Close acForm, strKeypadFormName, acSavePrompt
      DoCmd.OpenForm strKeypadFormName, , , strLinkCriteria, , , strOpenArg
    Case Else
  End Select
End Sub

Sub OpenSpeciespad(strKeypadFormName As String, frmFormToUpdate As Form, strControlToUpdate As String, strOpenArg As String)
  ' Comments   : Opens a keypad with a link established to a specific control
  ' Parameters : The name of the keypad form, the form where the control to update is located, the name of the control to update
  ' Returns    : Does not return a value
  ' Created    : Mark Lehman
  ' Modified   : 5/21/2007
  ' --------------------------------------------------------
    Dim strFormType As String
    Dim strFormParent As String
    Dim strFormParentUp1 As String
    Dim strFormParentUp2 As String
    Dim strLinkCriteria As String

  strFormType = GetFormType(frmFormToUpdate)
  strFormParent = frmFormToUpdate.Name
  Select Case strFormType
    Case "CtrlOnMain"
      Set CtrlToUpdate = Forms(strFormParent).Controls(strControlToUpdate)
      DoCmd.Close acForm, strKeypadFormName, acSavePrompt
      DoCmd.OpenForm strKeypadFormName, , , strLinkCriteria, , , strOpenArg
    Case "CtrlOnSUB"
      strFormParentUp1 = frmFormToUpdate.Parent.Name
      Set CtrlToUpdate = Forms(strFormParentUp1).Controls(strFormParent).Form.Controls(strControlToUpdate)
      DoCmd.Close acForm, strKeypadFormName, acSavePrompt
      DoCmd.OpenForm strKeypadFormName, , , strLinkCriteria, , , strOpenArg
    Case "CtrlOnSUBSUB"
      strFormParentUp1 = frmFormToUpdate.Parent.Name
      strFormParentUp2 = frmFormToUpdate.Parent.Parent.Name
      Set CtrlToUpdate = Forms(strFormParentUp2).Controls(strFormParentUp1).Form.Controls(strFormParent).Form.Controls(strControlToUpdate)
      DoCmd.Close acForm, strKeypadFormName, acSavePrompt
      DoCmd.OpenForm strKeypadFormName, , , strLinkCriteria, , , strOpenArg
    Case Else
  End Select
End Sub