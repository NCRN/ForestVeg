Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Change_History
' Level:        Form module
' Version:      1.01
'
' Description:  change history related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 11/5/2018 - 1.01 - added documentation, error handling
'                                          improved code efficiency
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ----------------
'  Methods
' ----------------

' ---------------------------------
' SUB:          OpenChangeHistory
' Description:  opens change history for various forms
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
' ---------------------------------
Public Sub OpenChangeHistory(frmFormToSave As Form, ctlControlToReset As Control, strTableName As String, strFieldName As String, strRecordIDFieldName As String, strRecordID As String, strOldValue As String, strNewValue As String, strDisplayTableName As String, strDisplayFieldName As String, strDisplayKeyFieldName As String, DisplayKeyFieldDataType As DAO.DataTypeEnum)
On Error GoTo Err_Handler
    Dim frm As Form
    Dim varNew As Variant
    Dim varOld As Variant
    
    'Dim strOpenargs As String
    'strOpenargs = XML_Tag("Form", strFormToSave) & XML_Tag("Control", strControlToReset)
    
    DoCmd.OpenForm "frm_Tags_History", acNormal, , , acFormAdd ', , strOpenargs
    Set frm = Forms("frm_Tags_History")
    
    With frm
        .cbxTag_ID = strRecordID
        !Table_Name = strTableName
        !Field_Name = strFieldName
        !Record_ID_Field_Name = strRecordIDFieldName
        !Display_Table_Name = strDisplayTableName
        !Display_Field_Name = strDisplayFieldName
        !Display_Key_Field_Name = strDisplayKeyFieldName
    
        If IsNothing(strDisplayTableName) Then
            .tbxValueOld = strOldValue
            .tbxValueNew = strNewValue
        Else
            Select Case DisplayKeyFieldDataType
                Case DAO.DataTypeEnum.dbText
                    varOld = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & CorrectText(strOldValue))
                    varNew = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & CorrectText(strNewValue))
                    .tbxValueOld = varOld
                    .tbxValueNew = varNew
                Case DAO.DataTypeEnum.dbLong, DAO.DataTypeEnum.dbBigInt, DAO.DataTypeEnum.dbByte, DAO.DataTypeEnum.dbInteger, DAO.DataTypeEnum.dbDouble
                    varNew = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strNewValue)
                    varOld = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strOldValue)
                    .tbxValueOld = varOld
                    .tbxValueNew = varNew
            End Select
        End If
        
        .tbxNetworkUserName = NetworkUserName()
        .tbxHistory_Notes.SetFocus
        
        Set frm.ctlToReset = ctlControlToReset
        Set frm.frmReferrer = frmFormToSave
    
    End With
    
Exit_Handler:
    Set frm = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenChangeHistory[mod_Change_History])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          OpenChangeValueAndLog
' Description:  opens form for tag value logging for various forms
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
' ---------------------------------
Public Sub OpenChangeValueAndLog(strChangeDescription As String, strChangeFieldType As String, frmFormToSave As Form, ctlControlToReset As Control, strTableName As String, strFieldName As String, strRecordIDFieldName As String, strRecordID As String, strOldValue As String)
On Error GoTo Err_Handler
    
    Dim frm As Form
    Dim varNew As Variant
    Dim varOld As Variant
    
    'Dim strOpenargs As String
    'strOpenargs = XML_Tag("Form", strFormToSave) & XML_Tag("Control", strControlToReset)
    
    DoCmd.OpenForm "frm_Tags_History_Update", acNormal, , , acFormAdd ', , strOpenargs
    
    Set frm = Forms("frm_Tags_History_Update")
    
    With frm
        .lblChange_Description.Caption = strChangeDescription
        .cbxTag_ID = strRecordID
        !Table_Name = strTableName
        !Field_Name = strFieldName
        !Record_ID_Field_Name = strRecordIDFieldName
    
        .tbxValueOld = strOldValue
        .tbxNetworkUserName = NetworkUserName()
        .tbxValueNew.SetFocus
        
        Set .ctlToReset = ctlControlToReset
        Set .frmReferrer = frmFormToSave
    End With

Exit_Handler:
    Set frm = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenChangeValueAndLog[mod_Change_History])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          OpenConfirmValueAndLog
' Description:  opens form for confirming tag value logging for various forms
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
' ---------------------------------
Public Sub OpenConfirmValueAndLog(strChangeDescription As String, strChangeFieldType As String, frmFormToSave As Form, ctlControlToReset As Control, strTableName As String, strFieldName As String, strRecordIDFieldName As String, strRecordID As String, strOldValue As String, strNewValue As String, strDisplayTableName As String, strDisplayFieldName As String, strDisplayKeyFieldName As String)
On Error GoTo Err_Handler
    Dim frm As Form
    Dim varNew As Variant
    Dim varOld As Variant
    
    'Dim strOpenargs As String
    'strOpenargs = XML_Tag("Form", strFormToSave) & XML_Tag("Control", strControlToReset)
    
    DoCmd.OpenForm "frm_Tags_History_Confirm", acNormal, , , acFormAdd ', , strOpenargs
    
    Set frm = Forms("frm_Tags_History_Confirm")
    
    With frm
    
        .lblDescription.Caption = strChangeDescription
        .cbxTag_ID = strRecordID
        !Table_Name = strTableName
        !Field_Name = strFieldName
        !Record_ID_Field_Name = strRecordIDFieldName
       
        If strFieldName = "TSN" Then
                    varOld = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strOldValue)
                    varNew = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strNewValue)
                    .tbxDescriptionOld = varOld
                    .tbxDescriptionNew = varNew
        End If
        
        .tbxValueOld = strOldValue
        .tbxValueNew = strNewValue
        .tbxNetworkUserName = NetworkUserName()
        .tbxHistoryNotes.SetFocus
        
        Set .ctlToReset = ctlControlToReset
        Set .frmReferrer = frmFormToSave
    End With

Exit_Handler:
    Set frm = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - OpenConfirmValueAndLog[mod_Change_History])"
    End Select
    Resume Exit_Handler
End Sub