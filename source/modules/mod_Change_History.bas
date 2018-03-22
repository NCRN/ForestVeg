Option Compare Database

Public Sub OpenChangeHistory(frmFormToSave As Form, ctlControlToReset As Control, strTableName As String, strFieldName As String, strRecordIDFieldName As String, strRecordID As String, strOldValue As String, strNewValue As String, strDisplayTableName As String, strDisplayFieldName As String, strDisplayKeyFieldName As String, DisplayKeyFieldDataType As DAO.DataTypeEnum)
Dim frm As Form
Dim varNew As Variant
Dim varOld As Variant

'Dim strOpenargs As String
'strOpenargs = XML_Tag("Form", strFormToSave) & XML_Tag("Control", strControlToReset)

DoCmd.OpenForm "frm_Tags_History", acNormal, , , acFormAdd ', , strOpenargs
'Set frm = Forms("frm_Tags_History")

Forms("frm_Tags_History").cboTag_ID = strRecordID
Forms("frm_Tags_History")!Table_Name = strTableName
Forms("frm_Tags_History")!Field_Name = strFieldName
Forms("frm_Tags_History")!Record_ID_Field_Name = strRecordIDFieldName
Forms("frm_Tags_History")!Display_Table_Name = strDisplayTableName
Forms("frm_Tags_History")!Display_Field_Name = strDisplayFieldName
Forms("frm_Tags_History")!Display_Key_Field_Name = strDisplayKeyFieldName

If IsNothing(strDisplayTableName) Then
    Forms("frm_Tags_History").txtValue_Old = strOldValue
    Forms("frm_Tags_History").txtValue_New = strNewValue
Else
    Select Case DisplayKeyFieldDataType
        Case DAO.DataTypeEnum.dbText
            varOld = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & CorrectText(strOldValue))
            varNew = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & CorrectText(strNewValue))
            Forms("frm_Tags_History").txtValue_Old = varOld
            Forms("frm_Tags_History").txtValue_New = varNew
        Case DAO.DataTypeEnum.dbLong, DAO.DataTypeEnum.dbBigInt, DAO.DataTypeEnum.dbByte, DAO.DataTypeEnum.dbInteger, DAO.DataTypeEnum.dbDouble
            varNew = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strNewValue)
            varOld = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strOldValue)
            Forms("frm_Tags_History").txtValue_Old = varOld
            Forms("frm_Tags_History").txtValue_New = varNew
    End Select
End If

Forms("frm_Tags_History").txtNetwork_User_Name = NetworkUserName()
Forms("frm_Tags_History").txtHistory_Notes.SetFocus
Set Forms("frm_Tags_History").ctlToReset = ctlControlToReset
Set Forms("frm_Tags_History").frmReferrer = frmFormToSave

'Set frm = Nothing
End Sub

Public Sub OpenChangeValueAndLog(strChangeDescription As String, strChangeFieldType As String, frmFormToSave As Form, ctlControlToReset As Control, strTableName As String, strFieldName As String, strRecordIDFieldName As String, strRecordID As String, strOldValue As String)
Dim frm As Form
Dim varNew As Variant
Dim varOld As Variant

'Dim strOpenargs As String
'strOpenargs = XML_Tag("Form", strFormToSave) & XML_Tag("Control", strControlToReset)

DoCmd.OpenForm "frm_Tags_History_Update", acNormal, , , acFormAdd ', , strOpenargs

Forms("frm_Tags_History_Update").lblChange_Description.Caption = strChangeDescription
Forms("frm_Tags_History_Update").cboTag_ID = strRecordID
Forms("frm_Tags_History_Update")!Table_Name = strTableName
Forms("frm_Tags_History_Update")!Field_Name = strFieldName
Forms("frm_Tags_History_Update")!Record_ID_Field_Name = strRecordIDFieldName

Forms("frm_Tags_History_Update").txtValue_Old = strOldValue

Forms("frm_Tags_History_Update").txtNetwork_User_Name = NetworkUserName()
Forms("frm_Tags_History_Update").txtValue_New.SetFocus
Set Forms("frm_Tags_History_Update").ctlToReset = ctlControlToReset
Set Forms("frm_Tags_History_Update").frmReferrer = frmFormToSave

End Sub

Public Sub OpenConfirmValueAndLog(strChangeDescription As String, strChangeFieldType As String, frmFormToSave As Form, ctlControlToReset As Control, strTableName As String, strFieldName As String, strRecordIDFieldName As String, strRecordID As String, strOldValue As String, strNewValue As String, strDisplayTableName As String, strDisplayFieldName As String, strDisplayKeyFieldName As String)
Dim frm As Form
Dim varNew As Variant
Dim varOld As Variant

'Dim strOpenargs As String
'strOpenargs = XML_Tag("Form", strFormToSave) & XML_Tag("Control", strControlToReset)

DoCmd.OpenForm "frm_Tags_History_Confirm", acNormal, , , acFormAdd ', , strOpenargs
'Set frm = Forms("frm_Tags_History")

Forms("frm_Tags_History_Confirm").lblChange_Description.Caption = strChangeDescription
Forms("frm_Tags_History_Confirm").cboTag_ID = strRecordID
Forms("frm_Tags_History_Confirm")!Table_Name = strTableName
Forms("frm_Tags_History_Confirm")!Field_Name = strFieldName
Forms("frm_Tags_History_Confirm")!Record_ID_Field_Name = strRecordIDFieldName

If strFieldName = "TSN" Then
            varOld = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strOldValue)
            varNew = DLookup(strDisplayFieldName, strDisplayTableName, strDisplayKeyFieldName & "=" & strNewValue)
            Forms("frm_Tags_History_Confirm").txtOld_Description = varOld
            Forms("frm_Tags_History_Confirm").txtNew_Description = varNew
End If

Forms("frm_Tags_History_Confirm").txtValue_Old = strOldValue
Forms("frm_Tags_History_Confirm").txtValue_New = strNewValue
Forms("frm_Tags_History_Confirm").txtNetwork_User_Name = NetworkUserName()
Forms("frm_Tags_History_Confirm").txtHistory_Notes.SetFocus
Set Forms("frm_Tags_History_Confirm").ctlToReset = ctlControlToReset
Set Forms("frm_Tags_History_Confirm").frmReferrer = frmFormToSave

End Sub