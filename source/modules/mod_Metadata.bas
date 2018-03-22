Option Compare Database
Option Explicit

Public Enum METADATA_TYPE
    LOCAL_FILE = 1
    NPS_DATA_STORE = 2
End Enum

Public Function fxnGetLocalMetadataFileName() As Variant
Dim varFileName As Variant

On Error GoTo Error_Handler

varFileName = fxnGetDbFileName("tbl_Db_Meta")
varFileName = GetPath(CStr(varFileName))
varFileName = varFileName + DLookup("Meta_File_Name", "tbl_Db_Meta")

Exit_Handler:
    fxnGetLocalMetadataFileName = varFileName
    Exit Function

Error_Handler:
    varFileName = Null
    Resume Exit_Handler

End Function

Public Function fxnGetNPSDataStoreMetadataLink() As Variant
Dim varURL As Variant

On Error GoTo Error_Handler

varURL = DLookup("Meta_MID", "tbl_Db_Meta")

Exit_Handler:
    fxnGetNPSDataStoreMetadataLink = varURL
    Exit Function

Error_Handler:
    varURL = Null
    Resume Exit_Handler

End Function

Public Function fxnLocalMetadataExists(Optional varFileNameIn As Variant) As Boolean
Dim varLocalFile As Variant

If IsMissing(varFileNameIn) Then
    varLocalFile = fxnGetLocalMetadataFileName
Else
    varLocalFile = varFileNameIn
End If

fxnLocalMetadataExists = FileExists(varLocalFile)

End Function

Public Function fxnNPSDataStoreMetadataExists(Optional varURLIn As Variant) As Boolean
Dim varURL As Variant

If IsMissing(varURLIn) Then
    varURL = fxnGetNPSDataStoreMetadataLink
Else
    varURL = varURLIn
End If

fxnNPSDataStoreMetadataExists = Not IsNothing(varURL)
End Function

Public Function fxnDBPurposeExists() As Boolean
fxnDBPurposeExists = Not IsNull(DLookup("Db_Desc", "tbl_Db_Meta"))
End Function

Public Function fxnAddMetadataLink(strValue As String, MetadataType As METADATA_TYPE) As Boolean
Dim rst As DAO.Recordset
Dim strFieldName As String
Dim intRecordCount As Integer

On Error GoTo Error_Handler

Set rst = CurrentDb.OpenRecordset("tbl_Db_Meta")

Select Case MetadataType
    Case LOCAL_FILE
        strFieldName = "Meta_File_Name"
    Case NPS_DATA_STORE
        strFieldName = "Meta_MID"
End Select

If rst.BOF And rst.BOF Then
    intRecordCount = 0
Else
    rst.MoveLast
    intRecordCount = rst.RecordCount
End If
    
Select Case intRecordCount
    Case 0
        'need to insert a record
        rst.AddNew
        rst(strFieldName) = strValue
        rst.Update
    Case 1
        'need to update existing record
        rst.Edit
        rst(strFieldName) = strValue
        rst.Update
    Case Else
        MsgBox "More than one metadata record exists in tbl_Db_Meta.  Only one metadata record should be in tbl_Db_Meta.", vbCritical, "Metadata Error - Not Updated"
        Err.Raise vbObjectError + 513
End Select

fxnAddMetadataLink = True

Exit_Handler:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Function

Error_Handler:
    Resume Exit_Handler

End Function