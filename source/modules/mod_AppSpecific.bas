Option Compare Database
Option Explicit

Public Function IsNetwork(varUnitCode As Variant) As Boolean
Select Case varUnitCode
    Case "ARCN", "CAKN", "CHDN", "CUPN", "ERMN", "GLKN", "GRYN", "GULN", "HTLN", "KLMN", "MEDN", "MIDN", "MOJN", "NCBN", "NCCN", "NCPN", "NCRN", "NETN", "NGPN", "PACN", "ROMN", "SCPN", "SEAN", "SECN", "SFAN", "SFCN", "SIEN", "SODN", "SOPN", "SWAN", "UCBN"
        IsNetwork = True
End Select
End Function

Public Function MakeTreeStemList(strEventID As String, strTreeDataID As String) As String
'Collapse all sapling stems into a single field   mel 8/21/06
    Dim rst As DAO.Recordset
    Dim strStemList As String
    Dim strStemListLive As String
    Dim strStemListDead As String
    
        Set rst = CurrentDb.OpenRecordset("SELECT tbl_Tree_DBH.DBH, tbl_Tree_DBH.Live, tbl_Tree_Data.Event_ID, tbl_Tree_Data.Tree_Data_ID FROM tbl_Tree_Data INNER JOIN tbl_Tree_DBH ON tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_ID WHERE tbl_Tree_Data.Event_ID= """ & strEventID & """ AND tbl_Tree_Data.Tree_Data_ID= """ & strTreeDataID & """;")

        Do Until rst.EOF
            If rst!Live = True Then
                strStemListLive = strStemListLive & ", " & Format(rst!DBH, "#0.0")
            Else
                strStemListDead = strStemListDead & ", " & Format(rst!DBH, "#0.0")
            End If
            rst.MoveNext
        Loop

    strStemListLive = Mid(strStemListLive, 3)
    strStemListDead = Mid(strStemListDead, 3)
    strStemList = "L: " & strStemListLive & " D: " & strStemListDead
    MakeTreeStemList = strStemList

End Function

Public Function MakeSaplingStemList(strEventID As String, strSaplingDataID As String) As String
'Collapse all sapling stems into a single field   mel 8/21/06
    Dim rst As DAO.Recordset
    Dim strStemList As String
    Dim strStemListLive As String
    Dim strStemListDead As String
    
        Set rst = CurrentDb.OpenRecordset("SELECT tbl_Sapling_DBH.DBH, tbl_Sapling_DBH.Live, tbl_Sapling_Data.Event_ID, tbl_Sapling_Data.Sapling_Data_ID FROM tbl_Sapling_Data INNER JOIN tbl_Sapling_DBH ON tbl_Sapling_Data.Sapling_Data_ID = tbl_Sapling_DBH.Sapling_Data_ID WHERE tbl_Sapling_Data.Event_ID= """ & strEventID & """ AND tbl_Sapling_Data.Sapling_Data_ID= """ & strSaplingDataID & """;")

Do Until rst.EOF
            If rst!Live = True Then
                strStemListLive = strStemListLive & ", " & Format(rst!DBH, "#0.0")
            Else
                strStemListDead = strStemListDead & ", " & Format(rst!DBH, "#0.0")
            End If
            rst.MoveNext
        Loop

    strStemListLive = Mid(strStemListLive, 3)
    strStemListDead = Mid(strStemListDead, 3)
    strStemList = "L: " & strStemListLive & " D: " & strStemListDead
    MakeSaplingStemList = strStemList


End Function

Public Function MakeTreeConditionList(strEventID As String, strTreeDataID As String) As String
'Collapse all tree conditions into a single field   mel 2/4/2011
    Dim rst As DAO.Recordset
    Dim strConditionList As String

        Set rst = CurrentDb.OpenRecordset("SELECT tbl_Tree_Conditions.Condition, tbl_Tree_Data.Event_ID, tbl_Tree_Data.Tree_Data_ID FROM tbl_Tree_Data INNER JOIN tbl_Tree_Conditions ON tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Conditions.Tree_Data_ID WHERE tbl_Tree_Data.Event_ID= """ & strEventID & """ AND tbl_Tree_Data.Tree_Data_ID= """ & strTreeDataID & """;")

        Do Until rst.EOF
            strConditionList = strConditionList & ", " & Format(rst!Condition, "#0.0")
            rst.MoveNext
        Loop

    strConditionList = Mid(strConditionList, 3)
    MakeTreeConditionList = strConditionList

End Function