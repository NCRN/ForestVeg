dbMemo "SQL" ="UPDATE tbl_Tags AS t SET t.Tag_Status = 'Inactive (In Field)'\015\012WHERE t.Tag"
    "_ID IN \015\012(SELECT t.Tag_ID\015\012FROM (tbl_Tags t\015\012INNER JOIN tbl_Tr"
    "ee_Data td ON td.Tag_ID = t.Tag_ID)\015\012WHERE \015\012td.Tree_Status = 'Dead "
    "- Too Small'\015\012AND\015\012t.Tag_Status <> 'Inactive (In Field)'\015\012ORDE"
    "R BY \015\012td.Updated_Date, td.Tree_Status);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
End
