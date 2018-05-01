dbMemo "SQL" ="UPDATE tbl_Tags AS t SET t.Tag_Status = 'Tree'\015\012WHERE t.Tag_ID IN \015\012"
    "(SELECT t.Tag_ID\015\012FROM (tbl_Tags t\015\012INNER JOIN tbl_Tree_Data td ON t"
    "d.Tag_ID = t.Tag_ID)\015\012WHERE \015\012td.Tree_Status = 'Dead Leaning'\015\012"
    "AND\015\012t.Tag_Status <> 'Tree'\015\012ORDER BY \015\012td.Updated_Date, td.Tr"
    "ee_Status);\015\012"
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
        dbText "Name" ="t.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stop_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RFS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Updated_Date"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
