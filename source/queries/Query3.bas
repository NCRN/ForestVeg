﻿dbMemo "SQL" ="SELECT t.Tag_ID\015\012FROM tbl_Sapling_Data AS td LEFT JOIN tbl_Tags AS t ON t."
    "Tag_ID = td.Tag_ID\015\012WHERE td.Sapling_Status LIKE 'Dead*'\015\012AND t.Tag_"
    "Status <> 'Retired (In Office)'\015\012AND t.Tag_Status NOT LIKE 'Inactive *';\015"
    "\012"
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
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
End
