dbMemo "SQL" ="SELECT DISTINCT t.Tag_ID\015\012FROM tbl_Sapling_Data AS td LEFT JOIN tbl_Tags A"
    "S t ON t.Tag_ID = td.Tag_ID\015\012WHERE td.Sapling_Status LIKE 'Dead*'\015\012A"
    "ND t.Tag_Status <> 'Retired (In Office)'\015\012AND t.Tag_Status NOT LIKE 'Inact"
    "ive *';\015\012"
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
End
