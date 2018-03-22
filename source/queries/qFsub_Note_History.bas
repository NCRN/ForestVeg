dbMemo "SQL" ="SELECT Location_ID, \"Event\" as Note_Type, Event_Date as Note_Date, Event_Notes"
    " as Notes\015\012FROM tbl_Events\015\012UNION ALL SELECT Location_ID, \"Plot Set"
    "up\", Install_Date, Location_Notes \015\012FROM tbl_Locations\015\012ORDER BY No"
    "te_Date DESC;\015\012"
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
        dbText "Name" ="Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Note_Type"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Note_Date"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes"
        dbInteger "ColumnWidth" ="13215"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
