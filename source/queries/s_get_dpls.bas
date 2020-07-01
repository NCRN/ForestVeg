dbMemo "SQL" ="SELECT Enum_Code, Enum_Description, Enum_Group, Sort_Order\015\012FROM tlu_Enume"
    "rations\015\012WHERE Enum_Group = 'DPL'\015\012ORDER BY Sort_Order;\015\012"
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
        dbText "Name" ="Enum_Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Enum_Code"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sort_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Enum_Group"
        dbLong "AggregateType" ="-1"
    End
End
