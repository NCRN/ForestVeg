dbMemo "SQL" ="SELECT t.Append_Order, t.Table_Name, t.Append, t.Append_Table\015\012FROM tsys_A"
    "ppend_Tables AS t\015\012WHERE t.Table_Name IN (\"tbl_Quadrat_Data\",\"tbl_Quadr"
    "at_Seedlings_Data\",\"tbl_Quadrat_Herbaceous_Data\",\"tbl_CWD_Data\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="t.Append_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Table_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Append"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Append_Table"
        dbLong "AggregateType" ="-1"
    End
End
