dbMemo "SQL" ="SELECT t.ID, t.Table_Name, t.Import\015\012FROM tsys_Import_Tables AS t\015\012W"
    "HERE t.[Table_Name] IN (\"tbl_Quadrat_Data\",\"tbl_Quadrat_Seedlings_Data\",\"tb"
    "l_Quadrat_Herbaceous_Data\",\"tbl_CWD_Data\");\015\012"
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
        dbText "Name" ="[tsys_Import_Tables].ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tsys_Import_Tables].Table_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tsys_Import_Tables].Import"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Import_Tables.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Import_Tables.Table_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Import_Tables.Import"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tsys_Import_Tables.[Table_Name]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Table_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Import"
        dbLong "AggregateType" ="-1"
    End
End
