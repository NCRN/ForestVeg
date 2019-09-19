dbMemo "SQL" ="SELECT tsys_Import_Tables.ID, tsys_Import_Tables.Table_Name, tsys_Import_Tables."
    "Import\015\012FROM tsys_Import_Tables\015\012WHERE tsys_Import_Tables.[Table_Nam"
    "e] NOT IN (\"tbl_Quadrat_Data\",\"tbl_Quadrat_Seedlings_Data\",\"tbl_Quadrat_Her"
    "baceous_Data\",\"tbl_CWD_Data\");\015\012"
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
End
