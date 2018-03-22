Operation =1
Option =0
Where ="(((tsys_Import_Tables.[Table_Name])=\"tbl_Quadrat_Data\" Or (tsys_Import_Tables."
    "[Table_Name])=\"tbl_Quadrat_Seedlings_Data\" Or (tsys_Import_Tables.[Table_Name]"
    ")=\"tbl_Quadrat_Herbaceous_Data\" Or (tsys_Import_Tables.[Table_Name])=\"tbl_CWD"
    "_Data\"))"
Begin InputTables
    Name ="tsys_Import_Tables"
End
Begin OutputColumns
    Expression ="tsys_Import_Tables.ID"
    Expression ="tsys_Import_Tables.Table_Name"
    Expression ="tsys_Import_Tables.Import"
End
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
End
Begin
    State =0
    Left =105
    Top =49
    Right =973
    Bottom =369
    Left =-1
    Top =-1
    Right =844
    Bottom =76
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tsys_Import_Tables"
        Name =""
    End
End
