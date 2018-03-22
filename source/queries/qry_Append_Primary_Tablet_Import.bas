Operation =1
Option =0
Where ="(((tsys_Import_Tables.[Table_Name])<>\"tbl_Quadrat_Data\" And (tsys_Import_Table"
    "s.[Table_Name])<>\"tbl_Quadrat_Seedlings_Data\" And (tsys_Import_Tables.[Table_N"
    "ame])<>\"tbl_Quadrat_Herbaceous_Data\" And (tsys_Import_Tables.[Table_Name])<>\""
    "tbl_CWD_Data\"))"
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
    Left =133
    Top =482
    Right =1001
    Bottom =813
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
