Operation =1
Option =0
Where ="(((tsys_Append_Tables.Table_Name)<>\"tbl_Locations\" And (tsys_Append_Tables.Tab"
    "le_Name)<>\"tbl_Events\"))"
Begin InputTables
    Name ="tsys_Append_Tables"
End
Begin OutputColumns
    Expression ="tsys_Append_Tables.Table_Name"
    Expression ="tsys_Append_Tables.Append"
    Expression ="tsys_Append_Tables.Append_Table"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="[tsys_Append_Tables].Append_Table"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tsys_Append_Tables].Table_Name"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =47
    Top =69
    Right =752
    Bottom =541
    Left =-1
    Top =-1
    Right =681
    Bottom =118
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tsys_Append_Tables"
        Name =""
    End
End
