Operation =1
Option =0
Where ="(((tlu_Enumerations.Enum_Group)=\"Crown Class\"))"
Begin InputTables
    Name ="tlu_Enumerations"
End
Begin OutputColumns
    Alias ="Crown_Class"
    Expression ="Int([Enum_Code])"
    Alias ="Crown_Description"
    Expression ="tlu_Enumerations.Enum_Description"
    Expression ="tlu_Enumerations.Enum_Group"
End
Begin OrderBy
    Expression ="Int([Enum_Code])"
    Flag =0
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
        dbText "Name" ="tlu_Enumerations.Enum_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Enumerations.Enum_Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Enumerations.Enum_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Crown_Description"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Crown_Class"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =114
    Top =178
    Right =1518
    Bottom =947
    Left =-1
    Top =-1
    Right =1372
    Bottom =452
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tlu_Enumerations"
        Name =""
    End
End
