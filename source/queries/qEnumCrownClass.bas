Operation =1
Option =0
Where ="(((tlu_Enumerations.Enum_Group)=\"Crown Class\"))"
Begin InputTables
    Name ="tlu_Enumerations"
End
Begin OutputColumns
    Alias ="CrownClassCode"
    Expression ="Val([Enum_Code])"
    Expression ="tlu_Enumerations.Enum_Description"
    Expression ="tlu_Enumerations.Enum_Group"
    Expression ="tlu_Enumerations.Sort_Order"
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
        dbText "Name" ="tlu_Enumerations.Sort_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Enumerations.Enum_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Enumerations.Enum_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Enumerations.Enum_Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CrownClassCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =86
    Top =142
    Right =1488
    Bottom =1030
    Left =-1
    Top =-1
    Right =1370
    Bottom =571
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =279
        Top =126
        Right =580
        Bottom =259
        Top =0
        Name ="tlu_Enumerations"
        Name =""
    End
End
