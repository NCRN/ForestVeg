dbMemo "SQL" ="SELECT Val([Enum_Code]) AS CrownClassCode, e.Enum_Description, e.Enum_Group, e.S"
    "ort_Order\015\012FROM tlu_Enumerations AS e\015\012WHERE e.Enum_Group=\"Crown Cl"
    "ass\";\015\012"
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
        dbText "Name" ="tlu_Enumerations.Sort_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Enumerations.Enum_Group"
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
    Begin
        dbText "Name" ="e.Enum_Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Enum_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Enumerations.Enum_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Sort_Order"
        dbLong "AggregateType" ="-1"
    End
End
