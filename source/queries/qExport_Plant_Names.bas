Operation =1
Option =0
Begin InputTables
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tlu_Plants.Latin_Name"
    Alias ="NCRN_Common"
    Expression ="tlu_Plants.Plants_Common"
    Expression ="tlu_Plants.Common"
    Expression ="tlu_Plants.Family"
    Expression ="tlu_Plants.Genus"
    Expression ="tlu_Plants.Species"
    Expression ="tlu_Plants.TSN"
    Expression ="tlu_Plants.Favorite"
    Expression ="tlu_Plants.Woody"
    Expression ="tlu_Plants.Herbaceous"
    Expression ="tlu_Plants.Targeted_Herb"
    Expression ="tlu_Plants.Tree"
    Expression ="tlu_Plants.Shrub"
    Expression ="tlu_Plants.Vine"
    Expression ="tlu_Plants.Exotic"
End
Begin OrderBy
    Expression ="tlu_Plants.Latin_Name"
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Family"
        dbInteger "ColumnWidth" ="1740"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Genus"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Species"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
        dbInteger "ColumnWidth" ="9030"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Tree"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Vine"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Shrub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NCRN_Common"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Favorite"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Woody"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Herbaceous"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Targeted_Herb"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =302
    Top =11
    Right =1855
    Bottom =927
    Left =-1
    Top =-1
    Right =1812
    Bottom =642
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =333
        Bottom =501
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
