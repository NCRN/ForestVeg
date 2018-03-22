Operation =1
Option =0
Having ="(((tlu_Plants.Genus)=\"Fraxinus\"))"
Begin InputTables
    Name ="qActive_Trees_and_Shrubs"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qActive_Trees_and_Shrubs.Location_ID"
    Expression ="qActive_Trees_and_Shrubs.Event_ID"
    Expression ="tlu_Plants.Genus"
    Alias ="Occurences"
    Expression ="Count(qActive_Trees_and_Shrubs.Sample_ID)"
End
Begin Joins
    LeftTable ="qActive_Trees_and_Shrubs"
    RightTable ="tlu_Plants"
    Expression ="qActive_Trees_and_Shrubs.TSN = tlu_Plants.TSN"
    Flag =2
End
Begin Groups
    Expression ="qActive_Trees_and_Shrubs.Location_ID"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.Event_ID"
    GroupLevel =0
    Expression ="tlu_Plants.Genus"
    GroupLevel =0
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
        dbText "Name" ="tlu_Plants.Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Location_ID"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Occurences"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =400
    Top =14
    Right =1334
    Bottom =842
    Left =-1
    Top =-1
    Right =902
    Bottom =545
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =432
        Top =12
        Right =584
        Bottom =339
        Top =0
        Name ="qActive_Trees_and_Shrubs"
        Name =""
    End
    Begin
        Left =632
        Top =12
        Right =776
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
