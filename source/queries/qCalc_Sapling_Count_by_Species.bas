Operation =1
Option =0
Begin InputTables
    Name ="tlu_Plants"
    Name ="qActive_Sapling_Data"
    Name ="qCalc_Basal_Area_per_Sapling"
    Name ="qFiltered_Events"
    Name ="qFiltered_Locations"
End
Begin OutputColumns
    Expression ="tlu_Plants.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.PLANTS_Common"
    Alias ="Sapling_Count"
    Expression ="Count(qActive_Sapling_Data.Sapling_Data_ID)"
    Alias ="Sapling_SumBasalArea_cm2"
    Expression ="Sum(Round([SumBasalArea_cm2]))"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="qActive_Sapling_Data"
    Expression ="tlu_Plants.TSN = qActive_Sapling_Data.TSN"
    Flag =1
    LeftTable ="qActive_Sapling_Data"
    RightTable ="qCalc_Basal_Area_per_Sapling"
    Expression ="qActive_Sapling_Data.Sapling_Data_ID = qCalc_Basal_Area_per_Sapling.Sapling_Data"
        "_ID"
    Flag =1
    LeftTable ="qActive_Sapling_Data"
    RightTable ="qFiltered_Events"
    Expression ="qActive_Sapling_Data.Event_ID = qFiltered_Events.Event_ID"
    Flag =1
    LeftTable ="qActive_Sapling_Data"
    RightTable ="qFiltered_Locations"
    Expression ="qActive_Sapling_Data.Location_ID = qFiltered_Locations.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tlu_Plants.Latin_Name"
    Flag =0
End
Begin Groups
    Expression ="tlu_Plants.TSN"
    GroupLevel =0
    Expression ="tlu_Plants.Latin_Name"
    GroupLevel =0
    Expression ="tlu_Plants.PLANTS_Common"
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
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.PLANTS_Common"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sapling_Count"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sapling_SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2865"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =125
    Top =21
    Right =1076
    Bottom =931
    Left =-1
    Top =-1
    Right =919
    Bottom =392
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =424
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =227
        Top =14
        Right =371
        Bottom =362
        Top =0
        Name ="qActive_Sapling_Data"
        Name =""
    End
    Begin
        Left =566
        Top =103
        Right =710
        Bottom =243
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
    Begin
        Left =563
        Top =262
        Right =707
        Bottom =406
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =714
        Top =15
        Right =858
        Bottom =159
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
End
