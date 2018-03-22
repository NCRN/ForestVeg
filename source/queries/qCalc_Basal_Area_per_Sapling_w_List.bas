Operation =1
Option =0
Begin InputTables
    Name ="qCalc_Basal_Area_per_Sapling"
    Name ="tbl_Stem_List"
End
Begin OutputColumns
    Expression ="qCalc_Basal_Area_per_Sapling.*"
    Expression ="tbl_Stem_List.StemList"
End
Begin Joins
    LeftTable ="qCalc_Basal_Area_per_Sapling"
    RightTable ="tbl_Stem_List"
    Expression ="qCalc_Basal_Area_per_Sapling.Sapling_Data_ID = tbl_Stem_List.Tree_Data_ID"
    Flag =2
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
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.FirstOfTag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Stem_List.StemList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.tbl_Sapling_Data.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.tbl_Sapling_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.StemsLive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.StemsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Sapling.Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-14
    Top =173
    Right =1470
    Bottom =703
    Left =-1
    Top =-1
    Right =1452
    Bottom =268
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
    Begin
        Left =228
        Top =17
        Right =372
        Bottom =161
        Top =0
        Name ="tbl_Stem_List"
        Name =""
    End
End
