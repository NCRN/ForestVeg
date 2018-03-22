Operation =1
Option =0
Begin InputTables
    Name ="qCalc_Basal_Area_per_Tree"
    Name ="tbl_Stem_List"
End
Begin OutputColumns
    Expression ="qCalc_Basal_Area_per_Tree.*"
    Expression ="tbl_Stem_List.StemList"
End
Begin Joins
    LeftTable ="qCalc_Basal_Area_per_Tree"
    RightTable ="tbl_Stem_List"
    Expression ="qCalc_Basal_Area_per_Tree.Tree_Data_ID = tbl_Stem_List.Tree_Data_ID"
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
        dbText "Name" ="qCalc_Basal_Area_per_Tree.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.FirstOfTag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Stem_List.StemList"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.tbl_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.tbl_Tree_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.StemsLive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.StemsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.CrownClass"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-3
    Top =100
    Right =1519
    Bottom =785
    Left =-1
    Top =-1
    Right =1490
    Bottom =197
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_Basal_Area_per_Tree"
        Name =""
    End
    Begin
        Left =325
        Top =14
        Right =469
        Bottom =158
        Top =0
        Name ="tbl_Stem_List"
        Name =""
    End
End
