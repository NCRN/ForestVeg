Operation =1
Option =0
Where ="(((tbl_Locations.Panel)=3) AND ((qCalc_Count_Events_by_Location.Event_Count)>1))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="qCalc_Count_Events_by_Location"
    Name ="qCalc_Basal_Area_Saplings_by_Event_2008"
    Name ="qCalc_Basal_Area_Saplings_by_Event_2012"
    Name ="qCalc_Basal_Area_Trees_by_Event_2008"
    Name ="qCalc_Basal_Area_Trees_by_Event_2012"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2008.Tree_Count_2008"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2012.Tree_Count_2012"
    Alias ="Tree_Count_Change"
    Expression ="[Tree_Count_2012]-[Tree_Count_2008]"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2008.Tree_Stem_Count_2008"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2012.Tree_Stem_Count_2012"
    Alias ="Tree_Stem_Count_Change"
    Expression ="[Tree_stem_count_2012]-[tree_stem_count_2008]"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2008.Tree_BasalArea_cm2_Sum_2008"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2012.Tree_BasalArea_cm2_Sum_2012"
    Alias ="Tree_BasalArea_cm2_Change"
    Expression ="[Tree_BasalArea_cm2_Sum_2012]-[Tree_BasalArea_cm2_Sum_2008]"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2008.Sapling_Count_2008"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2012.Sapling_Count_2012"
    Alias ="Sapling_Count_Change"
    Expression ="[Sapling_Count_2012]-[Sapling_Count_2008]"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2008.Sapling_Stem_Count_2008"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2012.Sapling_Stem_Count_2012"
    Alias ="Sapling_Stem_Count_Change"
    Expression ="[Sapling_stem_count_2012]-[sapling_stem_count_2008]"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2008.Sapling_BasalArea_cm2_Sum_2008"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2012.Sapling_BasalArea_cm2_Sum_2012"
    Alias ="Sapling_BasalArea_cm2_Change"
    Expression ="[Sapling_BasalArea_cm2_Sum_2012]-[Sapling_BasalArea_cm2_Sum_2008]"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Count_Events_by_Location"
    Expression ="tbl_Locations.Location_ID = qCalc_Count_Events_by_Location.Location_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Saplings_by_Event_2008"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Saplings_by_Event_2008.Location_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Saplings_by_Event_2012"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Saplings_by_Event_2012.Location_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Trees_by_Event_2008"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Trees_by_Event_2008.Location_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Trees_by_Event_2012"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Trees_by_Event_2012.Location_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Tree_Count_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Tree_Stem_Count_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sapling_BasalArea_cm2_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3240"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sapling_Count_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="14"
    End
    Begin
        dbText "Name" ="Sapling_Stem_Count_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="17"
    End
    Begin
        dbText "Name" ="Tree_BasalArea_cm2_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2008.Tree_Count_2008"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2012.Tree_Count_2012"
        dbInteger "ColumnWidth" ="1815"
        dbInteger "ColumnOrder" ="4"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2008.Tree_Stem_Count_2008"
        dbInteger "ColumnWidth" ="2460"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2012.Tree_Stem_Count_2012"
        dbInteger "ColumnWidth" ="2460"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2008.Tree_BasalArea_cm2_Sum_2008"
        dbInteger "ColumnWidth" ="3240"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2012.Tree_BasalArea_cm2_Sum_2012"
        dbInteger "ColumnWidth" ="3240"
        dbInteger "ColumnOrder" ="10"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2008.Sapling_Count_2008"
        dbInteger "ColumnWidth" ="2175"
        dbInteger "ColumnOrder" ="12"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2012.Sapling_Count_2012"
        dbInteger "ColumnWidth" ="2175"
        dbInteger "ColumnOrder" ="13"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2008.Sapling_Stem_Count_2008"
        dbInteger "ColumnWidth" ="2715"
        dbInteger "ColumnOrder" ="15"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2012.Sapling_Stem_Count_2012"
        dbInteger "ColumnWidth" ="2715"
        dbInteger "ColumnOrder" ="16"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2008.Sapling_BasalArea_cm2_Sum_2008"
        dbInteger "ColumnWidth" ="3495"
        dbInteger "ColumnOrder" ="18"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2012.Sapling_BasalArea_cm2_Sum_2012"
        dbInteger "ColumnWidth" ="3495"
        dbInteger "ColumnOrder" ="19"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =11
    Top =96
    Right =1290
    Bottom =909
    Left =-1
    Top =-1
    Right =1247
    Bottom =446
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =303
        Bottom =270
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =406
        Top =8
        Right =740
        Bottom =152
        Top =0
        Name ="qCalc_Count_Events_by_Location"
        Name =""
    End
    Begin
        Left =717
        Top =125
        Right =985
        Bottom =269
        Top =0
        Name ="qCalc_Basal_Area_Saplings_by_Event_2008"
        Name =""
    End
    Begin
        Left =48
        Top =276
        Right =325
        Bottom =420
        Top =0
        Name ="qCalc_Basal_Area_Saplings_by_Event_2012"
        Name =""
    End
    Begin
        Left =382
        Top =288
        Right =689
        Bottom =432
        Top =0
        Name ="qCalc_Basal_Area_Trees_by_Event_2008"
        Name =""
    End
    Begin
        Left =807
        Top =299
        Right =1099
        Bottom =443
        Top =0
        Name ="qCalc_Basal_Area_Trees_by_Event_2012"
        Name =""
    End
End
