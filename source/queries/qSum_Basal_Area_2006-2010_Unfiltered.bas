Operation =1
Option =0
Where ="(((qCalc_Count_Events_by_Location.Event_Count)>1))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="qCalc_Count_Events_by_Location"
    Name ="qCalc_Basal_Area_Saplings_by_Event_2006"
    Name ="qCalc_Basal_Area_Saplings_by_Event_2010"
    Name ="qCalc_Basal_Area_Trees_by_Event_2006"
    Name ="qCalc_Basal_Area_Trees_by_Event_2010"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2006.Tree_Count_2006"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2010.Tree_Count_2010"
    Alias ="Tree_Count_Change"
    Expression ="[Tree_Count_2010]-[Tree_Count_2006]"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2006.Tree_Stem_Count_2006"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2010.Tree_Stem_Count_2010"
    Alias ="Tree_Stem_Count_Change"
    Expression ="[Tree_stem_count_2010]-[tree_stem_count_2006]"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2006.Tree_BasalArea_cm2_Sum_2006"
    Expression ="qCalc_Basal_Area_Trees_by_Event_2010.Tree_BasalArea_cm2_Sum_2010"
    Alias ="Tree_BasalArea_cm2_Change"
    Expression ="[Tree_BasalArea_cm2_Sum_2010]-[Tree_BasalArea_cm2_Sum_2006]"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2006.Sapling_Count_2006"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2010.Sapling_Count_2010"
    Alias ="Sapling_Count_Change"
    Expression ="[Sapling_Count_2010]-[Sapling_Count_2006]"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2006.Sapling_Stem_Count_2006"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2010.Sapling_Stem_Count_2010"
    Alias ="Sapling_Stem_Count_Change"
    Expression ="[Sapling_stem_count_2010]-[sapling_stem_count_2006]"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2006.Sapling_BasalArea_cm2_Sum_2006"
    Expression ="qCalc_Basal_Area_Saplings_by_Event_2010.Sapling_BasalArea_cm2_Sum_2010"
    Alias ="Sapling_BasalArea_cm2_Change"
    Expression ="[Sapling_BasalArea_cm2_Sum_2010]-[Sapling_BasalArea_cm2_Sum_2006]"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Count_Events_by_Location"
    Expression ="tbl_Locations.Location_ID = qCalc_Count_Events_by_Location.Location_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Saplings_by_Event_2006"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Saplings_by_Event_2006.Location_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Saplings_by_Event_2010"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Saplings_by_Event_2010.Location_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Trees_by_Event_2006"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Trees_by_Event_2006.Location_ID"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="qCalc_Basal_Area_Trees_by_Event_2010"
    Expression ="tbl_Locations.Location_ID = qCalc_Basal_Area_Trees_by_Event_2010.Location_ID"
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
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2006.Tree_Count_2006"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2006.Tree_Stem_Count_2006"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2006.Tree_BasalArea_cm2_Sum_2006"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
        dbInteger "ColumnWidth" ="3240"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2010.Tree_Count_2010"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2010.Tree_Stem_Count_2010"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Trees_by_Event_2010.Tree_BasalArea_cm2_Sum_2010"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
        dbInteger "ColumnWidth" ="3240"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2006.Sapling_Count_2006"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2006.Sapling_Stem_Count_2006"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="14"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2006.Sapling_BasalArea_cm2_Sum_2006"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="17"
        dbInteger "ColumnWidth" ="3495"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2010.Sapling_Count_2010"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="12"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2010.Sapling_Stem_Count_2010"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="15"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_Saplings_by_Event_2010.Sapling_BasalArea_cm2_Sum_2010"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="18"
        dbInteger "ColumnWidth" ="3495"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Tree_Count_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Tree_Stem_Count_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
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
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="Sapling_Stem_Count_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2955"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="16"
    End
    Begin
        dbText "Name" ="Tree_BasalArea_cm2_Change"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="10"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =-1
    Top =-1
    Right =690
    Bottom =270
    Left =691
    Top =0
    ColumnsShown =539
    Begin
        Left =-643
        Top =12
        Right =-499
        Bottom =270
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =-327
        Top =1
        Right =7
        Bottom =145
        Top =0
        Name ="qCalc_Count_Events_by_Location"
        Name =""
    End
    Begin
        Left =-436
        Top =217
        Right =-201
        Bottom =445
        Top =0
        Name ="qCalc_Basal_Area_Saplings_by_Event_2006"
        Name =""
    End
    Begin
        Left =-77
        Top =231
        Right =104
        Bottom =382
        Top =0
        Name ="qCalc_Basal_Area_Saplings_by_Event_2010"
        Name =""
    End
    Begin
        Left =124
        Top =185
        Right =410
        Bottom =363
        Top =0
        Name ="qCalc_Basal_Area_Trees_by_Event_2006"
        Name =""
    End
    Begin
        Left =332
        Top =12
        Right =476
        Bottom =156
        Top =0
        Name ="qCalc_Basal_Area_Trees_by_Event_2010"
        Name =""
    End
End
