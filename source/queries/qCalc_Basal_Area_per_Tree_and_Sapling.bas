dbMemo "SQL" ="SELECT tbl_Locations.Plot_Name, tbl_Events.Event_Date, Year([Event_Date]) AS Sam"
    "ple_Year, tbl_Locations.Panel, tbl_Locations.Frame, tbl_Tags.Tag, tbl_Tags.Tag_I"
    "D, tbl_Tags.TSN,  \"Tree\" as Sampled_As,\"Tree\" as Habit, tbl_Tree_Data.Tree_S"
    "tatus as Status, tbl_Tree_Data.Tree_Notes as Notes, qCalc_Basal_Area_per_Tree.St"
    "ems, qCalc_Basal_Area_per_Tree.SumLiveBasalArea_cm2, qCalc_Basal_Area_per_Tree.S"
    "umDeadBasalArea_cm2\015\012FROM (qCalc_Basal_Area_per_Tree RIGHT JOIN ((tbl_Loca"
    "tions INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_I"
    "D) INNER JOIN tbl_Tree_Data ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) ON "
    "qCalc_Basal_Area_per_Tree.Tree_Data_ID = tbl_Tree_Data.Tree_Data_ID) INNER JOIN "
    "tbl_Tags ON tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID\015\012UNION ALL SELECT tbl_L"
    "ocations.Plot_Name, tbl_Events.Event_Date, Year([Event_Date]) AS Sample_Year, tb"
    "l_Locations.Panel, tbl_Locations.Frame, tbl_Tags.Tag, tbl_Tags.Tag_ID, tbl_Tags."
    "TSN, \"Sapling\", tbl_Sapling_Data.Habit, tbl_Sapling_Data.Sapling_Status, tbl_S"
    "apling_Data.Sapling_Notes, qCalc_Basal_Area_per_Sapling.Stems, qCalc_Basal_Area_"
    "per_Sapling.SumLiveBasalArea_cm2, qCalc_Basal_Area_per_Sapling.SumDeadBasalArea_"
    "cm2\015\012FROM (qCalc_Basal_Area_per_Sapling RIGHT JOIN ((tbl_Locations INNER J"
    "OIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) INNER JOIN"
    " tbl_Sapling_Data ON tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID) ON qCalc_B"
    "asal_Area_per_Sapling.Sapling_Data_ID = tbl_Sapling_Data.Sapling_Data_ID) INNER "
    "JOIN tbl_Tags ON tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID\015\012WHERE (((tbl_S"
    "apling_Data.Habit)=\"Tree\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sampled_As"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
End
