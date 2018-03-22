dbMemo "SQL" ="SELECT tbl_Locations.Plot_Name, tbl_Events.Event_Date, Year([Event_Date]) AS Sam"
    "ple_Year, tbl_Locations.Panel, tbl_Locations.Frame, tbl_Tags.Tag, tbl_Tags.Tag_I"
    "D, tbl_Tags.TSN, \"Tree\" as Sampled_As,\"Tree\" as Habit, tbl_Tree_Data.Tree_St"
    "atus as Status, tbl_Tree_Data.Tree_Notes as Notes, qCalc_Basal_Area_per_Tree_w_L"
    "ist.CrownClass, qCalc_Basal_Area_per_Tree_w_List.Stems, qCalc_Basal_Area_per_Tre"
    "e_w_List.SumBasalArea_cm2, qCalc_Basal_Area_per_Tree_w_List.StemList\015\012FROM"
    " (qCalc_Basal_Area_per_Tree_w_List RIGHT JOIN ((tbl_Locations INNER JOIN tbl_Eve"
    "nts ON tbl_Locations.Location_ID = tbl_Events.Location_ID) INNER JOIN tbl_Tree_D"
    "ata ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) ON qCalc_Basal_Area_per_Tre"
    "e_w_List.Tree_Data_ID = tbl_Tree_Data.Tree_Data_ID) INNER JOIN tbl_Tags ON tbl_T"
    "ree_Data.Tag_ID = tbl_Tags.Tag_ID\015\012UNION ALL SELECT tbl_Locations.Plot_Nam"
    "e, tbl_Events.Event_Date, Year([Event_Date]) AS Sample_Year, tbl_Locations.Panel"
    ", tbl_Locations.Frame, tbl_Tags.Tag, tbl_Tags.Tag_ID, tbl_Tags.TSN, \"Sapling\","
    " tbl_Sapling_Data.Habit, tbl_Sapling_Data.Sapling_Status, tbl_Sapling_Data.Sapli"
    "ng_Notes, qCalc_Basal_Area_per_Sapling_w_List.CrownClass,qCalc_Basal_Area_per_Sa"
    "pling_w_List.Stems, qCalc_Basal_Area_per_Sapling_w_List.SumBasalArea_cm2, qCalc_"
    "Basal_Area_per_Sapling_w_List.StemList\015\012FROM (qCalc_Basal_Area_per_Sapling"
    "_w_List RIGHT JOIN ((tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Locati"
    "on_ID = tbl_Events.Location_ID) INNER JOIN tbl_Sapling_Data ON tbl_Events.Event_"
    "ID = tbl_Sapling_Data.Event_ID) ON qCalc_Basal_Area_per_Sapling_w_List.Sapling_D"
    "ata_ID = tbl_Sapling_Data.Sapling_Data_ID) INNER JOIN tbl_Tags ON tbl_Sapling_Da"
    "ta.Tag_ID = tbl_Tags.Tag_ID\015\012WHERE (((tbl_Sapling_Data.Habit)=\"Tree\"));\015"
    "\012"
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
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_w_List.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_w_List.SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_w_List.StemList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_w_List.CrownClass"
        dbLong "AggregateType" ="-1"
    End
End
