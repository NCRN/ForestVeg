Operation =1
Option =0
Where ="(((qCalc_Basal_Area_per_Tree_and_Sapling_w_List.Sample_Year)=[Forms]![frm_Switch"
    "board]![cTimeframe]-4))"
Begin InputTables
    Name ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List"
End
Begin OutputColumns
    Expression ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.*"
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
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.Sample_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.Sampled_As"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="9"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.Habit"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="11"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.Notes"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="13"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.qCalc_Basal_Area_per_Tree_w_List.St"
            "ems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.qCalc_Basal_Area_per_Tree_w_List.Su"
            "mBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.qCalc_Basal_Area_per_Tree_w_List.St"
            "emList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List.qCalc_Basal_Area_per_Tree_w_List.Cr"
            "ownClass"
        dbInteger "ColumnOrder" ="12"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =27
    Top =0
    Right =1416
    Bottom =746
    Left =-1
    Top =-1
    Right =1357
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =15
        Top =15
        Right =332
        Bottom =238
        Top =0
        Name ="qCalc_Basal_Area_per_Tree_and_Sapling_w_List"
        Name =""
    End
End
