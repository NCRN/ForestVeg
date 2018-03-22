Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Sapling_Data"
    Name ="tbl_Tags"
    Name ="qCalc_Basal_Area_per_Sapling"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Expression ="qActive_Sapling_Data.TSN"
    Alias ="Habit"
    Expression ="\"Tree\""
    Alias ="Class"
    Expression ="\"Sapling\""
    Alias ="Habit-Class"
    Expression ="\"Tree/Sapling\""
    Alias ="Samp_Count"
    Expression ="Count(qActive_Sapling_Data.Sapling_Data_ID)"
    Alias ="Sum_BA"
    Expression ="Sum(qCalc_Basal_Area_per_Sapling.SumBasalArea_cm2)"
End
Begin Joins
    LeftTable ="qActive_Sapling_Data"
    RightTable ="tbl_Tags"
    Expression ="qActive_Sapling_Data.Tag_ID=tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Sapling_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID=qActive_Sapling_Data.Event"
        "_ID"
    Flag =1
    LeftTable ="qActive_Sapling_Data"
    RightTable ="qCalc_Basal_Area_per_Sapling"
    Expression ="qActive_Sapling_Data.Sapling_Data_ID=qCalc_Basal_Area_per_Sapling.Sapling_Data_I"
        "D"
    Flag =2
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Sapling_Data.TSN"
    GroupLevel =0
    Expression ="\"Tree\""
    GroupLevel =0
    Expression ="\"Sapling\""
    GroupLevel =0
    Expression ="\"Tree/Sapling\""
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
        dbText "Name" ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Sapling_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Samp_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sum_BA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit-Class"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =56
    Top =494
    Right =1303
    Bottom =951
    Left =-1
    Top =-1
    Right =1215
    Bottom =231
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qActive_Sapling_Data"
        Name =""
    End
    Begin
        Left =442
        Top =11
        Right =586
        Bottom =155
        Top =0
        Name ="qCalc_Basal_Area_per_Sapling"
        Name =""
    End
    Begin
        Left =604
        Top =15
        Right =748
        Bottom =159
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
