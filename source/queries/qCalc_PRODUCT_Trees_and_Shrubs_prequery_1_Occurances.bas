Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Trees_and_Shrubs"
    Name ="tbl_Tree_DBH"
    Name ="tbl_Sapling_DBH"
End
Begin OutputColumns
    Expression ="qActive_Trees_and_Shrubs.Plot_Name"
    Expression ="qActive_Trees_and_Shrubs.TSN"
    Expression ="qActive_Trees_and_Shrubs.Habit"
    Expression ="qActive_Trees_and_Shrubs.Class"
    Alias ="Habit-Class"
    Expression ="[Habit] & \"/\" & [Class]"
    Alias ="Samp_Count"
    Expression ="Count(qActive_Trees_and_Shrubs.Sample_ID)"
    Alias ="Tree_BA_cm2"
    Expression ="Round(Sum((3.1415926*(([tbl_Tree_DBH]![DBH]/2)^2))),2)"
    Alias ="Sapling_BA_cm2"
    Expression ="Round(Sum((3.1415926*(([tbl_Sapling_DBH]![DBH]/2)^2))),2)"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Trees_and_Shrubs"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID=qActive_Trees_and_Shrubs.E"
        "vent_ID"
    Flag =1
    LeftTable ="qActive_Trees_and_Shrubs"
    RightTable ="tbl_Tree_DBH"
    Expression ="qActive_Trees_and_Shrubs.Sample_ID=tbl_Tree_DBH.Tree_Data_ID"
    Flag =2
    LeftTable ="qActive_Trees_and_Shrubs"
    RightTable ="tbl_Sapling_DBH"
    Expression ="qActive_Trees_and_Shrubs.Sample_ID=tbl_Sapling_DBH.Sapling_Data_ID"
    Flag =2
End
Begin Groups
    Expression ="qActive_Trees_and_Shrubs.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.TSN"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.Habit"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.Class"
    GroupLevel =0
    Expression ="[Habit] & \"/\" & [Class]"
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
        dbText "Name" ="qActive_Trees_and_Shrubs.Plot_Name"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit-Class"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Samp_Count"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_BA_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sapling_BA_cm2"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =36
    Top =70
    Right =1170
    Bottom =752
    Left =-1
    Top =-1
    Right =1102
    Bottom =399
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
        Left =302
        Top =19
        Right =491
        Bottom =241
        Top =0
        Name ="qActive_Trees_and_Shrubs"
        Name =""
    End
    Begin
        Left =583
        Top =183
        Right =727
        Bottom =327
        Top =0
        Name ="tbl_Sapling_DBH"
        Name =""
    End
    Begin
        Left =583
        Top =25
        Right =727
        Bottom =169
        Top =0
        Name ="tbl_Tree_DBH"
        Name =""
    End
End
