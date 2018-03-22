Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="tbl_Tags"
    Name ="qActive_Tree_Data"
    Name ="qCalc_Basal_Area_per_Tree"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Expression ="tbl_Tags.TSN"
    Alias ="Habit"
    Expression ="\"Tree\""
    Alias ="Class"
    Expression ="\"Tree\""
    Alias ="Habit-Class"
    Expression ="\"Tree/Tree\""
    Alias ="Samp_Count"
    Expression ="Count(qActive_Tree_Data.Tree_Data_ID)"
    Alias ="Sum_BA"
    Expression ="Sum(qCalc_Basal_Area_per_Tree.SumBasalArea_cm2)"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="qActive_Tree_Data"
    Expression ="tbl_Tags.Tag_ID=qActive_Tree_Data.Tag_ID"
    Flag =3
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Tree_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID=qActive_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="qActive_Tree_Data"
    RightTable ="qCalc_Basal_Area_per_Tree"
    Expression ="qActive_Tree_Data.Tree_Data_ID=qCalc_Basal_Area_per_Tree.Tree_Data_ID"
    Flag =2
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    GroupLevel =0
    Expression ="tbl_Tags.TSN"
    GroupLevel =0
    Expression ="\"Tree\""
    GroupLevel =0
    Expression ="\"Tree\""
    GroupLevel =0
    Expression ="\"Tree/Tree\""
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
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit-Class"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Samp_Count"
        dbInteger "ColumnWidth" ="1515"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sum_BA"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =50
    Top =42
    Right =1263
    Bottom =479
    Left =-1
    Top =-1
    Right =1181
    Bottom =156
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =16
        Top =19
        Right =160
        Bottom =207
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =408
        Top =92
        Right =585
        Bottom =237
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =607
        Top =15
        Right =829
        Bottom =161
        Top =0
        Name ="qCalc_Basal_Area_per_Tree"
        Name =""
    End
    Begin
        Left =213
        Top =20
        Right =357
        Bottom =164
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
End
