Operation =1
Option =0
Begin InputTables
    Name ="qActive_Shrub_Data"
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Expression ="qActive_Shrub_Data.TSN"
    Alias ="Habit"
    Expression ="\"Shrub\""
    Alias ="Class"
    Expression ="\"Shrub\""
    Alias ="Habit-Class"
    Expression ="\"Shrub/Shrub\""
    Alias ="Samp_Count"
    Expression ="Count(qActive_Shrub_Data.Sapling_Data_ID)"
End
Begin Joins
    LeftTable ="qActive_Shrub_Data"
    RightTable ="tbl_Tags"
    Expression ="qActive_Shrub_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Shrub_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID = qActive_Shrub_Data.Event"
        "_ID"
    Flag =1
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Shrub_Data.TSN"
    GroupLevel =0
    Expression ="\"Shrub\""
    GroupLevel =0
    Expression ="\"Shrub\""
    GroupLevel =0
    Expression ="\"Shrub/Shrub\""
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
        dbText "Name" ="Samp_Count"
        dbInteger "ColumnWidth" ="1515"
        dbBoolean "ColumnHidden" ="0"
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
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =51
    Top =493
    Right =1298
    Bottom =916
    Left =-1
    Top =-1
    Right =1215
    Bottom =215
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =232
        Top =21
        Right =376
        Bottom =183
        Top =0
        Name ="qActive_Shrub_Data"
        Name =""
    End
    Begin
        Left =18
        Top =13
        Right =189
        Bottom =214
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
