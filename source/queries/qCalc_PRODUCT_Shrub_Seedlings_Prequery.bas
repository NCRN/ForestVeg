Operation =1
Option =0
Begin InputTables
    Name ="qActive_Shrub_Seedling_Data"
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
End
Begin OutputColumns
    Expression ="qActive_Shrub_Seedling_Data.Plot_Name"
    Expression ="qActive_Shrub_Seedling_Data.TSN"
    Alias ="Habit"
    Expression ="\"Shrub\""
    Alias ="Class"
    Expression ="\"Seedling\""
    Alias ="Habit-Class"
    Expression ="\"Shrub/Seedling\""
    Alias ="Samp_Count"
    Expression ="Count(qActive_Shrub_Seedling_Data.Quadrat_Seedlings_ID)"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Shrub_Seedling_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID = qActive_Shrub_Seedling_D"
        "ata.Event_ID"
    Flag =1
End
Begin Groups
    Expression ="qActive_Shrub_Seedling_Data.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Shrub_Seedling_Data.TSN"
    GroupLevel =0
    Expression ="\"Shrub\""
    GroupLevel =0
    Expression ="\"Seedling\""
    GroupLevel =0
    Expression ="\"Shrub/Seedling\""
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
        dbText "Name" ="qActive_Shrub_Seedling_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Seedling_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Seedling_Data.Quadrat_Seedlings_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfQuadrat_Seedlings_ID"
        dbLong "AggregateType" ="-1"
    End
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
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =576
    Right =1429
    Bottom =966
    Left =-1
    Top =-1
    Right =1321
    Bottom =183
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =247
        Top =16
        Right =459
        Bottom =160
        Top =0
        Name ="qActive_Shrub_Seedling_Data"
        Name =""
    End
    Begin
        Left =39
        Top =16
        Right =183
        Bottom =160
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
End
