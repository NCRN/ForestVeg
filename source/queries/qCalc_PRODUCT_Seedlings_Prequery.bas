Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Seedling_Data"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
    Expression ="qActive_Seedling_Data.Plot_Name"
    Expression ="qActive_Seedling_Data.TSN"
    Alias ="Habit"
    Expression ="\"Tree\""
    Alias ="Class"
    Expression ="\"Seedling\""
    Alias ="Habit-Class"
    Expression ="\"Tree/Seedling\""
    Alias ="Samp_Count"
    Expression ="Count(qActive_Seedling_Data.Quadrat_Seedlings_ID)"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Seedling_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID = qActive_Seedling_Data.Ev"
        "ent_ID"
    Flag =1
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qActive_Seedling_Data.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Seedling_Data.TSN"
    GroupLevel =0
    Expression ="\"Tree\""
    GroupLevel =0
    Expression ="\"Seedling\""
    GroupLevel =0
    Expression ="\"Tree/Seedling\""
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
        dbText "Name" ="qActive_Seedling_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Seedling_Data.TSN"
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
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =54
    Top =445
    Right =1263
    Bottom =921
    Left =-1
    Top =-1
    Right =1177
    Bottom =233
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =362
        Bottom =240
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =433
        Top =8
        Right =620
        Bottom =294
        Top =0
        Name ="qActive_Seedling_Data"
        Name =""
    End
End
