Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Herbaceous_Data"
End
Begin OutputColumns
    Expression ="qActive_Herbaceous_Data.TSN"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Expression ="qActive_Herbaceous_Data.Exotic"
    Alias ="SumOfPercent_Cover"
    Expression ="Sum(qActive_Herbaceous_Data.Percent_Cover)"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Herbaceous_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID=qActive_Herbaceous_Data.Ev"
        "ent_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qActive_Herbaceous_Data.TSN"
    Flag =0
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Flag =0
End
Begin Groups
    Expression ="qActive_Herbaceous_Data.TSN"
    GroupLevel =0
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Exotic"
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
        dbText "Name" ="qActive_Herbaceous_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Exotic"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle].Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPercent_Cover"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =93
    Top =193
    Right =1365
    Bottom =936
    Left =-1
    Top =-1
    Right =1240
    Bottom =509
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =350
        Bottom =233
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =380
        Top =15
        Right =602
        Bottom =311
        Top =0
        Name ="qActive_Herbaceous_Data"
        Name =""
    End
End
