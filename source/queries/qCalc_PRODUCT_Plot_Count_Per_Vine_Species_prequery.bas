Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Vine_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qActive_Vine_Data.TSN"
    Expression ="qActive_Vine_Data.Latin_Name"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="qActive_Vine_Data.Exotic"
    Expression ="qActive_Vine_Data.Location_ID"
    Alias ="Vine_Count"
    Expression ="Count(qActive_Vine_Data.Tree_Vine_ID)"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Vine_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID = qActive_Vine_Data.Event_"
        "ID"
    Flag =1
    LeftTable ="qActive_Vine_Data"
    RightTable ="tlu_Plants"
    Expression ="qActive_Vine_Data.TSN = tlu_Plants.TSN"
    Flag =1
End
Begin OrderBy
    Expression ="qActive_Vine_Data.Latin_Name"
    Flag =0
End
Begin Groups
    Expression ="qActive_Vine_Data.TSN"
    GroupLevel =0
    Expression ="qActive_Vine_Data.Latin_Name"
    GroupLevel =0
    Expression ="tlu_Plants.TaxonCode"
    GroupLevel =0
    Expression ="qActive_Vine_Data.Exotic"
    GroupLevel =0
    Expression ="qActive_Vine_Data.Location_ID"
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
        dbText "Name" ="qActive_Vine_Data.Latin_Name"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Exotic"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Location_ID"
        dbInteger "ColumnWidth" ="4170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vine_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =113
    Top =209
    Right =1336
    Bottom =960
    Left =-1
    Top =-1
    Right =1191
    Bottom =407
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =231
        Bottom =297
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =407
        Top =15
        Right =644
        Bottom =377
        Top =0
        Name ="qActive_Vine_Data"
        Name =""
    End
    Begin
        Left =846
        Top =-11
        Right =1017
        Bottom =367
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
