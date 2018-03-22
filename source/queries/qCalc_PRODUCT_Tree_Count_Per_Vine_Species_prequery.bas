Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Vine_Data"
End
Begin OutputColumns
    Expression ="qActive_Vine_Data.TSN"
    Expression ="qActive_Vine_Data.Latin_Name"
    Expression ="qActive_Vine_Data.Exotic"
    Expression ="qActive_Vine_Data.Location_ID"
    Expression ="qActive_Vine_Data.Tag_ID"
    Expression ="qActive_Vine_Data.Tree_Data_ID"
    Alias ="Vine_Count"
    Expression ="Count(qActive_Vine_Data.Tree_Vine_ID)"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Vine_Data"
    Expression ="[qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle].Event_ID=qActive_Vine_Data.Event_"
        "ID"
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
    Expression ="qActive_Vine_Data.Exotic"
    GroupLevel =0
    Expression ="qActive_Vine_Data.Location_ID"
    GroupLevel =0
    Expression ="qActive_Vine_Data.Tag_ID"
    GroupLevel =0
    Expression ="qActive_Vine_Data.Tree_Data_ID"
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
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vine_Count"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Tag_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1530"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Vine_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4185"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =63
    Top =196
    Right =1286
    Bottom =947
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
        Right =333
        Bottom =236
        Top =0
        Name ="qSum_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =701
        Top =14
        Right =889
        Bottom =376
        Top =0
        Name ="qActive_Vine_Data"
        Name =""
    End
End
