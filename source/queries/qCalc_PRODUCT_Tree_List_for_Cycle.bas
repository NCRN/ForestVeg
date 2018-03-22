Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Tree_Data"
End
Begin OutputColumns
    Alias ="Tree_Count_for_4_Year_Cycle"
    Expression ="qActive_Tree_Data.Tree_Data_ID"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Tree_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID = qActive_Tree_Data.Event_"
        "ID"
    Flag =1
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
        dbText "Name" ="Tree_Count_for_4_Year_Cycle"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =132
    Right =1478
    Bottom =997
    Left =-1
    Top =-1
    Right =1370
    Bottom =531
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =235
        Top =20
        Right =399
        Bottom =415
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
End
