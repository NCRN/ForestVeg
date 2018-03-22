Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
    Alias ="Plot_Count"
    Expression ="Count(qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Certified)"
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
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
        dbText "Name" ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =428
        Bottom =190
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
End
