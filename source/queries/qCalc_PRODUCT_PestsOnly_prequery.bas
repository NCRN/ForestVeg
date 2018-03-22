Operation =1
Option =0
Having ="(((tlu_Tree_Condition.Pest)=True))"
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_Conditions"
    Name ="tlu_Tree_Condition"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Alias ="CountOfTree_Condition_ID"
    Expression ="Count(tbl_Tree_Conditions.Tree_Condition_ID)"
    Expression ="tlu_Tree_Condition.Pest"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="tbl_Tree_Data"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_Conditions"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Conditions.Tree_Data_ID"
    Flag =1
    LeftTable ="tbl_Tree_Conditions"
    RightTable ="tlu_Tree_Condition"
    Expression ="tbl_Tree_Conditions.Condition = tlu_Tree_Condition.Description"
    Flag =1
End
Begin OrderBy
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Flag =0
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    GroupLevel =0
    Expression ="tlu_Tree_Condition.Pest"
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
        dbText "Name" ="tbl_Tree_Conditions.Condition"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfTree_Condition_ID"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Tree_Condition.Category"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Tree_Condition.Pest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =39
    Top =14
    Right =1479
    Bottom =925
    Left =-1
    Top =-1
    Right =1408
    Bottom =467
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =365
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =276
        Top =19
        Right =420
        Bottom =261
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =617
        Top =86
        Right =761
        Bottom =230
        Top =0
        Name ="tbl_Tree_Conditions"
        Name =""
    End
    Begin
        Left =839
        Top =87
        Right =983
        Bottom =231
        Top =0
        Name ="tlu_Tree_Condition"
        Name =""
    End
End
