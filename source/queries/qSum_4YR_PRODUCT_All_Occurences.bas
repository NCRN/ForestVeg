Operation =1
Option =0
Begin InputTables
    Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    Name ="qActive_Trees_Shrubs_Herbs_Vines"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Alias ="Admin Unit Code"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
    Alias ="Sample Year"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    Alias ="Latin Name"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Exotic YN"
    Expression ="IIf([tlu_Plants]![Exotic]=0,\"No\",\"Yes\")"
    Alias ="Habit-Class"
    Expression ="[Habit] & \"/\" & [Class]"
    Expression ="tlu_Plants.TaxonCode"
End
Begin Joins
    LeftTable ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    RightTable ="qActive_Trees_Shrubs_Herbs_Vines"
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Event_ID = qActive_Trees_Shrubs_Her"
        "bs_Vines.Event_ID"
    Flag =1
    LeftTable ="qActive_Trees_Shrubs_Herbs_Vines"
    RightTable ="tlu_Plants"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.TSN = tlu_Plants.TSN"
    Flag =1
End
Begin OrderBy
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    Flag =0
    Expression ="tlu_Plants.Latin_Name"
    Flag =0
    Expression ="[Habit] & \"/\" & [Class]"
    Flag =0
End
Begin Groups
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Plot_Name"
    GroupLevel =0
    Expression ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    GroupLevel =0
    Expression ="tlu_Plants.Latin_Name"
    GroupLevel =0
    Expression ="IIf([tlu_Plants]![Exotic]=0,\"No\",\"Yes\")"
    GroupLevel =0
    Expression ="[Habit] & \"/\" & [Class]"
    GroupLevel =0
    Expression ="tlu_Plants.TaxonCode"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit-Class"
        dbInteger "ColumnWidth" ="1455"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic YN"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latin Name"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample Year"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Admin Unit Code"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =10
    Top =213
    Right =958
    Bottom =775
    Left =-1
    Top =-1
    Right =916
    Bottom =279
    Left =0
    Top =96
    ColumnsShown =543
    Begin
        Left =9
        Top =-81
        Right =312
        Bottom =147
        Top =0
        Name ="qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
        Name =""
    End
    Begin
        Left =388
        Top =-81
        Right =632
        Bottom =171
        Top =0
        Name ="qActive_Trees_Shrubs_Herbs_Vines"
        Name =""
    End
    Begin
        Left =719
        Top =-84
        Right =921
        Bottom =242
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
