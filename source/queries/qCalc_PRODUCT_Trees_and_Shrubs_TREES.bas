Operation =1
Option =0
Having ="(((qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.[Habit-Class])=\"Tree/Tr"
    "ee\"))"
Begin InputTables
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.TSN"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Habit"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.[Habit-Class]"
    Alias ="Plots_w_Tr_Species"
    Expression ="Count(qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Plot_Name)"
    Alias ="Sample_Count"
    Expression ="Sum(qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Samp_Count)"
    Alias ="SumOfTree_BA_cm2"
    Expression ="Sum(qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Tree_BA_cm2)"
    Alias ="SumOfSapling_BA_cm2"
    Expression ="Sum(qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Sapling_BA_cm2)"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Alias ="Tr_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.070686),2)"
    Alias ="Tr_BA_per_ha"
    Expression ="Round([SumOfTree_BA_cm2]/([Plot_Count_Total]*0.070686),2)"
    Alias ="Perc_Plots_w_Tr_Species"
    Expression ="Round(([Plots_w_Tr_Species]*100)/[Plot_Count_Total],2)"
End
Begin Joins
    LeftTable ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances"
    RightTable ="tlu_Plants"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.TSN=tlu_Plants.TSN"
    Flag =1
End
Begin Groups
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Habit"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.[Habit-Class]"
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
        dbText "Name" ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.[Habit-Class]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Count"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count_Total"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfTree_BA_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSapling_BA_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Tr_Species"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tr_per_ha"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tr_BA_per_ha"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_Tr_Species"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =12
    Top =39
    Right =1452
    Bottom =501
    Left =-1
    Top =-1
    Right =1408
    Bottom =189
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =399
        Bottom =286
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances"
        Name =""
    End
    Begin
        Left =447
        Top =12
        Right =674
        Bottom =496
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
