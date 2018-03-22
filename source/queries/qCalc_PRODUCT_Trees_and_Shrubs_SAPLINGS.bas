Operation =1
Option =0
Having ="(((qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.[Habit-Class])=\"Tree/Sa"
    "pling\"))"
Begin InputTables
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.TSN"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Habit"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.[Habit-Class]"
    Alias ="Plots_w_Sa_Species"
    Expression ="Count(qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Plot_Name)"
    Alias ="Sample_Count"
    Expression ="Sum(qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Samp_Count)"
    Alias ="SumOfSapling_BA_cm2"
    Expression ="Sum(qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances.Sapling_BA_cm2)"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Alias ="Sa_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.008482),2)"
    Alias ="Sa_BA_per_ha"
    Expression ="Round([SumOfSapling_BA_cm2]/([Plot_Count_Total]*0.008482),2)"
    Alias ="Perc_Plots_w_Sa_Species"
    Expression ="Round(([Plots_w_Sa_Species]*100)/[Plot_Count_Total],2)"
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
        dbText "Name" ="SumOfSapling_BA_cm2"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Sa_Species"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sa_per_ha"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sa_BA_per_ha"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_Sa_Species"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =34
    Top =11
    Right =1260
    Bottom =442
    Left =-1
    Top =-1
    Right =1194
    Bottom =165
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =42
        Top =7
        Right =393
        Bottom =164
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_1_Occurances"
        Name =""
    End
    Begin
        Left =445
        Top =2
        Right =672
        Bottom =486
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
