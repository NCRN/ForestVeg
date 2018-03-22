Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Saplings_Prequery"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Saplings_Prequery.TSN"
    Expression ="qCalc_PRODUCT_Saplings_Prequery.Habit"
    Expression ="qCalc_PRODUCT_Saplings_Prequery.[Habit-Class]"
    Alias ="Plots_w_Sa_Species"
    Expression ="Count(qCalc_PRODUCT_Saplings_Prequery.Plot_Name)"
    Alias ="Sample_Count"
    Expression ="Sum(qCalc_PRODUCT_Saplings_Prequery.Samp_Count)"
    Alias ="Sapling_BA_cm2"
    Expression ="Sum(qCalc_PRODUCT_Saplings_Prequery.Sum_BA)"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Alias ="Sa_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.008482),2)"
    Alias ="Sa_BA_per_ha"
    Expression ="Round([Sapling_BA_cm2]/([Plot_Count_Total]*0.008482),2)"
    Alias ="Perc_Plots_w_Sa_Species"
    Expression ="Round(([Plots_w_Sa_Species]*100)/[Plot_Count_Total],2)"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Saplings_Prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Saplings_Prequery.Habit"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Saplings_Prequery.[Habit-Class]"
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
        dbText "Name" ="qCalc_PRODUCT_Saplings_Prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Saplings_Prequery.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Saplings_Prequery.[Habit-Class]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Sa_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sapling_BA_cm2"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count_Total"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_Sa_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sa_BA_per_ha"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sa_per_ha"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =513
    Right =1289
    Bottom =880
    Left =-1
    Top =-1
    Right =1215
    Bottom =129
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =43
        Top =-13
        Right =187
        Bottom =131
        Top =0
        Name ="qCalc_PRODUCT_Saplings_Prequery"
        Name =""
    End
End
