Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Trees_Prequery"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Trees_Prequery.TSN"
    Expression ="qCalc_PRODUCT_Trees_Prequery.Habit"
    Expression ="qCalc_PRODUCT_Trees_Prequery.[Habit-Class]"
    Alias ="Plots_w_Tr_Species"
    Expression ="Count(qCalc_PRODUCT_Trees_Prequery.Plot_Name)"
    Alias ="Sample_Count"
    Expression ="Sum(qCalc_PRODUCT_Trees_Prequery.Samp_Count)"
    Alias ="Tree_BA_cm2"
    Expression ="Sum(qCalc_PRODUCT_Trees_Prequery.Sum_BA)"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Alias ="Tr_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.070686),2)"
    Alias ="Tr_BA_per_ha"
    Expression ="Round([Tree_BA_cm2]/([Plot_Count_Total]*0.070686),2)"
    Alias ="Perc_Plots_w_Tr_Species"
    Expression ="Round(([Plots_w_Tr_Species]*100)/[Plot_Count_Total],2)"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Trees_Prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Trees_Prequery.Habit"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Trees_Prequery.[Habit-Class]"
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
        dbText "Name" ="qCalc_PRODUCT_Trees_Prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Trees_Prequery.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Trees_Prequery.[Habit-Class]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count_Total"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_Tr_Species"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tr_per_ha"
        dbInteger "ColumnWidth" ="1275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tr_BA_per_ha"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Count"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_BA_cm2"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Tr_Species"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =32
    Top =423
    Right =1241
    Bottom =931
    Left =-1
    Top =-1
    Right =1177
    Bottom =258
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =204
        Bottom =187
        Top =0
        Name ="qCalc_PRODUCT_Trees_Prequery"
        Name =""
    End
End
