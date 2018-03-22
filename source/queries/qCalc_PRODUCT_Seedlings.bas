Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Seedlings_prequery"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Seedlings_prequery.TSN"
    Expression ="qCalc_PRODUCT_Seedlings_prequery.Habit"
    Expression ="qCalc_PRODUCT_Seedlings_prequery.[Habit-Class]"
    Alias ="Plot_w_Se_Species"
    Expression ="Count(qCalc_PRODUCT_Seedlings_prequery.Plot_Name)"
    Alias ="Sample_Count"
    Expression ="Sum(qCalc_PRODUCT_Seedlings_prequery.Samp_Count)"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Alias ="Plots_w_Se_Species"
    Expression ="[Plot_w_Se_Species]"
    Alias ="Se_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.0012),2)"
    Alias ="Perc_Plots_w_Se_Species"
    Expression ="Round(([Plots_w_Se_Species]*100)/[Plot_Count_Total],2)"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Seedlings_prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Seedlings_prequery.Habit"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Seedlings_prequery.[Habit-Class]"
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
        dbText "Name" ="qCalc_PRODUCT_Seedlings_prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Seedlings_prequery.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Seedlings_prequery.[Habit-Class]"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_w_Se_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count_Total"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Se_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_Se_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Se_per_ha"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =30
    Top =437
    Right =1277
    Bottom =891
    Left =-1
    Top =-1
    Right =1215
    Bottom =277
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_PRODUCT_Seedlings_prequery"
        Name =""
    End
End
