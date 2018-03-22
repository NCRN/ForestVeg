Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Shrub_Seedlings_Prequery"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Shrub_Seedlings_Prequery.TSN"
    Expression ="qCalc_PRODUCT_Shrub_Seedlings_Prequery.[Habit-Class]"
    Alias ="Plots_w_ShSe_Species"
    Expression ="Count(qCalc_PRODUCT_Shrub_Seedlings_Prequery.Plot_Name)"
    Alias ="Sample_Count"
    Expression ="Sum(qCalc_PRODUCT_Shrub_Seedlings_Prequery.Samp_Count)"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Alias ="ShSe_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.0012),2)"
    Alias ="Perc_Plots_w_ShSe_Species"
    Expression ="Round(([Plots_w_ShSe_Species]*100)/[Plot_Count_Total],2)"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Shrub_Seedlings_Prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Shrub_Seedlings_Prequery.[Habit-Class]"
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
        dbText "Name" ="qCalc_PRODUCT_Shrub_Seedlings_Prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Shrub_Seedlings_Prequery.[Habit-Class]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_ShSe_Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ShSe_per_ha"
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
        dbText "Name" ="Perc_Plots_w_Sh_Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Shrub_Seedlings_Prequery.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfPlot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_ShSe_Species"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =59
    Top =511
    Right =1412
    Bottom =877
    Left =-1
    Top =-1
    Right =1321
    Bottom =142
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_PRODUCT_Shrub_Seedlings_Prequery"
        Name =""
    End
End
