Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Shrubs_Prequery"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Shrubs_Prequery.TSN"
    Expression ="qCalc_PRODUCT_Shrubs_Prequery.Habit"
    Expression ="qCalc_PRODUCT_Shrubs_Prequery.[Habit-Class]"
    Alias ="Plots_w_Sh_Species"
    Expression ="Count(qCalc_PRODUCT_Shrubs_Prequery.Plot_Name)"
    Alias ="Sample_Count"
    Expression ="Sum(qCalc_PRODUCT_Shrubs_Prequery.Samp_Count)"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Alias ="Sh_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.008482),2)"
    Alias ="Perc_Plots_w_Sh_Species"
    Expression ="Round(([Plots_w_Sh_Species]*100)/[Plot_Count_Total],2)"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Shrubs_Prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Shrubs_Prequery.Habit"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Shrubs_Prequery.[Habit-Class]"
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
        dbText "Name" ="qCalc_PRODUCT_Shrubs_Prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Shrubs_Prequery.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Shrubs_Prequery.[Habit-Class]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Sh_Species"
        dbInteger "ColumnWidth" ="2190"
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
        dbText "Name" ="Plot_Count_Total"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sh_per_ha"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_Sh_Species"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =558
    Right =1299
    Bottom =934
    Left =-1
    Top =-1
    Right =1215
    Bottom =123
    Left =9
    Top =0
    ColumnsShown =543
    Begin
        Left =32
        Top =0
        Right =176
        Bottom =144
        Top =0
        Name ="qCalc_PRODUCT_Shrubs_Prequery"
        Name =""
    End
End
