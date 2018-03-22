Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Herbaceous_prequery_2"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Herbaceous_prequery_2.TSN"
    Expression ="qCalc_PRODUCT_Herbaceous_prequery_2.Plot_Count_Present"
    Alias ="Plot_Count_Total"
    Expression ="CInt(DCount(\"[Event_ID]\",\"qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle\"))"
    Expression ="qCalc_PRODUCT_Herbaceous_prequery_2.SumOfSumOfPercent_Cover"
    Alias ="Percent_Cover_in_All_Plots"
    Expression ="Round([SumOfSumOfPercent_Cover]/([Plot_Count_Total]*12),2)"
    Alias ="Percent_Cover_Where_Present"
    Expression ="Round([SumOfSumOfPercent_Cover]/([Plot_Count_Present]*12),2)"
    Alias ="Perc_Plots_w_Herb_Species"
    Expression ="Round(([Plot_Count_Present]*100)/[Plot_Count_Total],2)"
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
        dbText "Name" ="Percent_Cover_Where_Present"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3030"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Herbaceous_prequery_2.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Herbaceous_prequery_2.Plot_Count_Present"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Herbaceous_prequery_2.SumOfSumOfPercent_Cover"
        dbInteger "ColumnWidth" ="2775"
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
        dbText "Name" ="Percent_Cover_in_All_Plots"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Perc_Plots_w_Herb_Species"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =187
    Top =174
    Right =1384
    Bottom =626
    Left =-1
    Top =-1
    Right =1165
    Bottom =155
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =14
        Top =12
        Right =273
        Bottom =135
        Top =0
        Name ="qCalc_PRODUCT_Herbaceous_prequery_2"
        Name =""
    End
End
