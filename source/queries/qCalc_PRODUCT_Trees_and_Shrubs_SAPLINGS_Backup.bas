Operation =1
Option =0
Where ="(((qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.[Habit-Class])=\"Tree/Sap"
    "ling\"))"
Begin InputTables
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.TSN"
    Alias ="Plots_w_Sa_Species"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.Plot_Count_Present"
    Alias ="Sa_per_ha"
    Expression ="Round([Sample_Count]/([Plot_Count_Total]*0.008482),2)"
    Alias ="Sa_BA_per_ha"
    Expression ="Round([SumOfSapling_BA_cm2]/([Plot_Count_Total]*0.008482),2)"
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
        dbText "Name" ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sa_per_ha"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Sa_Species"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sa_BA_per_ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =75
    Top =170
    Right =1326
    Bottom =771
    Left =-1
    Top =-1
    Right =1219
    Bottom =121
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =226
        Bottom =194
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_prequery_2_PlotCount"
        Name =""
    End
End
