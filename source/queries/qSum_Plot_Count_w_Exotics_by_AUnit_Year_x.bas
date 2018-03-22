Operation =6
Option =0
Begin InputTables
    Name ="qCalc_Exotic_Plot_Count"
End
Begin OutputColumns
    Expression ="qCalc_Exotic_Plot_Count.Admin_Unit_Code"
    GroupLevel =2
    Expression ="qCalc_Exotic_Plot_Count.Sample_Year"
    GroupLevel =1
    Alias ="FirstOfExoticPlots_and_Plots"
    Expression ="First(qCalc_Exotic_Plot_Count.ExoticPlots_and_Plots)"
End
Begin Groups
    Expression ="qCalc_Exotic_Plot_Count.Admin_Unit_Code"
    GroupLevel =2
    Expression ="qCalc_Exotic_Plot_Count.Sample_Year"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Count of Plots that contain any exotics (trees, shrubs, vines, herbs) and Total "
    "plot count crosstabulated by Admin Unit and Sampling Year."
Begin
    Begin
        dbText "Name" ="qCalc_Exotic_Plot_Count.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qCalc_Exotic_Plot_Count.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Exotic_Plot_Count.Percent_Plots_Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2007"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2009"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2010"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPercent_Plots_Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfExoticPlots_and_Plots"
    End
    Begin
        dbText "Name" ="2011"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_Exotic_Plot_Count"
        Name =""
    End
End
