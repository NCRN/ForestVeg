Operation =1
Option =0
Begin InputTables
    Name ="tlu_Plants"
    Name ="qCalc_Exotic_Herbs_by_Species"
    Name ="qCalc_Exotic_Herbs_by_Species_Plot_Count"
End
Begin OutputColumns
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.PLANTS_Common"
    Expression ="tlu_Plants.Exotic"
    Expression ="qCalc_Exotic_Herbs_by_Species_Plot_Count.Plot_Count"
    Expression ="qCalc_Exotic_Herbs_by_Species.Event_Count"
    Expression ="qCalc_Exotic_Herbs_by_Species.Mean_Cover_Where_Present"
    Alias ="Mean_Cover_Where_Present_in_Plot"
    Expression ="[Sum_Cover]/(12*[Plot_Count])"
    Expression ="qCalc_Exotic_Herbs_by_Species.Mean_Cover_in_All_Quadrats"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_Exotic_Herbs_by_Species"
    Expression ="tlu_Plants.TSN = qCalc_Exotic_Herbs_by_Species.TSN"
    Flag =1
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_Exotic_Herbs_by_Species_Plot_Count"
    Expression ="tlu_Plants.TSN = qCalc_Exotic_Herbs_by_Species_Plot_Count.TSN"
    Flag =1
End
Begin OrderBy
    Expression ="tlu_Plants.Latin_Name"
    Flag =0
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="tlu_Plants.PLANTS_Common"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="Mean_Cover_Where_Present_in_Plot"
        dbInteger "ColumnWidth" ="2385"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species_Plot_Count.Plot_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species.Event_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species.Mean_Cover_Where_Present"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species.Mean_Cover_in_All_Quadrats"
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
    ColumnsShown =539
    Begin
        Left =48
        Top =3
        Right =192
        Bottom =397
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =615
        Bottom =156
        Top =0
        Name ="qCalc_Exotic_Herbs_by_Species"
        Name =""
    End
    Begin
        Left =249
        Top =168
        Right =554
        Bottom =312
        Top =0
        Name ="qCalc_Exotic_Herbs_by_Species_Plot_Count"
        Name =""
    End
End
