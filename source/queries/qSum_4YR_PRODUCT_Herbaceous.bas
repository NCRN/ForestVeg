Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Herbaceous"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Alias ="Latin Name"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Common Name"
    Expression ="tlu_Plants.Common"
    Alias ="Exotic YN"
    Expression ="IIf([Exotic],\"Y\",\"N\")"
    Alias ="% Plots w Species"
    Expression ="Round([Perc_Plots_w_Herb_Species],2)"
    Alias ="% Cover in All Plots"
    Expression ="qCalc_PRODUCT_Herbaceous.Percent_Cover_in_All_Plots"
End
Begin Joins
    LeftTable ="qCalc_PRODUCT_Herbaceous"
    RightTable ="tlu_Plants"
    Expression ="qCalc_PRODUCT_Herbaceous.TSN = tlu_Plants.TSN"
    Flag =2
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
        dbText "Name" ="% Plots w Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Latin Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Common Name"
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic YN"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="% Cover in All Plots"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =-1
    Top =-1
    Right =690
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =8
        Top =10
        Right =238
        Bottom =189
        Top =0
        Name ="qCalc_PRODUCT_Herbaceous"
        Name =""
    End
    Begin
        Left =291
        Top =14
        Right =435
        Bottom =158
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
