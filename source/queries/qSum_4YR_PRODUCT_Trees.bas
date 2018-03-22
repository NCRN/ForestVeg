Operation =1
Option =0
Where ="(((qCalc_PRODUCT_Trees.[Habit-Class]) Is Not Null)) OR (((qCalc_PRODUCT_Saplings"
    ".[Habit-Class]) Is Not Null)) OR (((qCalc_PRODUCT_Seedlings.[Habit-Class]) Is No"
    "t Null))"
Begin InputTables
    Name ="tlu_Plants"
    Name ="qCalc_PRODUCT_Trees"
    Name ="qCalc_PRODUCT_Saplings"
    Name ="qCalc_PRODUCT_Seedlings"
End
Begin OutputColumns
    Alias ="Latin Name"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Common Name"
    Expression ="tlu_Plants.Common"
    Expression ="tlu_Plants.TaxonCode"
    Alias ="Exotic YN"
    Expression ="IIf([Exotic],\"Y\",\"N\")"
    Alias ="% Plots with Tree Species"
    Expression ="CDbl(Nz([Perc_Plots_w_Tr_Species],0))"
    Alias ="Trees / ha"
    Expression ="CDbl(Nz([Tr_per_ha],0))"
    Alias ="Tree BA cm2 / ha"
    Expression ="CDbl(Nz([Tr_BA_per_ha],0))"
    Alias ="% Plots with Sapling Species"
    Expression ="CDbl(Nz([Perc_Plots_w_Sa_Species],0))"
    Alias ="Saplings / ha"
    Expression ="CDbl(Nz([Sa_per_ha],0))"
    Alias ="Sapling BA cm2/ ha"
    Expression ="CDbl(Nz([Sa_BA_per_ha],0))"
    Alias ="% Plots with Seedling Species"
    Expression ="CDbl(Nz([Perc_Plots_w_Se_Species],0))"
    Alias ="Seedlings / ha"
    Expression ="CDbl(Nz([Se_per_ha],0))"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_PRODUCT_Trees"
    Expression ="tlu_Plants.TSN = qCalc_PRODUCT_Trees.TSN"
    Flag =2
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_PRODUCT_Saplings"
    Expression ="tlu_Plants.TSN = qCalc_PRODUCT_Saplings.TSN"
    Flag =2
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_PRODUCT_Seedlings"
    Expression ="tlu_Plants.TSN = qCalc_PRODUCT_Seedlings.TSN"
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
        dbText "Name" ="Tree BA cm2 / ha"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trees / ha"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Saplings / ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1500"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Seedlings / ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Common Name"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latin Name"
        dbInteger "ColumnWidth" ="2505"
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
        dbText "Name" ="% Plots with Tree Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2565"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="% Plots with Seedling Species"
        dbInteger "ColumnWidth" ="2925"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="% Plots with Sapling Species"
        dbInteger "ColumnWidth" ="2820"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sapling BA cm2/ ha"
        dbInteger "ColumnWidth" ="1665"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =114
    Right =1179
    Bottom =676
    Left =-1
    Top =-1
    Right =1059
    Bottom =228
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =20
        Top =10
        Right =164
        Bottom =154
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =226
        Top =5
        Right =392
        Bottom =185
        Top =0
        Name ="qCalc_PRODUCT_Trees"
        Name =""
    End
    Begin
        Left =428
        Top =4
        Right =572
        Bottom =148
        Top =0
        Name ="qCalc_PRODUCT_Saplings"
        Name =""
    End
    Begin
        Left =609
        Top =9
        Right =753
        Bottom =153
        Top =0
        Name ="qCalc_PRODUCT_Seedlings"
        Name =""
    End
End
