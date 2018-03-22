Operation =1
Option =0
Where ="(((qCalc_PRODUCT_Shrubs.[Habit-Class]) Is Not Null)) OR (((qCalc_PRODUCT_Shrub_S"
    "eedlings.[Habit-Class]) Is Not Null))"
Begin InputTables
    Name ="tlu_Plants"
    Name ="qCalc_PRODUCT_Shrubs"
    Name ="qCalc_PRODUCT_Shrub_Seedlings"
End
Begin OutputColumns
    Alias ="Latin Name"
    Expression ="tlu_Plants.Latin_Name"
    Alias ="Common Name"
    Expression ="tlu_Plants.Common"
    Expression ="tlu_Plants.TaxonCode"
    Alias ="Exotic YN"
    Expression ="IIf([Exotic],\"Y\",\"N\")"
    Alias ="% Plots w Shrub Species"
    Expression ="CDbl(Nz([Perc_Plots_w_Sh_Species],0))"
    Alias ="Shrubs / ha"
    Expression ="CDbl(Nz([Sh_per_ha],0))"
    Alias ="% Plots w Shrub Seedling Species"
    Expression ="CDbl(Nz([Perc_Plots_w_ShSe_Species],0))"
    Alias ="Shrub Seedlings / ha"
    Expression ="CDbl(Nz([ShSe_per_ha],0))"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_PRODUCT_Shrubs"
    Expression ="tlu_Plants.TSN = qCalc_PRODUCT_Shrubs.TSN"
    Flag =2
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_PRODUCT_Shrub_Seedlings"
    Expression ="tlu_Plants.TSN = qCalc_PRODUCT_Shrub_Seedlings.TSN"
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
        dbText "Name" ="% Plots w Shrub Species"
        dbInteger "ColumnWidth" ="2490"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="% Plots w Shrub Seedling Species"
        dbInteger "ColumnWidth" ="3270"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latin Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Common Name"
        dbInteger "ColumnWidth" ="2310"
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
        dbText "Name" ="Shrubs / ha"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Shrub Seedlings / ha"
        dbInteger "ColumnWidth" ="2550"
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
    Left =-159
    Top =205
    Right =1184
    Bottom =835
    Left =-1
    Top =-1
    Right =1311
    Bottom =217
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =695
        Bottom =208
        Top =0
        Name ="qCalc_PRODUCT_Shrubs"
        Name =""
    End
    Begin
        Left =764
        Top =33
        Right =1113
        Bottom =239
        Top =0
        Name ="qCalc_PRODUCT_Shrub_Seedlings"
        Name =""
    End
End
