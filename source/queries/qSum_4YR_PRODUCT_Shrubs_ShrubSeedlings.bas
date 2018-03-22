Operation =1
Option =0
Begin InputTables
    Name ="tlu_Plants"
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List"
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_SEEDLINGS"
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUBS"
End
Begin OutputColumns
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.Common"
    Alias ="Exotic_YN"
    Expression ="IIf([Exotic],\"Y\",\"N\")"
    Alias ="Plots_w_Shrub_Species"
    Expression ="CInt(Nz([Plots_w_Sh_Species],0))"
    Alias ="Shrubs_per_ha"
    Expression ="CDbl(Nz([Sh_per_ha],0))"
    Alias ="Plots_w_Shrub_Seedling_Species"
    Expression ="CInt(Nz([Plots_w_ShSe_Species],0))"
    Alias ="Shrub_Seedlings_per_ha"
    Expression ="CDbl(Nz([ShSe_per_ha],0))"
End
Begin Joins
    LeftTable ="tlu_Plants"
    RightTable ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List"
    Expression ="tlu_Plants.TSN = qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List.TSN"
    Flag =3
    LeftTable ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List"
    RightTable ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_SEEDLINGS"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List.TSN = qCalc_PRODUCT_Trees_and_"
        "Shrubs_SHRUB_SEEDLINGS.TSN"
    Flag =2
    LeftTable ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List"
    RightTable ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUBS"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List.TSN = qCalc_PRODUCT_Trees_and_"
        "Shrubs_SHRUBS.TSN"
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Exotic_YN"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Shrub_Species"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Shrubs_per_ha"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Shrub_Seedling_Species"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Shrub_Seedlings_per_ha"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
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
        Left =280
        Top =8
        Right =424
        Bottom =152
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =16
        Top =115
        Right =193
        Bottom =259
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_Species_List"
        Name =""
    End
    Begin
        Left =484
        Top =254
        Right =821
        Bottom =366
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUB_SEEDLINGS"
        Name =""
    End
    Begin
        Left =482
        Top =140
        Right =778
        Bottom =249
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_SHRUBS"
        Name =""
    End
End
