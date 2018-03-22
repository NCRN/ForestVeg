Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List"
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_TREES"
    Name ="tlu_Plants"
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_SAPLINGS"
    Name ="qCalc_PRODUCT_Trees_and_Shrubs_SEEDLINGS"
End
Begin OutputColumns
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.Common"
    Expression ="tlu_Plants.TaxonCode"
    Alias ="Exotic_YN"
    Expression ="IIf([Exotic],\"Y\",\"N\")"
    Alias ="Plots_w_Tree_Species"
    Expression ="CInt(Nz([Plots_w_Tr_Species],0))"
    Alias ="Trees_per_ha"
    Expression ="CDbl(Nz([Tr_per_ha],0))"
    Alias ="Tree_BA_cm2_per_ha"
    Expression ="CDbl(Nz([Tr_BA_per_ha],0))"
    Alias ="Plots_w_Sapling_Species"
    Expression ="CInt(Nz([Plots_w_Sa_Species],0))"
    Alias ="Saplings_per_ha"
    Expression ="CDbl(Nz([Sa_per_ha],0))"
    Alias ="Sapling_BA_per_ha"
    Expression ="CDbl(Nz([Sa_BA_per_ha],0))"
    Alias ="Plots_w_Seedling_Species"
    Expression ="CInt(Nz([Plots_w_Se_Species],0))"
    Alias ="Seedlings_per_ha"
    Expression ="CDbl(Nz([Se_per_ha],0))"
End
Begin Joins
    LeftTable ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List"
    RightTable ="qCalc_PRODUCT_Trees_and_Shrubs_TREES"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List.TSN = qCalc_PRODUCT_Trees_and_S"
        "hrubs_TREES.TSN"
    Flag =2
    LeftTable ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List"
    RightTable ="tlu_Plants"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List"
    RightTable ="qCalc_PRODUCT_Trees_and_Shrubs_SAPLINGS"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List.TSN = qCalc_PRODUCT_Trees_and_S"
        "hrubs_SAPLINGS.TSN"
    Flag =2
    LeftTable ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List"
    RightTable ="qCalc_PRODUCT_Trees_and_Shrubs_SEEDLINGS"
    Expression ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List.TSN = qCalc_PRODUCT_Trees_and_S"
        "hrubs_SEEDLINGS.TSN"
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
        dbText "Name" ="Trees_per_ha"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Tree_Species"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic_YN"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plots_w_Sapling_Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Saplings_per_ha"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tree_BA_cm2_per_ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sapling_BA_per_ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Plots_w_Seedling_Species"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Seedlings_per_ha"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =5
    Top =3
    Right =1562
    Bottom =565
    Left =-1
    Top =-1
    Right =1525
    Bottom =301
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =103
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_TREE_Species_List"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =501
        Bottom =113
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_TREES"
        Name =""
    End
    Begin
        Left =238
        Top =121
        Right =382
        Bottom =265
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =538
        Top =6
        Right =843
        Bottom =116
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_SAPLINGS"
        Name =""
    End
    Begin
        Left =868
        Top =90
        Right =1012
        Bottom =234
        Top =0
        Name ="qCalc_PRODUCT_Trees_and_Shrubs_SEEDLINGS"
        Name =""
    End
End
