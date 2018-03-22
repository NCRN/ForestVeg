Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.TSN"
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Latin_Name"
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.TaxonCode"
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Exotic"
    Alias ="Plot_Count"
    Expression ="Count(qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Location_ID)"
    Alias ="Vine_Count"
    Expression ="Sum(qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Vine_Count)"
End
Begin OrderBy
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Latin_Name"
    Flag =0
End
Begin Groups
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Latin_Name"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.TaxonCode"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Exotic"
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
        dbText "Name" ="Plot_Count"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vine_Count"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.Exotic"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery].Latin_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2625"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery].Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =171
    Top =51
    Right =1248
    Bottom =962
    Left =-1
    Top =-1
    Right =1045
    Bottom =509
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =357
        Bottom =212
        Top =0
        Name ="qCalc_PRODUCT_Plot_Count_Per_Vine_Species_prequery"
        Name =""
    End
End
