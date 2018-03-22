Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.TSN"
    Expression ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Latin_Name"
    Expression ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Exotic"
    Alias ="Tree_Count"
    Expression ="Count(qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Tree_Data_ID)"
    Alias ="Vine_Count"
    Expression ="Sum(qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Vine_Count)"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Latin_Name"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Exotic"
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
        dbText "Name" ="Tree_Count"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vine_Count"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery.Exotic"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =110
    Right =1444
    Bottom =975
    Left =-1
    Top =-1
    Right =1370
    Bottom =480
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_PRODUCT_Tree_Count_Per_Vine_Species_prequery"
        Name =""
    End
End
