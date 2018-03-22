Operation =1
Option =0
Begin InputTables
    Name ="qCalc_PRODUCT_Herbaceous_prequery_1"
End
Begin OutputColumns
    Expression ="qCalc_PRODUCT_Herbaceous_prequery_1.TSN"
    Expression ="qCalc_PRODUCT_Herbaceous_prequery_1.Exotic"
    Alias ="Plot_Count_Present"
    Expression ="Count(qCalc_PRODUCT_Herbaceous_prequery_1.Plot_Name)"
    Alias ="SumOfSumOfPercent_Cover"
    Expression ="Sum(qCalc_PRODUCT_Herbaceous_prequery_1.SumOfPercent_Cover)"
End
Begin Groups
    Expression ="qCalc_PRODUCT_Herbaceous_prequery_1.TSN"
    GroupLevel =0
    Expression ="qCalc_PRODUCT_Herbaceous_prequery_1.Exotic"
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
        dbText "Name" ="qCalc_PRODUCT_Herbaceous_prequery_1.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_PRODUCT_Herbaceous_prequery_1.Exotic"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfSumOfPercent_Cover"
        dbInteger "ColumnWidth" ="2775"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_count_Present"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =111
    Top =553
    Right =1401
    Bottom =946
    Left =-1
    Top =-1
    Right =1258
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_PRODUCT_Herbaceous_prequery_1"
        Name =""
    End
End
