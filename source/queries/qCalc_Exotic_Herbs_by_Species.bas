Operation =1
Option =0
Begin InputTables
    Name ="qCalc_Exotic_Herbs_by_Species_prequery"
End
Begin OutputColumns
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.TSN"
    Alias ="Mean_Cover_Where_Present"
    Expression ="Avg(qCalc_Exotic_Herbs_by_Species_prequery.Percent_Cover)"
    Alias ="Sum_Cover"
    Expression ="Sum(qCalc_Exotic_Herbs_by_Species_prequery.Percent_Cover)"
    Alias ="Event_Count"
    Expression ="DCount(\"[Event_ID]\",\"qFiltered_Events\")"
    Alias ="Mean_Cover_in_All_Quadrats"
    Expression ="[Sum_Cover]/([Event_Count]*12)"
End
Begin Groups
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.TSN"
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
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species_prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_Cover_Where_Present"
        dbInteger "ColumnWidth" ="2850"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sum_Cover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_Cover_in_All_Quadrats"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =260
    Right =1540
    Bottom =902
    Left =-1
    Top =-1
    Right =1478
    Bottom =349
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =273
        Bottom =334
        Top =0
        Name ="qCalc_Exotic_Herbs_by_Species_prequery"
        Name =""
    End
End
