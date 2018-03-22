Operation =1
Option =0
Begin InputTables
    Name ="qCalc_Exotic_Herbs_by_Species_prequery"
End
Begin OutputColumns
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.TSN"
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.Location_ID"
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.Plot_Name"
End
Begin Groups
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.TSN"
    GroupLevel =0
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.Location_ID"
    GroupLevel =0
    Expression ="qCalc_Exotic_Herbs_by_Species_prequery.Plot_Name"
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
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species_prequery.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species_prequery.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =94
    Right =1553
    Bottom =971
    Left =-1
    Top =-1
    Right =1501
    Bottom =598
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =423
        Bottom =292
        Top =0
        Name ="qCalc_Exotic_Herbs_by_Species_prequery"
        Name =""
    End
End
