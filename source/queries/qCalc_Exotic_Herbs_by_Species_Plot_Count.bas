Operation =1
Option =0
Begin InputTables
    Name ="qCalc_Exotic_Herbs_by_Species_Plot_Count_prequery"
End
Begin OutputColumns
    Expression ="qCalc_Exotic_Herbs_by_Species_Plot_Count_prequery.TSN"
    Alias ="Plot_Count"
    Expression ="Count(qCalc_Exotic_Herbs_by_Species_Plot_Count_prequery.[Location_ID])"
End
Begin Groups
    Expression ="qCalc_Exotic_Herbs_by_Species_Plot_Count_prequery.TSN"
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
        dbText "Name" ="qCalc_Exotic_Herbs_by_Species_Plot_Count_prequery.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =30
    Top =60
    Right =1555
    Bottom =937
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
        Right =192
        Bottom =156
        Top =0
        Name ="qCalc_Exotic_Herbs_by_Species_Plot_Count_prequery"
        Name =""
    End
End
