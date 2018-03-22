Operation =1
Option =0
Where ="(((qExport_All_Plots.Panel)=3))"
Begin InputTables
    Name ="qExport_All_Plots"
End
Begin OutputColumns
    Alias ="Plot"
    Expression ="\"'\" & [Plot_Name] & \"'\""
    Alias ="Date"
    Expression ="\"'\" & [Event_Earliest] & \"'\""
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
        dbText "Name" ="Plot"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =6
    Top =15
    Right =1273
    Bottom =949
    Left =-1
    Top =-1
    Right =1235
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =34
        Top =4
        Right =258
        Bottom =148
        Top =0
        Name ="qExport_All_Plots"
        Name =""
    End
End
