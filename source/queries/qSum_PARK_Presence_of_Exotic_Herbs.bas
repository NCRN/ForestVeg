Operation =1
Option =0
Begin InputTables
    Name ="qSum_EVENT_Presence_of_Exotic_Herbs"
End
Begin OutputColumns
    Expression ="qSum_EVENT_Presence_of_Exotic_Herbs.Admin_Unit_Code"
    Alias ="Plot_Count"
    Expression ="Count(qSum_EVENT_Presence_of_Exotic_Herbs.Plot_Name)"
    Alias ="Present"
    Expression ="Sum(IIf([Presence]=\"Present\",1,0))"
    Alias ="Absent"
    Expression ="Sum(IIf([Presence]=\"Absent\",1,0))"
    Alias ="Percent_Plots_Exotic"
    Expression ="Round(100*([Present]/[Plot_Count]))"
End
Begin Groups
    Expression ="qSum_EVENT_Presence_of_Exotic_Herbs.Admin_Unit_Code"
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
        dbText "Name" ="qSum_EVENT_Presence_of_Exotic_Herbs.Admin_Unit_Code"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Present"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Absent"
        dbInteger "ColumnWidth" ="990"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Count"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Percent_Plots_Exotic"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =40
    Top =504
    Right =1474
    Bottom =1011
    Left =-1
    Top =-1
    Right =1402
    Bottom =443
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =399
        Bottom =172
        Top =0
        Name ="qSum_EVENT_Presence_of_Exotic_Herbs"
        Name =""
    End
End
