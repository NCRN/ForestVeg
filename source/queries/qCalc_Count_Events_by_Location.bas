Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Locations.Plot_Name"
    Alias ="Event_Count"
    Expression ="Count(tbl_Events.Event_ID)"
    Alias ="Earliest_Event_Date"
    Expression ="Min(tbl_Events.Event_Date)"
    Alias ="Most_Recent_Event_Date"
    Expression ="Max(tbl_Events.Event_Date)"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =1
End
Begin Groups
    Expression ="tbl_Locations.Location_ID"
    GroupLevel =0
    Expression ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Count"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Most_Recent_Event_Date"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Earliest_Event_Date"
        dbInteger "ColumnWidth" ="2550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =58
    Top =194
    Right =1395
    Bottom =891
    Left =-1
    Top =-1
    Right =1305
    Bottom =191
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =193
        Bottom =195
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =421
        Bottom =179
        Top =0
        Name ="tbl_Events"
        Name =""
    End
End
