Operation =1
Option =0
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Locations.Unit_Code"
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="qFiltered_Locations.Panel"
    Expression ="qFiltered_Locations.Frame"
    Expression ="qFiltered_Events.Event_Date"
    Alias ="Event_Year"
    Expression ="CInt([qFiltered_Events].[Event_Year])"
    Expression ="qFiltered_Events.Certified"
    Expression ="qFiltered_Locations.Location_ID"
    Expression ="qFiltered_Events.Event_ID"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qFiltered_Locations.Plot_Name"
    Flag =0
    Expression ="qFiltered_Events.Event_Date"
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
        dbText "Name" ="qFiltered_Events_Cycle.Event_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="10"
    End
    Begin
        dbText "Name" ="qFiltered_Events_Cycle.Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="qFiltered_Events_Cycle.Certified"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Location_ID"
        dbInteger "ColumnWidth" ="1695"
        dbInteger "ColumnOrder" ="9"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1350"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Unit_Code"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Panel"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Frame"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Year"
        dbInteger "ColumnWidth" ="1350"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_Date"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_ID"
        dbInteger "ColumnOrder" ="10"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Certified"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =9
    Top =96
    Right =1024
    Bottom =658
    Left =-1
    Top =-1
    Right =983
    Bottom =328
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =8
        Top =17
        Right =152
        Bottom =161
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =218
        Top =12
        Right =362
        Bottom =156
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
End
