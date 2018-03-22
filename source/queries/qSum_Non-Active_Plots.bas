Operation =1
Option =0
Where ="(((qFiltered_Locations.Location_Status)<>\"Active\"))"
Begin InputTables
    Name ="qFiltered_Locations"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Location_Status"
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Locations.Unit_Code"
    Expression ="qFiltered_Locations.Subunit_Code"
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="qFiltered_Locations.X_Coord"
    Expression ="qFiltered_Locations.Y_Coord"
    Expression ="qFiltered_Locations.Location_Notes"
    Expression ="qFiltered_Locations.Panel"
    Expression ="qFiltered_Locations.Frame"
    Expression ="qFiltered_Locations.GRTS_Order"
    Expression ="qFiltered_Locations.Install_Date"
    Expression ="qFiltered_Locations.Location_ID"
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
dbText "Description" ="Query returns a list of plots that have been retired or that have not yet been s"
    "ampled"
Begin
    Begin
        dbText "Name" ="qFiltered_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.X_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Y_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Location_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Location_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.GRTS_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Install_Date"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =327
        Bottom =520
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
End
