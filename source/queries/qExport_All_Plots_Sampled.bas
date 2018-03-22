Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Unit_Group"
    Expression ="tbl_Locations.Subunit_Code"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Locations.Location_Status"
    Expression ="tbl_Locations.Location_Notes"
    Alias ="Event_Count"
    Expression ="Count(Format([Event_Date],\"yyyymmdd\"))"
    Alias ="Event_Earliest"
    Expression ="Min(Format([Event_Date],\"yyyymmdd\"))"
    Alias ="Event_Latest"
    Expression ="Max(Format([Event_Date],\"yyyymmdd\"))"
    Alias ="UTM_18N_NAD83_X"
    Expression ="tbl_Locations.X_Coord"
    Alias ="UTM_18N_NAD83_Y"
    Expression ="tbl_Locations.Y_Coord"
    Alias ="Latitude"
    Expression ="Format([Lat_WGS84],\"0.00000000\")"
    Alias ="Longitude"
    Expression ="Format([Lon_WGS84],\"0.00000000\")"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =0
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="tbl_Locations.Unit_Group"
    GroupLevel =0
    Expression ="tbl_Locations.Subunit_Code"
    GroupLevel =0
    Expression ="tbl_Locations.Panel"
    GroupLevel =0
    Expression ="tbl_Locations.Frame"
    GroupLevel =0
    Expression ="tbl_Locations.Location_Status"
    GroupLevel =0
    Expression ="tbl_Locations.Location_Notes"
    GroupLevel =0
    Expression ="tbl_Locations.X_Coord"
    GroupLevel =0
    Expression ="tbl_Locations.Y_Coord"
    GroupLevel =0
    Expression ="Format([Lat_WGS84],\"0.00000000\")"
    GroupLevel =0
    Expression ="Format([Lon_WGS84],\"0.00000000\")"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTM_18N_NAD83_Y"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UTM_18N_NAD83_X"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Count"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Earliest"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Short Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Latest"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Short Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latitude"
        dbInteger "ColumnWidth" ="1410"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Longitude"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =23
    Top =43
    Right =1517
    Bottom =937
    Left =-1
    Top =-1
    Right =1462
    Bottom =509
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =285
        Bottom =510
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =318
        Top =13
        Right =596
        Bottom =510
        Top =0
        Name ="tbl_Events"
        Name =""
    End
End
