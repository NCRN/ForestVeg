﻿dbMemo "SQL" ="SELECT l.Location_ID, e.Event_ID, l.Unit_Code, l.Unit_Group, l.Subunit_Code, l.A"
    "dmin_Unit_Code, l.Plot_Name, l.GRTS_Order, l.Install_Date, l.Panel, l.Frame, l.L"
    "ocation_Status, Year([Event_Date]) AS Event_Year, e.Event_Date, e.Protocol_Name,"
    " l.Updated_Date, e.PseudoEvent\015\012FROM tbl_Locations AS l LEFT JOIN tbl_Even"
    "ts AS e ON l.Location_ID = e.Location_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="List of sample locations and associated sampling events"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Event_Year"
        dbInteger "ColumnWidth" ="1176"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.GRTS_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Install_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.PseudoEvent"
        dbLong "AggregateType" ="-1"
    End
End
