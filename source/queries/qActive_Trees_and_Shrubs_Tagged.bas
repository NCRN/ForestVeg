﻿dbMemo "SQL" ="SELECT TSN, Plot_Name, \"Tree\" AS Habit, \"Tree\" AS Class, Event_Date, Sample_"
    "Year, Location_ID, Event_ID, Tag_ID, Tree_Status as Status\015\012FROM qActive_T"
    "ree_Data\015\012UNION ALL SELECT TSN, Plot_Name, \"Tree\" AS Habit, \"Sapling\" "
    "AS Class, Event_Date, Sample_Year, Location_ID, Event_ID, Tag_ID, Sapling_Status"
    "\015\012FROM qActive_Sapling_Data\015\012UNION ALL SELECT TSN, Plot_Name, \"Shru"
    "b\" AS Habit, \"Shrub\" AS Class, Event_Date, Sample_Year, Location_ID, Event_ID"
    ", Tag_ID, Sapling_Status\015\012FROM qActive_Shrub_Data;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
End
