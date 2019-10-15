dbMemo "SQL" ="SELECT td.Tag_ID, e.Event_Date, Year(e.Event_Date) AS Event_Year, 1+Int((Year(e."
    "Event_Date)-2006)/4) AS Cycle, \015\012l.Panel, \"Tree\" AS Class, td.[Tree_Stat"
    "us] AS Status, \015\012\"Tree - \" & [Tree_Status] AS Class_Status, td.[Tree_Not"
    "es] AS Notes\015\012FROM ((tbl_Locations l \015\012LEFT JOIN tbl_Events e ON l.L"
    "ocation_ID = e.Location_ID)\015\012LEFT JOIN tbl_Tree_Data td ON e.Event_ID = td"
    ".Event_ID) \015\012WHERE e.Event_ID IS NOT NULL\015\012AND td.Tree_Status <> 'Re"
    "moved from study'\015\012\015\012UNION ALL SELECT sd.Tag_ID, e.Event_Date, Year("
    "e.Event_Date),\015\012 1+Int((Year(e.Event_Date)-2006)/4) AS Cycle, \015\012 l.P"
    "anel, \"Sapling\" AS Class, sd.[Sapling_Status] AS Status, \015\012 \"Sapling - "
    "\" & [Sapling_Status] AS Class_Status, sd.[Sapling_Notes] AS Notes\015\012FROM ("
    "(tbl_Locations l\015\012LEFT JOIN tbl_Events e ON l.Location_ID = e.Location_ID)"
    "\015\012LEFT JOIN tbl_Sapling_Data sd ON e.Event_ID = sd.Event_ID)\015\012WHERE "
    "e.Event_ID IS NOT NULL\015\012AND sd.Sapling_Status <> 'Removed from study';\015"
    "\012"
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
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class_Status"
        dbInteger "ColumnWidth" ="2715"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Year"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
End
