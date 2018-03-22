dbMemo "SQL" ="SELECT tbl_Tree_Data.Tag_ID, tbl_Events.Event_Date, Year(tbl_Events.Event_Date) "
    "AS Event_Year, 1+Int((Year([Event_Date])-2006)/4) AS Cycle, tbl_Locations.Panel,"
    " \"Tree\" AS Class, tbl_Tree_Data.[Tree_Status] AS Status, \"Tree - \" & [Tree_S"
    "tatus] AS Class_Status, tbl_Tree_Data.[Tree_Notes] AS Notes\015\012FROM tbl_Loca"
    "tions INNER JOIN (tbl_Events INNER JOIN tbl_Tree_Data ON tbl_Events.Event_ID = t"
    "bl_Tree_Data.Event_ID) ON tbl_Locations.Location_ID = tbl_Events.Location_ID\015"
    "\012\015\012UNION ALL SELECT tbl_Sapling_Data.Tag_ID, tbl_Events.Event_Date, Yea"
    "r(tbl_Events.Event_Date), 1+Int((Year([Event_Date])-2006)/4) AS Cycle, tbl_Locat"
    "ions.Panel, \"Sapling\" AS Class, tbl_Sapling_Data.[Sapling_Status] AS Status, \""
    "Sapling - \" & [Sapling_Status] AS Class_Status, tbl_Sapling_Data.[Sapling_Notes"
    "] AS Notes\015\012FROM tbl_Locations INNER JOIN (tbl_Events INNER JOIN tbl_Sapli"
    "ng_Data ON tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID) ON tbl_Locations.Loc"
    "ation_ID = tbl_Events.Location_ID;\015\012"
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
        dbText "Name" ="tbl_Tree_Data.Tag_ID"
        dbInteger "ColumnWidth" ="4230"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
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
End
