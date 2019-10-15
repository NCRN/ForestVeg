dbMemo "SQL" ="SELECT l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year([Event_Date])-2006"
    ")/4) AS Cycle, l.Panel, l.Frame, Year(e.Event_Date) AS Sample_Year, e.Event_Date"
    ", CLng(Format(e.Event_Date,\"yyyymmdd\")) AS [Date], l.Plot_Name, t.Tag, t.Tag_S"
    "tatus, t.TSN, p.Latin_Name, sd.Sapling_Status AS Status, sc.Condition, tc.Pest, "
    "sd.Sapling_Notes AS Notes, e.Event_ID, l.Location_ID, t.Tag_ID, sd.Sapling_Data_"
    "ID, sc.Sapling_Condition_ID, l.Admin_Unit_Code\015\012FROM (((((tbl_Locations AS"
    " l INNER JOIN tbl_Tags AS t ON l.Location_ID = t.Location_ID) INNER JOIN tbl_Sap"
    "ling_Data AS sd ON t.Tag_ID = sd.Tag_ID) INNER JOIN tbl_Events AS e ON e.Event_I"
    "D = sd.Event_ID) INNER JOIN tbl_Sapling_Conditions AS sc ON sd.Sapling_Data_ID ="
    " sc.Sapling_Data_ID) INNER JOIN tlu_Plants AS p ON t.TSN = p.TSN) LEFT JOIN tlu_"
    "Tree_Condition AS tc ON sc.Condition = tc.Description\015\012WHERE sd.Sapling_St"
    "atus <> 'Removed from study'\015\012ORDER BY l.Plot_Name, t.Tag, p.Latin_Name;\015"
    "\012"
dbMemo "Connect" =""
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
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
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
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.Pest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sc.Sapling_Condition_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes"
        dbLong "AggregateType" ="-1"
    End
End
