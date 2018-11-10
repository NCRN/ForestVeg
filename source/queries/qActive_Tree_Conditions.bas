dbMemo "SQL" ="SELECT l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year([Event_Date])-2006"
    ")/4) AS Cycle, l.Panel, l.Frame, Year([Event_Date]) AS Sample_Year, e.Event_Date"
    ", CLng(Format([Event_Date],\"yyyymmdd\")) AS [Date], l.Plot_Name, t.Tag, t.Tag_S"
    "tatus, t.TSN, p.Latin_Name, td.Crown_Class, td.Tree_Status AS Status, tc.Conditi"
    "on, c.Pest, td.Tree_Notes AS Notes, e.Event_ID, l.Location_ID, t.Tag_ID, td.Tree"
    "_Data_ID, tc.Tree_Condition_ID, l.Admin_Unit_Code\015\012FROM (((((tbl_Locations"
    " AS l INNER JOIN tbl_Tags AS t ON l.Location_ID = t.Location_ID) INNER JOIN tbl_"
    "Tree_Data AS td ON t.Tag_ID = td.Tag_ID) INNER JOIN tbl_Events AS e ON e.Event_I"
    "D = td.Event_ID) INNER JOIN tbl_Tree_Conditions AS tc ON td.Tree_Data_ID = tc.Tr"
    "ee_Data_ID) INNER JOIN tlu_Plants AS p ON t.TSN = p.TSN) LEFT JOIN tlu_Tree_Cond"
    "ition AS c ON tc.Condition = c.Description\015\012ORDER BY l.Plot_Name, t.Tag, p"
    ".Latin_Name;\015\012"
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2700"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Tree_Condition.Pest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Conditions.Tree_Condition_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Conditions.Condition"
        dbInteger "ColumnWidth" ="1875"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Crown_Class"
        dbInteger "ColumnWidth" ="1560"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.Tree_Condition_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
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
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Crown_Class"
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
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tc.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="c.Pest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
