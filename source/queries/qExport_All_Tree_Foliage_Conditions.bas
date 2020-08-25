dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year(e.Eve"
    "nt_Date)-2006)/4) AS Cycle, l.Panel, l.Frame, Year([Event_Date]) AS Sample_Year,"
    " CLng(Format(e.Event_Date,\"yyyymmdd\")) AS [Date], t.Tag, t.Tag_Status, tfc.Con"
    "dition, en.Enum_Description AS Condition_Description, tfc.Percent_Afflicted, t.T"
    "SN, p.Latin_Name, td.Tree_Status AS Status\015\012FROM (((((tbl_Tree_Data AS td "
    "INNER JOIN tbl_Events AS e ON e.Event_ID = td.Event_ID) INNER JOIN tbl_Tags AS t"
    " ON t.Tag_ID = td.Tag_ID) INNER JOIN tlu_Plants AS p ON t.TSN = p.TSN) INNER JOI"
    "N tbl_Locations AS l ON l.Location_ID = t.Location_ID) INNER JOIN tbl_Tree_Folia"
    "ge_Conditions AS tfc ON td.Tree_Data_ID = tfc.Tree_Data_ID) INNER JOIN tlu_Enume"
    "rations AS en ON tfc.Condition = en.Enum_Code\015\012WHERE en.Enum_Group=\"Folia"
    "ge Condition\"\015\012AND td.Tree_Status <> 'Removed from study'\015\012ORDER BY"
    " l.Plot_Name, t.Tag;\015\012"
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
        dbText "Name" ="Condition_Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2250"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="1425"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Group"
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
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
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
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tfc.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tfc.Percent_Afflicted"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
End
