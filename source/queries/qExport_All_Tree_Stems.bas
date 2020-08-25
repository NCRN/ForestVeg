dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year([Even"
    "t_Date])-2006)/4) AS Cycle, l.Panel, l.Frame, Year([Event_Date]) AS Sample_Year,"
    " CLng(Format([e].[Event_Date],\"yyyymmdd\")) AS [Date], t.Tag, t.TSN, p.Latin_Na"
    "me, td.Tree_Status AS Status, dbh.Live, dbh.DBH, td.DBH_Check, \"Tree\" AS Habit"
    ", \"Tree\" AS Class\015\012FROM ((((tbl_Locations AS l RIGHT JOIN tbl_Events AS "
    "e ON l.Location_ID = e.Location_ID) INNER JOIN tbl_Tree_Data AS td ON e.Event_ID"
    " = td.Event_ID) INNER JOIN tbl_Tags AS t ON t.Tag_ID = td.Tag_ID) LEFT JOIN tlu_"
    "Plants AS p ON t.TSN = p.TSN) INNER JOIN tbl_Tree_DBH AS dbh ON td.Tree_Data_ID "
    "= dbh.Tree_Data_ID\015\012ORDER BY t.Tag;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[qExport_All_Tree_Stems].[Plot_Name], [qExport_All_Tree_Stems].[Sample_Year]"
Begin
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="p.Latin_Name"
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
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbh.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbh.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
End
