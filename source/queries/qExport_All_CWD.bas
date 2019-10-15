dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year(e.Eve"
    "nt_Date)-2006)/4) AS Cycle, l.Panel, l.Frame, Year(e.Event_Date) AS Sample_Year,"
    " CLng(Format(e.Event_Date,\"yyyymmdd\")) AS [Date], cwd.Transect_Azimuth, cwd.TS"
    "N, cwd.Decay_Class, cwd.Diameter, cwd.Hollow, cwd.CWD_Notes AS CWD_Notes, p.Lati"
    "n_Name\015\012FROM ((tbl_Locations AS l RIGHT JOIN tbl_Events AS e ON l.Location"
    "_ID = e.Location_ID) INNER JOIN tbl_CWD_Data AS cwd ON e.Event_ID = cwd.Event_ID"
    ") LEFT JOIN tlu_Plants AS p ON cwd.TSN = p.TSN\015\012ORDER BY l.Plot_Name;\015\012"
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
        dbText "Name" ="CWD_Notes"
        dbInteger "ColumnOrder" ="16"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnOrder" ="8"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cwd.Transect_Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cwd.Diameter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cwd.Decay_Class"
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
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
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
        dbText "Name" ="cwd.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cwd.Hollow"
        dbLong "AggregateType" ="-1"
    End
End
