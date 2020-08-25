dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year(e.Eve"
    "nt_Date)-2006)/4) AS Cycle, l.Panel, l.Frame, Year(e.Event_Date) AS Sample_Year,"
    " CLng(Format(e.Event_Date,\"yyyymmdd\")) AS [Date], t.Tag_Status, tv.TSN, pe.Tax"
    "onCode, pe.Latin_Name, t.Tag AS Host_Tag, t.TSN AS Host_TSN, ph.Latin_Name AS Ho"
    "st_Latin_Name, td.Tree_Status AS Host_Status, vic.Condition, pe.Exotic\015\012FR"
    "OM ((((((tbl_Tree_Data AS td INNER JOIN tbl_Events AS e ON e.Event_ID = td.Event"
    "_ID) INNER JOIN tbl_Tree_Vines AS tv ON td.Tree_Data_ID = tv.Tree_Data_ID) INNER"
    " JOIN tbl_Locations AS l ON l.Location_ID = e.Location_ID) INNER JOIN tbl_Tags A"
    "S t ON td.Tag_ID = t.Tag_ID) INNER JOIN tlu_Plants AS ph ON t.TSN = ph.TSN) LEFT"
    " JOIN qCalc_Trees_with_Vines_in_Crown AS vic ON td.Tree_Data_ID = vic.Tree_Data_"
    "ID) INNER JOIN tlu_Plants AS pe ON tv.TSN = pe.TSN\015\012WHERE td.Tree_Status <"
    "> 'Removed from study';\015\012"
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
        dbText "Name" ="Host_TSN"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Tag"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tv.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pe.Exotic"
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
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pe.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="pe.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="vic.Condition"
        dbLong "AggregateType" ="-1"
    End
End
