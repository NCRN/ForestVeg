dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year([Even"
    "t_Date])-2006)/4) AS Cycle, l.Panel, l.Frame, Year(e.Event_Date) AS Sample_Year,"
    " CLng(Format(e.Event_Date,\"yyyymmdd\")) AS [Date], qd.Quadrat_Number, qd.Percen"
    "t_Trees, qd.Percent_Bryophytes, qd.Percent_Rock, qd.Percent_Woody_Debris AS Perc"
    "ent_Coarse_Woody_Debris, qd.Percent_Fine_Woody_Debris, qd.Percent_Other, qd.Perc"
    "ent_Grasses, qd.Percent_Sedges, qd.Percent_Herbs, qd.Percent_Ferns, qd.Quadrat_N"
    "otes\015\012FROM (tbl_Locations AS l INNER JOIN tbl_Events AS e ON l.Location_ID"
    " = e.Location_ID) INNER JOIN tbl_Quadrat_Data AS qd ON e.Event_ID = qd.Event_ID\015"
    "\012ORDER BY l.Plot_Name, qd.Quadrat_Number;\015\012"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Trees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Bryophytes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Rock"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Other"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Grasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Sedges"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Herbs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Ferns"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Notes"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="Date"
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
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Percent_Coarse_Woody_Debris"
        dbInteger "ColumnWidth" ="2295"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Fine_Woody_Debris"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Ferns"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Rock"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Bryophytes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
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
        dbText "Name" ="qd.Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Trees"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Fine_Woody_Debris"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Other"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Grasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Sedges"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Percent_Herbs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Quadrat_Notes"
        dbLong "AggregateType" ="-1"
    End
End
