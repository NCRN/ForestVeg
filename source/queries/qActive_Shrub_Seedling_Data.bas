dbMemo "SQL" ="SELECT l.Panel, l.Frame, l.Plot_Name, l.Location_ID, e.Event_ID, qsd.*, p.Shrub,"
    " YEAR([e].[Event_Date]) AS Sample_Year, e.Event_Date, qd.Quadrat_Number, p.Exoti"
    "c, 1+INT((YEAR([Event_Date])-2006)/4) AS Cycle, l.Unit_Code, l.Subunit_Code, l.A"
    "dmin_Unit_Code\015\012FROM (((tbl_Locations AS l INNER JOIN tbl_Events AS e ON l"
    ".Location_ID = e.Location_ID) INNER JOIN tbl_Quadrat_Data AS qd ON e.Event_ID = "
    "qd.Event_ID) INNER JOIN tbl_Quadrat_Seedlings_Data AS qsd ON qd.Quadrat_Data_ID "
    "= qsd.Quadrat_Data_ID) INNER JOIN tlu_Plants AS p ON p.TSN = qsd.TSN\015\012WHER"
    "E p.Shrub=True;\015\012"
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
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Shrub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsd.Quadrat_Seedlings_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qd.Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsd.Quadrat_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsd.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsd.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsd.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsd.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qsd.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
