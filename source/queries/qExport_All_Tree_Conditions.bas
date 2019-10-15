dbMemo "SQL" ="SELECT atc.Plot_Name, atc.Unit_Code, atc.Unit_Group, atc.Subunit_Code, 1+Int((Ye"
    "ar(e.Event_Date)-2006)/4) AS Cycle, atc.Panel, atc.Frame, atc.Sample_Year, CLng("
    "Format(e.Event_Date,\"yyyymmdd\")) AS [Date], atc.Tag, atc.TSN, atc.Latin_Name, "
    "atc.Crown_Class, atc.Status, atc.Condition, atc.Pest\015\012FROM qActive_Tree_Co"
    "nditions AS atc\015\012ORDER BY atc.Plot_Name, atc.Sample_Year, atc.Tag;\015\012"
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
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Pest"
        dbLong "AggregateType" ="-1"
    End
End
