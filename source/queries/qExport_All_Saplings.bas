dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year(e.Eve"
    "nt_Date)-2006)/4) AS Cycle, l.Panel, l.Frame, Year(e.Event_Date) AS Sample_Year,"
    " Format(e.Event_Date,\"yyyymmdd\") AS [Date], t.Tag, t.TSN, p.TaxonCode, p.Latin"
    "_Name, ba.StemsLive, ba.SumLiveBasalArea_cm2, ba.Equiv_Live_DBH_cm, sd.DBH_Check"
    ", sd.Sapling_Status AS Status, sd.Habit, sd.Browsable, sd.Browsed, sd.SaplingVig"
    "or, tv.TreeVigorClass AS VigorClass\015\012FROM (((((tbl_Locations AS l INNER JO"
    "IN tbl_Events AS e ON l.Location_ID = e.Location_ID) INNER JOIN tbl_Sapling_Data"
    " AS sd ON e.Event_ID = sd.Event_ID) LEFT JOIN qCalc_Basal_Area_per_Sapling AS ba"
    " ON sd.Sapling_Data_ID = ba.Sapling_Data_ID) INNER JOIN tbl_Tags AS t ON sd.Tag_"
    "ID = t.Tag_ID) LEFT JOIN tlu_Plants AS p ON t.TSN = p.TSN) LEFT JOIN tluTreeVigo"
    "r AS tv ON tv.TreeVigorCode = sd.SaplingVigor\015\012WHERE sd.Sapling_Status<>\""
    "Removed from study\" AND (sd.Habit=\"Tree\" OR sd.Habit IS NULL)\015\012ORDER BY"
    " l.Plot_Name;\015\012"
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
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Browsable"
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
        dbText "Name" ="p.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.StemsLive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.SaplingVigor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="VigorClass"
        dbLong "AggregateType" ="-1"
    End
End
