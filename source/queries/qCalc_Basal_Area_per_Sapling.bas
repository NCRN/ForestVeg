dbMemo "SQL" ="SELECT sd.Sapling_Data_ID, sd.Event_ID, \"\" AS CrownClass, Count(dbh.DBH) AS St"
    "ems, Sum(IIf([Live]=True,1,0)) AS StemsLive, Sum(IIf([Live]=False,1,0)) AS Stems"
    "Dead, Round(Sum(3.1415926*(([DBH]/2)^2)),1) AS SumBasalArea_cm2, First(sd.Tag_ID"
    ") AS FirstOfTag_ID, Round((([SumBasalArea_cm2]/3.1415)^0.5)*2,1) AS Equiv_DBH_cm"
    ", Round(Sum(3.1415926*(((IIf([Live]=True,[DBH],0))/2)^2)),1) AS SumLiveBasalArea"
    "_cm2, Round(Sum(3.1415926*(((IIf([Live]=False,[DBH],0))/2)^2)),1) AS SumDeadBasa"
    "lArea_cm2, Round((([SumLiveBasalArea_cm2]/3.1415)^0.5)*2,1) AS Equiv_Live_DBH_cm"
    ", Round((([SumDeadBasalArea_cm2]/3.1415)^0.5)*2,1) AS Equiv_Dead_DBH_cm, sd.Sapl"
    "ing_Status, e.Event_Date, l.Plot_Name\015\012FROM ((tbl_Sapling_Data AS sd LEFT "
    "JOIN tbl_Sapling_DBH AS dbh ON sd.Sapling_Data_ID = dbh.Sapling_Data_ID) LEFT JO"
    "IN tbl_Events AS e ON e.Event_ID = sd.Event_ID) LEFT JOIN tbl_Locations AS l ON "
    "l.Location_ID = e.Location_ID\015\012GROUP BY sd.Sapling_Data_ID, sd.Event_ID, \""
    "\", sd.Sapling_Status, e.Event_Date, l.Plot_Name;\015\012"
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
dbMemo "Filter" ="([qCalc_Basal_Area_per_Sapling].[Stems]=0)"
dbMemo "OrderBy" ="[qCalc_Basal_Area_per_Sapling].[Event_Date] DESC, [qCalc_Basal_Area_per_Sapling]"
    ".[Sapling_Status]"
Begin
    Begin
        dbText "Name" ="Stems"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="FirstOfTag_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1800"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StemsDead"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StemsLive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CrownClass"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Status"
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
End
