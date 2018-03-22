dbMemo "SQL" ="SELECT Plot_Name & \"_\" & Quadrat_Number & \"_\" & Sample_Year as Unique_ID, Pl"
    "ot_Name, Quadrat_Number,  \"Tree\" AS Habit, \"Seedling\" AS Class,  tlu_Plants."
    "TSN, Latin_Name, Event_Date, Sample_Year, Panel, Frame, tlu_Plants.Exotic, Locat"
    "ion_ID, Event_ID\015\012FROM qActive_Seedling_Data INNER JOIN tlu_Plants ON qAct"
    "ive_Seedling_Data.TSN = tlu_Plants.TSN\015\012UNION ALL SELECT Plot_Name & \"_\""
    " & Quadrat_Number & \"_\" & Sample_Year as Unique_ID, Plot_Name, Quadrat_Number,"
    " \"Shrub\" AS Habit, \"Seedling\" AS Class,  tlu_Plants.TSN, Latin_Name, Event_D"
    "ate, Sample_Year, Panel, Frame, tlu_Plants.Exotic, Location_ID, Event_ID\015\012"
    "FROM qActive_Shrub_Seedling_Data INNER JOIN tlu_Plants ON qActive_Shrub_Seedling"
    "_Data.TSN = tlu_Plants.TSN\015\012UNION ALL SELECT Plot_Name & \"_\" & Quadrat_N"
    "umber & \"_\" & Sample_Year as Unique_ID, Plot_Name, Quadrat_Number, \"Herb\" AS"
    " Habit, \"Herb\" AS Class,  tlu_Plants.TSN, Latin_Name, Event_Date, Sample_Year,"
    " Panel, Frame, tlu_Plants.Exotic, Location_ID, Event_ID\015\012FROM qActive_Herb"
    "aceous_Data INNER JOIN tlu_Plants ON qActive_Herbaceous_Data.TSN = tlu_Plants.TS"
    "N;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbText "Description" ="Union of all tree and shrub records including species across size class. Expecte"
    "d to be used for specimen and species counts. Not appropriate for density calcul"
    "ation without correcting for the different sample areas of Trees, saplings and s"
    "eedlings."
Begin
    Begin
        dbText "Name" ="Plot_Name"
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
        dbText "Name" ="Event_Date"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unique_ID"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latin_Name"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
