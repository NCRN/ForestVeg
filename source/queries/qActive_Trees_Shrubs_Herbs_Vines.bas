dbMemo "SQL" ="SELECT TSN, Plot_Name, \"Tree\" AS Habit, \"Tree\" AS Class, Event_Date, Sample_"
    "Year, Location_ID, Event_ID\015\012FROM qActive_Tree_Data\015\012UNION ALL SELEC"
    "T TSN, Plot_Name, \"Tree\" AS Habit, \"Sapling\" AS Class, Event_Date, Sample_Ye"
    "ar, Location_ID, Event_ID\015\012FROM qActive_Sapling_Data\015\012UNION ALL SELE"
    "CT TSN, Plot_Name, \"Tree\" AS Habit, \"Seedling\" AS Class, Event_Date, Sample_"
    "Year, Location_ID, Event_ID\015\012FROM qActive_Seedling_Data\015\012UNION ALL S"
    "ELECT TSN, Plot_Name, \"Shrub\" AS Habit, \"Shrub\" AS Class, Event_Date, Sample"
    "_Year, Location_ID, Event_ID\015\012FROM qActive_Shrub_Data\015\012UNION ALL SEL"
    "ECT TSN, Plot_Name, \"Shrub\" AS Habit, \"Seedling\" AS Class, Event_Date, Sampl"
    "e_Year, Location_ID, Event_ID\015\012FROM qActive_Shrub_Seedling_Data\015\012UNI"
    "ON ALL SELECT TSN, Plot_Name, \"Herb\" AS Habit, \"Herb\" AS Class, Event_Date, "
    "Sample_Year, Location_ID, Event_ID\015\012FROM qActive_Herbaceous_Data\015\012UN"
    "ION ALL SELECT TSN, Plot_Name, \"Vine\" AS Habit, \"Vine\" AS Class, Event_Date,"
    " Sample_Year, Location_ID, Event_ID\015\012FROM qActive_Vine_Data;\015\012"
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
        dbText "Name" ="TSN"
        dbLong "AggregateType" ="-1"
    End
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
End
