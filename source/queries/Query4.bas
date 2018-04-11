dbMemo "SQL" ="PARAMETERS eid Text ( 50 );\015\012SELECT d.DBH, sd.Sapling_Data_ID, l.Plot_Name"
    ", e.Event_Date, e.Event_ID, p.Latin_Name, tg.Tag_Status, sd.Sapling_Status, sd.S"
    "tatus, tg.Azimuth, tg.Distance, tg.Microplot_Number, tg.Azimuth/tg.Distance AS A"
    "zi_Dist, tg.Tag_Notes\015\012FROM ((((tbl_Sapling_DBH AS d LEFT JOIN tbl_Sapling"
    "_Data AS sd ON sd.Sapling_Data_ID = d.Sapling_Data_ID) LEFT JOIN tbl_Events AS e"
    " ON e.Event_ID = sd.Event_ID) LEFT JOIN tbl_Locations AS l ON l.Location_ID = e."
    "Location_ID) LEFT JOIN tbl_Tags AS tg ON tg.Tag_ID = sd.Tag_ID) LEFT JOIN tlu_Pl"
    "ants AS p ON p.TSN = tg.TSN\015\012WHERE d.DBH > 10\015\012AND e.Event_ID = [eid"
    "];\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="d.Sapling_DBH_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Sapling_Data_ID"
        dbInteger "ColumnWidth" ="4155"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.DBH"
        dbInteger "ColumnWidth" ="2565"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="d.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tg.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tg.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tg.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tg.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Azi_Dist"
        dbLong "AggregateType" ="-1"
    End
End
