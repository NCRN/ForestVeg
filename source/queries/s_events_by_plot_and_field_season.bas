dbMemo "SQL" ="SELECT e.Event_ID, e.PseudoEvent, e.Event_Date, Year(e.Event_Date) AS FieldSeaso"
    "n, l.Plot_Name, 0 AS RecordPicker\015\012FROM tbl_Locations AS l INNER JOIN tbl_"
    "Events AS e ON l.Location_ID =e.Location_ID\015\012ORDER BY e.Event_Date, l.Plot"
    "_Name;\015\012"
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
        dbText "Name" ="e.PseudoEvent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FieldSeason"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RecordPicker"
        dbLong "AggregateType" ="-1"
    End
End
