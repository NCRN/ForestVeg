dbMemo "SQL" ="SELECT sd.Sapling_Data_ID\015\012FROM (((tbl_Sapling_Data AS sd LEFT JOIN tbl_Ev"
    "ents AS e ON e.Event_ID = sd.Event_ID) LEFT JOIN tbl_Locations AS l ON l.Locatio"
    "n_ID = e.Location_ID) LEFT JOIN tbl_Tags AS t ON t.Tag_ID = sd.Tag_ID) LEFT JOIN"
    " tlu_Plants AS p ON p.TSN = t.TSN\015\012WHERE sd.Habit IS NULL\015\012AND e.Pse"
    "udoEvent = 0\015\012AND p.Shrub = True\015\012ORDER BY e.Event_Date DESC , l.Plo"
    "t_Name, t.Tag;\015\012"
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
        dbText "Name" ="sd.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
End
