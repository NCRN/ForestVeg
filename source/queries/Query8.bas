dbMemo "SQL" ="SELECT e.Event_Date, e.PseudoEvent, l.Plot_Name, t.Tag, p.Latin_Name, sd.Habit, "
    "p.Tree, p.Shrub, sd.*\015\012FROM (((tbl_Sapling_Data AS sd LEFT JOIN tbl_Events"
    " AS e ON e.Event_ID = sd.Event_ID) LEFT JOIN tbl_Locations AS l ON l.Location_ID"
    " = e.Location_ID) LEFT JOIN tbl_Tags AS t ON t.Tag_ID = sd.Tag_ID) LEFT JOIN tlu"
    "_Plants AS p ON p.TSN = t.TSN\015\012WHERE p.Latin_Name IS NULL\015\012AND e.Pse"
    "udoEvent = 0\015\012ORDER BY e.Event_Date DESC , l.Plot_Name, t.Tag;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbMemo "OrderBy" ="[Query8].[Event_Date]"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="sd.SaplingVigor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.PseudoEvent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Event_ID"
        dbInteger "ColumnWidth" ="465"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Tag_ID"
        dbInteger "ColumnWidth" ="435"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Status"
        dbInteger "ColumnWidth" ="2070"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Status"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.DRC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Updated_Date"
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
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Tree"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Shrub"
        dbLong "AggregateType" ="-1"
    End
End
