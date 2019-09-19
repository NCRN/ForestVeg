dbMemo "SQL" ="SELECT DISTINCT p.TSN, t.TSN, p.Latin_Name, IIf(p.Latin_Name=\"Kalmia latifolia\""
    ",p.Latin_Name & \"***\",\015\012IIf(p.Latin_Name=\"Lindera benzoin\",p.Latin_Nam"
    "e & \"***\",\015\012IIf(p.Latin_Name=\"Ilex verticillata\",p.Latin_Name & \"***\""
    ",p.Latin_Name))) AS Name, p.Tree, p.Shrub, IIF([p].[Tree]=True,\"Tree\", IIF([p]"
    ".[Shrub]=True,\"Shrub\",\"\")) AS Habit, sd.Habit, t.Tag\015\012FROM (tbl_Saplin"
    "g_Data AS sd LEFT JOIN tbl_Tags AS t ON t.Tag_ID = sd.Tag_ID) LEFT JOIN tlu_Plan"
    "ts AS p ON p.TSN = t.TSN\015\012WHERE sd.Habit <> IIF([p].[Tree]=True,\"Tree\", "
    "IIF([p].[Shrub]=True,\"Shrub\",\"\"))\015\012ORDER BY sd.Habit, t.Tag;\015\012"
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
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Azi_Dist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tce.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tag_ID"
        dbInteger "ColumnWidth" ="3975"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Name"
        dbInteger "ColumnWidth" ="2235"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Habit"
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
    Begin
        dbText "Name" ="HHabit"
        dbLong "AggregateType" ="-1"
    End
End
