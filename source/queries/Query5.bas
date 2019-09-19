dbMemo "SQL" ="SELECT DISTINCT p.TSN, p.Latin_Name, IIf(p.Latin_Name=\"Kalmia latifolia\",p.Lat"
    "in_Name & \"***\",\015\012IIf(p.Latin_Name=\"Lindera benzoin\",p.Latin_Name & \""
    "***\",\015\012IIf(p.Latin_Name=\"Ilex verticillata\",p.Latin_Name & \"***\",p.La"
    "tin_Name))) AS Name, p.Tree, p.Shrub, IIF([p].[Tree]=True,\"Tree\", IIF([p].[Shr"
    "ub]=True,\"Shrub\",\"\")) AS Habit\015\012FROM tlu_Plants AS p\015\012WHERE p.Tr"
    "ee = True AND p.Shrub = True;\015\012"
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
        dbText "Name" ="p.TSN"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Name"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
End
