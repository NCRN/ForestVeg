dbMemo "SQL" ="SELECT DISTINCT t.ID, t.TSN, t.TargetList, p.Latin_Name, t.EstablishDate, t.Reti"
    "reDate\015\012FROM TargetedSpecies AS t INNER JOIN tlu_Plants AS p ON CLng(p.TSN"
    ")  = t.TSN;\015\012"
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
        dbText "Name" ="t.EstablishDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TargetList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RetireDate"
        dbLong "AggregateType" ="-1"
    End
End
