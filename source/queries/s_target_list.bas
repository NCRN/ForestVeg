dbMemo "SQL" ="SELECT DISTINCT TargetedSpecies.TargetList\015\012FROM TargetedSpecies;\015\012"
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
        dbText "Name" ="TargetedSpecies.TargetList"
        dbLong "AggregateType" ="-1"
    End
End
