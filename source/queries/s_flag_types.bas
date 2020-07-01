dbMemo "SQL" ="SELECT ' - Choose Type -' AS FlagTypePicker FROM Flags f\015\012UNION SELECT DIS"
    "TINCT f.FlagType AS FlagTypePicker FROM Flags f;\015\012"
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
        dbText "Name" ="FlagTypePicker"
        dbLong "AggregateType" ="-1"
    End
End
