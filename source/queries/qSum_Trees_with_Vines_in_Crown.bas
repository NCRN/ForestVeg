dbMemo "SQL" ="SELECT atc.Tree_Data_ID, atc.Plot_Name, atc.Panel, atc.Frame, atc.Sample_Year, a"
    "tc.Tag, atc.Condition, Count(atc.TSN) AS CountOfTSN\015\012FROM qActive_Tree_Con"
    "ditions AS atc\015\012GROUP BY atc.Tree_Data_ID, atc.Plot_Name, atc.Panel, atc.F"
    "rame, atc.Sample_Year, atc.Tag, atc.Condition\015\012HAVING (((atc.Condition)=\""
    "Vines in the crown\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="CountOfTSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="atc.Condition"
        dbLong "AggregateType" ="-1"
    End
End
