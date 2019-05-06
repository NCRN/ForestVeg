dbMemo "SQL" ="INSERT INTO tlu_Tree_Condition ( Code, Description, Details, Sequence )\015\012V"
    "ALUES (35, 'Vine stress', 'tree/sapling/shrub exhibits stress due to vine(s)', 1"
    "9);\015\012"
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
        dbText "Name" ="Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Azi_Dist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MP"
        dbLong "AggregateType" ="-1"
    End
End
