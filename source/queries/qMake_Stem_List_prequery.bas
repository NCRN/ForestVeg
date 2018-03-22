dbMemo "SQL" ="SELECT qActive_Tree_Data.Tree_Data_ID, qActive_Tree_Data.StemList, \"Tree\" AS S"
    "ize_Class\015\012FROM qActive_Tree_Data\015\012UNION ALL SELECT qActive_Sapling_"
    "Data.Sapling_Data_ID, qActive_Sapling_Data.StemList, \"Sapling\" \015\012FROM qA"
    "ctive_Sapling_Data;\015\012"
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
        dbText "Name" ="qActive_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4155"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.StemList"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Size_Class"
        dbLong "AggregateType" ="-1"
    End
End
