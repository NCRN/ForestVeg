dbMemo "SQL" ="SELECT tbl_Tags_History.Record_ID as Tag_ID, tbl_Tags_History.Change_Date, [Fiel"
    "d_Name] & \" was changed by \" & [First_Name] & \" \" & [Last_Name] & \" from \""
    " & [Value_Old] & \" to \" & [Value_New] AS Change_Desc\015\012FROM tbl_Tags_Hist"
    "ory LEFT JOIN tlu_Contacts ON tbl_Tags_History.Contact_ID = tlu_Contacts.Contact"
    "_ID\015\012UNION ALL SELECT qActive_Tree_Data.Tag_ID, qActive_Tree_Data.Event_Da"
    "te, \"Measured as \" & [Tree_Status] & \" TREE with stems(s) of \" & [stemlist] "
    "& \" cm DBH\" AS Change_Desc\015\012FROM qActive_Tree_Data\015\012UNION ALL SELE"
    "CT qActive_Sapling_Data.Tag_ID, qActive_Sapling_Data.Event_Date, \"Measured as \""
    " & [Sapling_Status] & \" SAPLING with stems(s) of \" & [stemlist] & \" cm DBH\" "
    "AS Change_Desc\015\012FROM qActive_Sapling_Data;\015\012"
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
        dbText "Name" ="Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Change_Date"
        dbInteger "ColumnWidth" ="2670"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Change_Desc"
        dbInteger "ColumnWidth" ="6780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
