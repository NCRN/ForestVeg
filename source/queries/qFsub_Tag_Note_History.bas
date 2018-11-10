﻿dbMemo "SQL" ="SELECT th.Record_ID as Tag_ID, \015\012th.Change_Date AS TagChangeDate, \015\012"
    "[Field_Name] & \" was changed by \" & [First_Name] & \" \" & [Last_Name] & \" fr"
    "om \" & [Value_Old] & \" to \" & [Value_New] AS Change_Desc\015\012FROM tbl_Tags"
    "_History th\015\012LEFT JOIN tlu_Contacts c ON th.Contact_ID = c.Contact_ID\015\012"
    "\015\012UNION ALL \015\012\015\012SELECT td.Tag_ID, \015\012e.Event_Date AS TagC"
    "hangeDate, \015\012\"Recorded as \" & [Tree_Status] & \" TREE\" AS Change_Desc\015"
    "\012FROM tbl_Events e\015\012INNER JOIN tbl_Tree_Data td ON e.Event_ID = td.Even"
    "t_ID\015\012\015\012UNION ALL \015\012\015\012SELECT sd.Tag_ID, \015\012e.Event_"
    "Date AS TagChangeDate, \015\012\"Recorded as \" & [Sapling_Status] & \" SAPLING\""
    " AS Change_Desc\015\012FROM tbl_Events e\015\012INNER JOIN tbl_Sapling_Data sd O"
    "N e.Event_ID = sd.Event_ID\015\012\015\012UNION ALL \015\012\015\012SELECT td.Ta"
    "g_ID, \015\012e.Event_Date AS TagChangeDate, \015\012\"Noted during event: \" & "
    "[Tree_Notes] AS Change_Desc\015\012FROM tbl_Events e \015\012INNER JOIN tbl_Tree"
    "_Data td ON e.Event_ID = td.Event_ID\015\012WHERE Not(IsNull(td.Tree_Notes))\015"
    "\012\015\012UNION ALL SELECT sd.Tag_ID, \015\012e.Event_Date AS TagChangeDate, \015"
    "\012\"Noted during event: \" & [Sapling_Notes] AS Change_Desc\015\012FROM tbl_Ev"
    "ents e\015\012INNER JOIN tbl_Sapling_Data sd ON e.Event_ID = sd.Event_ID\015\012"
    "WHERE Not(IsNull(sd.Sapling_Notes))\015\012ORDER BY TagChangeDate DESC;\015\012"
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
        dbInteger "ColumnWidth" ="4170"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Change_Desc"
        dbInteger "ColumnWidth" ="6780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TagChangeDate"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
