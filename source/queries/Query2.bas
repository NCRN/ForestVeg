dbMemo "SQL" ="SELECT t.Tag_ID, t.Tag, t.Tag_Status, IIf(IsNull([azimuth]),\"\",[Azimuth] & \" "
    "/ \" & [distance] & \"m\") AS Azi_Dist, t.Microplot_Number AS MP, t.Location_ID\015"
    "\012FROM (tbl_Tags AS t LEFT JOIN qry_Status_Sapling_Current_Event ON t.Tag_ID ="
    " qry_Status_Sapling_Current_Event.Tag_ID) LEFT JOIN qry_Status_Tree_Current_Even"
    "t ON t.Tag_ID = qry_Status_Tree_Current_Event.Tag_ID\015\012WHERE (\015\012((qry"
    "_Status_Sapling_Current_Event.Event_ID) Is Null) \015\012AND ((qry_Status_Tree_C"
    "urrent_Event.Event_ID) Is Null))\015\012ORDER BY t.Tag_Status, t.Tag;\015\012"
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
        dbText "Name" ="tbl_Tags.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
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
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Azi_Dist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.MP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
End
