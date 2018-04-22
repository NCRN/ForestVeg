dbMemo "SQL" ="SELECT t.Tag_ID, t.Tag, t.Tag_Status, t.Microplot_Number AS MP, IIf(IsNull([azim"
    "uth]),\"\",[Azimuth] & \" / \" & [distance] & \"m\") AS Azi_Dist, td.Tree_Data_I"
    "D, sd.Sapling_Data_ID, t.Location_ID, sd.Event_ID\015\012FROM (((tbl_Tags AS t L"
    "EFT JOIN qry_Status_Tree_Current_Event ON t.Tag_ID = qry_Status_Tree_Current_Eve"
    "nt.Tag_ID) LEFT JOIN qry_Status_Sapling_Current_Event ON t.Tag_ID = qry_Status_S"
    "apling_Current_Event.Tag_ID) INNER JOIN tbl_Tree_Data AS td ON t.Tag_ID = td.Tag"
    "_ID) INNER JOIN tbl_Sapling_Data AS sd ON t.Tag_ID = sd.Tag_ID\015\012WHERE (qry"
    "_Status_Sapling_Current_Event.Event_ID Is Null) \015\012AND (qry_Status_Tree_Cur"
    "rent_Event.Event_ID Is Null)\015\012ORDER BY t.Tag_Status, t.Tag;\015\012"
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
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EquivDBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Azi_Dist"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Event_ID"
        dbLong "AggregateType" ="-1"
    End
End
