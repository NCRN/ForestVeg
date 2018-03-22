dbMemo "SQL" ="SELECT CStr(Year([Start_date])) AS Event_year, tbl_Events.Location_ID, tbl_Event"
    "s.Start_date, Count(tbl_Point_Counts.Observation_ID) AS N_obs_recs\015\012FROM t"
    "bl_Events LEFT JOIN tbl_Point_Counts ON tbl_Events.Event_ID=tbl_Point_Counts.Eve"
    "nt_ID\015\012GROUP BY CStr(Year([Start_date])), tbl_Events.Location_ID, tbl_Even"
    "ts.Start_date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbMemo "Filter" ="((N_obs_recs=0))"
dbMemo "OrderBy" ="Lookup_Location__ID.Location"
dbText "Description" ="Number of point count records by year, location and date"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="Event_year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_obs_recs"
        dbLong "AggregateType" ="-1"
    End
End
