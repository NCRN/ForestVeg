dbMemo "SQL" ="SELECT e.*, l.ShowLocMsg, l.LocMessage\015\012FROM tbl_Locations AS l INNER JOIN"
    " tbl_Events AS e ON l.Location_ID = e.Location_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Data entry form record source"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="e.Updated_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Plot_Maint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Certified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Group_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Is_Excluded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.LocMessage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Deer_Impact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Pictures_Taken"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Early_Detect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.PseudoEvent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Certified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.CWD_Check_360"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.CWD_Check_120"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.ShowLocMsg"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.CWD_Check_240"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Rare_Spp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Entered_On_Tablet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Entered_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Entered_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Verified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Verified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Verified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Certified_Date"
        dbLong "AggregateType" ="-1"
    End
End
