dbMemo "SQL" ="SELECT e.*, l.Plot_Name, l.Unit_Code, l.Subunit_Code, l.Admin_Unit_Code, l.X_Coo"
    "rd, l.Y_Coord, l.Location_Notes, l.Location_Status, l.Panel, l.Frame, l.GRTS_Ord"
    "er, l.Install_Date, l.Lon_WGS84, l.Lat_WGS84, l.Slope, l.Aspect\015\012FROM tbl_"
    "Locations AS l INNER JOIN qUnfiltered_Events AS e ON l.Location_ID = e.Location_"
    "ID;\015\012"
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
        dbText "Name" ="e.tbl_Events.Verified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Rare_Spp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Early_Detect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Event_Group_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Entered_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Event_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Event_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Entered_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Lon_WGS84"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Lat_WGS84"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Updated_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.PseudoEvent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Pictures_Taken"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.CWD_Check_360"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.CWD_Check_120"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.CWD_Check_240"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Deer_Impact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Is_Excluded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Plot_Maint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Entered_On_Tablet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Verified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Verified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Certified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Certified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.tbl_Events.Certified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.X_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Y_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Location_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.GRTS_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Install_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Slope"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Aspect"
        dbLong "AggregateType" ="-1"
    End
End
