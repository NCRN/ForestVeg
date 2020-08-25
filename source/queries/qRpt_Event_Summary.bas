Operation =1
Option =0
Where ="(((qFiltered_Events.Event_ID) Like [Forms]![frm_Data_Summary_Advanced]![cbxEvent"
    "Selection]))"
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
End
Begin OutputColumns
    Expression ="qFiltered_Events.*"
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="qFiltered_Locations.Unit_Code"
    Expression ="qFiltered_Locations.Subunit_Code"
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="qFiltered_Locations.X_Coord"
    Expression ="qFiltered_Locations.Y_Coord"
    Expression ="qFiltered_Locations.Location_Notes"
    Expression ="qFiltered_Locations.Location_Status"
    Expression ="qFiltered_Locations.Panel"
    Expression ="qFiltered_Locations.Frame"
    Expression ="qFiltered_Locations.GRTS_Order"
    Expression ="qFiltered_Locations.Install_Date"
    Expression ="qFiltered_Locations.Lon_WGS84"
    Expression ="qFiltered_Locations.Lat_WGS84"
    Expression ="qFiltered_Locations.Slope"
    Expression ="qFiltered_Locations.Aspect"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbInteger "RowHeight" ="510"
Begin
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4050"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Location_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4125"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Event_Group_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Event_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Event_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Pictures_Taken"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Entered_On_Tablet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Entered_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Entered_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Updated_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Verified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Verified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Verified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Certified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Certified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.Certified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.CWD_Check_360"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.CWD_Check_120"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.tbl_Events.CWD_Check_240"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.X_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Y_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Location_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Location_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.GRTS_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Install_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Lon_WGS84"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Lat_WGS84"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Slope"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Aspect"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =51
    Top =25
    Right =1004
    Bottom =488
    Left =-1
    Top =-1
    Right =929
    Bottom =248
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =47
        Top =-7
        Right =210
        Bottom =319
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =456
        Bottom =335
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
End
