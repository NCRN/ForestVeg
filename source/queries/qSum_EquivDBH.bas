dbMemo "SQL" ="SELECT tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_"
    "Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag, Round(((("
    "Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2,6) AS EquivDBH\015"
    "\012FROM ((tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tb"
    "l_Events.Location_ID) INNER JOIN (tbl_Tree_Data INNER JOIN tbl_Tags ON tbl_Tree_"
    "Data.Tag_ID = tbl_Tags.Tag_ID) ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) "
    "INNER JOIN tbl_Tree_DBH ON tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_I"
    "D\015\012GROUP BY tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations."
    "Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag"
    "\015\012HAVING (((tbl_Locations.Location_ID)={guid {41FECD57-8CF4-4E1C-B7E1-E2AE"
    "BD210292}}) AND ((tbl_Tags.Tag)=19198))\015\012ORDER BY tbl_Events.Event_Date;\015"
    "\012"
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
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EquivDBH"
        dbLong "AggregateType" ="-1"
    End
End
