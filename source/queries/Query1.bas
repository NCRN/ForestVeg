dbMemo "SQL" ="SELECT l.Location_ID, e.Event_ID, l.Admin_Unit_Code, l.Subunit_Code, e.Event_Dat"
    "e, t.Tag, Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5"
    ")*2,6) AS EquivDBH\015\012FROM ((tbl_Locations AS l INNER JOIN tbl_Events AS e O"
    "N l.Location_ID = e.Location_ID) INNER JOIN (tbl_Sapling_Data AS sd INNER JOIN t"
    "bl_Tags AS t ON sd.Tag_ID = t.Tag_ID) ON e.Event_ID = sd.Event_ID) INNER JOIN tb"
    "l_Sapling_DBH AS sbh ON sd.Sapling_Data_ID = sbh.Sapling_Data_ID\015\012GROUP BY"
    " l.Location_ID, e.Event_ID, l.Admin_Unit_Code, l.Subunit_Code, e.Event_Date, t.T"
    "ag\015\012HAVING (((l.Location_ID) = \"20170725080104-227781593.799591\") AND (("
    "t.Tag) = 23102))\015\012ORDER BY e.Event_Date;\015\012"
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
End
