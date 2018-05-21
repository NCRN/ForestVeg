dbMemo "SQL" ="SELECT l.Location_ID, e.Event_ID, l.Admin_Unit_Code, l.Subunit_Code, e.Event_Dat"
    "e, t.Tag, Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5"
    ")*2,6) AS EquivDBH\015\012FROM (((tbl_Locations AS l INNER JOIN tbl_Events AS e "
    "ON l.Location_ID = e.Location_ID) INNER JOIN tbl_Tree_Data AS sd ON e.Event_ID ="
    " sd.Event_ID) INNER JOIN tbl_Tags AS t ON sd.Tag_ID = t.Tag_ID) INNER JOIN tbl_T"
    "ree_DBH AS sbh ON sd.Tree_Data_ID = sbh.Tree_Data_ID\015\012GROUP BY l.Location_"
    "ID, e.Event_ID, l.Admin_Unit_Code, l.Subunit_Code, e.Event_Date, t.Tag\015\012HA"
    "VING (((l.Location_ID) = \"{FAC4340D-EF33-4BAA-AFBE-090EECE4BC5D}\") \015\012AND"
    " ((t.Tag) = 133))\015\012ORDER BY e.Event_Date;\015\012"
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
