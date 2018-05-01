dbMemo "SQL" ="SELECT t.Tag_Status AS New_Tag_Status, td.Tree_Status, t.Tag_ID, t.Location_ID, "
    "t.Tag_Status, td.*\015\012FROM tbl_Tags AS t INNER JOIN tbl_Tree_Data AS td ON t"
    "d.Tag_ID = t.Tag_ID\015\012WHERE td.Tree_Status = 'Dead Leaning'\015\012ORDER BY"
    " td.Updated_Date, td.Tree_Status;\015\012"
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
        dbText "Name" ="td.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="New_Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Wind_Lightning_Damage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.TreeVigor"
        dbLong "AggregateType" ="-1"
    End
End
