dbMemo "SQL" ="SELECT e.Event_Date, t.Tag, dbh.*\015\012FROM ((tbl_Tags AS t LEFT JOIN tbl_Tree"
    "_Data AS td ON td.Tag_ID = t.Tag_ID) LEFT JOIN tbl_Tree_DBH AS dbh ON dbh.Tree_D"
    "ata_ID = td.Tree_Data_ID) LEFT JOIN tbl_Events AS e ON e.Event_ID = td.Event_ID\015"
    "\012WHERE t.Tag = 23290\015\012AND\015\012td.Event_ID  = '{5DD03496-502A-462F-AF"
    "9C-34C036D06379}';\015\012"
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
        dbText "Name" ="tbl_Tree_DBH.Tree_DBH_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.Tree_DBH_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stop_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RFS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tag_ID"
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
        dbText "Name" ="td.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Vines_Checked"
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
    Begin
        dbText "Name" ="DBH.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.Updated_Date"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Date"
        dbLong "AggregateType" ="-1"
    End
End
