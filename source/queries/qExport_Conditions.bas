dbMemo "SQL" ="SELECT Plot_Name, Unit_Code, Unit_Group, Subunit_Code, Cycle, Panel, Frame, Samp"
    "le_Year, Date, Tag, Tag_Status, TSN, Latin_Name, Status, Condition, Pest\015\012"
    "FROM qActive_Tree_Conditions\015\012UNION ALL SELECT Plot_Name, Unit_Code, Unit_"
    "Group, Subunit_Code, Cycle, Panel, Frame, Sample_Year, Date, Tag, Tag_Status, TS"
    "N, Latin_Name, Status, Condition, Pest\015\012FROM qActive_Sapling_Conditions;\015"
    "\012"
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
        dbText "Name" ="Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pest"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tag_Status"
        dbLong "AggregateType" ="-1"
    End
End
