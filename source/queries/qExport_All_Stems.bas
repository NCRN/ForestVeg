dbMemo "SQL" ="SELECT Plot_Name, Unit_Code, Unit_Group, Subunit_Code, Cycle, Panel, Frame, Samp"
    "le_Year, Date, Tag, TSN, Latin_Name, Status, DBH, Live, Habit, Class\015\012FROM"
    " qExport_All_Tree_Stems\015\012UNION ALL SELECT Plot_Name, Unit_Code, Unit_Group"
    ", Subunit_Code, Cycle, Panel, Frame, Sample_Year, Date, Tag, TSN, Latin_Name, St"
    "atus, DBH, Live, Habit, Class\015\012FROM qExport_All_Sapling_Stems;\015\012"
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
        dbText "Name" ="DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Class"
        dbLong "AggregateType" ="-1"
    End
End
