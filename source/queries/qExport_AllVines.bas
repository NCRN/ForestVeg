dbMemo "SQL" ="SELECT Plot_Name, Unit_Code, Unit_Group, Subunit_Code, Cycle, Panel, Frame, Samp"
    "le_Year, Date, TSN, Latin_Name, Tag_Status, Host_Tag, Host_Latin_Name, Host_Stat"
    "us, Condition, Exotic\015\012FROM qExport_All_Sapling_Vines\015\012UNION ALL SEL"
    "ECT Plot_Name, Unit_Code, Unit_Group, Subunit_Code, Cycle, Panel, Frame, Sample_"
    "Year, Date, TSN, Latin_Name, Tag_Status, Host_Tag, Host_Latin_Name, Host_Status,"
    " Condition, Exotic\015\012FROM qExport_All_Tree_Vines;\015\012"
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
        dbText "Name" ="Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Host_Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exotic"
        dbLong "AggregateType" ="-1"
    End
End
