dbMemo "SQL" ="SELECT *\015\012FROM (SELECT Plot_Name, Unit_Code, Unit_Group, Subunit_Code, Cyc"
    "le, Panel, Frame, Sample_Year, Date, Tag, TSN, Latin_Name, Status, Live, \015\012"
    "DBH, DBH_Check, Habit, Class\015\012FROM qExport_All_Tree_Stems\015\012WHERE Sta"
    "tus <> 'Removed from study'\015\012UNION ALL \015\012SELECT Plot_Name, Unit_Code"
    ", Unit_Group, Subunit_Code, Cycle, Panel, Frame, Sample_Year, Date, Tag, TSN, La"
    "tin_Name, Status, Live, \015\012DBH, DBH_Check, Habit, Class\015\012FROM qExport"
    "_All_Sapling_Stems\015\012WHERE Status <> 'Removed from study')  AS [%$##@_Alias"
    "]\015\012ORDER BY Class, Plot_Name, Tag;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbMemo "Filter" ="([qExport_All_Stems].[Status] In (\"Removed from study\") Or [qExport_All_Stems]"
    ".[Status] IS Null)"
Begin
    Begin
        dbText "Name" ="%$##@_Alias.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="%$##@_Alias.Habit"
        dbLong "AggregateType" ="-1"
    End
End
