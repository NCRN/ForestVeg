dbMemo "SQL" ="SELECT Plot_Name, Unit_Code, Unit_Group, Subunit_Code, Cycle, Panel, Frame, Samp"
    "le_Year, Date, Tag, TSN, Latin_Name, Status, DBH, Live, Habit, Class\015\012FROM"
    " qExport_All_Tree_Stems\015\012UNION ALL SELECT Plot_Name, Unit_Code, Unit_Group"
    ", Subunit_Code, Cycle, Panel, Frame, Sample_Year, Date, Tag, TSN, Latin_Name, St"
    "atus, DBH, Live, Habit, Class\015\012FROM qExport_All_Sapling_Stems;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
