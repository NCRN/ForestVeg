dbMemo "SQL" ="SELECT qFiltered_Locations.Plot_Name, qFiltered_Locations.Unit_Code, qFiltered_L"
    "ocations.Admin_Unit_Code, qFiltered_Locations.Panel, qFiltered_Locations.Frame, "
    "qFiltered_Events_Cycle.Event_Date, CInt([qFiltered_Events_Cycle].[Event_Year]) A"
    "S Event_Year, qFiltered_Events_Cycle.Certified, qFiltered_Locations.Location_ID,"
    " qFiltered_Events_Cycle.Event_ID\015\012FROM qFiltered_Locations INNER JOIN qFil"
    "tered_Events_Cycle ON qFiltered_Locations.Location_ID = qFiltered_Events_Cycle.L"
    "ocation_ID\015\012ORDER BY qFiltered_Locations.Plot_Name, qFiltered_Events_Cycle"
    ".Event_Date;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
