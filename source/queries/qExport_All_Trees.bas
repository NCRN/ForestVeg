dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year([Even"
    "t_Date])-2006)/4) AS Cycle, l.Panel, l.Frame, Year([Event_Date]) AS Sample_Year,"
    " CLng(Format(e.Event_Date,\"yyyymmdd\")) AS [Date], t.Tag, t.TSN, p.TaxonCode, p"
    ".Latin_Name, ba.Stems, ba.SumLiveBasalArea_cm2, ba.SumDeadBasalArea_cm2, ba.Equi"
    "v_Live_DBH_cm, ba.Equiv_Dead_DBH_cm, td.DBH_Check, v.Condition, td.Tree_Status A"
    "S Status, td.Crown_Class, cc.Crown_Description\015\012FROM ((((((tbl_Locations A"
    "S l RIGHT JOIN tbl_Events AS e ON l.Location_ID = e.Location_ID) INNER JOIN tbl_"
    "Tree_Data AS td ON e.Event_ID = td.Event_ID) INNER JOIN tbl_Tags AS t ON t.Tag_I"
    "D = td.Tag_ID) LEFT JOIN qList_Crown_Class_Descriptions AS cc ON td.Crown_Class "
    "= cc.Crown_Class) LEFT JOIN qCalc_Basal_Area_Per_Tree AS ba ON td.Tree_Data_ID ="
    " ba.Tree_Data_ID) LEFT JOIN qSum_Trees_with_Vines_in_Crown AS v ON td.Tree_Data_"
    "ID = v.Tree_Data_ID) LEFT JOIN tlu_Plants AS p ON t.TSN = p.TSN\015\012ORDER BY "
    "t.Tag;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="cc.Crown_Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
End
