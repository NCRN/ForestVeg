dbMemo "SQL" ="SELECT td.Tree_Data_ID, td.Event_ID, cc.Enum_Description AS CrownClass, Count(db"
    "h.DBH) AS Stems, Sum(IIf([Live]=True,1,0)) AS StemsLive, Sum(IIf([Live]=False,1,"
    "0)) AS StemsDead, Round(Sum(3.1415926*(([DBH]/2)^2)),1) AS SumBasalArea_cm2, Fir"
    "st(td.Tag_ID) AS FirstOfTag_ID, Round((([SumBasalArea_cm2]/3.1415)^0.5)*2,1) AS "
    "Equiv_DBH_cm, Round(Sum(3.1415926*(((IIf([Live]=True,[DBH],0))/2)^2)),1) AS SumL"
    "iveBasalArea_cm2, Round(Sum(3.1415926*(((IIf([Live]=False,[DBH],0))/2)^2)),1) AS"
    " SumDeadBasalArea_cm2, Round((([SumLiveBasalArea_cm2]/3.1415)^0.5)*2,1) AS Equiv"
    "_Live_DBH_cm, Round((([SumDeadBasalArea_cm2]/3.1415)^0.5)*2,1) AS Equiv_Dead_DBH"
    "_cm, td.DBH_Check, td.Crown_Class, td.Tree_Status, td.Tag_ID\015\012FROM (tbl_Tr"
    "ee_Data AS td LEFT JOIN tbl_Tree_DBH AS dbh ON td.Tree_Data_ID = dbh.Tree_Data_I"
    "D) LEFT JOIN qEnumCrownClass AS cc ON td.Crown_Class = cc.CrownClassCode\015\012"
    "GROUP BY td.Tree_Data_ID, td.Event_ID, cc.Enum_Description, td.DBH_Check, td.Cro"
    "wn_Class, td.Tree_Status, td.Tag_ID;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "Filter" ="([qCalc_Basal_Area_per_Tree].[Tag_ID]=\"\")"
dbMemo "OrderBy" ="[qCalc_Basal_Area_per_Tree].[Event_ID], [qCalc_Basal_Area_per_Tree].[Tree_Data_I"
    "D]"
Begin
    Begin
        dbText "Name" ="Stems"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="5"
    End
    Begin
        dbText "Name" ="SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
        dbInteger "ColumnOrder" ="8"
    End
    Begin
        dbText "Name" ="FirstOfTag_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4185"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1725"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2160"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="StemsDead"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="7"
    End
    Begin
        dbText "Name" ="StemsLive"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="6"
    End
    Begin
        dbText "Name" ="CrownClass"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2310"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="td.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
End
