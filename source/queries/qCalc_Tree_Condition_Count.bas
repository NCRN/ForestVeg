dbMemo "SQL" ="SELECT qCalc_Tree_Condition_Count_Prequery.*, MakeTreeConditionList([Event_ID],["
    "Tree_Data_ID]) AS ConditionAndPest_List\015\012FROM qCalc_Tree_Condition_Count_P"
    "requery;\015\012"
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
        dbText "Name" ="ConditionAndPest_List"
        dbInteger "ColumnWidth" ="3120"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Locations.Location_ID"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.tbl_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qFiltered_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.qActive_Tree_Data.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.ConditionCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.PestCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.ConditionPresentYN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qCalc_Tree_Condition_Count_Prequery.PestPresentYN"
        dbLong "AggregateType" ="-1"
    End
End
