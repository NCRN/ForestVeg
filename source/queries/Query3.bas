dbMemo "SQL" ="SELECT t.Tag, p.Latin_Name, q.Stems, q.Equiv_DBH_cm, td.Crown_Class & \" \" & q."
    "CrownClass AS CC, td.TreeVigor & \" \" & tv.TreeVigorClass AS Vig, td.Vines_Chec"
    "ked, td.Conditions_Checked, td.Foliage_Conditions_Checked, td.Tree_Status, t.Azi"
    "muth, t.Distance, td.Tree_Notes, td.Tree_Data_ID, td.Event_ID, MakeStemList('Tre"
    "e', td.Event_ID,td.Tree_Data_Id) AS StemList, MakeLiveFlag('Tree',td.Event_ID,td"
    ".Tree_Data_Id) AS LiveFlag\015\012FROM (((tbl_Tree_Data AS td LEFT JOIN qCalc_Ba"
    "sal_Area_per_Tree AS q ON td.Tree_Data_ID = q.Tree_Data_ID) LEFT JOIN tbl_Tags A"
    "S t ON td.Tag_ID = t.Tag_ID) LEFT JOIN tlu_Plants AS p ON t.TSN = p.TSN) LEFT JO"
    "IN tluTreeVigor AS tv ON td.TreeVigor = tv.TreeVigorCode\015\012ORDER BY t.Tag;\015"
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
        dbText "Name" ="Vig"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="q.Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Notes"
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
        dbText "Name" ="StemList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LiveFlag"
        dbLong "AggregateType" ="-1"
    End
End
