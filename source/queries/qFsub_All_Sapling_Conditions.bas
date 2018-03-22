﻿dbMemo "SQL" ="SELECT \"VINE\" AS Category, tlu_Plants.Latin_Name as Description, \"na\" AS Per"
    "centAfflicted\015\012FROM (tbl_Sapling_Data INNER JOIN tbl_Sapling_Vines ON tbl_"
    "Sapling_Data.Sapling_Data_ID = tbl_Sapling_Vines.Sapling_Data_ID) INNER JOIN tlu"
    "_Plants ON tbl_Sapling_Vines.TSN = tlu_Plants.TSN\015\012WHERE (((tbl_Sapling_Da"
    "ta.Sapling_Data_ID)=[Forms]![frm_Events]![fsub_Sapling_Data]![Sapling_Data_ID]))"
    "\015\012UNION ALL\015\012SELECT IIf([Pest]=True,\"PEST\",\"CONDITION\") AS Descr"
    "iption, tbl_Sapling_Conditions.Condition, \"na\" AS PercentAff\015\012FROM (tbl_"
    "Sapling_Data INNER JOIN tbl_Sapling_Conditions ON tbl_Sapling_Data.Sapling_Data_"
    "ID = tbl_Sapling_Conditions.Sapling_Data_ID) INNER JOIN tlu_Tree_Condition ON tb"
    "l_Sapling_Conditions.Condition = tlu_Tree_Condition.Description\015\012WHERE ((("
    "tbl_Sapling_Conditions.Sapling_Data_ID)=[Forms]![frm_Events]![fsub_Sapling_Data]"
    "![Sapling_Data_ID]))\015\012UNION ALL SELECT \"FOLIAGE\" AS Description, tbl_Sap"
    "ling_Foliage_Conditions.Condition, tbl_Sapling_Foliage_Conditions.Percent_Afflic"
    "ted\015\012FROM tbl_Sapling_Data INNER JOIN tbl_Sapling_Foliage_Conditions ON tb"
    "l_Sapling_Data.Sapling_Data_ID = tbl_Sapling_Foliage_Conditions.Sapling_Data_ID\015"
    "\012WHERE (((tbl_Sapling_Data.Sapling_Data_ID)=[Forms]![frm_Events]![fsub_Saplin"
    "g_Data]![Sapling_Data_ID]));\015\012"
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
        dbText "Name" ="Description"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Category"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PercentAfflicted"
        dbInteger "ColumnWidth" ="1680"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
