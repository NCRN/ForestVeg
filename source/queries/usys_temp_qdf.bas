﻿dbMemo "SQL" ="SELECT SumDBH\015\012FROM (SELECT TOP 1 e.Event_Date, dbh.Tree_Data_ID, SUM(dbh."
    "DBH) AS SumDBH FROM ((tbl_Events e INNER JOIN tbl_Tree_Data d ON d.Event_ID = e."
    "Event_ID) INNER JOIN tbl_Tree_DBH dbh ON dbh.Tree_Data_ID = d.Tree_Data_ID) WHER"
    "E d.Tree_Data_ID <> No Tree_Data_ID AND d.Tag_ID = No Tag_ID AND YEAR(e.Event_Da"
    "te) < (SELECT YEAR(ee.Event_Date) FROM (tbl_Events ee INNER JOIN tbl_Tree_Data d"
    "d ON dd.Event_ID = ee.Event_ID) WHERE dd.Tree_Data_ID = No Tree_Data_ID ) GROUP "
    "BY dbh.Tree_Data_ID, e.Event_Date  ORDER BY e.Event_Date DESC)  AS [%$##@_Alias]"
    ";\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbText "Description" ="Query used throughout the application for handling template SQL. QueryDef is upd"
    "ated based on desired template. (Hidden to avoid removal)"
Begin
End
