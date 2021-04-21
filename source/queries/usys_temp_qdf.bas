﻿dbMemo "SQL" ="SELECT SumDBH\015\012FROM (SELECT TOP 1 e.Event_Date, dbh.Tree_Data_ID, SUM(dbh."
    "DBH) AS SumDBH FROM ((tbl_Events e INNER JOIN tbl_Tree_Data d ON d.Event_ID = e."
    "Event_ID) INNER JOIN tbl_Tree_DBH dbh ON dbh.Tree_Data_ID = d.Tree_Data_ID) WHER"
    "E d.Tree_Data_ID <> {80C849AF-C864-40B4-9804-213E2853E308} AND d.Tag_ID = {22423"
    "06D-6A83-4130-9CCF-7B780A89B865} AND YEAR(e.Event_Date) < (SELECT YEAR(ee.Event_"
    "Date) FROM (tbl_Events ee INNER JOIN tbl_Tree_Data dd ON dd.Event_ID = ee.Event_"
    "ID) WHERE dd.Tree_Data_ID = {80C849AF-C864-40B4-9804-213E2853E308} ) GROUP BY db"
    "h.Tree_Data_ID, e.Event_Date  ORDER BY e.Event_Date DESC)  AS [%$##@_Alias];\015"
    "\012"
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
