﻿dbMemo "SQL" ="SELECT l.Plot_Name, l.Unit_Code, l.Unit_Group, l.Subunit_Code, 1+Int((Year([Even"
    "t_Date])-2006)/4) AS Cycle, l.Panel, l.Frame, Year(e.Event_Date) AS Sample_Year,"
    " Format(e.Event_Date,\"yyyymmdd\") AS [Date], t.Tag, t.TSN, p.TaxonCode, p.Latin"
    "_Name, ba.StemsLive, ba.SumLiveBasalArea_cm2, ba.Equiv_Live_DBH_cm, sd.DBH_Check"
    ", sd.Sapling_Status AS Status, sd.Habit, sd.Browsable, sd.Browsed\015\012FROM (("
    "((tbl_Locations AS l INNER JOIN tbl_Events AS e ON l.Location_ID = e.Location_ID"
    ") INNER JOIN tbl_Sapling_Data AS sd ON e.Event_ID = sd.Event_ID) LEFT JOIN qCalc"
    "_Basal_Area_per_Sapling AS ba ON sd.Sapling_Data_ID = ba.Sapling_Data_ID) INNER "
    "JOIN tbl_Tags AS t ON sd.Tag_ID = t.Tag_ID) LEFT JOIN tlu_Plants AS p ON t.TSN ="
    " p.TSN\015\012WHERE sd.Sapling_Status<>\"Removed from study\" AND sd.Habit=\"Tre"
    "e\"\015\012AND YEAR(e.Event_Date) = 2017\015\012ORDER BY l.Plot_Name;\015\012"
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
