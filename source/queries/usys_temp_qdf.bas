﻿dbMemo "SQL" ="SELECT l.Location_ID, e.Event_ID, l.Admin_Unit_Code, l.Subunit_Code, e.Event_Dat"
    "e, t.Tag, Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5"
    ")*2,6) AS EquivDBH\015\012FROM ((tbl_Locations AS l INNER JOIN tbl_Events AS e O"
    "N l.Location_ID = e.Location_ID) INNER JOIN (tbl_Tree_Data AS sd INNER JOIN tbl_"
    "Tags AS t ON sd.Tag_ID = t.Tag_ID) ON e.Event_ID = sd.Event_ID) INNER JOIN tbl_T"
    "ree_DBH AS sbh ON sd.Tree_Data_ID = sbh.Tree_Data_ID\015\012GROUP BY l.Location_"
    "ID, e.Event_ID, l.Admin_Unit_Code, l.Subunit_Code, e.Event_Date, t.Tag\015\012HA"
    "VING (((l.Location_ID) = \"20170725080104-227781593.799591\") AND ((t.Tag) = 232"
    "90))\015\012ORDER BY e.Event_Date;\015\012"
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
