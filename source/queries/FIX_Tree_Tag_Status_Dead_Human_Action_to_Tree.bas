﻿dbMemo "SQL" ="UPDATE tbl_Tags AS t SET t.Tag_Status = 'Tree'\015\012WHERE t.Tag_ID IN \015\012"
    "(SELECT t.Tag_ID\015\012FROM (tbl_Tags t\015\012INNER JOIN tbl_Tree_Data td ON t"
    "d.Tag_ID = t.Tag_ID)\015\012WHERE \015\012td.Tree_Status = 'Dead - Human Action'"
    "\015\012AND\015\012t.Tag_Status <> 'Tree'\015\012ORDER BY \015\012td.Updated_Dat"
    "e, td.Tree_Status);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
