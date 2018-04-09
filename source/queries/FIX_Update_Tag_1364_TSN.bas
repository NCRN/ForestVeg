dbMemo "SQL" ="UPDATE tbl_Tags SET TSN = 999999\015\012WHERE Tag IN (SELECT t.Tag \015\012FROM "
    "((tbl_Tags t\015\012INNER JOIN tbl_Locations l ON t.Location_ID = l.Location_ID)"
    "\015\012LEFT JOIN tlu_Plants p ON t.TSN = p.TSN)\015\012WHERE t.Tag = 1364 AND l"
    ".Plot_Name = 'CHOH-0788');\015\012"
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
