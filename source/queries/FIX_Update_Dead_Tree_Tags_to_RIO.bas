dbMemo "SQL" ="UPDATE tbl_Tags AS tt SET tt.Tag_Status = 'Retired (In Office)'\015\012WHERE tt."
    "Tag_ID IN\015\012(\015\012SELECT DISTINCT t.Tag_ID\015\012FROM tbl_Tree_Data td\015"
    "\012LEFT JOIN tbl_Tags t ON t.Tag_ID = td.Tag_ID \015\012WHERE\015\012td.Tree_St"
    "atus LIKE 'Dead*'\015\012AND t.Tag_Status <> 'Retired (In Office)'\015\012AND t."
    "Tag_Status NOT LIKE 'Inactive*'\015\012);\015\012"
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
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
End
