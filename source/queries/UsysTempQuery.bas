dbMemo "SQL" ="SELECT TemplateName, Count(TemplateName) AS NumberOfDupes\015\012FROM tsys_Db_Te"
    "mplates\015\012GROUP BY TemplateName\015\012HAVING Count(TemplateName) > 1;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
