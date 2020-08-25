dbMemo "SQL" ="SELECT SumDBH\015\012FROM (SELECT TOP 1 e.Event_Date, dbh.Sapling_Data_ID, SUM(d"
    "bh.DBH) AS SumDBH FROM ((tbl_Events e INNER JOIN tbl_Sapling_Data d ON d.Event_I"
    "D = e.Event_ID) INNER JOIN tbl_Sapling_DBH dbh ON dbh.Sapling_Data_ID = d.Saplin"
    "g_Data_ID) WHERE d.Sapling_Data_ID <> {1DEAAAA1-A96C-4414-8695-CC8E3FE97A83} AND"
    " d.Tag_ID = {270EA0D3-3471-4A8E-B652-A22D3465D8C9} AND YEAR(e.Event_Date) < (SEL"
    "ECT YEAR(ee.Event_Date) FROM (tbl_Events ee INNER JOIN tbl_Sapling_Data dd ON dd"
    ".Event_ID = ee.Event_ID) WHERE dd.Sapling_Data_ID = {1DEAAAA1-A96C-4414-8695-CC8"
    "E3FE97A83} ) GROUP BY dbh.Sapling_Data_ID, e.Event_Date  ORDER BY e.Event_Date D"
    "ESC)  AS [%$##@_Alias];\015\012"
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
        dbText "Name" ="SumDBH"
        dbLong "AggregateType" ="-1"
    End
End
