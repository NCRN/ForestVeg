dbMemo "SQL" ="PARAMETERS park Text ( 255 ), tgtYear Short;\015\012SELECT tbl_Target_Species.*,"
    " tbl_Target_Species.Park_Code, tbl_Target_Species.Target_Year, *\015\012FROM tbl"
    "_Target_Species\015\012WHERE (((tbl_Target_Species.Target_Year)=CInt(tgtYear)) A"
    "nd ((LCase(tbl_Target_Species.Park_Code))=LCase(park)))\015\012ORDER BY tbl_Targ"
    "et_Species.Species_Name;\015\012"
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
dbText "Description" =" Target species list for a park & year\015\012(Target List Tool update)"
Begin
    Begin
        dbText "Name" ="tbl_Target_Species.Tgt_Species_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Park_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Transect_Only"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Area_ID"
        dbLong "AggregateType" ="-1"
    End
End
