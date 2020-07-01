dbMemo "SQL" ="SELECT *\015\012FROM qry_Tgt_Species_List_Park_Summary_Data\015\012ORDER BY Fami"
    "ly, utah_species;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
dbText "Description" ="Target species summary for all parks for a given year  (Target List Tool update)"
dbMemo "OrderBy" ="[qry_Tgt_Species_List_Park_Summary].[Family], [qry_Tgt_Species_List_Park_Summary"
    "].[Species_Name]"
Begin
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.ParkYearPriorities"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="8136"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.MinYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.MaxYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.Master_Plant_Code_FK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.LU_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Tgt_Species_List_Park_Summary_Data.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
End
