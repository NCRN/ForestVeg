dbMemo "SQL" ="SELECT tbl_Target_List.Park_Code AS Park, tbl_Target_List.Target_Year AS TgtYear"
    ", tbl_Target_Species.LU_Code, tbl_Target_Species.Master_Plant_Code_FK, tbl_Targe"
    "t_Species.Species_Name, tbl_Target_Species.Priority, tbl_Target_Species.Transect"
    "_Only, tbl_Target_Species.Target_Area_ID, tbl_Target_Areas.Target_Area AS Tgt_Ar"
    "ea, tlu_NCPN_Plants.Master_Family AS Family, tlu_NCPN_Plants.Master_Common_Name,"
    " tlu_NCPN_Plants.Utah_Species, tlu_NCPN_Plants.Co_Species, tlu_NCPN_Plants.Wy_Sp"
    "ecies, tbl_Target_List.Park_Code & \"-\" & tbl_Target_List.Target_Year AS TgtLis"
    "t, tbl_Target_List.Created, tbl_Target_List.Last_Modified\015\012FROM ((tbl_Targ"
    "et_Species LEFT JOIN tbl_Target_Areas ON tbl_Target_Species.Target_Area_ID = tbl"
    "_Target_Areas.Target_Area_ID) LEFT JOIN tbl_Target_List ON tbl_Target_Species.Tg"
    "t_List_ID_FK = tbl_Target_List.Tgt_List_ID) LEFT JOIN tlu_NCPN_Plants ON tbl_Tar"
    "get_Species.Master_Plant_Code_FK = tlu_NCPN_Plants.Master_PLANT_Code;\015\012"
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
dbText "Description" ="Park target species listings including priority, transect_only, and target_area "
    "(Target List Tool update)"
Begin
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Park"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TgtYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Target_Area_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Master_Plant_Code_FK"
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
        dbText "Name" ="Tgt_Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TgtList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.LU_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_List.Created"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_List.Last_Modified"
        dbLong "AggregateType" ="-1"
    End
End
