dbMemo "SQL" ="PARAMETERS park Text ( 255 ), tgtYear Short;\015\012SELECT [park] AS Park, [tgtY"
    "ear] AS TgtYear, temp_Listbox_Recordset.LUCode AS LU_code, temp_Listbox_Recordse"
    "t.Code, temp_Listbox_Recordset.Species, CInt(1) AS Priority, temp_Listbox_Record"
    "set.Transect_Only, temp_Listbox_Recordset.[Extra_Area_ID], tbl_Target_Areas.Targ"
    "et_Area AS Tgt_Area, tlu_NCPN_Plants.Master_Family AS Family, tlu_NCPN_Plants.Ma"
    "ster_Common_Name, tlu_NCPN_Plants.utah_species, tlu_NCPN_Plants.Co_Species, tlu_"
    "NCPN_Plants.Wy_Species, Park & \"-\" & TgtYear AS TgtList, Now() AS Last_Modifie"
    "d INTO temp_List_Preview\015\012FROM (temp_Listbox_Recordset LEFT JOIN tbl_Targe"
    "t_Areas ON temp_Listbox_Recordset.Extra_Area_ID = tbl_Target_Areas.Target_Area_I"
    "D) LEFT JOIN tlu_NCPN_Plants ON temp_Listbox_Recordset.Code = tlu_NCPN_Plants.Ma"
    "ster_Plant_Code\015\012ORDER BY tlu_NCPN_Plants.Master_Family, Species;\015\012"
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
dbText "Description" ="parameterized query for list previews\015\012 (Target List Tool update)"
Begin
    Begin
        dbText "Name" ="park"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tgtYear"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Listbox_Recordset.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Listbox_Recordset.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Listbox_Recordset.LUCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Listbox_Recordset.Transect_Only"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Listbox_Recordset.Target_Area_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="[park]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tgtYear]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TgtList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="temp_Listbox_Recordset.LU_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Target_Species.Species_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Priority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Tgt_Area"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Master_Common_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.utah_species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Co_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_NCPN_Plants.Wy_Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LU_code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Last_Modified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Extra_Area"
        dbLong "AggregateType" ="-1"
    End
End
