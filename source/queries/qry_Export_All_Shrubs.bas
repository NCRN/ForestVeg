dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, tbl_Locations.Admin_Unit_Code, tbl_Locations.Pan"
    "el, qry_Active_Shrub_Data.Sample_Year, qry_Active_Shrub_Data.Plot_Name, qry_Acti"
    "ve_Shrub_Data.TSN, tlu_Plants.Latin_Name, qry_Active_Shrub_Data.Sapling_Status, "
    "qry_Active_Shrub_Data.Browse, qry_Active_Shrub_Data.DRC, qry_Calc_Basal_Area_Per"
    "_Shrub.Equiv_DBH_cm, qry_Calc_Basal_Area_Per_Shrub.SumBasalArea_cm2, qry_Active_"
    "Shrub_Data.Tree_Tag\015\012FROM tbl_Locations RIGHT JOIN (((qry_Active_Shrub_Dat"
    "a INNER JOIN tlu_Plants ON qry_Active_Shrub_Data.TSN=tlu_Plants.TSN) INNER JOIN "
    "tbl_Events ON qry_Active_Shrub_Data.Event_ID=tbl_Events.Event_ID) LEFT JOIN qry_"
    "Calc_Basal_Area_Per_Shrub ON qry_Active_Shrub_Data.Sapling_Data_ID=qry_Calc_Basa"
    "l_Area_Per_Shrub.Sapling_Data_ID) ON tbl_Locations.Location_ID=tbl_Events.Locati"
    "on_ID\015\012ORDER BY qry_Active_Shrub_Data.Tree_Tag;\015\012"
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
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Active_Shrub_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Active_Shrub_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Shrub_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Shrub_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Shrub_Data.DRC"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1155"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Active_Shrub_Data.Browse"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1080"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qry_Calc_Basal_Area_Per_Shrub.Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Calc_Basal_Area_Per_Shrub.SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Shrub_Data.Tree_Tag"
        dbLong "AggregateType" ="-1"
    End
End
