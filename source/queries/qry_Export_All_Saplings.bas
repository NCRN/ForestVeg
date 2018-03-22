dbMemo "SQL" ="SELECT tbl_Locations.Unit_Code, tbl_Locations.Admin_Unit_Code, tbl_Locations.Pan"
    "el, qActive_Sapling_Data.Sample_Year, qActive_Sapling_Data.Plot_Name, qActive_Sa"
    "pling_Data.Microplot_Number, qActive_Sapling_Data.Tree_Tag, tbl_Locations.X_Coor"
    "d, tbl_Locations.Y_Coord, tlu_Plants.Latin_Name, qCalc_Basal_Area_Per_Sapling.St"
    "ems, qCalc_Basal_Area_Per_Sapling.SumBasalArea_cm2, qCalc_Basal_Area_Per_Sapling"
    ".Equiv_DBH_cm, qActive_Sapling_Data.DRC, qActive_Sapling_Data.Browse\015\012FROM"
    " qCalc_Basal_Area_Per_Sapling RIGHT JOIN (((tlu_Plants RIGHT JOIN qActive_Saplin"
    "g_Data ON tlu_Plants.TSN=qActive_Sapling_Data.TSN) LEFT JOIN tbl_Locations ON qA"
    "ctive_Sapling_Data.Location_ID=tbl_Locations.Location_ID) LEFT JOIN qry_Sapling_"
    "Summary ON qActive_Sapling_Data.Sapling_Data_ID=qry_Sapling_Summary.Sapling_Data"
    "_ID) ON qCalc_Basal_Area_Per_Sapling.Sapling_Data_ID=qActive_Sapling_Data.Saplin"
    "g_Data_ID\015\012ORDER BY qry_Active_Sapling_Data.Tree_Tag;\015\012"
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
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.X_Coord"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Y_Coord"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="2"
    End
    Begin
        dbText "Name" ="qry_Active_Sapling_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="3"
    End
    Begin
        dbText "Name" ="qry_Active_Sapling_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnOrder" ="4"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Sapling_Data.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Sapling_Data.Tree_Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Sapling_Data.DRC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Active_Sapling_Data.Browse"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Calc_Basal_Area_Per_Sapling.Stems"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Calc_Basal_Area_Per_Sapling.Equiv_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qry_Calc_Basal_Area_Per_Sapling.SumBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
End
