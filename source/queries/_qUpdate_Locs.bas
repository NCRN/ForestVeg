﻿dbMemo "SQL" ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, tbl_Locations.Slope, "
    "tbl_Locations.Aspect, [_tbl_Locations_Import_20180706_PRIMARY].Slope, [_tbl_Loca"
    "tions_Import_20180706_PRIMARY].Aspect\015\012FROM _tbl_Locations_Import_20180706"
    "_PRIMARY INNER JOIN tbl_Locations ON [_tbl_Locations_Import_20180706_PRIMARY].Lo"
    "cation_ID = tbl_Locations.Location_ID\015\012WHERE (((tbl_Locations.Slope) Is Nu"
    "ll) AND ((tbl_Locations.Aspect) Is Null) AND (([_tbl_Locations_Import_20180706_P"
    "RIMARY].Slope) Is Not Null) AND (([_tbl_Locations_Import_20180706_PRIMARY].Aspec"
    "t) Is Not Null)) OR (((tbl_Locations.Slope) Is Null) AND (([_tbl_Locations_Impor"
    "t_20180706_PRIMARY].Slope) Is Not Null)) OR (((tbl_Locations.Aspect) Is Null) AN"
    "D (([_tbl_Locations_Import_20180706_PRIMARY].Aspect) Is Not Null));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
Begin
End
