dbMemo "SQL" ="SELECT tbl_Locations.*\015\012FROM tbl_Locations\015\012WHERE (((tbl_Locations.U"
    "nit_Code) Like Nz(IIf(Abs([Forms]![frm_Data_Summary_Advanced]![togFilterByPark])"
    "=1,[Forms]![frm_Data_Summary_Advanced]![cboParkFilter],Null),\"*\")) AND ((tbl_L"
    "ocations.Admin_Unit_Code) Like Nz(IIf(Abs([Forms]![frm_Data_Summary_Advanced]![t"
    "ogFilterByAdminPark])=1,[Forms]![frm_Data_Summary_Advanced]![cboAdminParkFilter]"
    ",Null),\"*\")) AND ((tbl_Locations.Subunit_Code) Like Nz(IIf(Abs([Forms]![frm_Da"
    "ta_Summary_Advanced]![togFilterBySubunit])=1,[Forms]![frm_Data_Summary_Advanced]"
    "![cboSubunitFilter],Null),\"*\")) AND ((tbl_Locations.Panel) Like Nz(IIf(Abs([Fo"
    "rms]![frm_Data_Summary_Advanced]![togFilterByPanel])=1,[Forms]![frm_Data_Summary"
    "_Advanced]![cboPanelFilter],Null),\"*\")) AND ((tbl_Locations.Frame) Like Nz(IIf"
    "(Abs([Forms]![frm_Data_Summary_Advanced]![togFilterByFrame])=1,[Forms]![frm_Data"
    "_Summary_Advanced]![cboFrameFilter],Null),\"*\")) AND ((tbl_Locations.Location_S"
    "tatus) Like Nz(IIf(Abs([Forms]![frm_Data_Summary_Advanced]![togFilterByStatus])="
    "1,[Forms]![frm_Data_Summary_Advanced]![cboStatusFilter],Null),\"*\")) AND ((tbl_"
    "Locations.Location_ID) Like Nz(IIf(Abs([Forms]![frm_Data_Summary_Advanced]![togF"
    "ilterByLocation])=1,[Forms]![frm_Data_Summary_Advanced]![cboLocationFilter],Null"
    "),\"*\")));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbText "Description" ="Standard subquery to filter location records based on filter values in frm_Summa"
    "ry_Tool"
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
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.X_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Y_Coord"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Coord_Units"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Coord_System"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.UTM_Zone"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Datum"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.GRTS_Order"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Install_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Lon_WGS84"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Lat_WGS84"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
End
