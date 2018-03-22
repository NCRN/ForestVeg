Operation =1
Option =0
Where ="(((tbl_Locations.Unit_Code) Like Nz(IIf(Abs([Forms]![frm_Data_Summary]![togFilte"
    "rByPark])=1,[Forms]![frm_Data_Summary]![cboParkFilter],Null),\"*\")))"
Begin InputTables
    Name ="tbl_Locations"
End
Begin OutputColumns
    Expression ="tbl_Locations.*"
End
Begin Parameters
    Name ="[Forms]![frm_Data_Summary]![cboParkFilter]"
    Flag =10
    Name ="[Forms]![frm_Data_Summary]![togFilterByPark]"
    Flag =10
End
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
Begin
    State =0
    Left =6
    Top =348
    Right =1206
    Bottom =883
    Left =-1
    Top =-1
    Right =1176
    Bottom =140
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =43
        Top =7
        Right =247
        Bottom =191
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
