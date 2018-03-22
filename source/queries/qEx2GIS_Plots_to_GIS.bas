Operation =1
Option =0
Begin InputTables
    Name ="qExport_All_Plots"
End
Begin OutputColumns
    Expression ="qExport_All_Plots.*"
    Alias ="X_MP_060"
    Expression ="[UTM_18N_NAD83_X]+8.66025"
    Alias ="Y_MP_060"
    Expression ="[UTM_18N_NAD83_Y]+5"
    Alias ="X_MP_180"
    Expression ="qExport_All_Plots.UTM_18N_NAD83_X"
    Alias ="Y_MP_180"
    Expression ="[UTM_18N_NAD83_Y]-10"
    Alias ="X_MP_300"
    Expression ="[UTM_18N_NAD83_X]-8.66025"
    Alias ="Y_MP_300"
    Expression ="[UTM_18N_NAD83_Y]+5"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="-1"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
dbMemo "OrderBy" ="[qEx2GIS_Plots_to_GIS].[Plot_Name], [qEx2GIS_Plots_to_GIS].[Location_Status]"
Begin
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.Location_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.Location_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.Event_Count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.Event_Earliest"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.Event_Latest"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.UTM_18N_NAD83_X"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.UTM_18N_NAD83_Y"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Y_MP_060"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X_MP_180"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Y_MP_180"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X_MP_300"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Y_MP_300"
        dbInteger "ColumnWidth" ="2100"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="X_MP_060"
        dbInteger "ColumnWidth" ="1380"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qExport_All_Plots.tbl_Locations.GRTS_Order"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =6
    Top =15
    Right =1273
    Bottom =949
    Left =-1
    Top =-1
    Right =1235
    Bottom =185
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =34
        Top =4
        Right =258
        Bottom =148
        Top =0
        Name ="qExport_All_Plots"
        Name =""
    End
End
