Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="qActive_Shrub_Seedling_Data"
    Name ="tlu_Plants"
    Name ="tbl_Quadrat_Data"
End
Begin OutputColumns
    Expression ="qActive_Shrub_Seedling_Data.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Panel"
    Expression ="qActive_Shrub_Seedling_Data.Frame"
    Expression ="qActive_Shrub_Seedling_Data.Sample_Year"
    Alias ="Date"
    Expression ="CLng(Format([tbl_Events].[Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="qActive_Shrub_Seedling_Data.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qActive_Shrub_Seedling_Data.Height"
End
Begin Joins
    LeftTable ="tbl_Events"
    RightTable ="qActive_Shrub_Seedling_Data"
    Expression ="tbl_Events.Event_ID = qActive_Shrub_Seedling_Data.Event_ID"
    Flag =1
    LeftTable ="qActive_Shrub_Seedling_Data"
    RightTable ="tlu_Plants"
    Expression ="qActive_Shrub_Seedling_Data.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="qActive_Shrub_Seedling_Data"
    RightTable ="tbl_Quadrat_Data"
    Expression ="qActive_Shrub_Seedling_Data.Quadrat_Data_ID = tbl_Quadrat_Data.Quadrat_Data_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
End
Begin OrderBy
    Expression ="qActive_Shrub_Seedling_Data.Plot_Name"
    Flag =0
    Expression ="qActive_Shrub_Seedling_Data.Sample_Year"
    Flag =0
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Flag =0
End
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
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbInteger "ColumnWidth" ="2280"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Seedling_Data.Sample_Year"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Seedling_Data.Plot_Name"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Seedling_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Seedling_Data.Height"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Shrub_Seedling_Data.Frame"
        dbInteger "ColumnWidth" ="960"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =14
    Top =83
    Right =1195
    Bottom =702
    Left =-1
    Top =-1
    Right =1149
    Bottom =245
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =884
        Top =1
        Right =1028
        Bottom =307
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =267
        Top =3
        Right =411
        Bottom =147
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =290
        Top =0
        Name ="qActive_Shrub_Seedling_Data"
        Name =""
    End
    Begin
        Left =297
        Top =164
        Right =441
        Bottom =308
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =562
        Top =100
        Right =753
        Bottom =244
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
End
