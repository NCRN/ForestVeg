Operation =1
Option =0
Having ="(((Avg(tbl_Soil_Samples.pH)) Is Not Null))"
Begin InputTables
    Name ="tbl_Events"
    Name ="tbl_Soil_Samples"
    Name ="tbl_Soil_Data"
    Name ="tbl_Locations"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.X_Coord"
    Expression ="tbl_Locations.Y_Coord"
    Alias ="Samp_Year"
    Expression ="Year([Event_Date])"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="Mean_pH"
    Expression ="Avg(tbl_Soil_Samples.pH)"
    Alias ="Mean_EqpH"
    Expression ="Avg(tbl_Soil_Samples.Equiv_Water_pH)"
    Alias ="Mean_BSat"
    Expression ="Avg(tbl_Soil_Samples.[Base_Saturation_%])"
    Alias ="Mean_T_C"
    Expression ="Avg(tbl_Soil_Samples.[Total_C_%])"
End
Begin Joins
    LeftTable ="tbl_Soil_Data"
    RightTable ="tbl_Soil_Samples"
    Expression ="tbl_Soil_Data.Soil_Data_ID = tbl_Soil_Samples.Soil_Data_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Soil_Data"
    Expression ="tbl_Events.Event_ID = tbl_Soil_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
End
Begin Groups
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =0
    Expression ="tbl_Locations.X_Coord"
    GroupLevel =0
    Expression ="tbl_Locations.Y_Coord"
    GroupLevel =0
    Expression ="Year([Event_Date])"
    GroupLevel =0
    Expression ="tbl_Locations.Panel"
    GroupLevel =0
    Expression ="tbl_Locations.Frame"
    GroupLevel =0
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
        dbText "Name" ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Samp_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_pH"
        dbInteger "ColumnWidth" ="2445"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_EqpH"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_BSat"
        dbInteger "ColumnWidth" ="2535"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mean_T_C"
        dbInteger "ColumnWidth" ="1905"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1555
    Bottom =992
    Left =-1
    Top =-1
    Right =1523
    Bottom =669
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =228
        Top =31
        Right =372
        Bottom =456
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =675
        Top =38
        Right =844
        Bottom =535
        Top =0
        Name ="tbl_Soil_Samples"
        Name =""
    End
    Begin
        Left =436
        Top =40
        Right =579
        Bottom =287
        Top =0
        Name ="tbl_Soil_Data"
        Name =""
    End
    Begin
        Left =34
        Top =30
        Right =178
        Bottom =373
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
