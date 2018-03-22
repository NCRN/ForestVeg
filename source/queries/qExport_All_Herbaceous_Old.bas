Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="qActive_Herbaceous_Data"
    Name ="tbl_Quadrat_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="qActive_Herbaceous_Data.Panel"
    Expression ="qActive_Herbaceous_Data.Frame"
    Expression ="qActive_Herbaceous_Data.Sample_Year"
    Alias ="Date"
    Expression ="CLng(Format([tbl_events].[Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="qActive_Herbaceous_Data.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="qActive_Herbaceous_Data.Exotic"
    Expression ="qActive_Herbaceous_Data.Percent_Cover"
End
Begin Joins
    LeftTable ="qActive_Herbaceous_Data"
    RightTable ="tbl_Quadrat_Data"
    Expression ="qActive_Herbaceous_Data.Quadrat_Data_ID = tbl_Quadrat_Data.Quadrat_Data_ID"
    Flag =1
    LeftTable ="qActive_Herbaceous_Data"
    RightTable ="tlu_Plants"
    Expression ="qActive_Herbaceous_Data.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="tbl_Events.Event_ID = tbl_Quadrat_Data.Event_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Percent_Cover"
        dbInteger "ColumnWidth" ="1635"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =60
    Right =1181
    Bottom =809
    Left =-1
    Top =-1
    Right =1132
    Bottom =364
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =1029
        Top =5
        Right =1125
        Bottom =119
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =767
        Top =5
        Right =863
        Bottom =119
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =304
        Bottom =156
        Top =0
        Name ="qActive_Herbaceous_Data"
        Name =""
    End
    Begin
        Left =375
        Top =11
        Right =643
        Bottom =346
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =779
        Top =197
        Right =923
        Bottom =341
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
