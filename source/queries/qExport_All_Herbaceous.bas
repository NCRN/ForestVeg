Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
    Name ="tlu_Plants"
    Name ="tbl_Quadrat_Herbaceous_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Unit_Group"
    Expression ="tbl_Locations.Subunit_Code"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="Sample_Year"
    Expression ="Year([Event_Date])"
    Alias ="Date"
    Expression ="CLng(Format([tbl_events].[Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tbl_Quadrat_Herbaceous_Data.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.Exotic"
    Expression ="tbl_Quadrat_Herbaceous_Data.Percent_Cover"
    Expression ="tlu_Plants.TaxonCode"
End
Begin Joins
    LeftTable ="tbl_Quadrat_Herbaceous_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_Quadrat_Herbaceous_Data.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="tbl_Events.Event_ID = tbl_Quadrat_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Quadrat_Data"
    RightTable ="tbl_Quadrat_Herbaceous_Data"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID = tbl_Quadrat_Herbaceous_Data.Quadrat_Data_ID"
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
        dbText "Name" ="Date"
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
        dbText "Name" ="tbl_Quadrat_Herbaceous_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Herbaceous_Data.Percent_Cover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =15
    Top =159
    Right =1179
    Bottom =758
    Left =-1
    Top =-1
    Right =1132
    Bottom =320
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =890
        Top =5
        Right =986
        Bottom =119
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =704
        Top =5
        Right =862
        Bottom =205
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =410
        Top =-2
        Right =674
        Bottom =290
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =254
        Top =114
        Right =401
        Bottom =258
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =22
        Top =6
        Right =221
        Bottom =150
        Top =0
        Name ="tbl_Quadrat_Herbaceous_Data"
        Name =""
    End
End
