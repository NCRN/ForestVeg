Operation =1
Option =0
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
    Name ="tbl_Quadrat_Herbaceous_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Quadrat_Herbaceous_Data.Quadrat_Herbaceous_ID"
    Expression ="tbl_Quadrat_Herbaceous_Data.Quadrat_Data_ID"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Locations.Location_ID"
    Alias ="Sample_Year"
    Expression ="Year([tbl_Events].[Event_Date])"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tbl_Quadrat_Herbaceous_Data.TSN"
    Expression ="tbl_Quadrat_Herbaceous_Data.Percent_Cover"
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Admin_Unit_Code"
    Expression ="tbl_Locations.Subunit_Code"
    Expression ="tlu_Plants.Exotic"
    Expression ="tbl_Events.Event_Date"
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
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Herbaceous_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Herbaceous_Data.Percent_Cover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Herbaceous_Data.Quadrat_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Herbaceous_Data.Quadrat_Herbaceous_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =30
    Top =34
    Right =1138
    Bottom =581
    Left =-1
    Top =-1
    Right =1076
    Bottom =229
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =611
        Top =10
        Right =759
        Bottom =221
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =399
        Top =9
        Right =543
        Bottom =153
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =213
        Top =13
        Right =357
        Bottom =241
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =23
        Top =9
        Right =183
        Bottom =205
        Top =0
        Name ="tbl_Quadrat_Herbaceous_Data"
        Name =""
    End
    Begin
        Left =807
        Top =12
        Right =951
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
