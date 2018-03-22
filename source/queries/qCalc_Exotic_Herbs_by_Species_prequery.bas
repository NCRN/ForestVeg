Operation =1
Option =0
Where ="(((tlu_Plants.Exotic)=True))"
Begin InputTables
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
    Name ="tbl_Quadrat_Data"
    Name ="tbl_Quadrat_Herbaceous_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Location_ID"
    Expression ="qFiltered_Events.Event_ID"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID"
    Expression ="tbl_Quadrat_Herbaceous_Data.Quadrat_Herbaceous_ID"
    Expression ="qFiltered_Events.Event_Year"
    Expression ="qFiltered_Locations.Panel"
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tbl_Quadrat_Herbaceous_Data.TSN"
    Expression ="tbl_Quadrat_Herbaceous_Data.Percent_Cover"
    Expression ="qFiltered_Locations.Unit_Code"
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Expression ="tlu_Plants.Exotic"
End
Begin Joins
    LeftTable ="tbl_Quadrat_Data"
    RightTable ="tbl_Quadrat_Herbaceous_Data"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID=tbl_Quadrat_Herbaceous_Data.Quadrat_Data_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="qFiltered_Events.Event_ID=tbl_Quadrat_Data.Event_ID"
    Flag =1
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID=qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Quadrat_Herbaceous_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_Quadrat_Herbaceous_Data.TSN=tlu_Plants.TSN"
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
        dbText "Name" ="qFiltered_Locations.Location_ID"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Herbaceous_Data.Quadrat_Herbaceous_ID"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Events.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
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
        dbText "Name" ="qFiltered_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =11
    Top =140
    Right =1272
    Bottom =802
    Left =-1
    Top =-1
    Right =1237
    Bottom =341
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =315
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =404
        Bottom =329
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =263
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =265
        Top =0
        Name ="tbl_Quadrat_Herbaceous_Data"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
