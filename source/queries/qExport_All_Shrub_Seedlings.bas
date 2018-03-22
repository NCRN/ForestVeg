Operation =1
Option =0
Where ="(((tlu_Plants.Shrub)=True))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
    Name ="tlu_Plants"
    Name ="tbl_Quadrat_Seedlings_Data"
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
    Expression ="CLng(Format([tbl_Events].[Event_Date],\"yyyymmdd\"))"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="tbl_Quadrat_Seedlings_Data.Height"
    Expression ="tbl_Quadrat_Seedlings_Data.Browsable"
    Expression ="tbl_Quadrat_Seedlings_Data.Browsed"
End
Begin Joins
    LeftTable ="tbl_Quadrat_Seedlings_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_Quadrat_Seedlings_Data.TSN = tlu_Plants.TSN"
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
    RightTable ="tbl_Quadrat_Seedlings_Data"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID = tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID"
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
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="tbl_Locations.Frame"
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
    Begin
        dbText "Name" ="Sample_Year"
        dbInteger "ColumnWidth" ="1545"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Height"
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
    Begin
        dbText "Name" ="tbl_Locations.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =164
    Right =1595
    Bottom =839
    Left =-1
    Top =-1
    Right =1555
    Bottom =230
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =922
        Top =7
        Right =1075
        Bottom =121
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =770
        Top =5
        Right =866
        Bottom =119
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =436
        Top =3
        Right =700
        Bottom =117
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =249
        Top =38
        Right =393
        Bottom =182
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =25
        Top =9
        Right =204
        Bottom =243
        Top =0
        Name ="tbl_Quadrat_Seedlings_Data"
        Name =""
    End
End
