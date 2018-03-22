Operation =1
Option =0
Where ="(((tbl_Sapling_Data.Sapling_Status)<>\"Missing\" And (tbl_Sapling_Data.Sapling_S"
    "tatus)<>\"Removed from Study\" And (tbl_Sapling_Data.Sapling_Status)<>\"Downgrad"
    "ed to Non-Sampled\") AND ((tbl_Sapling_Data.Habit)=\"Tree\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Sapling_Data"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tlu_Plants.TSN"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Sapling_Data.*"
    Alias ="Sample_Year"
    Expression ="Year([tbl_Events].[Event_Date])"
    Expression ="tbl_Events.Event_Date"
    Alias ="StemList"
    Expression ="MakeSaplingStemList([tbl_Sapling_Data].[Event_ID],[tbl_Sapling_Data].[Sapling_Da"
        "ta_ID])"
    Expression ="tlu_Plants.Shrub"
    Expression ="tbl_Tags.Microplot_Number"
    Alias ="Cycle"
    Expression ="1+Int((Year([Event_Date])-2006)/4)"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Locations.Subunit_Code"
    Expression ="tlu_Plants.Exotic"
    Alias ="Dead"
    Expression ="IIf([Sapling_Status]=\"Dead\" Or [Sapling_Status]=\"Dead Standing\" Or [Sapling_"
        "Status]=\"Dead Fallen\",\"Y\",\"N\")"
End
Begin Joins
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =3
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
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="735"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2940"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.DRC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Shrub"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="StemList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="900"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="Dead"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =51
    Top =18
    Right =1329
    Bottom =566
    Left =-1
    Top =-1
    Right =1246
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =183
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =282
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =625
        Top =163
        Right =769
        Bottom =350
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
