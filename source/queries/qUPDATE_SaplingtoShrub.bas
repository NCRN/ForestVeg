Operation =4
Option =0
Where ="(((Year([Event_Date]))>2014) AND ((tlu_Plants.Latin_Name)=\"Kalmia latifolia\" O"
    "r (tlu_Plants.Latin_Name)=\"Lindera benzoin\" Or (tlu_Plants.Latin_Name)=\"Ilex "
    "verticillata\"))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Sapling_Data"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Name ="tbl_Tags.Tag_Status"
    Expression ="\"Retired (In Office)\""
End
Begin Joins
    LeftTable ="tbl_Sapling_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Sapling_Data"
    Expression ="tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Stop_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Sapling_Data.Sapling_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =7
    Top =2
    Right =1470
    Bottom =828
    Left =-1
    Top =-1
    Right =1431
    Bottom =509
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =69
        Top =103
        Right =277
        Bottom =513
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =388
        Top =20
        Right =616
        Bottom =529
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =646
        Top =103
        Right =902
        Bottom =448
        Top =0
        Name ="tbl_Sapling_Data"
        Name =""
    End
    Begin
        Left =1013
        Top =27
        Right =1157
        Bottom =361
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =1210
        Top =43
        Right =1398
        Bottom =476
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
