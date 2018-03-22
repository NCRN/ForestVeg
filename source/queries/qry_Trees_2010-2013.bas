Operation =1
Option =0
Having ="(((tbl_Locations.Unit_Code)=\"wotr\") AND ((Year([Event_Date])) Between 2009 And"
    " 2013))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tbl_Tree_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Alias ="Year"
    Expression ="Year([Event_Date])"
    Expression ="tlu_Plants.TaxonCode"
    Expression ="tlu_Plants.TSN"
    Expression ="tlu_Plants.Family"
    Expression ="tlu_Plants.Genus"
    Expression ="tlu_Plants.Species"
    Expression ="tlu_Plants.Subspecies"
    Expression ="tlu_Plants.Common"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Tree_Data.Tag_ID"
    Flag =1
    LeftTable ="tlu_Plants"
    RightTable ="tbl_Tags"
    Expression ="tlu_Plants.TSN = tbl_Tags.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
End
Begin Groups
    Expression ="tbl_Locations.Unit_Code"
    GroupLevel =0
    Expression ="Year([Event_Date])"
    GroupLevel =0
    Expression ="tlu_Plants.TaxonCode"
    GroupLevel =0
    Expression ="tlu_Plants.TSN"
    GroupLevel =0
    Expression ="tlu_Plants.Family"
    GroupLevel =0
    Expression ="tlu_Plants.Genus"
    GroupLevel =0
    Expression ="tlu_Plants.Species"
    GroupLevel =0
    Expression ="tlu_Plants.Subspecies"
    GroupLevel =0
    Expression ="tlu_Plants.Common"
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
        dbText "Name" ="tbl_Locations.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Family"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Genus"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Species"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Subspecies"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Common"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =182
    Top =45
    Right =1185
    Bottom =1002
    Left =-1
    Top =-1
    Right =971
    Bottom =538
    Left =384
    Top =0
    ColumnsShown =543
    Begin
        Left =-336
        Top =12
        Right =-192
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =-112
        Top =258
        Right =138
        Bottom =536
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =462
        Top =155
        Right =666
        Bottom =551
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =426
        Bottom =401
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =781
        Top =82
        Right =1081
        Bottom =562
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
