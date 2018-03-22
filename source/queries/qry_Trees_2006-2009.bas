Operation =1
Option =0
Having ="(((tbl_Locations.Unit_Code)=\"WOTR\") AND ((Year([Event_Date]))<=2009))"
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
Begin OrderBy
    Expression ="tlu_Plants.Family"
    Flag =0
    Expression ="tlu_Plants.Genus"
    Flag =0
    Expression ="tlu_Plants.Species"
    Flag =0
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
    Left =2
    Top =3
    Right =1233
    Bottom =960
    Left =-1
    Top =-1
    Right =1199
    Bottom =487
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =272
        Top =258
        Right =522
        Bottom =536
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =858
        Top =155
        Right =1050
        Bottom =392
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =810
        Bottom =401
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =1072
        Top =47
        Right =1372
        Bottom =571
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
