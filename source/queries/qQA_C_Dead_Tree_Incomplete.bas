Operation =1
Option =0
Where ="(((tbl_Tags.TSN) Is Null) AND ((tbl_Events.Event_Date) Is Not Null) AND ((tbl_Tr"
    "ee_Data.Tree_Status)=\"Dead Standing\" Or (tbl_Tree_Data.Tree_Status)=\"Dead Lea"
    "ning\")) OR (((tbl_Tags.Tag) Is Null) AND ((tbl_Events.Event_Date) Is Not Null) "
    "AND ((tbl_Tree_Data.Tree_Status)=\"Dead Standing\" Or (tbl_Tree_Data.Tree_Status"
    ")=\"Dead Leaning\")) OR (((tbl_Tree_Data.Tree_Status) Is Null)) OR (((tbl_Tree_D"
    "ata.Tree_Status)=\"Dead Standing\" Or (tbl_Tree_Data.Tree_Status)=\"Dead Leaning"
    "\") AND ((tbl_Tags.Azimuth) Is Null)) OR (((tbl_Tree_Data.Tree_Status)=\"Dead St"
    "anding\" Or (tbl_Tree_Data.Tree_Status)=\"Dead Leaning\") AND ((tbl_Tags.Distanc"
    "e) Is Null)) OR (((tbl_Tree_Data.Tree_Status)=\"Dead Standing\" Or (tbl_Tree_Dat"
    "a.Tree_Status)=\"Dead Leaning\") AND ((tbl_Tree_DBH.DBH) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_DBH"
End
Begin OutputColumns
    Expression ="tbl_Tags.TSN"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tree_Data.Tree_Status"
    Expression ="tbl_Locations.Panel"
    Alias ="EventTxt"
    Expression ="StringFromGUID([tbl_Tree_Data]![Event_ID])"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tbl_Tree_DBH.DBH"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Tags"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = tbl_Tree_Data.Tag_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =3
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_DBH"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Tags.Tag"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbText "Description" ="Tree sampling record is incomplete"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="EventTxt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Azimuth"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =2
    Top =196
    Right =1497
    Bottom =759
    Left =-1
    Top =-1
    Right =1463
    Bottom =256
    Left =0
    Top =0
    ColumnsShown =539
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
        Left =825
        Top =93
        Right =969
        Bottom =237
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =434
        Top =91
        Right =578
        Bottom =235
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =615
        Top =9
        Right =759
        Bottom =153
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =999
        Top =12
        Right =1293
        Bottom =217
        Top =0
        Name ="tbl_Tree_DBH"
        Name =""
    End
End
