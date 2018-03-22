Operation =6
Option =0
Begin InputTables
    Name ="qSum_Tree_and_Sapling_Cycles"
    Name ="tbl_Tags"
    Name ="tbl_Locations"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =2
    Expression ="tbl_Locations.Location_Status"
    GroupLevel =2
    Expression ="tbl_Tags.Tag"
    GroupLevel =2
    Expression ="qSum_Tree_and_Sapling_Cycles.Cycle"
    GroupLevel =1
    Alias ="Earliest_Event"
    Expression ="Min(qSum_Tree_and_Sapling_Cycles.Event_Year)"
    GroupLevel =2
    Alias ="Latest_Event"
    Expression ="Max(qSum_Tree_and_Sapling_Cycles.Event_Year)"
    GroupLevel =2
    Alias ="FirstOfClass_Status"
    Expression ="First(qSum_Tree_and_Sapling_Cycles.Class_Status)"
End
Begin Joins
    LeftTable ="qSum_Tree_and_Sapling_Cycles"
    RightTable ="tbl_Tags"
    Expression ="qSum_Tree_and_Sapling_Cycles.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =3
End
Begin Groups
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =2
    Expression ="tbl_Locations.Location_Status"
    GroupLevel =2
    Expression ="tbl_Tags.Tag"
    GroupLevel =2
    Expression ="tbl_Tags.Tag_Status"
    GroupLevel =2
    Expression ="qSum_Tree_and_Sapling_Cycles.Cycle"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Most_recent_Event"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Class_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Earliest_Event"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_Status"
        dbInteger "ColumnWidth" ="1830"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1"
        dbInteger "ColumnWidth" ="3300"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfEvent_Year"
        dbInteger "ColumnWidth" ="1845"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2"
        dbInteger "ColumnWidth" ="4290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Event_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qSum_Tree_and_Sapling_Cycles.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Latest_Event"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-16
    Top =-10
    Right =1502
    Bottom =915
    Left =-1
    Top =-1
    Right =1486
    Bottom =540
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =78
        Top =12
        Right =279
        Bottom =307
        Top =0
        Name ="qSum_Tree_and_Sapling_Cycles"
        Name =""
    End
    Begin
        Left =323
        Top =12
        Right =498
        Bottom =338
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =537
        Top =12
        Right =758
        Bottom =267
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
End
