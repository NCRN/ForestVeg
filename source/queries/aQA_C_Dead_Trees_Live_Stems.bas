Operation =1
Option =0
Where ="(((tbl_Tree_DBH.Live)=True) AND ((tbl_Tree_Data.Tree_Status) Like \"Dead*\"))"
Begin InputTables
    Name ="tbl_Events"
    Name ="tbl_Locations"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tree_DBH"
    Name ="tbl_Tags"
End
Begin OutputColumns
    Expression ="tbl_Locations.Unit_Code"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tree_DBH.DBH"
    Expression ="tbl_Tree_DBH.Live"
    Expression ="tbl_Tree_Data.Tree_Status"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_DBH"
    Expression ="tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbMemo "OrderBy" ="[Query1].[Tag]"
dbBoolean "OrderByOn" ="-1"
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
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_DBH.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1391
    Bottom =724
    Left =-1
    Top =-1
    Right =1359
    Bottom =401
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =249
        Top =57
        Right =393
        Bottom =201
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =0
        Top =197
        Right =144
        Bottom =341
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =615
        Bottom =360
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =833
        Top =15
        Right =1077
        Bottom =402
        Top =0
        Name ="tbl_Tree_DBH"
        Name =""
    End
    Begin
        Left =1158
        Top =98
        Right =1302
        Bottom =242
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
