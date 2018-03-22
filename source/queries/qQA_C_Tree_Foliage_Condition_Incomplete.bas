Operation =1
Option =0
Where ="(((tbl_Tree_Foliage_Conditions.Condition) Is Null)) OR (((tbl_Tree_Foliage_Condi"
    "tions.Percent_Afflicted) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tags"
    Name ="tbl_Tree_Foliage_Conditions"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Events.Event_Notes"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tree_Foliage_Conditions.Condition"
    Expression ="tbl_Tree_Foliage_Conditions.Percent_Afflicted"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID=tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID=tbl_Tree_Data.Event_ID"
    Flag =1
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tree_Foliage_Conditions"
    Expression ="tbl_Tree_Data.Tree_Data_ID=tbl_Tree_Foliage_Conditions.Tree_Data_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Locations.Panel"
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
dbText "Description" ="Foliage condition record is incomplete"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
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
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Foliage_Conditions.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Foliage_Conditions.Percent_Afflicted"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =28
    Top =104
    Right =1333
    Bottom =966
    Left =-1
    Top =-1
    Right =1281
    Bottom =576
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =683
        Top =239
        Right =827
        Bottom =547
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =387
        Top =303
        Right =531
        Bottom =447
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
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tbl_Tree_Foliage_Conditions"
        Name =""
    End
    Begin
        Left =486
        Top =20
        Right =630
        Bottom =164
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
End
