Operation =1
Option =0
Where ="(((tbl_Tree_Data.Tree_Status)<>\"Missing\" And (tbl_Tree_Data.Tree_Status)<>\"Re"
    "moved from Study\" And (tbl_Tree_Data.Tree_Status)<>\"Downgraded to Non-Sampled\""
    "))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Tree_Data"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Tags.TSN"
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tree_Data.*"
    Alias ="Sample_Year"
    Expression ="Year([tbl_Events].[Event_Date])"
    Expression ="tbl_Events.Event_Date"
    Alias ="StemList"
    Expression ="MakeTreeStemList([tbl_Tree_Data].[Event_ID],[tbl_Tree_Data].[Tree_Data_ID])"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Alias ="ConditionAndPest_List"
    Expression ="MakeTreeConditionList([tbl_Tree_Data].[Event_ID],[tbl_Tree_Data].[Tree_Data_ID])"
    Expression ="tlu_Plants.Exotic"
    Alias ="Dead"
    Expression ="IIf([Tree_Status]=\"Dead\" Or [Tree_Status]=\"Dead Standing\" Or [Tree_Status]=\""
        "Dead Fallen\",\"Y\",\"N\")"
End
Begin Joins
    LeftTable ="tbl_Tree_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_Tree_Data"
    Expression ="tbl_Events.Event_ID = tbl_Tree_Data.Event_ID"
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
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
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
        dbText "Name" ="tbl_Tree_Data.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Wind_Lightning_Damage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Tree_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StemList"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tree_Data.Updated_Date"
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
        dbText "Name" ="ConditionAndPest_List"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Dead"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =5
    Top =19
    Right =1299
    Bottom =535
    Left =-1
    Top =-1
    Right =1262
    Bottom =218
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
        Bottom =223
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =224
        Top =0
        Name ="tbl_Tree_Data"
        Name =""
    End
    Begin
        Left =436
        Top =12
        Right =580
        Bottom =237
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =625
        Top =166
        Right =769
        Bottom =302
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
