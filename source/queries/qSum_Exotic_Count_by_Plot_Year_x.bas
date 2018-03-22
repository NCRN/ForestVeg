Operation =6
Option =0
Where ="(((tlu_Plants.Exotic)=True))"
Begin InputTables
    Name ="qActive_Trees_Shrubs_Herbs_Vines"
    Name ="tlu_Plants"
    Name ="tbl_Locations"
    Name ="tbl_Events"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =2
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    GroupLevel =1
    Alias ="CountOfTSN"
    Expression ="Count(qActive_Trees_Shrubs_Herbs_Vines.TSN)"
End
Begin Joins
    LeftTable ="qActive_Trees_Shrubs_Herbs_Vines"
    RightTable ="tlu_Plants"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="qActive_Trees_Shrubs_Herbs_Vines"
    RightTable ="tbl_Locations"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Location_ID = tbl_Locations.Location_ID"
    Flag =1
    LeftTable ="qActive_Trees_Shrubs_Herbs_Vines"
    RightTable ="tbl_Events"
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID = tbl_Events.Event_ID"
    Flag =1
End
Begin Groups
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =2
    Expression ="tlu_Plants.Exotic"
    GroupLevel =2
    Expression ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Returns the count of exotic specimens (trees, shrubs, herbs and vines) in each p"
    "lot for each year."
Begin
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfExotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2006"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2007"
        dbInteger "ColumnOrder" ="3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2008"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2009"
        dbInteger "ColumnOrder" ="5"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2010"
        dbInteger "ColumnOrder" ="6"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfTSN"
    End
    Begin
        dbText "Name" ="2011"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =231
    Top =96
    Right =953
    Bottom =658
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =205
        Bottom =230
        Top =0
        Name ="qActive_Trees_Shrubs_Herbs_Vines"
        Name =""
    End
    Begin
        Left =271
        Top =4
        Right =415
        Bottom =234
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
    Begin
        Left =517
        Top =12
        Right =661
        Bottom =156
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =774
        Top =89
        Right =918
        Bottom =233
        Top =0
        Name ="tbl_Events"
        Name =""
    End
End
