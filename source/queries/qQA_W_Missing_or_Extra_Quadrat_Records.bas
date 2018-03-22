Operation =1
Option =0
Having ="(((Count(tbl_Quadrat_Data.Quadrat_Data_ID))<>12))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Location_ID"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Events.Event_Date"
    Alias ="CountOfQuadrat_Data_ID"
    Expression ="Count(tbl_Quadrat_Data.Quadrat_Data_ID)"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =1
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="tbl_Events.Event_ID=tbl_Quadrat_Data.Event_ID"
    Flag =1
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Events.Event_Date"
    Flag =0
End
Begin Groups
    Expression ="tbl_Locations.Plot_Name"
    GroupLevel =0
    Expression ="tbl_Locations.Location_ID"
    GroupLevel =0
    Expression ="tbl_Events.Event_ID"
    GroupLevel =0
    Expression ="tbl_Events.Event_Date"
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
dbText "Description" ="Returns events where more or less than 12 quadrat data records exist."
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CountOfQuadrat_Data_ID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =233
    Right =870
    Bottom =662
    Left =-1
    Top =-1
    Right =838
    Bottom =-1
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
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
End
