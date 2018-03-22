Operation =1
Option =0
Where ="(((tbl_Plot_Floor_Condition_Data.Rock_Cover) Is Null)) OR (((tbl_Plot_Floor_Cond"
    "ition_Data.Bare_Soil_Cover) Is Null)) OR (((tbl_Plot_Floor_Condition_Data.Trampl"
    "ed) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Plot_Floor_Condition_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Plot_Floor_Condition_Data.Rock_Cover"
    Expression ="tbl_Plot_Floor_Condition_Data.Bare_Soil_Cover"
    Expression ="tbl_Plot_Floor_Condition_Data.Trampled"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Plot_Floor_Condition_Data.Event_ID"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID = tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Plot_Floor_Condition_Data"
    Expression ="tbl_Events.Event_ID = tbl_Plot_Floor_Condition_Data.Event_ID"
    Flag =2
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
    Flag =0
    Expression ="tbl_Events.Event_Date"
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
dbText "Description" ="The Plot Floor sampling fields are incomplete"
Begin
    Begin
        dbText "Name" ="tbl_Plot_Floor_Condition_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Plot_Floor_Condition_Data.Rock_Cover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Plot_Floor_Condition_Data.Bare_Soil_Cover"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Plot_Floor_Condition_Data.Trampled"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =233
    Right =870
    Bottom =662
    Left =0
    Top =0
    Right =0
    Bottom =0
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =425
        Top =-2
        Right =593
        Bottom =112
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =230
        Top =2
        Right =326
        Bottom =116
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =40
        Top =1
        Right =152
        Bottom =115
        Top =0
        Name ="tbl_Plot_Floor_Condition_Data"
        Name =""
    End
End
