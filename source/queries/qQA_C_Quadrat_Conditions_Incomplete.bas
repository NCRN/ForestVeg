Operation =1
Option =0
Where ="(((tbl_Quadrat_Data.Percent_Ferns) Is Null)) OR (((tbl_Quadrat_Data.Percent_Tree"
    "s) Is Null)) OR (((tbl_Quadrat_Data.Percent_Bryophytes) Is Null)) OR (((tbl_Quad"
    "rat_Data.Percent_Rock) Is Null)) OR (((tbl_Quadrat_Data.Percent_Woody_Debris) Is"
    " Null)) OR (((tbl_Quadrat_Data.Percent_Other) Is Null)) OR (((tbl_Quadrat_Data.P"
    "ercent_Grasses) Is Null)) OR (((tbl_Quadrat_Data.Percent_Sedges) Is Null)) OR (("
    "(tbl_Quadrat_Data.Percent_Herbs) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tbl_Quadrat_Data.Browse"
    Expression ="tbl_Quadrat_Data.Percent_Trees"
    Expression ="tbl_Quadrat_Data.Percent_Bryophytes"
    Expression ="tbl_Quadrat_Data.Percent_Rock"
    Expression ="tbl_Quadrat_Data.Percent_Woody_Debris"
    Expression ="tbl_Quadrat_Data.Percent_Other"
    Expression ="tbl_Quadrat_Data.Percent_Grasses"
    Expression ="tbl_Quadrat_Data.Percent_Sedges"
    Expression ="tbl_Quadrat_Data.Percent_Herbs"
    Expression ="tbl_Quadrat_Data.Percent_Ferns"
    Expression ="tbl_Quadrat_Data.Quadrat_Notes"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID"
    Expression ="tbl_Quadrat_Data.Event_ID"
    Expression ="tbl_Locations.Panel"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="tbl_Events.Event_ID=tbl_Quadrat_Data.Event_ID"
    Flag =3
End
Begin OrderBy
    Expression ="tbl_Locations.Plot_Name"
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
dbText "Description" ="Quadrat conditions are incomplete"
Begin
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Data_ID"
        dbInteger "ColumnWidth" ="1485"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Event_ID"
        dbInteger "ColumnWidth" ="1290"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Trees"
        dbInteger "ColumnWidth" ="915"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Bryophytes"
        dbInteger "ColumnWidth" ="795"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Rock"
        dbInteger "ColumnWidth" ="690"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Woody_Debris"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Other"
        dbInteger "ColumnWidth" ="885"
        dbBoolean "ColumnHidden" ="0"
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
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Browse"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Grasses"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Sedges"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Herbs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Percent_Ferns"
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
    ColumnsShown =539
    Begin
        Left =462
        Top =3
        Right =558
        Bottom =117
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =265
        Top =6
        Right =361
        Bottom =120
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =153
        Bottom =120
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
End
