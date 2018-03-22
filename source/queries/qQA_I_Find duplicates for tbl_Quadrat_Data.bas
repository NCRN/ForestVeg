Operation =1
Option =0
Having ="(((Count(tbl_Quadrat_Data.Quadrat_Number))>1) AND ((Count(tbl_Quadrat_Data.Event"
    "_ID))>1))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
End
Begin OutputColumns
    Alias ="FirstOfPlot_Name"
    Expression ="First(tbl_Locations.Plot_Name)"
    Alias ="FirstOfStart_Date"
    Expression ="First(tbl_Events.Event_Date)"
    Alias ="Quadrat_Number Field"
    Expression ="First(tbl_Quadrat_Data.Quadrat_Number)"
    Alias ="NumberOfDups"
    Expression ="Count(tbl_Quadrat_Data.Quadrat_Number)"
    Alias ="Event_ID Field"
    Expression ="First(tbl_Quadrat_Data.Event_ID)"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_Events"
    RightTable ="tbl_Quadrat_Data"
    Expression ="tbl_Events.Event_ID=tbl_Quadrat_Data.Event_ID"
    Flag =1
End
Begin OrderBy
    Expression ="First(tbl_Locations.Plot_Name)"
    Flag =0
    Expression ="First(tbl_Events.Event_Date)"
    Flag =0
End
Begin Groups
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    GroupLevel =0
    Expression ="tbl_Quadrat_Data.Event_ID"
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
dbText "Description" ="Duplicate quadrat data record exists"
Begin
    Begin
        dbText "Name" ="Event_ID Field"
        dbInteger "ColumnWidth" ="4275"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfStart_Date"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="FirstOfPlot_Name"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Quadrat_Number Field"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NumberOfDups"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =74
    Top =214
    Right =936
    Bottom =643
    Left =-1
    Top =-1
    Right =838
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
End
