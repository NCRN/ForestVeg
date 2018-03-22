Operation =1
Option =0
Having ="(((qActive_Trees_and_Shrubs.Class)=\"Seedling\"))"
Begin InputTables
    Name ="qFiltered_Events"
    Name ="qFiltered_Locations"
    Name ="qSum_PARK_Event_Count"
    Name ="qActive_Trees_and_Shrubs"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Alias ="Event_Count"
    Expression ="Max(qSum_PARK_Event_Count.Plot_Count)"
    Expression ="qActive_Trees_and_Shrubs.Habit"
    Expression ="qActive_Trees_and_Shrubs.Class"
    Alias ="Seedling_Count"
    Expression ="Count(qActive_Trees_and_Shrubs.Sample_ID)"
    Alias ="Seedlings_per_ha"
    Expression ="Round([Seedling_Count]/([Event_Count]*0.0012),0)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Locations"
    RightTable ="qSum_PARK_Event_Count"
    Expression ="qFiltered_Locations.Admin_Unit_Code = qSum_PARK_Event_Count.Admin_Unit_Code"
    Flag =2
    LeftTable ="qFiltered_Events"
    RightTable ="qActive_Trees_and_Shrubs"
    Expression ="qFiltered_Events.Event_ID = qActive_Trees_and_Shrubs.Event_ID"
    Flag =1
End
Begin Groups
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.Habit"
    GroupLevel =0
    Expression ="qActive_Trees_and_Shrubs.Class"
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
dbText "Description" ="Returns the total basal area of all tree saplings in a plot, and the percentage "
    "of that total that is made up of exotic species. Created for IAN NRCA reports."
Begin
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Event_Count"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_and_Shrubs.Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Seedling_Count"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Seedlings_per_ha"
        dbInteger "ColumnWidth" ="2010"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =452
    Right =1181
    Bottom =1014
    Left =-1
    Top =-1
    Right =1132
    Bottom =319
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =248
        Top =16
        Right =426
        Bottom =282
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
    Begin
        Left =57
        Top =17
        Right =201
        Bottom =364
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =794
        Top =143
        Right =938
        Bottom =287
        Top =0
        Name ="qSum_PARK_Event_Count"
        Name =""
    End
    Begin
        Left =514
        Top =9
        Right =713
        Bottom =249
        Top =0
        Name ="qActive_Trees_and_Shrubs"
        Name =""
    End
End
