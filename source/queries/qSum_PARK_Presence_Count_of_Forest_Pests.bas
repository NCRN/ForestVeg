Operation =1
Option =0
Begin InputTables
    Name ="qCalc_Tree_Condition_Count"
    Name ="qFiltered_Locations"
    Name ="qFiltered_Events"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Alias ="Trees_in_Park"
    Expression ="Count(qCalc_Tree_Condition_Count.Tree_Data_ID)"
    Alias ="Trees_with_Pests_in_Park"
    Expression ="Sum(Nz([PestPresentYN]))"
    Alias ="Total_Pests_in_Park"
    Expression ="Sum(qCalc_Tree_Condition_Count.PestCount)"
    Alias ="Percent_Trees_With_Pests"
    Expression ="Round((100*[Trees_with_Pests_in_Park])/[Trees_in_Park],1)"
End
Begin Joins
    LeftTable ="qFiltered_Locations"
    RightTable ="qFiltered_Events"
    Expression ="qFiltered_Locations.Location_ID = qFiltered_Events.Location_ID"
    Flag =1
    LeftTable ="qFiltered_Events"
    RightTable ="qCalc_Tree_Condition_Count"
    Expression ="qFiltered_Events.Event_ID = qCalc_Tree_Condition_Count.Event_ID"
    Flag =2
End
Begin OrderBy
    Expression ="qFiltered_Locations.Admin_Unit_Code"
    Flag =0
End
Begin Groups
    Expression ="qFiltered_Locations.Admin_Unit_Code"
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
dbText "Description" ="Reports the number of trees, the number of trees identified to have a pest, and "
    "the total number of pests reported for each plot. Created for IAN NRCA reports."
Begin
    Begin
        dbText "Name" ="Percent_Trees_With_Pests"
        dbInteger "ColumnWidth" ="2685"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trees_in_Park"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Trees_with_Pests_in_Park"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Total_Pests_in_Park"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =84
    Top =341
    Right =1090
    Bottom =903
    Left =-1
    Top =-1
    Right =974
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =396
        Top =12
        Right =634
        Bottom =201
        Top =0
        Name ="qCalc_Tree_Condition_Count"
        Name =""
    End
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
    Begin
        Left =222
        Top =12
        Right =366
        Bottom =156
        Top =0
        Name ="qFiltered_Events"
        Name =""
    End
End
