Operation =1
Option =0
Having ="(((qActive_Herbaceous_Data.Exotic)=True))"
Begin InputTables
    Name ="qActive_Herbaceous_Data"
End
Begin OutputColumns
    Expression ="qActive_Herbaceous_Data.Admin_Unit_Code"
    Expression ="qActive_Herbaceous_Data.Plot_Name"
    Expression ="qActive_Herbaceous_Data.Sample_Year"
    Expression ="qActive_Herbaceous_Data.Quadrat_Number"
    Expression ="qActive_Herbaceous_Data.Exotic"
    Alias ="MaxOfPercent_Cover"
    Expression ="Max(qActive_Herbaceous_Data.Percent_Cover)"
    Expression ="qActive_Herbaceous_Data.Location_ID"
    Expression ="qActive_Herbaceous_Data.Event_ID"
End
Begin Groups
    Expression ="qActive_Herbaceous_Data.Admin_Unit_Code"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Sample_Year"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Quadrat_Number"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Exotic"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Location_ID"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Event_ID"
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
Begin
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Admin_Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Quadrat_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MaxOfPercent_Cover"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =170
    Top =113
    Right =1143
    Bottom =675
    Left =-1
    Top =-1
    Right =941
    Bottom =291
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =73
        Top =9
        Right =342
        Bottom =398
        Top =0
        Name ="qActive_Herbaceous_Data"
        Name =""
    End
End
