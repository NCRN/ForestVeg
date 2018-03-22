Operation =1
Option =0
Having ="(((qActive_Herbaceous_Data.Exotic)=True))"
Begin InputTables
    Name ="qActive_Herbaceous_Data"
End
Begin OutputColumns
    Expression ="qActive_Herbaceous_Data.Event_ID"
    Expression ="qActive_Herbaceous_Data.Plot_Name"
    Expression ="qActive_Herbaceous_Data.Sample_Year"
    Expression ="qActive_Herbaceous_Data.Exotic"
    Alias ="Count_of_Specimens"
    Expression ="Count(qActive_Herbaceous_Data.TSN)"
End
Begin OrderBy
    Expression ="qActive_Herbaceous_Data.Plot_Name"
    Flag =0
    Expression ="qActive_Herbaceous_Data.Sample_Year"
    Flag =0
End
Begin Groups
    Expression ="qActive_Herbaceous_Data.Event_ID"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Plot_Name"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Sample_Year"
    GroupLevel =0
    Expression ="qActive_Herbaceous_Data.Exotic"
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
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Trees_Shrubs_Herbs_Vines.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Count_of_Specimens"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Exotic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Herbaceous_Data.Exotic"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =176
    Top =39
    Right =1203
    Bottom =916
    Left =-1
    Top =-1
    Right =995
    Bottom =530
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =47
        Top =71
        Right =526
        Bottom =351
        Top =0
        Name ="qActive_Herbaceous_Data"
        Name =""
    End
End
