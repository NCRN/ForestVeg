Operation =1
Option =0
Where ="(((tbl_Quadrat_Seedlings_Data.Height) Is Null Or (tbl_Quadrat_Seedlings_Data.Hei"
    "ght)=0 Or (tbl_Quadrat_Seedlings_Data.Height)>300))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_Quadrat_Data"
    Name ="tbl_Quadrat_Seedlings_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_Quadrat_Data.Quadrat_Number"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Quadrat_Seedlings_Data.Height"
    Expression ="tbl_Events.Event_ID"
    Expression ="tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID"
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
    LeftTable ="tbl_Quadrat_Seedlings_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_Quadrat_Seedlings_Data.TSN=tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Quadrat_Data"
    RightTable ="tbl_Quadrat_Seedlings_Data"
    Expression ="tbl_Quadrat_Data.Quadrat_Data_ID=tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID"
    Flag =3
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
dbText "Description" ="Seddling height record is incomplete or outside of expected range"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1335"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Data.Quadrat_Number"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Height"
        dbInteger "ColumnWidth" ="1110"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbInteger "ColumnWidth" ="1650"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
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
        Left =440
        Top =6
        Right =536
        Bottom =120
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="tbl_Events"
        Name =""
    End
    Begin
        Left =172
        Top =3
        Right =295
        Bottom =120
        Top =0
        Name ="tbl_Quadrat_Data"
        Name =""
    End
    Begin
        Left =19
        Top =2
        Right =145
        Bottom =117
        Top =0
        Name ="tbl_Quadrat_Seedlings_Data"
        Name =""
    End
    Begin
        Left =584
        Top =12
        Right =728
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
