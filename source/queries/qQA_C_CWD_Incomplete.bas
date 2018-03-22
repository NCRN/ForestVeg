Operation =1
Option =0
Where ="(((tbl_CWD_Data.Decay_Class) Is Null)) OR (((tbl_CWD_Data.Diameter) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Events"
    Name ="tbl_CWD_Data"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Events.Event_Date"
    Expression ="tbl_CWD_Data.CWD_Data_ID"
    Expression ="tbl_CWD_Data.Event_ID"
    Expression ="tbl_CWD_Data.Transect_Azimuth"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_CWD_Data.Decay_Class"
    Expression ="tbl_CWD_Data.Diameter"
    Expression ="tbl_CWD_Data.Hollow"
    Expression ="tbl_CWD_Data.CWD_Notes"
    Expression ="tbl_Locations.Panel"
End
Begin Joins
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Events"
    Expression ="tbl_Locations.Location_ID=tbl_Events.Location_ID"
    Flag =3
    LeftTable ="tbl_CWD_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_CWD_Data.TSN=tlu_Plants.TSN"
    Flag =2
    LeftTable ="tbl_Events"
    RightTable ="tbl_CWD_Data"
    Expression ="tbl_Events.Event_ID=tbl_CWD_Data.Event_ID"
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
dbText "Description" ="Coarse woody debris sampling record is incomplete."
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbInteger "ColumnWidth" ="1170"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.CWD_Data_ID"
        dbInteger "ColumnWidth" ="1095"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Event_ID"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Transect_Azimuth"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Diameter"
        dbInteger "ColumnWidth" ="945"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Decay_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Hollow"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
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
        dbText "Name" ="tbl_CWD_Data.CWD_Notes"
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
        Bottom =230
        Top =0
        Name ="tbl_CWD_Data"
        Name =""
    End
    Begin
        Left =450
        Top =12
        Right =594
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
