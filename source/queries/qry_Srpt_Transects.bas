Operation =1
Option =0
Begin InputTables
    Name ="tbl_CWD_Data"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_CWD_Data.Event_ID"
    Expression ="tbl_CWD_Data.Transect_Azimuth"
    Expression ="tbl_CWD_Data.Decay_Class"
    Expression ="tbl_CWD_Data.TSN"
    Expression ="tbl_CWD_Data.Diameter"
    Expression ="tbl_CWD_Data.Hollow"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_CWD_Data.CWD_Notes"
End
Begin Joins
    LeftTable ="tbl_CWD_Data"
    RightTable ="tbl_Tags"
    Expression ="tbl_CWD_Data.Tag_ID = tbl_Tags.Tag_ID"
    Flag =2
    LeftTable ="tbl_CWD_Data"
    RightTable ="tlu_Plants"
    Expression ="tbl_CWD_Data.TSN = tlu_Plants.TSN"
    Flag =2
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
        dbText "Name" ="tbl_CWD_Data.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Transect_Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Decay_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Diameter"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.Hollow"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_CWD_Data.CWD_Notes"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =106
    Top =118
    Right =1120
    Bottom =661
    Left =-1
    Top =-1
    Right =982
    Bottom =254
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =406
        Bottom =247
        Top =0
        Name ="tbl_CWD_Data"
        Name =""
    End
    Begin
        Left =803
        Top =114
        Right =947
        Bottom =258
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =646
        Top =12
        Right =790
        Bottom =156
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
