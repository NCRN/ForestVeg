Operation =1
Option =0
Where ="(((tbl_Tags.TSN)<>tlu_Plants!TSN_Accepted)) Or (((tbl_Tags.Tag_Status) Is Null))"
    " Or (((tbl_Tags.Tag_Status)=\"Tree\") And ((tbl_Tags.Azimuth) Is Null)) Or (((tb"
    "l_Tags.Tag_Status)=\"Tree\") And ((tbl_Tags.Distance) Is Null)) Or (((tbl_Tags.T"
    "ag_Status)=\"Sapling\") And ((tbl_Tags.Microplot_Number) Is Null))"
Begin InputTables
    Name ="tbl_Locations"
    Name ="tbl_Tags"
    Name ="tlu_Plants"
End
Begin OutputColumns
    Expression ="tbl_Locations.Plot_Name"
    Expression ="tbl_Locations.Panel"
    Expression ="tbl_Locations.Frame"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.TSN"
    Expression ="tlu_Plants.Latin_Name"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tbl_Tags.Microplot_Number"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="tlu_Plants"
    Expression ="tbl_Tags.TSN = tlu_Plants.TSN"
    Flag =1
    LeftTable ="tbl_Locations"
    RightTable ="tbl_Tags"
    Expression ="tbl_Locations.Location_ID = tbl_Tags.Location_ID"
    Flag =1
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
dbText "Description" ="The Tag record is missing information or contains a non-accepted TSN"
Begin
    Begin
        dbText "Name" ="tbl_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Locations.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tlu_Plants.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =8
    Top =233
    Right =1385
    Bottom =662
    Left =-1
    Top =-1
    Right =1345
    Bottom =162
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =390
        Top =0
        Name ="tbl_Locations"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =273
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =448
        Top =0
        Name ="tlu_Plants"
        Name =""
    End
End
