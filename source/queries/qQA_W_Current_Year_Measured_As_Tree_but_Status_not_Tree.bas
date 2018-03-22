Operation =1
Option =0
Where ="(((qActive_Tree_Data.Sample_Year)=2013) AND ((tbl_Tags.Tag_Status)<>\"Tree\"))"
Begin InputTables
    Name ="tbl_Tags"
    Name ="qActive_Tree_Data"
End
Begin OutputColumns
    Expression ="qActive_Tree_Data.Sample_Year"
    Expression ="qActive_Tree_Data.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tbl_Tags.Microplot_Number"
    Expression ="tbl_Tags.Tag_Status"
End
Begin Joins
    LeftTable ="tbl_Tags"
    RightTable ="qActive_Tree_Data"
    Expression ="tbl_Tags.Tag_ID = qActive_Tree_Data.Tag_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qActive_Tree_Data.Plot_Name"
    Flag =0
    Expression ="tbl_Tags.Tag"
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
dbText "Description" ="Tree record exists but Tag Status is not set to Tree"
Begin
    Begin
        dbText "Name" ="qActive_Tree_Data.Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
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
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qActive_Tree_Data.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1202
    Bottom =1036
    Left =-1
    Top =-1
    Right =1132
    Bottom =514
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =267
        Top =38
        Right =503
        Bottom =398
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =26
        Top =36
        Right =211
        Bottom =490
        Top =0
        Name ="qActive_Tree_Data"
        Name =""
    End
End
