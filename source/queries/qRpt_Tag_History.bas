Operation =1
Option =0
Where ="(((tbl_Tags_History.Change_Date)>=Nz(IIf(Abs(Forms!frm_Data_Summary!togFilterByR"
    "ange)=1,Forms!frm_Data_Summary!txtStartDateFilter,#1/1/1800#)) And (tbl_Tags_His"
    "tory.Change_Date)<=Nz(IIf(Abs(Forms!frm_Data_Summary!togFilterByRange)=1,Forms!f"
    "rm_Data_Summary!txtEndDateFilter,#12/31/2200#))) And ((Year([Change_Date])) Like"
    " Nz(IIf(Abs(Forms!frm_Data_Summary!togFilterByYear)=1,Forms!frm_Data_Summary!cbo"
    "YearFilter,Null),\"*\")))"
Begin InputTables
    Name ="tbl_Tags_History"
    Name ="tlu_Contacts"
    Name ="tbl_Tags"
    Name ="qFiltered_Locations"
End
Begin OutputColumns
    Expression ="qFiltered_Locations.Plot_Name"
    Expression ="tbl_Tags.Tag"
    Alias ="Change_Desc"
    Expression ="IIf([Field_Name]=\"TSN\",[Field_Name] & \" was changed by \" & [First_Name] & \""
        " \" & [Last_Name] & \" from \" & [Value_Old] & \" (\" & DLookUp(\"[Latin_Name]\""
        ",\"tlu_Plants\",\"[TSN] =\" & [Value_Old]) & \")\" & \" to \" & [Value_New] & \""
        " (\" & DLookUp(\"[Latin_Name]\",\"tlu_Plants\",\"[TSN] =\" & [Value_New]) & \")\""
        ",[Field_Name] & \" was changed by \" & [First_Name] & \" \" & [Last_Name] & \" f"
        "rom \" & [Value_Old] & \" to \" & [Value_New])"
    Expression ="tbl_Tags_History.Change_Date"
    Alias ="Change_Year"
    Expression ="Year([Change_Date])"
    Expression ="tbl_Tags.Azimuth"
    Expression ="tbl_Tags.Distance"
    Expression ="tbl_Tags.Microplot_Number"
    Expression ="tbl_Tags.TSN"
    Expression ="tbl_Tags.Tag_Notes"
    Expression ="tbl_Tags.Tag_Status"
    Expression ="tbl_Tags.Updated_Date"
    Alias ="Tag_ID"
    Expression ="tbl_Tags_History.Record_ID"
    Expression ="tbl_Tags.Location_ID"
End
Begin Joins
    LeftTable ="tbl_Tags_History"
    RightTable ="tlu_Contacts"
    Expression ="tbl_Tags_History.Contact_ID=tlu_Contacts.Contact_ID"
    Flag =2
    LeftTable ="tbl_Tags_History"
    RightTable ="tbl_Tags"
    Expression ="tbl_Tags_History.Record_ID=tbl_Tags.Tag_ID"
    Flag =1
    LeftTable ="tbl_Tags"
    RightTable ="qFiltered_Locations"
    Expression ="tbl_Tags.Location_ID=qFiltered_Locations.Location_ID"
    Flag =1
End
Begin OrderBy
    Expression ="qFiltered_Locations.Plot_Name"
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
Begin
    Begin
        dbText "Name" ="Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags_History.Change_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Change_Desc"
        dbInteger "ColumnWidth" ="8430"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag"
        dbInteger "ColumnWidth" ="825"
        dbBoolean "ColumnHidden" ="0"
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
    Begin
        dbText "Name" ="tbl_Tags.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Tags.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qFiltered_Locations.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Change_Year"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-19
    Top =148
    Right =1344
    Bottom =765
    Left =-1
    Top =-1
    Right =1331
    Bottom =291
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =7
        Top =7
        Right =189
        Bottom =327
        Top =0
        Name ="tbl_Tags_History"
        Name =""
    End
    Begin
        Left =256
        Top =190
        Right =400
        Bottom =328
        Top =0
        Name ="tlu_Contacts"
        Name =""
    End
    Begin
        Left =445
        Top =55
        Right =589
        Bottom =321
        Top =0
        Name ="tbl_Tags"
        Name =""
    End
    Begin
        Left =633
        Top =8
        Right =777
        Bottom =152
        Top =0
        Name ="qFiltered_Locations"
        Name =""
    End
End
