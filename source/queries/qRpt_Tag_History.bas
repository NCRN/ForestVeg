dbMemo "SQL" ="SELECT fl.Plot_Name, t.Tag, IIf([Field_Name]=\"TSN\",[Field_Name] & \" was chang"
    "ed by \" & [First_Name] & \" \" & [Last_Name] & \" from \" & [Value_Old] & \" (\""
    " & DLookUp(\"[Latin_Name]\",\"tlu_Plants\",\"[TSN] =\" & [Value_Old]) & \")\" & "
    "\" to \" & [Value_New] & \" (\" & DLookUp(\"[Latin_Name]\",\"tlu_Plants\",\"[TSN"
    "] =\" & [Value_New]) & \")\",[Field_Name] & \" was changed by \" & [First_Name] "
    "& \" \" & [Last_Name] & \" from \" & [Value_Old] & \" to \" & [Value_New]) AS Ch"
    "ange_Desc, th.Change_Date, Year([Change_Date]) AS Change_Year, t.Azimuth, t.Dist"
    "ance, t.Microplot_Number, t.TSN, t.Tag_Notes, t.Tag_Status, t.Updated_Date, th.R"
    "ecord_ID AS Tag_ID, t.Location_ID\015\012FROM ((tbl_Tags_History AS th LEFT JOIN"
    " tbl_Tags AS t ON th.Record_ID = t.Tag_ID) LEFT JOIN tlu_Contacts AS c ON th.Con"
    "tact_ID = c.Contact_ID) LEFT JOIN qFiltered_Locations AS fl ON t.Location_ID = f"
    "l.Location_ID\015\012WHERE (((th.Change_Date)>=Nz(IIf(Abs([Forms]![frm_Data_Summ"
    "ary_Advanced]![tglFilterByRange])=1,[Forms]![frm_Data_Summary_Advanced]![tbxStar"
    "tDateFilter],#1/1/1800#)) \015\012AND (th.Change_Date)<=Nz(IIf(Abs([Forms]![frm_"
    "Data_Summary_Advanced]![tglFilterByRange])=1,[Forms]![frm_Data_Summary_Advanced]"
    "![tbxEndDateFilter],#12/31/2200#))) \015\012AND ((Year([Change_Date])) LIKE Nz(I"
    "If(Abs([Forms]![frm_Data_Summary_Advanced]![tglFilterByYear])=1,[Forms]![frm_Dat"
    "a_Summary_Advanced]![cbxYearFilter],Null),\"*\")))\015\012AND t.Tag_Status <> 'R"
    "emoved from study'\015\012ORDER BY fl.Plot_Name, t.Tag;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="Change_Desc"
        dbInteger "ColumnWidth" ="8430"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Change_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="fl.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="th.Change_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
End
