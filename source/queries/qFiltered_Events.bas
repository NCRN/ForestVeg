dbMemo "SQL" ="SELECT tbl_Events.*, CStr(Year([Event_Date])) AS Event_Year\015\012FROM tbl_Even"
    "ts\015\012WHERE (((tbl_Events.Event_Date)>=Nz(IIf(Abs([Forms]![frm_Data_Summary_"
    "Advanced]![togFilterByRange])=1,[Forms]![frm_Data_Summary_Advanced]![txtStartDat"
    "eFilter],#1/1/1800#)) And (tbl_Events.Event_Date)<=Nz(IIf(Abs([Forms]![frm_Data_"
    "Summary_Advanced]![togFilterByRange])=1,[Forms]![frm_Data_Summary_Advanced]![txt"
    "EndDateFilter],#12/31/2200#))) AND ((Year([Event_Date])) Like Nz(IIf(Abs([Forms]"
    "![frm_Data_Summary_Advanced]![togFilterByYear])=1,[Forms]![frm_Data_Summary_Adva"
    "nced]![cboYearFilter],Null),\"*\")) AND ((IIf(IsNull([Certified_date])=False And"
    " ([Certified_date]>=[Updated_date] Or IsNull([Updated_date])),2,0))=Nz([Forms]!["
    "frm_Data_Summary_Advanced]![optgScope]) Or (IIf(IsNull([Certified_date])=False A"
    "nd ([Certified_date]>=[Updated_date] Or IsNull([Updated_date])),2,0))=Nz(IIf([Fo"
    "rms]![frm_Data_Summary_Advanced]![optgScope]=1,0,1)) Or (IIf(IsNull([Certified_d"
    "ate])=False And ([Certified_date]>=[Updated_date] Or IsNull([Updated_date])),2,0"
    "))=Nz(IIf([Forms]![frm_Data_Summary_Advanced]![optgScope]=1,2,1))));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbText "Description" ="Standard subquery to filter event records based on filter values in frm_Summary_"
    "Tool"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Event_year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1245"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tbl_Events.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Group_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Event_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Pictures_Taken"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Entered_On_Tablet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Entered_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Entered_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Updated_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Verified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Verified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Verified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Certified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Certified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.Certified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.CWD_Check_360"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.CWD_Check_120"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tbl_Events.CWD_Check_240"
        dbLong "AggregateType" ="-1"
    End
End
