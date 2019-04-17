dbMemo "SQL" ="UPDATE tbl_Events AS e SET e.PseudoEvent = 0, e.Updated_Date = Now(), e.Updated_"
    "By = 12345, e.Event_Notes = e.Event_Notes & CHR(13) & CHR(10) & CHR(13) & CHR(10"
    ") & 'Converted ' & e.Event_Date & ' rehab (pseudoevent)'\015\012WHERE e.Event_ID"
    " = '{A67F44B9-F32E-48AF-B0E2-B46395871093}'\015\012AND e.PseudoEvent = 1;\015\012"
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
        dbText "Name" ="l.Panel"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Plot_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Frame"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Unit_Group"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="l.Subunit_Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cycle"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.Latin_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Sample_Year"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TSN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="p.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.StemsLive"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.SumLiveBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.Equiv_Live_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Habit"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Browsed"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="sd.Browsable"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.Equiv_Dead_DBH_cm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.DBH_Check"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Distance"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Early_Detect"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.CrownClass"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Rare_Spp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Updated_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Entered_On_Tablet"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Verified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Certified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Pictures_Taken"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Verified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="v.Condition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Entered_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ba.SumDeadBasalArea_cm2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Crown_Class"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Group_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Azimuth"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Crown_Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Microplot_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Protocol_Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Location_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.PseudoEvent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Event_Time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.CWD_Check_360"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Notes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.CWD_Check_120"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.CWD_Check_240"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Deer_Impact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Is_Excluded"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Plot_Maint"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Entered_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Verified"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Certified_By"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="e.Certified_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Start_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Stop_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Tag_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.RFS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="t.TaxonCode"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Data_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Event_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tag_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Wind_Lightning_Damage"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Tree_Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Vines_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Foliage_Conditions_Checked"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="td.TreeVigor"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.Tree_DBH_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.DBH"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.Live"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DBH.Updated_Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stems"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CrownClass"
        dbLong "AggregateType" ="-1"
    End
End
