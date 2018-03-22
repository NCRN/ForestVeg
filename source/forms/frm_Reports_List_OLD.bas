Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11160
    DatasheetFontHeight =10
    ItemSuffix =89
    Left =2460
    Top =285
    Right =13905
    Bottom =6900
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf1f79facd5fde240
    End
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Section
            Height =6600
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =9480
                    Top =5580
                    Width =1140
                    Name ="cmbo_Location"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_LU_Locations_Active.Location_ID, qry_LU_Locations_Active.Plot_Name FR"
                        "OM qry_LU_Locations_Active WHERE (((qry_LU_Locations_Active.Unit_Code) Like Form"
                        "s!frm_Reports_List!cmbo_Field_Pick_Park) And ((qry_LU_Locations_Active.Panel) Li"
                        "ke Forms!frm_Reports_list!cmbo_Field_Pick_Panel)) ORDER BY qry_LU_Locations_Acti"
                        "ve.Plot_Name; "
                    ColumnWidths ="0;1440"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =9480
                    LayoutCachedTop =5580
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =5820
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =9060
                            Top =5580
                            Width =360
                            Height =245
                            Name ="cmbo_Location_Label"
                            Caption ="Plot"
                            LayoutCachedLeft =9060
                            LayoutCachedTop =5580
                            LayoutCachedWidth =9420
                            LayoutCachedHeight =5825
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6660
                    Top =5940
                    Width =840
                    Height =299
                    TabIndex =1
                    Name ="cmd_Field_Trees"
                    Caption ="Trees"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6660
                    LayoutCachedTop =5940
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =6239
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =7560
                    Top =5940
                    Width =945
                    Height =299
                    TabIndex =2
                    Name ="cmd_Field_Saplings"
                    Caption ="Saplings"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =5940
                    LayoutCachedWidth =8505
                    LayoutCachedHeight =6239
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8580
                    Top =5940
                    Width =990
                    Height =299
                    TabIndex =3
                    Name ="cmd_Field_Pt1"
                    Caption ="Other Pt 1"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =5940
                    LayoutCachedWidth =9570
                    LayoutCachedHeight =6239
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =9660
                    Top =5940
                    Width =990
                    Height =299
                    TabIndex =4
                    Name ="cmd_Field_Other_Pt2"
                    Caption ="Other Pt 2"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9660
                    LayoutCachedTop =5940
                    LayoutCachedWidth =10650
                    LayoutCachedHeight =6239
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =6540
                    Top =5460
                    Width =4380
                    Height =1020
                    Name ="Box6"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =5460
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =6480
                End
                Begin Label
                    OverlapFlags =85
                    Left =6540
                    Top =5220
                    Width =1545
                    Height =210
                    Name ="Label7"
                    Caption ="FIELD DATA SHEETS"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =5220
                    LayoutCachedWidth =8085
                    LayoutCachedHeight =5430
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =3
                    ListRows =24
                    ListWidth =2664
                    Left =2700
                    Top =1260
                    TabIndex =5
                    Name ="cmbo_Event_Selection"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_LU_Events.Event_ID, qry_LU_Events.Plot_Name, qry_LU_Events.Start_Date"
                        " FROM qry_LU_Events WHERE (((qry_LU_Events.Unit_Code) Like Forms!frm_Reports_Lis"
                        "t!cmbo_Summary_Pick_Park) And ((qry_LU_Events.Panel) Like Forms!frm_Reports_List"
                        "!cmbo_Summary_Pick_Panel)) ORDER BY qry_LU_Events.Plot_Name; "
                    ColumnWidths ="0;1224;720"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =2160
                            Top =1260
                            Width =480
                            Height =240
                            Name ="Label9"
                            Caption ="Event"
                            LayoutCachedLeft =2160
                            LayoutCachedTop =1260
                            LayoutCachedWidth =2640
                            LayoutCachedHeight =1500
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =720
                    Width =6180
                    Height =5760
                    Name ="Box10"
                    LayoutCachedLeft =120
                    LayoutCachedTop =720
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =6480
                End
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =480
                    Width =1545
                    Height =210
                    Name ="Label11"
                    Caption ="SUMMARY REPORTS"
                    LayoutCachedLeft =120
                    LayoutCachedTop =480
                    LayoutCachedWidth =1665
                    LayoutCachedHeight =690
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =240
                    Top =1260
                    Width =1800
                    Height =300
                    TabIndex =6
                    Name ="cmd_Event_Summary"
                    Caption ="Event Summary"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =240
                    LayoutCachedTop =1260
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =1560
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6600
                    Top =780
                    Width =2100
                    Height =299
                    TabIndex =7
                    Name ="cmd_Sampling_Cycle"
                    Caption ="Sampling Panels"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =780
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1079
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6600
                    Top =3540
                    Width =2100
                    Height =300
                    TabIndex =8
                    Name ="cmd_Export_Trees"
                    Caption ="Export Trees to GIS"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =3540
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =3840
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =6540
                    Top =3480
                    Width =4380
                    Height =1680
                    Name ="Box17"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =3480
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =5160
                End
                Begin Label
                    OverlapFlags =85
                    Left =6540
                    Top =3240
                    Width =1545
                    Height =210
                    Name ="Label18"
                    Caption ="MACROS"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =3240
                    LayoutCachedWidth =8085
                    LayoutCachedHeight =3450
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    Left =60
                    Width =11040
                    Height =420
                    FontSize =12
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="Label19"
                    Caption ="Reports"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =10140
                    Top =60
                    Width =780
                    Height =300
                    TabIndex =9
                    Name ="cmd_close_form"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =60
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =360
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8760
                    Top =780
                    Width =2100
                    Height =299
                    TabIndex =10
                    Name ="cmd_Events"
                    Caption ="Events by Year"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =780
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1079
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8760
                    Top =1140
                    Width =2100
                    Height =299
                    TabIndex =11
                    Name ="cmd_Plot_Setup"
                    Caption ="Plot Setup  by Panel"
                    OnEnter ="[Event Procedure]"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =1140
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1439
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6600
                    Top =1140
                    Width =2100
                    Height =299
                    TabIndex =12
                    Name ="cmd_Review"
                    Caption ="Events to Review"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =1140
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1439
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6600
                    Top =3900
                    Width =2100
                    Height =300
                    TabIndex =13
                    Name ="cmd_Export_Plots"
                    Caption ="Export Plots to GIS"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =3900
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =4200
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8760
                    Top =3540
                    Width =2100
                    Height =300
                    TabIndex =14
                    Name ="cmd_Export_Trees_1yr"
                    Caption ="Export Trees (1 yr)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =3540
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =3840
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8760
                    Top =3900
                    Width =2100
                    Height =300
                    TabIndex =15
                    Name ="cmd_Export_Plots_1yr"
                    Caption ="Export Plots (1 yr)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =3900
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =4200
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =864
                    Left =8400
                    Top =5580
                    Width =540
                    TabIndex =16
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbo_Field_Pick_Panel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Panel FROM tbl_Locations GROUP BY tbl_Locations.Panel HAVIN"
                        "G (((tbl_Locations.Panel) Is Not Null))  UNION SELECT \"*\" as Panel From tbl_Lo"
                        "cations;"
                    ColumnWidths ="864"
                    DefaultValue ="=[Forms]![frm_Switchboard]![cPanel]"

                    LayoutCachedLeft =8400
                    LayoutCachedTop =5580
                    LayoutCachedWidth =8940
                    LayoutCachedHeight =5820
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =7860
                            Top =5580
                            Width =555
                            Height =240
                            ForeColor =3355443
                            Name ="Label69"
                            Caption =" Panel"
                            ControlTipText ="0=Undetermined, 1=2006, 2=2007, 3=2008, 4=2009"
                            LayoutCachedLeft =7860
                            LayoutCachedTop =5580
                            LayoutCachedWidth =8415
                            LayoutCachedHeight =5820
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =864
                    Left =7080
                    Top =5580
                    Width =780
                    TabIndex =17
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbo_Field_Pick_Park"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Unit_Code FROM tbl_Locations GROUP BY tbl_Locations.Unit_Co"
                        "de UNION SELECT \"*\" as Unit_Code From tbl_Locations;"
                    ColumnWidths ="864"
                    DefaultValue ="\"*\""

                    LayoutCachedLeft =7080
                    LayoutCachedTop =5580
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =5820
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =6660
                            Top =5580
                            Width =435
                            Height =240
                            ForeColor =3355443
                            Name ="Label71"
                            Caption ="Park"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =5580
                            LayoutCachedWidth =7095
                            LayoutCachedHeight =5820
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =864
                    Left =4380
                    Top =840
                    Width =540
                    TabIndex =18
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbo_Summary_Pick_Panel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Panel FROM tbl_Locations GROUP BY tbl_Locations.Panel HAVIN"
                        "G (((tbl_Locations.Panel) Is Not Null))  UNION SELECT \"*\" as Panel From tbl_Lo"
                        "cations;"
                    ColumnWidths ="864"
                    DefaultValue ="=[Forms]![frm_Switchboard]![cPanel]"

                    LayoutCachedLeft =4380
                    LayoutCachedTop =840
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =3840
                            Top =840
                            Width =480
                            Height =240
                            ForeColor =10040879
                            Name ="Label28"
                            Caption =" Panel"
                            ControlTipText ="0=Undetermined, 1=2006, 2=2007, 3=2008, 4=2009"
                            LayoutCachedLeft =3840
                            LayoutCachedTop =840
                            LayoutCachedWidth =4320
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =864
                    Left =1680
                    Top =840
                    Width =900
                    TabIndex =19
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbo_Summary_Pick_Park"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Unit_Code FROM tbl_Locations GROUP BY tbl_Locations.Unit_Co"
                        "de UNION SELECT \"*\" as Unit_Code From tbl_Locations;"
                    ColumnWidths ="864"
                    DefaultValue ="\"*\""

                    LayoutCachedLeft =1680
                    LayoutCachedTop =840
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =1200
                            Top =840
                            Width =420
                            Height =240
                            ForeColor =10040879
                            Name ="Label30"
                            Caption ="Park"
                            LayoutCachedLeft =1200
                            LayoutCachedTop =840
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6600
                    Top =1500
                    Width =2100
                    Height =299
                    TabIndex =20
                    Name ="cmd_Audit_by_Panel"
                    Caption ="Audit by Panel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =1500
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =1799
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2355
                    Top =2460
                    Width =1800
                    Height =300
                    TabIndex =21
                    Name ="cmd_rpt_Trees_by_Plot"
                    Caption ="Trees"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2355
                    LayoutCachedTop =2460
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =2760
                End
                Begin Line
                    OverlapFlags =119
                    Left =240
                    Top =1140
                    Width =5940
                    Name ="Line33"
                    LeftPadding =30
                    TopPadding =30
                    RightPadding =30
                    BottomPadding =30
                    GridlineStyleLeft =0
                    GridlineStyleTop =0
                    GridlineStyleRight =0
                    GridlineStyleBottom =0
                    GridlineWidthLeft =1
                    GridlineWidthTop =1
                    GridlineWidthRight =1
                    GridlineWidthBottom =1
                    LayoutCachedLeft =240
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1140
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =180
                    Top =1920
                    Width =1875
                    Height =240
                    Name ="Label34"
                    Caption ="----By Park (1 Year)----"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1920
                    LayoutCachedWidth =2055
                    LayoutCachedHeight =2160
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =2340
                    Top =1920
                    Width =1800
                    Height =240
                    Name ="Label37"
                    Caption ="-------By Plot-------"
                    LayoutCachedLeft =2340
                    LayoutCachedTop =1920
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =2160
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =225
                    Top =2460
                    Width =1800
                    Height =300
                    TabIndex =22
                    Name ="cmd_rpt_Trees_by_Park"
                    Caption ="Trees"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =225
                    LayoutCachedTop =2460
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =2760
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4395
                    Top =2460
                    Width =1800
                    Height =300
                    TabIndex =23
                    Name ="cmd_rpt_Trees_by_Species"
                    Caption ="Trees"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4395
                    LayoutCachedTop =2460
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =2760
                End
                Begin Label
                    OverlapFlags =247
                    Left =240
                    Top =840
                    Width =960
                    Height =239
                    FontWeight =700
                    ForeColor =10040879
                    Name ="Label40"
                    Caption ="Filter By ->"
                    LayoutCachedLeft =240
                    LayoutCachedTop =840
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =1079
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =4380
                    Top =1920
                    Width =1800
                    Height =240
                    Name ="Label41"
                    Caption ="-----By Species-----"
                    LayoutCachedLeft =4380
                    LayoutCachedTop =1920
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =2160
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2355
                    Top =2820
                    Width =1800
                    Height =300
                    TabIndex =24
                    Name ="cmd_rpt_Shrubs_by_Plot"
                    Caption ="Shrubs"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2355
                    LayoutCachedTop =2820
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =3120
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =6600
                    Top =2820
                    Width =2100
                    Height =299
                    TabIndex =25
                    Name ="cmd_Species_Summary"
                    Caption ="Species Summary (Draft)"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6600
                    LayoutCachedTop =2820
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =3119
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2355
                    Top =3180
                    Width =1800
                    Height =300
                    TabIndex =26
                    Name ="cmd_rpt_Vines_by_Plot"
                    Caption ="Vines"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2355
                    LayoutCachedTop =3180
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =3480
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2355
                    Top =3540
                    Width =1800
                    Height =300
                    TabIndex =27
                    Name ="cmd_rpt_Herbs_by_Plot"
                    Caption ="Exotic Herbs"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2355
                    LayoutCachedTop =3540
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =3840
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4395
                    Top =2820
                    Width =1800
                    Height =300
                    TabIndex =28
                    Name ="cmd_rpt_Shrubs_by_Species"
                    Caption ="Shrubs"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4395
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =3120
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4395
                    Top =3180
                    Width =1800
                    Height =300
                    TabIndex =29
                    Name ="cmd_rpt_Vines_by_Species"
                    Caption ="Vines"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4395
                    LayoutCachedTop =3180
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =3480
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4395
                    Top =3540
                    Width =1800
                    Height =300
                    TabIndex =30
                    Name ="cmd_rpt_Herbs_by_Species"
                    Caption ="Exotic Herbs"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4395
                    LayoutCachedTop =3540
                    LayoutCachedWidth =6195
                    LayoutCachedHeight =3840
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =225
                    Top =3540
                    Width =1800
                    Height =300
                    TabIndex =31
                    Name ="cmd_rpt_Shrubs_by_Park"
                    Caption ="Shrubs"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =225
                    LayoutCachedTop =3540
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =3840
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =240
                    Top =4260
                    Width =1800
                    Height =300
                    TabIndex =32
                    Name ="cmd_rpt_Vines_by_Park"
                    Caption ="Vines"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =4260
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =4560
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =240
                    Top =4620
                    Width =1800
                    Height =300
                    TabIndex =33
                    Name ="cmd_rpt_Herbs_by_Park"
                    Caption ="Exotic Herbs"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =4620
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =4920
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =247
                    Left =240
                    Top =4980
                    Width =1800
                    Height =300
                    TabIndex =34
                    Name ="cmd_rpt_Woody_by_Park"
                    Caption ="Woody Debris"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =4980
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =5280
                End
                Begin ComboBox
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =864
                    Left =5460
                    Top =840
                    Width =720
                    TabIndex =35
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbo_Summary_Pick_Year"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tsys_App_Defaults.Sample_Year FROM tsys_App_Defaults GROUP BY tsys_App_De"
                        "faults.Sample_Year UNION SELECT \"*\" as Sample_Year From tsys_App_Defaults;"
                    ColumnWidths ="1440"
                    DefaultValue ="\"*\""

                    LayoutCachedLeft =5460
                    LayoutCachedTop =840
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =4980
                            Top =840
                            Width =420
                            Height =240
                            ForeColor =10040879
                            Name ="Label54"
                            Caption ="Year"
                            ControlTipText ="0=Undetermined, 1=2006, 2=2007, 3=2008, 4=2009"
                            LayoutCachedLeft =4980
                            LayoutCachedTop =840
                            LayoutCachedWidth =5400
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin Label
                    OverlapFlags =255
                    Left =945
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label55"
                    Caption ="C"
                    LayoutCachedLeft =945
                    LayoutCachedTop =2160
                    LayoutCachedWidth =1110
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =1125
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label56"
                    Caption ="P"
                    LayoutCachedLeft =1125
                    LayoutCachedTop =2160
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =1305
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =1643706
                    Name ="Label57"
                    Caption ="Y"
                    LayoutCachedLeft =1305
                    LayoutCachedTop =2160
                    LayoutCachedWidth =1470
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =255
                    Left =3105
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label58"
                    Caption ="C"
                    LayoutCachedLeft =3105
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3270
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =3285
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label59"
                    Caption ="P"
                    LayoutCachedLeft =3285
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3450
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =3465
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =5026082
                    Name ="Label60"
                    Caption ="Y"
                    LayoutCachedLeft =3465
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3630
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =4920
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =5026082
                    Name ="Label61"
                    Caption ="P"
                    LayoutCachedLeft =4920
                    LayoutCachedTop =2160
                    LayoutCachedWidth =5085
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =5280
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label62"
                    Caption ="P"
                    LayoutCachedLeft =5280
                    LayoutCachedTop =2160
                    LayoutCachedWidth =5445
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =5460
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =5026082
                    Name ="Label63"
                    Caption ="Y"
                    LayoutCachedLeft =5460
                    LayoutCachedTop =2160
                    LayoutCachedWidth =5625
                    LayoutCachedHeight =2370
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =225
                    Top =2820
                    Width =1800
                    Height =300
                    TabIndex =36
                    Name ="cmd_rpt_Saplings_by_Park"
                    Caption ="Saplings"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =225
                    LayoutCachedTop =2820
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =3120
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =225
                    Top =3180
                    Width =1800
                    Height =300
                    TabIndex =37
                    Name ="cmd_rpt_Seedlings_by_Park"
                    Caption ="Seedlings"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =225
                    LayoutCachedTop =3180
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =3480
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =6540
                    Top =720
                    Width =4380
                    Height =2460
                    Name ="Box66"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =720
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =3180
                End
                Begin Label
                    OverlapFlags =85
                    Left =6540
                    Top =480
                    Width =1545
                    Height =210
                    Name ="Label67"
                    Caption ="OTHER"
                    LayoutCachedLeft =6540
                    LayoutCachedTop =480
                    LayoutCachedWidth =8085
                    LayoutCachedHeight =690
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =2355
                    Top =3900
                    Width =1800
                    Height =300
                    TabIndex =38
                    Name ="cmd_rpt_Soils_by_Park"
                    Caption ="Soils"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2355
                    LayoutCachedTop =3900
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =4200
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =240
                    Top =6060
                    Width =825
                    Height =270
                    ForeColor =1643706
                    Name ="Label70"
                    Caption ="Required"
                    LayoutCachedLeft =240
                    LayoutCachedTop =6060
                    LayoutCachedWidth =1065
                    LayoutCachedHeight =6330
                End
                Begin Label
                    OverlapFlags =255
                    TextAlign =2
                    Left =1140
                    Top =6060
                    Width =675
                    Height =270
                    ForeColor =5026082
                    Name ="Label72"
                    Caption ="Optional"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =6060
                    LayoutCachedWidth =1815
                    LayoutCachedHeight =6330
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =180
                    Top =6000
                    Width =1740
                    Height =420
                    Name ="Box73"
                    LayoutCachedLeft =180
                    LayoutCachedTop =6000
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =6420
                End
                Begin Label
                    OverlapFlags =247
                    Left =300
                    Top =5760
                    Width =600
                    Height =210
                    Name ="Label74"
                    Caption ="Legend"
                    LayoutCachedLeft =300
                    LayoutCachedTop =5760
                    LayoutCachedWidth =900
                    LayoutCachedHeight =5970
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =240
                    Top =3900
                    Width =1800
                    Height =300
                    TabIndex =39
                    Name ="cmd_rpt_Shrub_Seedlings_by_Park"
                    Caption ="Shrub Seedlings"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =3900
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =4200
                End
                Begin Label
                    OverlapFlags =247
                    Left =780
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label76"
                    Caption ="P"
                    LayoutCachedLeft =780
                    LayoutCachedTop =2160
                    LayoutCachedWidth =945
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =2940
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =5026082
                    Name ="Label77"
                    Caption ="P"
                    LayoutCachedLeft =2940
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3105
                    LayoutCachedHeight =2370
                End
                Begin Label
                    OverlapFlags =247
                    Left =5100
                    Top =2160
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label78"
                    Caption ="C"
                    LayoutCachedLeft =5100
                    LayoutCachedTop =2160
                    LayoutCachedWidth =5265
                    LayoutCachedHeight =2370
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =864
                    Left =3240
                    Top =840
                    Width =540
                    TabIndex =40
                    Name ="cmbo_Summary_Pick_Cycle"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="864"
                    DefaultValue ="\"1\""

                    LayoutCachedLeft =3240
                    LayoutCachedTop =840
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =1080
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =2700
                            Top =840
                            Width =480
                            Height =240
                            ForeColor =10040879
                            Name ="Label80"
                            Caption ="Cycle"
                            ControlTipText ="0=Undetermined, 1=2006, 2=2007, 3=2008, 4=2009"
                            LayoutCachedLeft =2700
                            LayoutCachedTop =840
                            LayoutCachedWidth =3180
                            LayoutCachedHeight =1080
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8760
                    Top =1500
                    Width =2100
                    Height =299
                    TabIndex =41
                    Name ="cmd_Tree_Summary_by_Panel"
                    Caption ="Tree Summary by Panel"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8760
                    LayoutCachedTop =1500
                    LayoutCachedWidth =10860
                    LayoutCachedHeight =1799
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4410
                    Top =4980
                    Width =1800
                    Height =300
                    TabIndex =42
                    Name ="cmd_rpt_IVI_Trees"
                    Caption ="Trees"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4410
                    LayoutCachedTop =4980
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =5280
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =4365
                    Top =4440
                    Width =1845
                    Height =240
                    Name ="Label83"
                    Caption ="--Importance Values--"
                    LayoutCachedLeft =4365
                    LayoutCachedTop =4440
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =4680
                End
                Begin Label
                    OverlapFlags =247
                    Left =4935
                    Top =4680
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label84"
                    Caption ="P"
                    LayoutCachedLeft =4935
                    LayoutCachedTop =4680
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =4890
                End
                Begin Label
                    OverlapFlags =247
                    Left =5295
                    Top =4680
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label85"
                    Caption ="P"
                    LayoutCachedLeft =5295
                    LayoutCachedTop =4680
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =4890
                End
                Begin Label
                    OverlapFlags =247
                    Left =5475
                    Top =4680
                    Width =165
                    Height =210
                    ForeColor =5026082
                    Name ="Label86"
                    Caption ="Y"
                    LayoutCachedLeft =5475
                    LayoutCachedTop =4680
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =4890
                End
                Begin Label
                    OverlapFlags =247
                    Left =5115
                    Top =4680
                    Width =165
                    Height =210
                    ForeColor =9211020
                    Name ="Label87"
                    Caption ="C"
                    LayoutCachedLeft =5115
                    LayoutCachedTop =4680
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =4890
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4410
                    Top =5340
                    Width =1800
                    Height =300
                    TabIndex =43
                    Name ="cmd_Rpt_IVI_Saplings"
                    Caption ="Saplings"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4410
                    LayoutCachedTop =5340
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =5640
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbo_Event_Selection_GotFocus()
    Me!cmbo_Event_Selection.Requery
End Sub

Private Sub cmbo_Location_GotFocus()
    Me!cmbo_Location.Requery
End Sub

Private Sub cmd_Audit_by_Panel_Click()
On Error GoTo Err_cmd_Audit_by_Panel_Click

    Dim stDocName As String

    stDocName = "rpt_Audit_by_Panel"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Audit_by_Panel_Click:
    Exit Sub
Err_cmd_Audit_by_Panel_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Audit_by_Panel_Click
End Sub

Private Sub cmd_Events_Click()
On Error GoTo Err_cmd_Event_Summary_Click

    Dim stDocName As String

    stDocName = "rpt_Yearly_Events"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Event_Summary_Click:
    Exit Sub
Err_cmd_Event_Summary_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Event_Summary_Click
End Sub

Private Sub cmd_Field_Trees_Click()
On Error GoTo Err_cmd_Field_Trees_Click

    Dim stDocName As String

    stDocName = "rpt_Field_Sheet_Trees"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Field_Trees_Click:
    Exit Sub
Err_cmd_Field_Trees_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Field_Trees_Click
End Sub
Private Sub cmd_Field_Saplings_Click()
On Error GoTo Err_cmd_Field_Saplings_Click

    Dim stDocName As String

    stDocName = "rpt_Field_Sheets_Saplings"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Field_Saplings_Click:
    Exit Sub
Err_cmd_Field_Saplings_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Field_Saplings_Click
End Sub
Private Sub cmd_Field_Pt1_Click()
On Error GoTo Err_cmd_Field_Pt1_Click

    Dim stDocName As String

    stDocName = "rpt_Field_Sheet_Part1"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Field_Pt1_Click:
    Exit Sub
Err_cmd_Field_Pt1_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Field_Pt1_Click
End Sub
Private Sub cmd_Field_Other_Pt2_Click()
On Error GoTo Err_cmd_Field_Other_Pt2_Click

    Dim stDocName As String

    stDocName = "rpt_Field_Sheet_Part2"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Field_Other_Pt2_Click:
    Exit Sub
Err_cmd_Field_Other_Pt2_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Field_Other_Pt2_Click
End Sub
Private Sub cmd_Event_Summary_Click()
On Error GoTo Err_cmd_Event_Summary_Click

    Dim stDocName As String

    stDocName = "rpt_Event_Summary"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Event_Summary_Click:
    Exit Sub
Err_cmd_Event_Summary_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Event_Summary_Click
End Sub

Private Sub cmd_Plot_Setup_Enter()
On Error GoTo Err_cmd_Event_Summary_Click

    Dim stDocName As String

    stDocName = "rpt_Yearly_Plot_Setup"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Event_Summary_Click:
    Exit Sub
Err_cmd_Event_Summary_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Event_Summary_Click
End Sub

Private Sub cmd_Review_Click()
On Error GoTo Err_cmd_Sampling_Cycle_Click

    Dim stDocName As String

    stDocName = "rpt_Review_Issues"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Sampling_Cycle_Click:
    Exit Sub
Err_cmd_Sampling_Cycle_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Sampling_Cycle_Click
End Sub

Private Sub cmd_Sampling_Cycle_Click()
On Error GoTo Err_cmd_Sampling_Cycle_Click

    Dim stDocName As String

    stDocName = "rpt_Sampling_Panel"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Sampling_Cycle_Click:
    Exit Sub
Err_cmd_Sampling_Cycle_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Sampling_Cycle_Click
End Sub

Private Sub cmd_Tree_Summary_by_Panel_Click()
On Error GoTo Err_cmd_Tree_Summary_by_Panel_Click

    Dim stDocName As String

    stDocName = "rpt_Tree_Summary_by_Panel"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Tree_Summary_by_Panel_Click:
    Exit Sub
Err_cmd_Tree_Summary_by_Panel_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Tree_Summary_by_Panel_Click
End Sub

Private Sub cmd_Export_Trees_1yr_Click()
On Error GoTo Err_cmd_Export_Trees_1yr_Click

    Dim stDocName As String

    stDocName = "macro_Export_Trees_To_GIS_1yr"
    DoCmd.RunMacro stDocName
Exit_cmd_Export_Trees_1yr_Click:
    Exit Sub
Err_cmd_Export_Trees_1yr_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Export_Trees_1yr_Click
End Sub

Private Sub cmd_Export_Trees_Click()
On Error GoTo Err_cmd_Export_Trees_Click

    Dim stDocName As String

    stDocName = "macro_Export_Trees_To_GIS"
    DoCmd.RunMacro stDocName
Exit_cmd_Export_Trees_Click:
    Exit Sub
Err_cmd_Export_Trees_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Export_Trees_Click
End Sub


Private Sub cmd_Export_Plots_1yr_Click()
On Error GoTo Err_cmd_Export_Plots_1yr_Click

    Dim stDocName As String

    stDocName = "macro_Export_Plots_To_GIS_1yr"
    DoCmd.RunMacro stDocName
Exit_cmd_Export_Plots_1yr_Click:
    Exit Sub
Err_cmd_Export_Plots_1yr_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Export_Plots_1yr_Click
End Sub

Private Sub cmd_Export_Plots_Click()
On Error GoTo Err_cmd_Export_Plots_Click

    Dim stDocName As String

    stDocName = "macro_Export_Plots_To_GIS"
    DoCmd.RunMacro stDocName
Exit_cmd_Export_Plots_Click:
    Exit Sub
Err_cmd_Export_Plots_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Export_Plots_Click
End Sub

Private Sub cmd_Close_Form_Click()
On Error GoTo Err_cmd_close_form_Click

    DoCmd.Close
Exit_cmd_close_form_Click:
    Exit Sub
Err_cmd_close_form_Click:
    MsgBox Err.Description
    Resume Exit_cmd_close_form_Click
End Sub

Private Sub cmd_rpt_Shrubs_by_Plot_Click()
On Error GoTo Err_cmd_rpt_Shrubs_by_Plot_Click
 
    stDocName = "Rpt_Sum_Shrubs_by_Plot"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Shrubs_by_Plot_Click:
    Exit Sub
Err_cmd_rpt_Shrubs_by_Plot_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Shrubs_by_Plot_Click
End Sub

Private Sub cmd_Species_Summary_Click()
On Error GoTo Err_cmd_Species_Summary_Click
    
    stDocName = "rpt_Event_Summary_Species"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Species_Summary_Click:
    Exit Sub
Err_cmd_Species_Summary_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Species_Summary_Click
End Sub

Private Sub cmd_rpt_Trees_by_Plot_Click()
On Error GoTo Err_cmd_rpt_Trees_by_Plot_Click

    stDocName = "Rpt_Sum_Trees_by_Plot"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Trees_by_Plot_Click:
    Exit Sub
Err_cmd_rpt_Trees_by_Plot_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Trees_by_Plot_Click
End Sub

Private Sub cmd_rpt_Trees_by_Species_Click()
On Error GoTo Err_cmd_rpt_Trees_by_Species_Click

'If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Park] = "*" Then
'    MsgBox ("Please select a park from the dropdown list above")
'    GoTo Exit_cmd_rpt_Trees_by_Species_Click
'End If

    stDocName = "Rpt_Sum_Trees_by_Species"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Trees_by_Species_Click:
    Exit Sub
Err_cmd_rpt_Trees_by_Species_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Trees_by_Species_Click
End Sub

Private Sub cmd_rpt_Shrubs_by_Species_Click()
On Error GoTo Err_cmd_rpt_Shrubs_by_Species_Click

    stDocName = "Rpt_Sum_Shrubs_by_Species"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Shrubs_by_Species_Click:
    Exit Sub
Err_cmd_rpt_Shrubs_by_Species_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Shrubs_by_Species_Click
End Sub

Private Sub cmd_rpt_Vines_by_Species_Click()
On Error GoTo Err_cmd_rpt_Vines_by_Species_Click

    stDocName = "Rpt_Sum_Vines_by_Species"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Vines_by_Species_Click:
    Exit Sub
Err_cmd_rpt_Vines_by_Species_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Vines_by_Species_Click
End Sub

Private Sub cmd_rpt_Herbs_by_Species_Click()
On Error GoTo Err_cmd_rpt_Herbs_by_Species_Click

    stDocName = "Rpt_Sum_Exotic_Herbs_by_Species"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Herbs_by_Species_Click:
    Exit Sub
Err_cmd_rpt_Herbs_by_Species_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Herbs_by_Species_Click
End Sub

Private Sub cmd_rpt_Vines_by_Plot_Click()
On Error GoTo Err_cmd_rpt_Vines_by_Plot_Click

    stDocName = "Rpt_Sum_Vines_by_Plot"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Vines_by_Plot_Click:
    Exit Sub
Err_cmd_rpt_Vines_by_Plot_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Vines_by_Plot_Click
End Sub

Private Sub cmd_rpt_Herbs_by_Plot_Click()
On Error GoTo Err_cmd_rpt_Herbs_by_Plot_Click

    stDocName = "Rpt_Sum_Exotic_Herbs_by_Plot"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Herbs_by_Plot_Click:
    Exit Sub
Err_cmd_rpt_Herbs_by_Plot_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Herbs_by_Plot_Click
End Sub

Private Sub cmd_rpt_Trees_by_Park_Click()
On Error GoTo Err_cmd_rpt_Trees_by_Park_Click

If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
    MsgBox ("Please select a sampling year from the dropdown list above")
    GoTo Exit_cmd_rpt_Trees_by_Park_Click
End If

    stDocName = "Rpt_Sum_Trees_by_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Trees_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Trees_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Trees_by_Park_Click
End Sub

Private Sub cmd_rpt_Saplings_by_Park_Click()
On Error GoTo Err_cmd_rpt_Saplings_by_Park_Click

If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
    MsgBox ("Please select a sampling year from the dropdown list above")
    GoTo Exit_cmd_rpt_Saplings_by_Park_Click
End If

    stDocName = "Rpt_Sum_Saplings_by_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Saplings_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Saplings_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Saplings_by_Park_Click
End Sub

Private Sub cmd_rpt_Seedlings_by_Park_Click()
On Error GoTo Err_cmd_rpt_Seedlings_by_Park_Click

If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
    MsgBox ("Please select a sampling year from the dropdown list above")
    GoTo Exit_cmd_rpt_Seedlings_by_Park_Click
End If

    stDocName = "Rpt_Sum_Seedlings_by_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Seedlings_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Seedlings_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Seedlings_by_Park_Click
End Sub

Private Sub cmd_rpt_Soils_by_Park_Click()
On Error GoTo Err_cmd_rpt_Soils_by_Park_Click

'If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
'    MsgBox ("Please select a sampling year from the dropdown list above")
'    GoTo Exit_cmd_rpt_Soils_by_Park_Click
'End If

    stDocName = "Rpt_Sum_Soil_by_Plot_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Soils_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Soils_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Soils_by_Park_Click
End Sub

Private Sub cmd_rpt_Shrubs_by_Park_Click()
On Error GoTo Err_cmd_rpt_Shrubs_by_Park_Click

If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
    MsgBox ("Please select a sampling year from the dropdown list above")
    GoTo Exit_cmd_rpt_Shrubs_by_Park_Click
End If

    stDocName = "Rpt_Sum_Shrubs_by_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Shrubs_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Shrubs_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Shrubs_by_Park_Click
End Sub
Private Sub cmd_rpt_Shrub_Seedlings_by_Park_Click()
On Error GoTo Err_cmd_rpt_Shrub_Seedlings_by_Park_Click

If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
    MsgBox ("Please select a sampling year from the dropdown list above")
    GoTo Exit_cmd_rpt_Shrub_Seedlings_by_Park_Click
End If

    stDocName = "Rpt_Sum_Shrub_Seedlings_by_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Shrub_Seedlings_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Shrub_Seedlings_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Shrub_Seedlings_by_Park_Click
End Sub

Private Sub cmd_rpt_Vines_by_Park_Click()
On Error GoTo Err_cmd_rpt_Vines_by_Park_Click

If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
    MsgBox ("Please select a sampling year from the dropdown list above")
    GoTo Exit_cmd_rpt_Vines_by_Park_Click
End If

    stDocName = "Rpt_Sum_Vines_by_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Vines_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Vines_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Vines_by_Park_Click
End Sub

Private Sub cmd_rpt_Herbs_by_Park_Click()
On Error GoTo Err_cmd_rpt_Herbs_by_Park_Click

If [Forms]![frm_Reports_List]![cmbo_Summary_Pick_Year] = "*" Then
    MsgBox ("Please select a sampling year from the dropdown list above")
    GoTo Exit_cmd_rpt_Herbs_by_Park_Click
End If

    stDocName = "Rpt_Sum_Exotic_Herbs_by_AdminPark_SY"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_rpt_Herbs_by_Park_Click:
    Exit Sub
Err_cmd_rpt_Herbs_by_Park_Click:
    MsgBox Err.Description
    Resume Exit_cmd_rpt_Herbs_by_Park_Click
End Sub

Private Sub cmd_Rpt_IVI_Trees_Click()
On Error GoTo Err_cmd_Rpt_IVI_Trees_Click

    stDocName = "rpt_Importance_Value_Trees"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Rpt_IVI_Trees_Click:
    Exit Sub
Err_cmd_Rpt_IVI_Trees_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Rpt_IVI_Trees_Click
End Sub

Private Sub cmd_Rpt_IVI_Saplings_Click()
On Error GoTo Err_cmd_Rpt_IVI_Saplings_Click

    stDocName = "rpt_Importance_Value_Saplings"
    DoCmd.OpenReport stDocName, acPreview
Exit_cmd_Rpt_IVI_Saplings_Click:
    Exit Sub
Err_cmd_Rpt_IVI_Saplings_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Rpt_IVI_Saplings_Click
End Sub
