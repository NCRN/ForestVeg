Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =10
    ItemSuffix =56
    Left =4680
    Top =990
    Right =14760
    Bottom =11850
    DatasheetGridlinesColor =12632256
    OrderBy ="Append_Order"
    RecSrcDt = Begin
        0x117d3d3a0f5ae540
    End
    RecordSource ="tsys_Append_Tables"
    Caption ="Append Data"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =4515
            BackColor =5394044
            Name ="FormHeader"
            Begin
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =4800
                    Top =3960
                    Width =4815
                    Height =480
                    LeftMargin =36
                    TopMargin =36
                    RightMargin =36
                    ForeColor =16711680
                    Name ="lblPseudoEventsDeleted"
                    Caption ="** PseudoEvents WILL be DELETED from import tables && won't be included in data "
                        "appends/updates **"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =3960
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =4440
                End
                Begin Label
                    OverlapFlags =93
                    Left =60
                    Width =2520
                    Height =480
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Append Data"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =480
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8040
                    Top =120
                    Width =1320
                    Height =600
                    FontWeight =700
                    ForeColor =0
                    Name ="cmd_AppendLog"
                    Caption ="View Append Log"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8040
                    LayoutCachedTop =120
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =720
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =2880
                    Left =5715
                    Top =3300
                    Width =3360
                    ColumnOrder =5
                    TabIndex =3
                    BackColor =13434879
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"0\""
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2200220000000000
                    End
                    Name ="cmbo_Select_Event"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Events.Event_ID, [Plot_Name] & \" \" & \" \" & [Event_Date] AS PickSt"
                        "ring FROM tbl_Locations  INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tb"
                        "l_Events.Location_ID WHERE (((Year([Event_Date]))=Year(Now()))) ORDER BY tbl_Eve"
                        "nts.Event_Date DESC"
                    ColumnWidths ="0;2880"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"
                    LayoutCachedLeft =5715
                    LayoutCachedTop =3300
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =3540
                    ConditionalFormat14 = Begin
                        0x01000100000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =660
                            Top =3300
                            Width =4980
                            Height =240
                            FontWeight =700
                            ForeColor =9868950
                            Name ="lblMasterEventAppend"
                            Caption ="Select the Event to Append  to in Master Database -->"
                            ControlTipText ="Select the Event from the main data set that you wish to append the secondary ta"
                                "blet  data to"
                            LayoutCachedLeft =660
                            LayoutCachedTop =3300
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =3540
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5715
                    Top =2700
                    Width =3360
                    ColumnOrder =3
                    TabIndex =1
                    BackColor =13434879
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2200220000000000
                    End
                    Name ="cmbo_Select_Import_Event_Table"
                    RowSourceType ="Value List"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =5715
                    LayoutCachedTop =2700
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =2940
                    ConditionalFormat14 = Begin
                        0x01000100000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =660
                            Top =2700
                            Width =4980
                            Height =240
                            FontWeight =700
                            ForeColor =9868950
                            Name ="lblEventsSecondaryImport"
                            Caption ="Select Events Table to import from Secondary Tablet -->"
                            LayoutCachedLeft =660
                            LayoutCachedTop =2700
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =2940
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5715
                    Top =3000
                    Width =3360
                    ColumnOrder =4
                    TabIndex =2
                    BackColor =13434879
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000003000000000000000300000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2200220000000000
                    End
                    Name ="cmbo_Select_Import_Events"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;0;1440"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =5715
                    LayoutCachedTop =3000
                    LayoutCachedWidth =9075
                    LayoutCachedHeight =3240
                    ConditionalFormat14 = Begin
                        0x01000100000000000000030000000100000000000000ffffff00020000002200 ,
                        0x2200000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =660
                            Top =3000
                            Width =4980
                            Height =240
                            FontWeight =700
                            ForeColor =9868950
                            Name ="lblEventSecondary"
                            Caption ="Select the Event to Append from Secondary Tablet -->"
                            LayoutCachedLeft =660
                            LayoutCachedTop =3000
                            LayoutCachedWidth =5640
                            LayoutCachedHeight =3240
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =8040
                    Top =780
                    Width =1320
                    Height =600
                    FontWeight =700
                    TabIndex =4
                    ForeColor =0
                    Name ="cmd_ViewUpdateLog"
                    Caption ="View Update Log"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8040
                    LayoutCachedTop =780
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1380
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =540
                    Top =600
                    Width =6660
                    Height =1140
                    ColumnOrder =6
                    TabIndex =5
                    Name ="optframe_Step1Append"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =540
                    LayoutCachedTop =600
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1740
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =215
                            Left =660
                            Top =480
                            Width =1440
                            Height =420
                            FontSize =16
                            FontWeight =700
                            BackColor =5394044
                            ForeColor =8454143
                            Name ="lblStepOne"
                            Caption ="Step 1"
                            LayoutCachedLeft =660
                            LayoutCachedTop =480
                            LayoutCachedWidth =2100
                            LayoutCachedHeight =900
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =4620
                            Top =780
                            Width =1140
                            Height =390
                            FontWeight =700
                            OptionValue =1
                            ForeColor =0
                            Name ="tglOne"
                            Caption ="One Tablet"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =4620
                            LayoutCachedTop =780
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =1170
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Gradient =12
                            BackColor =8289145
                            BackThemeColorIndex =4
                            BorderColor =8289145
                            BorderThemeColorIndex =4
                            HoverColor =65280
                            PressedColor =16711680
                            HoverForeColor =0
                            HoverForeThemeColorIndex =0
                            PressedForeColor =16711680
                            Shadow =-1
                            QuickStyle =23
                            QuickStyleMask =-1
                            WebImagePaddingTop =1
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =247
                            Left =5760
                            Top =780
                            Width =1140
                            Height =390
                            FontWeight =700
                            TabIndex =1
                            OptionValue =2
                            ForeColor =0
                            Name ="tglTwo"
                            Caption ="Two Tablets"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =5760
                            LayoutCachedTop =780
                            LayoutCachedWidth =6900
                            LayoutCachedHeight =1170
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Gradient =12
                            BackColor =8289145
                            BackThemeColorIndex =4
                            BorderColor =8289145
                            BorderThemeColorIndex =4
                            HoverColor =65280
                            PressedColor =16711680
                            HoverForeColor =0
                            HoverForeThemeColorIndex =0
                            PressedForeColor =16711680
                            Shadow =-1
                            QuickStyle =23
                            QuickStyleMask =-1
                            WebImagePaddingTop =1
                            Overlaps =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =600
                    Top =960
                    Width =3960
                    Height =240
                    FontWeight =700
                    ForeColor =8454143
                    Name ="lblNumberOfTablets"
                    Caption ="How many tablets were the data collected on?"
                    LayoutCachedLeft =600
                    LayoutCachedTop =960
                    LayoutCachedWidth =4560
                    LayoutCachedHeight =1200
                End
                Begin OptionGroup
                    Enabled = NotDefault
                    OverlapFlags =255
                    Left =540
                    Top =1980
                    Width =8817
                    Height =1920
                    ColumnOrder =7
                    TabIndex =6
                    Name ="optframe_Step2Append"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =540
                    LayoutCachedTop =1980
                    LayoutCachedWidth =9357
                    LayoutCachedHeight =3900
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =247
                            Left =660
                            Top =1860
                            Width =1500
                            Height =360
                            FontSize =16
                            FontWeight =700
                            BackColor =5394044
                            ForeColor =8454143
                            Name ="lblStepTwo"
                            Caption ="Step 2"
                            LayoutCachedLeft =660
                            LayoutCachedTop =1860
                            LayoutCachedWidth =2160
                            LayoutCachedHeight =2220
                        End
                        Begin ToggleButton
                            OverlapFlags =127
                            Left =4620
                            Top =2160
                            Width =1140
                            Height =390
                            FontWeight =700
                            OptionValue =1
                            ForeColor =0
                            Name ="tglTabletOne"
                            Caption ="Tablet One"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =4620
                            LayoutCachedTop =2160
                            LayoutCachedWidth =5760
                            LayoutCachedHeight =2550
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Gradient =12
                            BackColor =8289145
                            BackThemeColorIndex =4
                            BorderColor =8289145
                            BorderThemeColorIndex =4
                            HoverColor =65280
                            PressedColor =16711680
                            HoverForeColor =0
                            HoverForeThemeColorIndex =0
                            PressedForeColor =16711680
                            Shadow =-1
                            QuickStyle =23
                            QuickStyleMask =-1
                            WebImagePaddingTop =1
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =247
                            Left =5760
                            Top =2160
                            Width =1140
                            Height =390
                            FontWeight =700
                            TabIndex =1
                            OptionValue =2
                            ForeColor =0
                            Name ="tglTabletTwo"
                            Caption ="Tablet Two"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =5760
                            LayoutCachedTop =2160
                            LayoutCachedWidth =6900
                            LayoutCachedHeight =2550
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Gradient =12
                            BackColor =8289145
                            BackThemeColorIndex =4
                            BorderColor =8289145
                            BorderThemeColorIndex =4
                            HoverColor =65280
                            PressedColor =16711680
                            HoverForeColor =0
                            HoverForeThemeColorIndex =0
                            PressedForeColor =16711680
                            Shadow =-1
                            QuickStyle =23
                            QuickStyleMask =-1
                            WebImagePaddingTop =1
                            Overlaps =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =600
                    Top =2280
                    Width =3720
                    Height =240
                    FontWeight =700
                    ForeColor =8454143
                    Name ="lblTablet"
                    Caption ="On which tablet was this data collected?"
                    LayoutCachedLeft =600
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =2520
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =1
                    Left =660
                    Top =3600
                    Width =7020
                    Height =240
                    FontWeight =700
                    ForeColor =9868950
                    Name ="lblStepTwoFinish"
                    Caption ="Click 'Append Data' below and then REPEAT Step 2 for each event to be appended"
                    ControlTipText ="Select the Event from the main data set that you wish to append the secondary ta"
                        "blet  data to"
                    LayoutCachedLeft =660
                    LayoutCachedTop =3600
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =3840
                End
                Begin TextBox
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3060
                    Top =120
                    Width =4140
                    Height =255
                    ColumnOrder =2
                    TabIndex =7
                    ForeColor =16777215
                    Name ="tbxImportFile"

                    LayoutCachedLeft =3060
                    LayoutCachedTop =120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =375
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =1860
                    Top =4020
                    Width =2646
                    Height =373
                    ColumnOrder =1
                    TabIndex =8
                    Name ="optgSelectTables"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1860
                    LayoutCachedTop =4020
                    LayoutCachedWidth =4506
                    LayoutCachedHeight =4393
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            Left =600
                            Top =4080
                            Width =1245
                            Height =240
                            BackColor =0
                            ForeColor =16777215
                            Name ="lblSelect"
                            Caption ="Select Tables ..."
                            LayoutCachedLeft =600
                            LayoutCachedTop =4080
                            LayoutCachedWidth =1845
                            LayoutCachedHeight =4320
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =2040
                            Top =4123
                            OptionValue =1
                            Name ="chkALL"

                            LayoutCachedLeft =2040
                            LayoutCachedTop =4123
                            LayoutCachedWidth =2300
                            LayoutCachedHeight =4363
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =2270
                                    Top =4095
                                    Width =720
                                    Height =240
                                    ForeColor =16777215
                                    Name ="lblALL"
                                    Caption ="Select All"
                                    LayoutCachedLeft =2270
                                    LayoutCachedTop =4095
                                    LayoutCachedWidth =2990
                                    LayoutCachedHeight =4335
                                End
                            End
                        End
                        Begin CheckBox
                            OverlapFlags =87
                            Left =3174
                            Top =4123
                            TabIndex =1
                            OptionValue =2
                            Name ="chkNone"

                            LayoutCachedLeft =3174
                            LayoutCachedTop =4123
                            LayoutCachedWidth =3434
                            LayoutCachedHeight =4363
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3404
                                    Top =4095
                                    Width =930
                                    Height =240
                                    ForeColor =16777215
                                    Name ="lblNone"
                                    Caption ="Select None"
                                    LayoutCachedLeft =3404
                                    LayoutCachedTop =4095
                                    LayoutCachedWidth =4334
                                    LayoutCachedHeight =4335
                                End
                            End
                        End
                    End
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =5460
                    Top =1320
                    Width =1380
                    Height =360
                    ColumnOrder =0
                    FontWeight =600
                    TabIndex =9
                    ForeColor =0
                    Name ="tglImportPseudoEvents"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Caption ="??"
                    FontName ="Segoe UI"
                    ControlTipText ="Click to toggle to INCLUDE or EXCLUDE pseudoevents"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5460
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =1680
                    ForeThemeColorIndex =0
                    UseTheme =1
                    OldBorderStyle =0
                    BorderColor =1796857
                    BorderThemeColorIndex =5
                    HoverColor =10092492
                    PressedColor =10092492
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =16724787
                    QuickStyle =3
                    QuickStyleMask =-369
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =5
                    WebImagePaddingBottom =8
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =1
                            Left =2280
                            Top =1380
                            Width =3120
                            Height =285
                            FontSize =10
                            FontWeight =600
                            ForeColor =10092543
                            Name ="lblPseudoEvents"
                            Caption ="Append/Update PseudoEvents?"
                            LayoutCachedLeft =2280
                            LayoutCachedTop =1380
                            LayoutCachedWidth =5400
                            LayoutCachedHeight =1665
                        End
                    End
                End
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4800
                    Top =4080
                    Width =4815
                    Height =240
                    LeftMargin =36
                    TopMargin =36
                    RightMargin =36
                    BackColor =10092543
                    ForeColor =255
                    Name ="lblPseudoEventsIncluded"
                    Caption ="** PseudoEvents WILL be INCLUDED in data appends/updates **"
                    LayoutCachedLeft =4800
                    LayoutCachedTop =4080
                    LayoutCachedWidth =9615
                    LayoutCachedHeight =4320
                End
            End
        End
        Begin Section
            Height =360
            BackColor =15527148
            Name ="Detail"
            OnClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1245
                    Top =60
                    Width =2535
                    ColumnWidth =4260
                    Name ="txt_Table_Name"
                    ControlSource ="Table_Name"

                    LayoutCachedLeft =1245
                    LayoutCachedTop =60
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =165
                            Top =60
                            Width =960
                            Height =240
                            Name ="lblTableName"
                            Caption ="Table Name:"
                            LayoutCachedLeft =165
                            LayoutCachedTop =60
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =300
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =3840
                    Top =60
                    TabIndex =1
                    Name ="chk_Append"
                    ControlSource ="Append"
                    DefaultValue ="0"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =60
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =300
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListRows =30
                    ListWidth =5760
                    Left =5580
                    Top =60
                    Width =4380
                    TabIndex =2
                    Name ="cmbo_Append_Table"
                    ControlSource ="Append_Table"
                    RowSourceType ="Value List"
                    RowSource =" "
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextAlign =3
                            Left =4080
                            Top =60
                            Width =1440
                            Height =240
                            Name ="lblAppendFrom"
                            Caption ="Append data from:"
                            LayoutCachedLeft =4080
                            LayoutCachedTop =60
                            LayoutCachedWidth =5520
                            LayoutCachedHeight =300
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =600
            BackColor =5394044
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =6240
                    Top =120
                    Width =1620
                    FontWeight =700
                    ForeColor =0
                    Name ="cmd_Append_Event_Data"
                    Caption ="Append Data"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6240
                    LayoutCachedTop =120
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8039
                    Top =120
                    Width =1319
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="cmd_Close"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8039
                    LayoutCachedTop =120
                    LayoutCachedWidth =9358
                    LayoutCachedHeight =480
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =255
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =120
                    Width =2400
                    Height =330
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="btnSelectFile"
                    Caption ="Select Another Import File"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =450
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    PressedColor =6644321
                    PressedThemeColorIndex =4
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
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
Option Explicit

' =================================
' MODULE:       frm_Append_Append_Data
' Level:        Application module
' Version:      1.07
'
' Description:  field data import related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:    ML/GS - unknown   - 1.00 - initial version
'               BLC   - 5/21/2019 - 1.01 - added documentation, error handling, option explicit,
'                                          fixed CWD data import failure issue
'               BLC   - 8/26/2019 - 1.02 - shifted append to AppendFieldData() & re-tooled
'               BLC   - 8/27/2019 - 1.03 - commented out unused opt_frame_Select_Append_AfterUpdate
'               BLC   - 9/1/2019  - 1.04 - added checkboxes to select ALL tables or None,
'                                          added pre-set selections for tablets based on file name
'                                          (primary or secondary)
'               BLC   - 9/15/2019 - 1.05 - populated rsAppend to get accurate recordcount, updated strAppendSQL
'                                          to use single quotes to find Event_ID
'               BLC   - 9/20/2019 - 1.06 - add import pseudoevents toggle
'               BLC   - 9/24/2019 - 1.07 - added PseudoEvents deleted warning notice
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(value As String)
Public Event InvalidDirections(value As String)
Public Event InvalidCallingForm(value As String)

'---------------------
' Properties
'---------------------
Public Property Let title(value As String)
    If Len(value) > 0 Then
        m_Title = value

        'set the form title & caption
        'Me.lblTitle.Caption = m_Title
        'Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(value)
    End If
End Property

Public Property Get title() As String
    title = m_Title
End Property

Public Property Let Directions(value As String)
    If Len(value) > 0 Then
        m_Directions = value

        'set the form directions
        'Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(value As String)
        m_CallingForm = value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Events
'---------------------
' ----------------
'  Form
' ----------------
' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
'   BLC  - 8/27/2019 - set recordsource to ensure this hasn't been changed inadvertently in design view
'   BLC  - 8/31/2019 - determine if primary/secondary tablet based on import filename
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim ImportFile As String
    Dim Tablet As String
    
    'ensure form loads w/ proper data source!
    Me.RecordSource = "tsys_Append_Tables"
    Me.OrderBy = "Append_Order"
    Me.OrderByOn = True
    
    'defaults
    Me.btnSelectFile.Visible = False
    Me.cmbo_Select_Import_Event_Table.Enabled = False
    Me.cmbo_Select_Import_Events.Enabled = False
    Me.cmbo_Select_Event.Enabled = False
    Me.lblPseudoEventsIncluded.Visible = False
    Me.lblPseudoEventsDeleted.Visible = False
    
    'fetch import filename & determine if secondary or primary
    ImportFile = Nz(Me.OpenArgs, "")
    Me.tbxImportFile = ImportFile
    'Me.tbxImportFile.Locked = True
    
    'determine if primary or secondary tablet
    If Len(ImportFile) > Len(Replace(ImportFile, "Primary", "")) Then
        'primary tablet
        Me.optframe_Step1Append = 1
        Me.optframe_Step2Append = 1
                
    ElseIf Len(ImportFile) > Len(Replace(ImportFile, "Secondary", "")) Then
        'secondary tablet
        Me.optframe_Step1Append = 2
        Me.optframe_Step2Append = 2
        
    End If
    
    'set chk append for all records
    Me.optgSelectTables = 1 'ALL tables
    Me.chk_Append = True

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdef As TableDef
    
    Set db = CurrentDb
    Set rs = Me.RecordsetClone
     
    rs.MoveFirst
    
    Do While Not rs.EOF
    
        DoCmd.RunCommand acCmdSaveRecord
        'Until all of the Slope and Aspect data are updated in the master locations table we want to update the locations
        'table with the slope and aspect data collected in the field on the field data bases.
        '    If Me!txt_Table_Name = "tbl_Locations" Then
        '        Me!chk_Append.Value = 0
        '    Else
        '        Me!chk_Append.Value = 1
        '    End If
        '**************************************************
        DoCmd.RunCommand acCmdSaveRecord
    
        For Each tdef In db.TableDefs
            Dim iTableName As Long
            'iTableName = Len(rs![Table_Name]) '(Me!txt_Table_Name.Value)
            
            Dim strTableName As String
            strTableName = rs![Table_Name]
            iTableName = Len(strTableName)
            Dim strAppTableName As String
            
            If Left(tdef.Name, 1) = "_" Then
                
                strAppTableName = Right(Left(tdef.Name, iTableName + 1), iTableName)
            
                'If it is tbl_Events make sure we are grabbing the events table from the primary tablet.
                If strAppTableName = "tbl_Events" Then
                    If Right(tdef.Name, 9) = "SECONDARY" Then
                        GoTo NextRecord:
                    End If
                End If
                
                If strAppTableName = strTableName Then
                    rs.Edit
                    rs![Append_Table] = tdef.Name
                    rs.Update
                End If
            
            Else
                
                GoTo NextRecord:
            
            End If
        
NextRecord:
        Next
        
        rs.MoveNext
    Loop
     
    Me.OrderBy = "Append_Order ASC"
    Me.OrderByOn = True
    
    Me.optframe_Step1Append.SetFocus
    Me.optframe_Step2Append.Enabled = False
    
    Set db = Nothing
    Set rs = Nothing
    Set tdef = Nothing

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
'   BLC  - 8/27/2019 - shifted code to ClearAppendTables
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    ClearAppendTables

'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim ctrlCombo As ComboBox
'
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset("tsys_Append_Tables")
'    Set ctrlCombo = Me!cmbo_Append_Table
'    Me!cmbo_Append_Table.RowSource = " "
'
'    rs.MoveFirst
'
'    Do While Not rs.EOF
'        rs.Edit
'        rs![Append_Table] = ""
'        rs.Update
'    rs.MoveNext
'    Loop
'
'    Set db = Nothing
'    Set rs = Nothing

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  GotFocus Events
' ----------------

' ---------------------------------
' Sub:          cbxAppendTable_GotFocus
' Description:  combobox focus actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub cmbo_Append_Table_GotFocus()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdef As TableDef
    
    Set db = CurrentDb
    Me!cmbo_Append_Table.RowSource = ""
    For Each tdef In db.TableDefs
        Dim iTableName As Long
            iTableName = Len(Me!txt_Table_Name.value)
        Dim strTableName As String
            strTableName = Me!txt_Table_Name.value
        Dim strAppTableName As String
       
        If Left(tdef.Name, 1) = "_" Then
            strAppTableName = Right(Left(tdef.Name, iTableName + 1), iTableName)
            If strAppTableName = strTableName Then
                Me!cmbo_Append_Table.AddItem tdef.Name
            End If
        Else
            GoTo NextRecord:
        End If
NextRecord:
    Next
    Me!cmbo_Append_Table.Requery
    
    'Cleanup
    Set db = Nothing
    Set tdef = Nothing

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAppendTable_GotFocus[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSelectImportEventTable_GotFocus
' Description:  combobox focus actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub cmbo_Select_Import_Event_Table_GotFocus()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdef As TableDef
    
    Set db = CurrentDb
    
    Me!cmbo_Select_Import_Event_Table.RowSource = ""
    
    For Each tdef In db.TableDefs
        If Left(tdef.Name, 11) = "_tbl_Events" Then
            Me!cmbo_Select_Import_Event_Table.AddItem tdef.Name
        Else
            GoTo NextRecord:
        End If
NextRecord:
    Next
    Me!cmbo_Select_Import_Event_Table.Requery
    Set db = Nothing
    Set tdef = Nothing

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectImportEventTable_GotFocus[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSelectImportEvents_GotFocus
' Description:  combobox focus actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub cmbo_Select_Import_Events_GotFocus()
On Error GoTo Err_Handler

    Me!cmbo_Select_Import_Events.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectImportEvents_GotFocus[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSelectEvent_GotFocus
' Description:  combobox focus actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 25, 2019
' Adapted:      -
' Revisions:
'   BLC - 8/25/2019 - initial version
' ---------------------------------
Private Sub cmbo_Select_Event_GotFocus()
On Error GoTo Err_Handler

    'NOTE - make sure to requery the DDL otherwise no data is visible though the SQL is correct
    '       if Me.Requery was used it would refer to the FORM not the DDL & no data would show
    Me!cmbo_Select_Event.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectEvent_GotFocus[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click Events
' ----------------
' ---------------------------------
' Sub:          btnAppendEventData_Click
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling, fixed strUpdateEventIDSQL to include quotes around EventID,
'                      declare variables (esp. GUIDs)
'   BLC  - 8/26/2019 - shift code to AppendFieldData()
' ---------------------------------
Private Sub cmd_Append_Event_Data_Click()
On Error GoTo Err_Handler

    AppendFieldData

'    Dim db As DAO.Database
'    Set db = CurrentDb
'
'    Dim rsMain As DAO.Recordset  'rsMain is the master dataset in the database to which data is being appended
'    Dim rsAppend As DAO.Recordset 'rsAppend is the dataset with new records to be appended to the
'    Dim rsForm As DAO.Recordset 'the recordset for the append form.
'    'Dim rsAppendLog As DAO.Recordset 'the recordset that stores the information about the records appended to each table
'    Dim rsAppendLog As DAO.Recordset 'the recordset that stores the information about the records appended to each table
'
'    Dim strMain As String
'    Dim strAppend As String
'
'    'This is required
'    DoCmd.RunCommand acCmdSaveRecord
'
'    Set rsForm = Me.RecordsetClone
'
'    rsForm.MoveFirst
'
'    Do Until rsForm.EOF
'
'        'Cycle through the tables to see which ones have been chosen
'        'to have new data appended.
'        If rsForm![Append] = True Then
'
'            'rsMain is the main dataset in the database
'             strMain = rsForm![Table_Name]
'
'            Set rsMain = db.OpenRecordset(strMain)
'
'                'Get the length of the table name to check and make sure that the Main Table and the Append Table names match
'                    Dim iLength As Long
'                    Dim iLength2 As Long
'                    iLength = Len(rsForm![Table_Name])
'                    iLength2 = iLength + 1
'
'                    'Root of the Append Table name
'                    Dim strAppTableName As String
'                    strAppTableName = Right(Left(rsForm![Append_Table], iLength2), iLength)
'
'                'Check to make sure that an Append Table is specified if the Append box is checked
'                If rsForm![Append_Table] = "" Or IsNull(rsForm![Append_Table]) Then
'                    MsgBox "Make sure that you have properly selected all of the data you wish to append!", vbCritical, "Append Data"
'                    Exit Sub
'                'Check to make sure that the Append Table Name matches the Main Table Name
'                ElseIf strAppTableName <> rsForm![Table_Name] Then
'                    MsgBox "Make sure you have properly selected the data set to append to " & rsForm![Table_Name] & ".", vbCritical, "Append Data"
'                    Exit Sub
'                End If
'
'            strAppend = rsForm![Append_Table]
'            Dim strAppendTableName As String
'            strAppendTableName = strAppend
'
'            'Capture the imported events table to use when appending new tree and sapling data.
'            If strMain = "tbl_Events" Then
'                Dim rsEvents As DAO.Recordset
'                Set rsEvents = db.OpenRecordset(strAppend)
'            End If
'
'            'Check to see if the table is tbl_Locations. If so, send it to a special functionto update the locations table with newly collected slope and aspect.
'            If strMain = "tbl_Locations" Then
'                Dim rsLoc As DAO.Recordset
'                Set rsLoc = db.OpenRecordset(strAppend)
'                'send it to a specail function to check to see if anything needs updating. If so, update it and return.
'                fxnUpdateLocInfo rsLoc, strAppend
'
'            End If
'
'
'        'Determine if the data being appended is for tbl_Tags.
'        'If it is send it to a special function to update the data in these tables tables prior to appending new data.
'
'            If strMain = "tbl_Tags" Then
'                Set rsAppend = db.OpenRecordset(strAppend)
'                UpdateTags rsMain, rsAppend, rsEvents, strAppendTableName
'                GoTo NextRecord:
'            End If
'
'     'If you are appending records to an existing Event_ID:
'     'First figure out if you are appending data from the Main Tablet or Secondary Tablet
'
'            'If it is from the secondary tablet run it through this code to replace Event_IDs
'            If Me!optframe_Step1Append.Value = 2 Then
'                If Me!optframe_Step2Append.Value = 2 Then
'                    If Me!cmbo_Select_Event = "" Or IsNull(Me!cmbo_Select_Event) Then
'                        MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
'                        Me!cmbo_Select_Event.SetFocus
'                        Exit Sub
'                    ElseIf Me!cmbo_Select_Import_Event_Table = "" Or IsNull(Me!cmbo_Select_Import_Event_Table) Then
'                            MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
'                            Me!cmbo_Select_Import_Event_Table.SetFocus
'                            Exit Sub
'                    ElseIf Me!cmbo_Select_Import_Events = "" Or IsNull(Me!cmbo_Select_Import_Events) Then
'                                MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
'                                Me!cmbo_Select_Import_Events.SetFocus
'                                Exit Sub
'                    End If
'
'        'Declare and set the variables for the Event ID's in both the Main (master) dataset
'        'as well as in the imported data set
'
'                    Dim GUIDMain As Variant 'GUID 'String
'                    Dim GUIDReplace As Variant 'GUID 'String
'
'                    GUIDMain = Me!cmbo_Select_Event.Column(0)
'                    GUIDReplace = Me!cmbo_Select_Import_Events.Column(0)
'
'     'Check to see if the table contains an Event_ID field
'
'                    Dim boolEvent As Boolean
'                    boolEvent = False
'
'                    Dim tdef As DAO.TableDef
'                    Dim lCount As Long
'                    Dim lCtr As Long
'                    Dim strFieldName As String
'                    Set tdef = db.TableDefs(strAppend)
'
'                    With tdef
'                        lCount = .Fields.Count
'                            For lCtr = 0 To lCount - 1
'                                strFieldName = .Fields(lCtr).Name
'
'                                If strFieldName = "Event_ID" Then
'                                    boolEvent = True
'                                End If
'
'                            Next
'                    End With
'
'     'if the table contains an Event_ID pass the data set to the update event id function
'
'                    If boolEvent = True Then
'
'        'Pass the appending data set, the master GUID as well as the GUID that needs to be
'        'replaced to the function to update the Event ID to the Master Event ID
'                        Dim strUpdateEventIDSQL As String
'                        Dim qdefUpdateEventID As QueryDef
'                        Dim strTableName As String
'                        Dim strFindUpdate As String
'
'                        strFindUpdate = GUIDReplace
'                        strTableName = Me!cmbo_Append_Table.Value
'
'                   'Only select those records where the EventID = the GUID to be replaced
'
'                        strUpdateEventIDSQL = "SELECT [" & strAppend & "].* " _
'                        & "FROM [" & strAppend & "] " _
'                        & "WHERE ((([" & strAppend & "].Event_ID) = " & strFindUpdate & ")); "
'
'                   ' MsgBox strUpdateEventIDSQL
'
'                    'save the SQL as a qdef and then create a recordset to pass to the UPDATE EVENT ID Function
'                        Set qdefUpdateEventID = db.CreateQueryDef("_Qry_UpdateEventID", strUpdateEventIDSQL)
'                        Set rsAppend = db.OpenRecordset("_Qry_UpdateEventID")
'
'                        If rsAppend.RecordCount > 0 Then
'                            UpdateEventID rsAppend, GUIDMain, GUIDReplace, strTableName
'                        End If
'                    'Delete the Select Query from the database after it has been used in the update event function
'                        DoCmd.DeleteObject acQuery, qdefUpdateEventID.Name
'
'        'Had to insert this code because when the record set was passed to the Append
'        'function it was attempting to append too many records.
'        'If events already exist additional information must be appended on a
'        'Event by Event basis.
'
'                        Dim strAppendSQL As String
'                        Dim qdefAppend As QueryDef
'
'            'main event id
'                        Dim strFind As String
'                        strFind = GUIDMain
'
'            'query only the records that were collected on this event
'
'                        strAppendSQL = "SELECT [" & strAppend & "].* " _
'                        & " FROM [" & strAppend & "] " _
'                        & " WHERE ((([" & strAppend & "].Event_ID) = " & strFind & "));"
'
'                        Set qdefAppend = db.CreateQueryDef("_Qry_AppendEventRecs", strAppendSQL)
'
'            'reset the rsAppend variable to equal the query that only contains records from
'            'targeted event
'
'                        Set rsAppend = db.OpenRecordset("_Qry_AppendEventRecs")
'           ' MsgBox rsAppend.Name
'
'            'pass the event specific recordset to the append function
'
'                        AppendtoTable rsAppend, rsMain, strAppendTableName
'
'                  'Delete the select query once the append function has been completed.
'
'                        DoCmd.DeleteObject acQuery, qdefAppend.Name
'
'                    Else
'
'                        GoTo AppendData:
'
'                    End If
'
'                Else
'
'                    GoTo AppendData:
'
'                End If
'
'            Else
'
'     'skip all of that crap about updating the event id and just append the damn records
'
'AppendData:
'
'     'We want to make sure to select only those records that are associated witht the events being imported.
'      'skip this query if the current append table is tbl_Events
'
'     If strAppTableName <> "tbl_Events" Then
'
'     'Check to see if the table contains an Event_ID field
'
'                    'Dim boolEvent As Boolean
'                    boolEvent = False
'
'                    'Dim tdef As DAO.TableDef
'                    'Dim lCount As Long
'                    'Dim lCtr As Long
'                    'Dim strFieldName As String
'                    Set tdef = db.TableDefs(strAppend)
'
'                    With tdef
'                        lCount = .Fields.Count
'                            For lCtr = 0 To lCount - 1
'                                strFieldName = .Fields(lCtr).Name
'
'                                If strFieldName = "Event_ID" Then
'                                    boolEvent = True
'                                End If
'
'                            Next
'                    End With
'            'If the data table has an Event_ID field, run it through the query that only selects data collected on the imported events. _
'            If it does not have an Event_ID field, send it through the standard Append function for now.
'
'            If boolEvent = True Then
'
'                    Dim strEvents As String, strSQL_FindImportedRecs As String
'                    strEvents = rsEvents.Name
'
'                    strSQL_FindImportedRecs = "SELECT [" & strAppend & "].* FROM [" & strEvents & "] INNER JOIN [" & strAppend & "] " _
'                                            & "ON [" & strEvents & "].Event_ID = [" & strAppend & "].Event_ID;"
'
'    'Turn the SQL statement into a query to be used in the following append fuctions as the Append data set
'
'                    Dim qdef_NewEventRecs As QueryDef
'                    Set qdef_NewEventRecs = db.CreateQueryDef("_qry_NewEventData", strSQL_FindImportedRecs)
'
'                    Set rsAppend = db.OpenRecordset("_qry_NewEventData")
'
'                    'Send the two recordsets (rsMain and rsAppend) to the Append Function
'
'                    AppendtoTable rsAppend, rsMain, strAppendTableName
'
'                    'delete the query def so that it can be recreated as the code loops
'
'                    DoCmd.DeleteObject acQuery, qdef_NewEventRecs.Name
'            Else
'
'       'Run the append function for any table that does not have an Event_ID
'            Set rsAppend = db.OpenRecordset(strAppend)
'
'            AppendtoTable rsAppend, rsMain, strAppendTableName
'
'            End If
'
'    Else
'
'    'Run the Append function for tbl_Events
'
'        Set rsAppend = db.OpenRecordset(strAppend)
'
'
'        AppendtoTable rsAppend, rsMain, strAppendTableName
'
'    End If
'
'    End If
'    'Make sure that the Events table is checked even if you do not need to append any data to it.
'
'    ElseIf rsForm![Table_Name] = "tbl_Events" And rsForm![Append] = False Then
'                    MsgBox "The Events table needs to be included in the append operation." & vbNewLine & vbNewLine & _
'                    "Please go back and check the append box and select an imported events table.", , "Append Data"
'                    Exit Sub
'    Else
'
'            GoTo NextRecord:
'
'    End If
'
'NextRecord:
'
'        rsForm.MoveNext
'
'    Loop
'
'    MsgBox "Update and Appending complete!", , "Update and Append Data"
'
'CleanUp:
'
'    Set rsAppend = Nothing
'    Set rsMain = Nothing
'    Set rsAppendLog = Nothing
'    Set rsEvents = Nothing
'
'    Set db = Nothing
'    Set rsForm = Nothing
'    Set qdefAppend = Nothing
'    Set qdefUpdateEventID = Nothing
'    Set qdef_NewEventRecs = Nothing

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAppendEventData_Click[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnViewUpdateLog_Click
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub cmd_ViewUpdateLog_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Update_Log"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewUpdateLog_Click[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Detail_Click
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub Detail_Click()
On Error GoTo Err_Handler

    If Me!optframe_Step2Append.value = 2 Then
        If Me!cmbo_Select_Event = "" Then
            MsgBox "You must complete the necessary information above.", , "Append Data"
            Me!cmbo_Select_Event.SetFocus
        ElseIf Me!cmbo_Select_Import_Event_Table.value = "" Then
            MsgBox "You must complete the necessary information above.", , "Append Data"
            Me!cmbo_Select_Import_Event_Table.SetFocus
        ElseIf Me!cmbo_Select_Import_Events.value = "" Then
            MsgBox "You must complete the necessary information above.", , "Append Data"
            Me!cmbo_Select_Import_Events.SetFocus
        End If
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Click[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAppendLog_Click
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub cmd_AppendLog_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Log"
    DoCmd.OpenForm stDocName, acFormDS, , stLinkCriteria

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAppendLog_Click[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          optgSelectTables_AfterUpdate
' Description:  option group after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 1, 2019
' Adapted:      -
' Revisions:
'   BLC - 9/1/2019  - initial version
' ---------------------------------
Private Sub optgSelectTables_AfterUpdate()
On Error GoTo Err_Handler

    Dim rsSelect As DAO.Recordset
    Set rsSelect = Me.RecordsetClone
    
    'ensure there are records
    If Not (rsSelect.EOF And rsSelect.BOF) Then
    
        rsSelect.MoveFirst
        
        Do Until rsSelect.EOF
        
        If Me!optgSelectTables.value = 1 Then
             
            rsSelect.Edit
            rsSelect![Append] = True
            rsSelect.Update
            
        ElseIf Me!optgSelectTables.value = 2 Then
            rsSelect.Edit
            rsSelect![Append] = False
            rsSelect.Update
            
        Else: GoTo NextRecord:
        
        End If
        
NextRecord:
        rsSelect.MoveNext
        
        Loop
        
        'Me.btnSelect.Enabled = True
        
    Else
        'no records to select
        'Me.btnSelect.Enabled = False

    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optgSelectTables_AfterUpdate[frm_Append_Append_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnClose_Click
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub cmd_Close_Click()
On Error GoTo Err_Handler

    Dim strResponse As String
    
    strResponse = MsgBox("Would you like to delete any of the imported tables?", vbYesNoCancel, "Delete Tables?")
    
    If strResponse = vbYes Then
        DoCmd.Close
        DoCmd.OpenForm "frm_Append_Delete_Tables"
    ElseIf strResponse = vbCancel Then
        Exit Sub
    Else
        DoCmd.Close
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnSelectFile_Click
' Description:  file selection click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 27, 2019
' Adapted:      -
' Revisions:
'   BLC - 8/27/2019 - initial version
' ---------------------------------
Private Sub btnSelectFile_Click()
On Error GoTo Err_Handler

    ClearAppendTables

    DoCmd.Close
    DoCmd.OpenForm "frm_Append_Select_Import_File", , , , , acWindowNormal
    
'    Dim db As DAO.Database
'    Dim rs As DAO.Recordset
'    Dim ctrlCombo As ComboBox
'
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset("tsys_Append_Tables")
'    Set ctrlCombo = Me!cmbo_Append_Table
'    Me!cmbo_Append_Table.RowSource = " "
'
'    rs.MoveFirst
'
'    Do While Not rs.EOF
'        rs.Edit
'        rs![Append_Table] = ""
'        rs.Update
'    rs.MoveNext
'    Loop
'
'    Set db = Nothing
'    Set rs = Nothing

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSelectFile_Click[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  AfterUpdate Events
' ----------------

'' ---------------------------------
'' Sub:          optgSelectAppend_AfterUpdate
'' Description:  option after update actions
'' Assumptions:
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
'' Adapted:      Bonnie Campbell, May 21, 2019
'' Revisions:
''   MEL/GS - unknown - initial version
''   BLC  - 5/21/2019 - added documentation, error handling
'' ---------------------------------
'Private Sub opt_frame_Select_Append_AfterUpdate()
'On Error GoTo Err_Handler
'
'    '-----------------------
'    '  Primary Tablet
'    '-----------------------
'    If Me!opt_frame_Select_Append.Value = 1 Then
'
'        Me!cmbo_Select_Event.Enabled = False
'        Me!cmbo_Select_Import_Event_Table.Enabled = False
'        Me!cmbo_Select_Import_Events.Enabled = False
'        Me!Lbl_Step2_Finish.Visible = False
'
'        Me.RecordSource = "tsys_Append_Tables"
'        Me!cmbo_Select_Event = ""
'        Me!cmbo_Select_Import_Event_Table = ""
'        Me!cmbo_Select_Import_Events = ""
'
'    '-----------------------
'    '  Secondary Tablet
'    '-----------------------
'    ElseIf Me!opt_frame_Select_Append.Value = 2 Then
'
'        Me!cmbo_Select_Event.Enabled = True
'        Me!cmbo_Select_Import_Event_Table.Enabled = True
'        Me!cmbo_Select_Import_Events.Enabled = True
'        Me!Lbl_Step2_Finish.Visible = True
'
'        Me.RecordSource = "qry_Append"
'        Me!cmbo_Select_Event = ""
'        Me!cmbo_Select_Import_Event_Table = ""
'        Me!cmbo_Select_Import_Events = ""
'
'    End If
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - optgSelectAppend_AfterUpdate[frm_Append_Append_Data form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Sub:          optStep1Append_AfterUpdate
' Description:  option after update actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
' ---------------------------------
Private Sub optframe_Step1Append_AfterUpdate()
On Error GoTo Err_Handler

    Select Case optframe_Step1Append.value
    
        '-----------------------
        '  One Tablet
        '-----------------------
        Case 1
            Me.Detail.Visible = True
            
             Me!optframe_Step2Append.value = 0
             optframe_Step2Append.Enabled = False
             
             Me.tglImportPseudoEvents.Enabled = True
             
             Me!cmbo_Select_Event.Enabled = False
             Me!cmbo_Select_Import_Event_Table.Enabled = False
             Me!cmbo_Select_Import_Events.Enabled = False
             Me!lblStepTwoFinish.Visible = False
            
             Me.RecordSource = "tsys_Append_Tables"
             Me!cmbo_Select_Event = ""
             Me!cmbo_Select_Import_Event_Table = ""
             Me!cmbo_Select_Import_Events = ""
             
             'Order the append tables in the proper order so that there are no errors during the append sequence
             Me.OrderBy = "Append_Order ASC"
             Me.OrderByOn = True
             
             Me!cmd_Append_Event_Data.Enabled = True
         
        '-----------------------
        '  Two Tablets
        '-----------------------
         Case 2
        
            Me.tglImportPseudoEvents.Enabled = True
            
            optframe_Step2Append.Enabled = True
            Me!optframe_Step2Append.value = 0
            Me!optframe_Step2Append.SetFocus
            Me.Detail.Visible = False
            
            Me!cmd_Append_Event_Data.Enabled = False
            
        Case Else
        
            Me.tglImportPseudoEvents.Enabled = False
            'optframe_Step2Import.Enabled = False
            optframe_Step2Append.Enabled = False
            Me!cmd_Append_Event_Data.Enabled = False
        
    End Select

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optStep1Append_AfterUpdate[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          optStep2Append_AfterUpdate
' Description:  option after update actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
'   BLC  - 8/27/2019 - added fore color change for stepping through each item
' ---------------------------------
Private Sub optframe_Step2Append_AfterUpdate()
On Error GoTo Err_Handler

    Me.Detail.Visible = True
    
    'default
    Me.lblEventsSecondaryImport.ForeColor = lngLtGray
    Me.lblEventSecondary.ForeColor = lngLtGray
    Me.lblMasterEventAppend.ForeColor = lngLtGray
    
    '-----------------------
    '  Primary Tablet
    '-----------------------
    If Me!optframe_Step2Append.value = 1 Then
      
        Me.RecordSource = "qry_Append_Primary_Tablet_Append"
        
        Me.OrderBy = "Append_Order ASC"
        Me.OrderByOn = True
        
        Me!cmbo_Append_Table.SetFocus
        
        Me!cmd_Append_Event_Data.Enabled = True
        
    '-----------------------
    '  Secondary Tablet
    '-----------------------
    ElseIf Me!optframe_Step2Append.value = 2 Then
           
        Me!cmbo_Select_Event.Enabled = True
        Me!cmbo_Select_Import_Event_Table.Enabled = True
        Me!cmbo_Select_Import_Events.Enabled = True
        Me!lblStepTwoFinish.Visible = True
        
        Me.RecordSource = "qry_Append_Secondary_Tablet_Append"
        Me!cmbo_Select_Event = ""
        Me!cmbo_Select_Import_Event_Table = ""
        Me!cmbo_Select_Import_Events = ""
    
        Me!cmbo_Select_Import_Event_Table.SetFocus
        
        Me!cmd_Append_Event_Data.Enabled = True
        
        Me.lblEventsSecondaryImport.ForeColor = lngLtBlue
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optStep2Append_AfterUpdate[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglImportPseudoEvents_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  True = IMPORT pseudo events from the import tables
'               False = EXCLUDE pseudo events from import tables
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 20, 2019
' Adapted:      -
' Revisions:
'   BLC - 9/20/2019 - initial version
'   BLC - 9/24/2019 - added PseudoEvents deleted warning notice
' ---------------------------------
Private Sub tglImportPseudoEvents_AfterUpdate()
On Error GoTo Err_Handler

    'Debug.Print Abs(tglImportPseudoEvents.Value)
    
    'default
    Me.lblPseudoEventsIncluded.Visible = False
    Me.lblPseudoEventsDeleted.Visible = False
    
    SetTempVar "ImportPseudoEvents", tglImportPseudoEvents.value
    
    With tglImportPseudoEvents
        Select Case .value
            Case True
                .Caption = "YES, INCLUDE"
                .BackColor = lngLtLime
                .ForeColor = lngBlue
                Me.lblPseudoEventsIncluded.Visible = True
            Case False
                .Caption = "NO, EXCLUDE"
                .BackColor = lngWhite
                .ForeColor = lngRed
                Me.lblPseudoEventsDeleted.Visible = True
            Case Else
                .Caption = "??"
        End Select
    End With
    
    'trigger the after update event if the table has been selected to refresh the import events combobox
    If Len(Me.cmbo_Select_Import_Event_Table.value) > 0 Then
        cmbo_Select_Import_Event_Table_AfterUpdate
        Me.cmbo_Select_Import_Events.Requery
    End If
    
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglImportPseudoEvents_AfterUpdate[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSelectImportEventTable_AfterUpdate
' Description:  combobox after update actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 21, 2019
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC  - 5/21/2019 - added documentation, error handling
'   BLC  - 8/27/2019 - added fore color change to guide workflow
'   BLC  - 9/20/2019 - adjust for include/exclude pseudoevents
' ---------------------------------
Private Sub cmbo_Select_Import_Event_Table_AfterUpdate()
On Error GoTo Err_Handler

    Dim strTableName As String
    Dim EventSQL As String
    
    'default
    Me.lblEventsSecondaryImport.ForeColor = lngLtGray
    Me.lblEventSecondary.ForeColor = lngLtGray
    Me.lblMasterEventAppend.ForeColor = lngLtGray
    Me.lblStepTwoFinish.ForeColor = lngLtGray
    
    If Me!cmbo_Select_Import_Event_Table = "" Or IsNull(Me!cmbo_Select_Import_Event_Table) Then
        
        Exit Sub
    
    Else
        strTableName = Me!cmbo_Select_Import_Event_Table.value
        
        Dim strExclude As String
        strExclude = IIf(Me.tglImportPseudoEvents = True, "", " WHERE PseudoEvent = 0 ")
        
        EventSQL = "SELECT [" & strTableName & "].Event_ID, [" & strTableName & "].Location_ID, " _
            & "[tbl_Locations].[Plot_Name] &" & """  """ & "& [" & strTableName & "].[Event_Date] " _
            & "AS [Pick String] " _
            & "FROM [" & strTableName & "] " _
            & "LEFT JOIN tbl_Locations " _
            & "ON [" & strTableName & "].Location_ID = tbl_Locations.Location_ID " _
            & strExclude _
            & "ORDER BY Event_Date DESC;"
            
Debug.Print "toggle pseudos = " & Me.tglImportPseudoEvents
Debug.Print "Import EventTable - select import events - EventSQL = " & EventSQL
           'MsgBox EventSQL
          
        Me!cmbo_Select_Import_Events.RowSource = EventSQL

        Me.lblEventSecondary.ForeColor = lngLtBlue

    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectImportEventTable_AfterUpdate[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSelectImportEvents_AfterUpdate
' Description:  combobox after update actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 27, 2019
' Adapted:      -
' Revisions:
'   BLC  - 8/27/2019 - initial version
'   BLC  - 9/12/2019 - set GUIDReplace TempVar for use in AppendFieldData
' ---------------------------------
Private Sub cmbo_Select_Import_Events_AfterUpdate()
On Error GoTo Err_Handler

    Dim strTableName As String
    Dim EventSQL As String
    
    'default
    Me.lblEventsSecondaryImport.ForeColor = lngLtGray
    Me.lblEventSecondary.ForeColor = lngLtGray
    Me.lblMasterEventAppend.ForeColor = lngLtGray
    Me.lblStepTwoFinish.ForeColor = lngLtGray
    
    If Me!cmbo_Select_Import_Events = "" Or IsNull(Me!cmbo_Select_Import_Events) Then
        Exit Sub
    
    Else
        strTableName = Me!cmbo_Select_Import_Events.value
        
        EventSQL = "SELECT e.Event_ID, [Plot_Name] & "" "" & "" "" & [Event_Date] AS PickString " _
                    & "FROM tbl_Locations l " _
                    & "INNER JOIN tbl_Events e ON l.Location_ID = e.Location_ID " _
                    & "WHERE (((Year([Event_Date])) = Year(Now()))) " _
                    & "ORDER BY e.Event_Date DESC;"
                             
    Debug.Print EventSQL
          
        Me!cmbo_Select_Event.RowSource = EventSQL
               
        Me.lblMasterEventAppend.ForeColor = lngLtBlue
               
    End If

Debug.Print "cmbo_Select_Import_Events.Value: " & Me.cmbo_Select_Import_Events.value
    
    'set GUIDReplace
    SetTempVar "GUIDReplace", Me.cmbo_Select_Import_Events.value

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectImportEvents_AfterUpdate[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxSelectEvent_AfterUpdate
' Description:  combobox after update actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 27, 2019
' Adapted:      -
' Revisions:
'   BLC  - 8/27/2019 - initial verison
'   BLC  - 9/12/2019 - set GUIDMain TempVar for use in AppendFieldData
' ---------------------------------
Private Sub cmbo_Select_Event_AfterUpdate()
On Error GoTo Err_Handler

    Dim strTableName As String
    Dim EventSQL As String
    
    'default
    Me.lblEventsSecondaryImport.ForeColor = lngLtGray
    Me.lblEventSecondary.ForeColor = lngLtGray
    Me.lblMasterEventAppend.ForeColor = lngLtGray
    Me.lblStepTwoFinish.ForeColor = lngLtGray
    
    If Me!cmbo_Select_Event = "" Or IsNull(Me!cmbo_Select_Event) Then
        Exit Sub
    
    Else
        
        Me.lblStepTwoFinish.ForeColor = lngLtBlue
               
    End If
    
    'set GUIDMain
    SetTempVar "GUIDMain", Me.cmbo_Select_Event.value

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectEvent_AfterUpdate[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------
' ---------------------------------
' Sub:          AppendFieldData
' Description:  button click actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 25, 2019
' Adapted:      -
' Revisions:
'   BLC - 8/25/2019 - initial version
'   BLC - 8/28/2019 - fixed tbl_Quadrat_Data_Import fail to import issue (added single quotes around Event_ID)
'   BLC - 9/15/2019 - populated rsAppend to get accurate recordcount, updated strAppendSQL to use single quotes to find Event_ID
' ---------------------------------
Private Sub AppendFieldData()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' ---------------------------------------------------------------
    ' recordsets:
    ' ---------------------------------------------------------------
    ' rsMain        = master cumulative being appended to
    ' rsAppend      = new records to append to master
    ' rsForm        = append form's recordset
    ' rsAppendLog   = stores info @ records appended to each table
    ' rsEvents      = new records which have an Events_ID
    ' ---------------------------------------------------------------
    Dim rsMain As DAO.Recordset
    Dim rsAppend As DAO.Recordset
    Dim rsForm As DAO.Recordset
    Dim rsAppendLog As DAO.Recordset
    
    Dim rsEvents As DAO.Recordset
    
    Dim strMain As String
    Dim strAppend As String
    
'    Dim iLength As Long
'    Dim iLength2 As Long
    Dim strAppTableName As String
    
    'handle pseudoevents
    Dim strExcludePseudos As String
    
    strExcludePseudos = ""
    If Me.tglImportPseudoEvents = False Then
        strExcludePseudos = " AND PseudoEvent = 0 "
    
    End If
    
    'This is required
    DoCmd.RunCommand acCmdSaveRecord
    
    Set rsForm = Me.RecordsetClone
    
    rsForm.MoveFirst

'may not require this section if append order honored
EventIdentify:
'-----------------------------------
' Event Loop
'-----------------------------------
    Do Until rsForm.EOF
        
        'identify table
        strMain = rsForm![Table_Name]
Debug.Print "strMain = " & strMain
        Set rsMain = db.OpenRecordset(strMain)
        
        strAppTableName = CompareTables(rsForm)
    
        If rsForm![Table_Name] = "tbl_Events" Then
                        
            Dim AppTableName As String
            strAppend = rsForm![Append_Table]
            
            '---------------
            ' pseudoevents?
            '---------------
            ' Excluded? --> delete from import table BEFORE appending data
            ' Included? --> do nothing
            If Me.tglImportPseudoEvents = False Then
                'delete pseudoevents from import event table
                DeletePseudoEvents strAppend
                
                Debug.Print "pseudoevents deleted from " & strAppend
                
                'remove pseudo-event related records from other tables
                DeleteRelatedPseudoEventRecords strAppend
            
            End If
            
            Set rsEvents = db.OpenRecordset(strAppend)
            
            Exit Do
        End If
        
        rsForm.MoveNext
        
    Loop
    Debug.Print "strAppend=" & strAppend

MainAppend:
    Debug.Print "MainAppend loop"

'-----------------------------------
' Main Append Loop
'-----------------------------------
    'cycle through tables to see which chosen to have new data appended
    'start @ first record
    rsForm.MoveFirst
    
    'give a message if there aren't any records
    If (rsForm.EOF And rsForm.BOF) Then
        MsgBox "No records to append.", vbInformation, "Nothing to Append"
        GoTo Exit_Handler
    End If
    
    Do Until rsForm.EOF
        
        'check if data should be appended
        '-----------------------------------
        ' Append Data? --> Yes
        '-----------------------------------
        If rsForm![Append] = True Then
        
            'identify table
            strMain = rsForm![Table_Name]
            
            Set rsMain = db.OpenRecordset(strMain)
            
            strAppTableName = CompareTables(rsForm)
        
'            'identify table
'            strMain = rsForm![Table_Name]
'
'            Set rsMain = db.OpenRecordset(strMain)
'
'            'compare main vs. append table names by length
'            Dim iLength As Long
'            Dim iLength2 As Long
'            iLength = Len(rsForm![Table_Name])
'            iLength2 = iLength + 1
'
'            'Root of the Append Table name
'            Dim strAppTableName As String
'            strAppTableName = Right(Left(rsForm![Append_Table], iLength2), iLength)

            'Ensure Append Table is specified if Append is checked
            If rsForm![Append_Table] = "" Or IsNull(rsForm![Append_Table]) Then
                MsgBox "Make sure that you have properly selected all of the data you wish to append!", vbCritical, "Append Data"
                Exit Sub
            'Ensure Append Table Name matches the Main Table Name
            ElseIf strAppTableName <> rsForm![Table_Name] Then
                MsgBox "Make sure you have properly selected the data set to append to " & rsForm![Table_Name] & ".", vbCritical, "Append Data"
                Exit Sub
            End If
            
            strAppend = rsForm![Append_Table]
            Dim strAppendTableName As String
            strAppendTableName = strAppend

Debug.Print strMain

            Select Case strMain
                Case "tbl_Events"
                    'capture imported events table to use when appending new tree & sapling data
 '                   Dim rsEvents As DAO.Recordset
                    Set rsEvents = db.OpenRecordset(strAppend)
                
                Case "tbl_Locations"
                    'update locations table w/ newly collected slope & aspect
                    Dim rsLoc As DAO.Recordset
                    Set rsLoc = db.OpenRecordset(strAppend)
                    
                    'check if anything needs updating. If so, update it & return
                    fxnUpdateLocInfo rsLoc, strAppend
            
                Case "tbl_Tags"
                    'update other tables prior to appending new data
                    Set rsAppend = db.OpenRecordset(strAppend)
'FIX THIS!!! rsEvents is NOT created therefore ERROR
                    UpdateTags rsMain, rsAppend, rsEvents, strAppendTableName
                    GoTo NextRecord:
                
            End Select
            
            
     'Appending records to an existing Event_ID:
     'Determine if appending data from the Main Tablet or Secondary Tablet
            
            '-------------------------------------
            ' Secondary Tablet: replace Event_IDs
            '-------------------------------------
            If Me!optframe_Step1Append.value = 2 Then
                If Me!optframe_Step2Append.value = 2 Then
                    If Me!cmbo_Select_Event = "" Or IsNull(Me!cmbo_Select_Event) Then
                        MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
                        Me!cmbo_Select_Event.SetFocus
                        Exit Sub
                    ElseIf Me!cmbo_Select_Import_Event_Table = "" Or IsNull(Me!cmbo_Select_Import_Event_Table) Then
                            MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
                            Me!cmbo_Select_Import_Event_Table.SetFocus
                            Exit Sub
                    ElseIf Me!cmbo_Select_Import_Events = "" Or IsNull(Me!cmbo_Select_Import_Events) Then
                                MsgBox "You must complete the appending criteria", vbExclamation, "Append Data"
                                Me!cmbo_Select_Import_Events.SetFocus
                                Exit Sub
                    End If
                            
'                    ' Declare & set variables for the Event ID's in both Main (master) & imported datasets
'                    Dim GUIDMain As Variant 'GUID 'String
'                    Dim GUIDReplace As Variant 'GUID 'String
'
'                    GUIDMain = Me!cmbo_Select_Event.Column(0)
'                    GUIDReplace = Me!cmbo_Select_Import_Events.Column(0)
        
                    ' Check to see if the table contains an Event_ID field
                    Dim boolEvent As Boolean
                    boolEvent = False
                
                    Dim tdef As DAO.TableDef
                    Dim lCount As Long
                    Dim lCtr As Long
                    Dim strFieldName As String
                    Set tdef = db.TableDefs(strAppend)
            
                    With tdef
                        lCount = .Fields.Count
                            For lCtr = 0 To lCount - 1
                                strFieldName = .Fields(lCtr).Name
                                                  
                                If strFieldName = "Event_ID" Then
                                    boolEvent = True
                                    
                                    'move on since it's established an EventID is present
                                    Exit For
                                End If
             
                            Next
                    End With
Debug.Print strMain
                    'if the table contains an Event_ID pass the data set to the update event id function
                    If boolEvent = True Then
       
                        'Pass appending data set, master GUID & GUID that needs to be replaced
                        'to function to update the Event ID to the Master Event ID
                        Dim strUpdateEventIDSQL As String
                        Dim qdefUpdateEventID As QueryDef
                        Dim strTableName As String
'                        Dim strFindUpdate As String
                    
'                        strFindUpdate = GUIDReplace
                        strTableName = Me!cmbo_Append_Table.value
                    
                        ' Only select those records where the EventID = the GUID to be replaced
''                        strUpdateEventIDSQL = "SELECT [" & strAppend & "].* " _
''                        & "FROM [" & strAppend & "] " _
''                        & "WHERE ((([" & strAppend & "].Event_ID) = '" & strFindUpdate & "')); "
'
'                        strUpdateEventIDSQL = "SELECT [" & strAppend & "].* " _
'                        & "FROM [" & strAppend & "] " _
'                        & "WHERE ((([" & strAppend & "].Event_ID) = '" & GUIDReplace & "')); "
                      
                        strUpdateEventIDSQL = "SELECT [" & strAppend & "].* " _
                        & "FROM [" & strAppend & "] " _
                        & "WHERE ((([" & strAppend & "].Event_ID) = '" & TempVars("GUIDReplace") & "')); "
                      
                      
'Debug.Print "GUIDMain: " & GUIDMain
'Debug.Print "GUIDReplace: " & GUIDReplace

Debug.Print "TempVarGUIDMain: " & TempVars("GUIDMain")
Debug.Print "TempVarGUIDReplace: " & TempVars("GUIDReplace")

Debug.Print "strUpdateEventIDSQL: " & strUpdateEventIDSQL
                    
                        'ensure _Qry_UpdateEventID doesn't already exist, if it does archive it & delete it
                        If qryExists("_Qry_UpdateEventID") = True Then
                            'rename it (remember new name first, old last in DoCmd.Rename)
                            'DoCmd.Rename "_OLD_Qry_UpdateEventID_" & Format(Now, "YYYYMMDD_hhmmss"), acQuery, "_Qry_UpdateEventID"
                            DoCmd.DeleteObject acQuery, "_Qry_UpdateEventID"
                        End If
                        
                        ' save SQL as a qdef & create a recordset to pass to the UPDATE EVENT ID Function
                        Set qdefUpdateEventID = db.CreateQueryDef("_Qry_UpdateEventID", strUpdateEventIDSQL)
                        Set rsAppend = db.OpenRecordset("_Qry_UpdateEventID")
                        
                        'get accurate record count
                        If Not (rsAppend.BOF = True And rsAppend.EOF = True) Then
                            rsAppend.MoveLast
                            rsAppend.MoveFirst
                        End If
                    
                        If rsAppend.RecordCount > 0 Then
'                            UpdateEventID rsAppend, GUIDMain, GUIDReplace, strTableName
                            UpdateEventID rsAppend, TempVars("GUIDMain"), TempVars("GUIDReplace"), strTableName
                        End If
                    
                        ' Delete Select Query from the database after it has been used in the update event function
                        DoCmd.DeleteObject acQuery, qdefUpdateEventID.Name
        
        'Had to insert this code because when the record set was passed to the Append
        'function it was attempting to append too many records.
        'If events already exist additional information must be appended on a
        'Event by Event basis.
            
                        Dim strAppendSQL As String
                        Dim qdefAppend As QueryDef
            
'            'main event id
'                        Dim strFind As String
'                        strFind = GUIDMain
                        
            'query only the records that were collected on this event
            
'                        strAppendSQL = "SELECT [" & strAppend & "].* " _
'                        & " FROM [" & strAppend & "] " _
'                        & " WHERE ((([" & strAppend & "].Event_ID) = " & strFind & "));"
                    
                        strAppendSQL = "SELECT [" & strAppend & "].* " _
                        & " FROM [" & strAppend & "] " _
                        & " WHERE ((([" & strAppend & "].Event_ID) = '" & TempVars("GUIDMain") & "'));"
                    
                    
                        Set qdefAppend = db.CreateQueryDef("_Qry_AppendEventRecs", strAppendSQL)
Debug.Print "strAppendSQL = " & strAppendSQL
            'reset the rsAppend variable to equal the query that only contains records from
            'targeted event
            
                        Set rsAppend = db.OpenRecordset("_Qry_AppendEventRecs")
           ' MsgBox rsAppend.Name
            
            'pass the event specific recordset to the append function
            
                        AppendtoTable rsAppend, rsMain, strAppendTableName
                    
                  'Delete the select query once the append function has been completed.
                    
                        DoCmd.DeleteObject acQuery, qdefAppend.Name
                    
                    Else
                    
                        GoTo AppendData:
                    
                    End If
                
                Else
                    
                    GoTo AppendData:
                    
                End If
                
            Else
            
                '-------------------------------------
                ' Primary Tablet: use Event_IDs
                '-------------------------------------
                
                ' skip all of that crap about updating the event id and just append the damn records
        
AppendData:
        
                '----------------------------------------
                ' Appending tbl_Events? --> No
                '----------------------------------------
                If strAppTableName <> "tbl_Events" Then
                 
                    'Select only records associated with events being imported
             
                    'Check if table contains an Event_ID field
                 
                    'Dim boolEvent As Boolean
                    boolEvent = False
                    
                    'Dim tdef As DAO.TableDef
                    'Dim lCount As Long
                    'Dim lCtr As Long
                    'Dim strFieldName As String
                    Set tdef = db.TableDefs(strAppend)
                    
                    With tdef
                        lCount = .Fields.Count
                            For lCtr = 0 To lCount - 1
                                strFieldName = .Fields(lCtr).Name
                                                  
                                If strFieldName = "Event_ID" Then
                                    boolEvent = True
                                    'established Event_ID is present in table, so exit loop
                                    Exit For
                                End If
                    
                            Next
                    End With
                   
                    '------------------------------------------------------------------------------------
                    ' Event_ID field in Table?   Yes --> Only select data collected on imported events
                    '                            No  --> Send to standard append function
                    '------------------------------------------------------------------------------------
                    If boolEvent = True Then
                            
                        ' Yes --> Only select data collected on imported events
                        Dim strEvents As String, strSQL_FindImportedRecs As String
                        strEvents = rsEvents.Name
                        
                        strSQL_FindImportedRecs = "SELECT [" & strAppend & "].* FROM [" & strEvents & "] INNER JOIN [" & strAppend & "] " _
                                                & "ON [" & strEvents & "].Event_ID = [" & strAppend & "].Event_ID;"
  Debug.Print strSQL_FindImportedRecs
  
                        'convert SQL to query for following append functions as the Append data set
                        Dim qdef_NewEventRecs As QueryDef
                        Set qdef_NewEventRecs = db.CreateQueryDef("_qry_NewEventData", strSQL_FindImportedRecs)
                        
                        Set rsAppend = db.OpenRecordset("_qry_NewEventData")
                        
                        'Send the two recordsets (rsMain and rsAppend) to the Append Function
                        AppendtoTable rsAppend, rsMain, strAppendTableName
                        
                        'delete the query def so that it can be recreated as the code loops
                        DoCmd.DeleteObject acQuery, qdef_NewEventRecs.Name
                    
                    Else
                    
                        'No  --> Send to standard append function
                        Set rsAppend = db.OpenRecordset(strAppend)
                        
                        AppendtoTable rsAppend, rsMain, strAppendTableName
                    
                    End If
                    
                '----------------------------------------
                ' Appending tbl_Events? --> Yes
                '----------------------------------------
                Else
                
                    'Skip query if current append table is tbl_Events
                
                    'Run Append function for tbl_Events
                    Set rsAppend = db.OpenRecordset(strAppend)
                    
                    AppendtoTable rsAppend, rsMain, strAppendTableName
                
                End If
            
            End If
        
        'Make sure that the Events table is checked even if you do not need to append any data to it.
        '-----------------------------------
        ' Append Data? --> No, but include tbl_Events!
        '-----------------------------------
        ElseIf rsForm![Table_Name] = "tbl_Events" And rsForm![Append] = False Then
                    MsgBox "The Events table needs to be included in the append operation." & vbNewLine & vbNewLine & _
                    "Please go back and check the append box and select an imported events table.", , "Append Data"
                    Exit Sub
        Else
                
            GoTo NextRecord:
            
        End If
    
NextRecord:
    
        rsForm.MoveNext
    
    Loop
    
    MsgBox "Update and Appending complete!", , "Update and Append Data"
    
    'make the file selection visible
    Me.btnSelectFile.Visible = True
'CompareTables:
'    'identify table
'    strMain = rsForm![Table_Name]
'
'    Set rsMain = db.OpenRecordset(strMain)
'
'    'compare main vs. append table names by length
''    Dim iLength As Long
''    Dim iLength2 As Long
'    iLength = Len(rsForm![Table_Name])
'    iLength2 = iLength + 1
'
'    'Root of the Append Table name
''    Dim strAppTableName As String
'    strAppTableName = Right(Left(rsForm![Append_Table], iLength2), iLength)
'
'    Resume Next
    
CleanUp:
    Set rsAppend = Nothing
    Set rsMain = Nothing
    Set rsAppendLog = Nothing
    Set rsEvents = Nothing
    
    Set db = Nothing
    Set rsForm = Nothing
    Set qdefAppend = Nothing
    Set qdefUpdateEventID = Nothing
    Set qdef_NewEventRecs = Nothing

Exit_Handler:
    'delete _Qry_UpdateEventID
    Dim qdf As QueryDef
    For Each qdf In CurrentDb.QueryDefs
        If qdf.Name = "_Qry_UpdateEventID" Then
            DoCmd.DeleteObject acQuery, qdf.Name
            Exit For
        End If
    Next
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AppendFieldData[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Function:     CompareTables
' Description:  compare tables that are being appended with appendeee tables
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 25, 2019
' Adapted:      -
' Revisions:
'   BLC - 8/25/2019 - initial version
' ---------------------------------
Private Function CompareTables(rsForm As DAO.Recordset) As String
On Error GoTo Err_Handler
    
    Dim iLength As Long
    Dim iLength2 As Long
    Dim strAppTableName As String
    
    'compare main vs. append table names by length
'    Dim iLength As Long
'    Dim iLength2 As Long
    iLength = Len(rsForm![Table_Name])
    iLength2 = iLength + 1
    
    'Root of the Append Table name
'    Dim strAppTableName As String
    strAppTableName = Right(Left(rsForm![Append_Table], iLength2), iLength)

    CompareTables = strAppTableName

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CompareTables[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Sub:          ClearAppendTables
' Description:  clears tsys_Append_Tables & resets to default
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 27, 2019
' Adapted:      -
' Revisions:
'   BLC - 8/27/2019 - initial version
' ---------------------------------
Private Sub ClearAppendTables()
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim ctrlCombo As ComboBox
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tsys_Append_Tables")
    Set ctrlCombo = Me!cmbo_Append_Table
    Me!cmbo_Append_Table.RowSource = " "
    
    rs.MoveFirst
    
    Do While Not rs.EOF
        rs.Edit
        rs![Append_Table] = ""
        rs.Update
    rs.MoveNext
    Loop
    
    Set db = Nothing
    Set rs = Nothing

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ClearAppendTables[frm_Append_Append_Data form])"
    End Select
    Resume Exit_Handler
End Sub
