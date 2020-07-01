Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =11
    ItemSuffix =279
    Left =5925
    Top =2115
    Right =12120
    Bottom =7620
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x9733b3777046e540
    End
    RecordSource ="tbl_Events"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =2
            Bevel =1
            BackColor =-1
            BackThemeColorIndex =4
            BackTint =60.0
            OldBorderStyle =0
            BorderLineStyle =0
            BorderColor =-1
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverColor =0
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedColor =0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeColor =0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeColor =0
            PressedForeThemeColorIndex =1
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =5040
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =4320
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =275078
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Create New Event"
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =540
                    ThemeFontIndex =-1
                    BackThemeColorIndex =5
                    BackShade =50.0
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    Width =3360
                    Height =615
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="dirs"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =615
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =1125
                    Top =1320
                    Width =2475
                    Height =510
                    ColumnOrder =5
                    FontSize =18
                    FontWeight =700
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cbxLocationID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, tbl_Locations.Panel, "
                        "tbl_Locations.Frame, tbl_Locations.Unit_Code FROM tbl_Locations WHERE (((tbl_Loc"
                        "ations.Panel)=[Forms]![frm_Switchboard]![Panel])) ORDER BY tbl_Locations.Plot_Na"
                        "me;"
                    ColumnWidths ="0;2160"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1125
                    LayoutCachedTop =1320
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1830
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =1320
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblPlot"
                            Caption ="Plot"
                            FontName ="Franklin Gothic Book"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1320
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =1835
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =2160
                    Left =1125
                    Top =720
                    Width =2475
                    Height =510
                    ColumnOrder =6
                    FontSize =18
                    FontWeight =700
                    TabIndex =1
                    BorderColor =12632256
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxParkCode"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Unit Code\")) ORDER BY tlu_Enumerations.Enum_Code;"
                    ColumnWidths ="2160"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"\""
                    FontName ="Franklin Gothic Book"

                    LayoutCachedLeft =1125
                    LayoutCachedTop =720
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =1230
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeShade =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =720
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblPark"
                            Caption ="Park"
                            FontName ="Franklin Gothic Book"
                            LayoutCachedLeft =180
                            LayoutCachedTop =720
                            LayoutCachedWidth =1050
                            LayoutCachedHeight =1235
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1140
                    Top =1920
                    Width =2460
                    Height =510
                    ColumnOrder =4
                    FontSize =18
                    FontWeight =700
                    TabIndex =2
                    BorderColor =12632256
                    Name ="tbxEventDate"
                    Format ="Short Date"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Click in this field & use the date picker that appears to set the date"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =1920
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =2430
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =1920
                            Width =885
                            Height =510
                            FontSize =18
                            FontWeight =700
                            Name ="lblEventDate"
                            Caption ="Date"
                            FontName ="Franklin Gothic Book"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1920
                            LayoutCachedWidth =1065
                            LayoutCachedHeight =2430
                            ThemeFontIndex =-1
                            BackThemeColorIndex =-1
                            BorderThemeColorIndex =-1
                            BorderTint =100.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =93
                    Left =180
                    Top =2520
                    Width =3840
                    Height =1260
                    BackColor =13754087
                    BorderColor =10921638
                    Name ="rctPseudoEvent"
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4020
                    LayoutCachedHeight =3780
                    BackThemeColorIndex =-1
                    BackTint =40.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =300
                    Top =3060
                    Width =3600
                    Height =660
                    FontSize =8
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblHintPseudoEvent"
                    Caption ="Bush-hogged or other non-data collecting visit that may impact analysis"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Bush-hogged or other non-data collecting visit that may impact analysis"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =3060
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =3720
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin ToggleButton
                    Enabled = NotDefault
                    OverlapFlags =215
                    Left =360
                    Top =2640
                    Width =270
                    Height =299
                    ColumnOrder =2
                    TabIndex =3
                    Name ="tglPseudoEvent"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Bush-hogged or other non-data collecting visit record?"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =2640
                    LayoutCachedWidth =630
                    LayoutCachedHeight =2939
                    ForeTint =100.0
                    Shape =0
                    Bevel =0
                    Gradient =12
                    BackColor =8289145
                    BackTint =100.0
                    OldBorderStyle =1
                    BorderColor =8289145
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =80.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeTint =100.0
                    PressedForeThemeColorIndex =0
                    QuickStyle =23
                    QuickStyleMask =-5
                    WebImagePaddingLeft =4
                    WebImagePaddingTop =2
                    WebImagePaddingRight =4
                    WebImagePaddingBottom =7
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =720
                            Top =2640
                            Width =2160
                            Height =315
                            BorderColor =8355711
                            ForeColor =16711680
                            Name ="lblPseudoEvent"
                            Caption ="Pseudo Event?"
                            FontName ="Franklin Gothic Book"
                            ControlTipText ="Bush-hogged or other non-data collecting visit that may impact analysis"
                            GridlineColor =10921638
                            LayoutCachedLeft =720
                            LayoutCachedTop =2640
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =2955
                            ForeThemeColorIndex =-1
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3180
                    Top =2640
                    Width =720
                    Height =300
                    ColumnOrder =3
                    FontSize =9
                    TabIndex =4
                    BorderColor =8355711
                    ForeColor =255
                    Name ="tbxPseudoEvent"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =2640
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =2940
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =204
                    Left =420
                    Top =3840
                    Width =2325
                    Height =1080
                    FontSize =14
                    TabIndex =5
                    Name ="btnCreate"
                    Caption ="Create Event"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =420
                    LayoutCachedTop =3840
                    LayoutCachedWidth =2745
                    LayoutCachedHeight =4920
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =9226162
                    HoverThemeColorIndex =7
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontCharSet =204
                    Left =2820
                    Top =3840
                    Width =1020
                    Height =1080
                    FontSize =14
                    TabIndex =6
                    Name ="btnCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =2820
                    LayoutCachedTop =3840
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =4920
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =7775995
                    HoverThemeColorIndex =5
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3900
                    Top =540
                    Width =420
                    Height =300
                    ColumnOrder =0
                    FontSize =9
                    TabIndex =7
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxDevMode"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =3900
                    LayoutCachedTop =540
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =840
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ffffff00050000004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin Section
            Height =870
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =720
                    Height =216
                    ColumnWidth =4215
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Event_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"
                    FontName ="Franklin Gothic Book"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =780
                    LayoutCachedHeight =336
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =120
                    Width =720
                    Height =216
                    ColumnWidth =4200
                    TabIndex =2
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Location_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"
                    FontName ="Franklin Gothic Book"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =120
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =336
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =8
                    ListWidth =7200
                    Left =3180
                    Top =120
                    Width =720
                    Height =216
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =2171426
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="Protocol_Name"
                    ControlSource ="Protocol_Name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"prot"
                        "ocol\" ORDER BY Sort_Order; "
                    ColumnWidths ="2160;5040"
                    StatusBarText ="M. The name or code of the protocol governing the event (Protocol_Name)"
                    FontName ="Franklin Gothic Book"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22

                    LayoutCachedLeft =3180
                    LayoutCachedTop =120
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =336
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =510
                    Width =720
                    Height =210
                    TabIndex =5
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PseudoEvent"
                    ControlSource ="PseudoEvent"
                    StatusBarText ="Flag identifying non-visit events impacting plots (e.g. bushhogging or other sit"
                        "uation where a plot was not visited & more data was not collected, but informati"
                        "on (put in event notes) is known that may impact analysis of this plot)"
                    FontName ="Franklin Gothic Book"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =510
                    LayoutCachedWidth =780
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1620
                    Top =510
                    Width =720
                    Height =210
                    ColumnWidth =2520
                    TabIndex =7
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Event_Date"
                    ControlSource ="Event_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Event_Date)"
                    FontName ="Franklin Gothic Book"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =510
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =840
                    Top =120
                    Width =720
                    Height =216
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Entered_On_Tablet"
                    ControlSource ="Entered_On_Tablet"
                    StatusBarText ="Was field data collection done on tablet pc"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =840
                    LayoutCachedTop =120
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =336
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2400
                    Top =120
                    Width =720
                    Height =216
                    TabIndex =3
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Entered_By"
                    ControlSource ="Entered_By"
                    StatusBarText ="Contact ID of person creating event"
                    FontName ="Franklin Gothic Book"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2400
                    LayoutCachedTop =120
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =336
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3180
                    Top =510
                    Width =720
                    Height =210
                    TabIndex =9
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Entered_Date"
                    ControlSource ="Entered_Date"
                    Format ="Short Date"
                    StatusBarText ="Event creation date"
                    FontName ="Franklin Gothic Book"
                    InputMask ="99.99.0000;0;_"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =3180
                    LayoutCachedTop =510
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =840
                    Top =510
                    Width =720
                    Height =210
                    TabIndex =6
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Updated_By"
                    ControlSource ="Updated_By"
                    StatusBarText ="Contact ID of person updating this event"
                    FontName ="Franklin Gothic Book"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =840
                    LayoutCachedTop =510
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2400
                    Top =510
                    Width =720
                    Height =210
                    TabIndex =8
                    LeftMargin =44
                    TopMargin =22
                    RightMargin =44
                    BottomMargin =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Updated_Date"
                    ControlSource ="Updated_Date"
                    Format ="Short Date"
                    StatusBarText ="Data update date"
                    FontName ="Franklin Gothic Book"
                    InputMask ="99.99.0000;0;_"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =2400
                    LayoutCachedTop =510
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =720
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
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
' MODULE:       EventAdd
' Level:        Application module
' Version:      1.04
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/5/2018 - 1.01 - added documentation, error handling
'               BLC   - 10/23/2018 - 1.02 - added Form_Open event, PseudoEvent handling
'               BLC   - 3/18/2019 - 1.03 - accommodate calling form park code
'               BLC   - 4/16/2019 - 1.04 - revise from table form to allow record creation
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
        Me.lblTitle.Caption = m_Title
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
        Me.lblDirections.Caption = m_Directions
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

' ----------------
'  Events
' ----------------
' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
'   BLC - 3/18/2019  - accommodate passing park code from calling form
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
'    Me.CallingForm = "Main"
'
'    If Len(Me.OpenArgs) > 0 Then Me.CallingForm = Me.OpenArgs
'
'    'minimize calling form
'    ToggleForm Me.CallingForm, -1
    
    'dev mode
    tbxDevMode = DEV_MODE
                
    title = "Create New Event"
    'lblTitle.Caption = "" 'clear header title
    Directions = "dirs"
    
    ' open as new record
    DoCmd.GoToRecord Record:=acNewRec
Debug.Print "newrecord = " & Me.NewRecord

    'defaults
    rctPseudoEvent.BackColor = lngLtTan
    
    'disable until data allows
    cbxLocationID.Enabled = False
    tbxEventDate.Enabled = False
    tglPseudoEvent.Enabled = False
    btnCreate.Enabled = False
    
    'hints
    lblPseudoEvent.Caption = "Pseudo Event?"
    lblPseudoEvent.ForeColor = lngBlue
    lblPseudoEvent.ControlTipText = "Bush-hogged or other non-data collecting visit that may impact analysis"
    lblPseudoEvent.Visible = True
    lblHintPseudoEvent.Caption = "Bush-hogged or other non-data collecting visit that may impact analysis"
    lblHintPseudoEvent.ForeColor = lngBlue
    lblHintPseudoEvent.ControlTipText = "Bush-hogged or other non-data collecting visit that may impact analysis"
    lblHintPseudoEvent.Visible = True
    
    'set hover
    tglPseudoEvent.HoverColor = lngGreen
       
    'set park code
    If IsEmpty(Me.OpenArgs) = False And Me.OpenArgs <> "Choose Park" Then
        Me.cbxParkCode = Me.OpenArgs
        Me.cbxLocationID.Enabled = True
        SetPlots Nz(Me.OpenArgs, "")
    Else
        Me.cbxParkCode = Me.OpenArgs
        Me.cbxLocationID.Enabled = False
    End If
    
    'set entered/modified by to current user
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[EventAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxParkCode_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cboPark_Code > cbxParkCode
'               BLC   - 10/23/2018 - revised to avoid error #2448 "can't assign value to this object"
' ---------------------------------
Private Sub cbxParkCode_AfterUpdate()
On Error GoTo Err_Handler

    SetPlots cbxParkCode
'    Me.cbxLocationID.RowSource = "SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, " _
'            & "tbl_Locations.Panel, tbl_Locations.Frame, tbl_Locations.Unit_Code " _
'            & "FROM tbl_Locations " _
'            & "WHERE (((tbl_Locations.Panel) = [Forms]![frm_Switchboard]![Panel]) " _
'            & "AND ((tbl_Locations.Unit_Code) = '" & Me.cbxParkCode & "')) " _
'            & "ORDER BY tbl_Locations.Plot_Name;"
'
'    'enable plot
'    cbxLocationID.Enabled = True
'
'    'set focus on next field
'    cbxLocationID.SetFocus
'
'    'Me.cbxLocationID = Me.cbxLocationID.ItemData(0) #Error 2448 - can't assign value to this object

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxParkCode_AfterUpdate[EventAdd])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxLocationID_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:    BLC   - 10/23/2018 - initial version
' ---------------------------------
Private Sub cbxLocationID_AfterUpdate()
On Error GoTo Err_Handler

    'set record value
    Me.Location_ID = cbxLocationID
'    tbxRecordLocationID = cbxLocationID
    
    'set the location
'    tbxPlot = tbxRecordLocationID

    'check
    ReadyForSave
    
    'set focus on next field
    tbxEventDate.SetFocus
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxLocationID_AfterUpdate[EventAdd])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxEventDate_AfterUpdate
' Description:  Textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Private Sub tbxEventDate_AfterUpdate()
On Error GoTo Err_Handler

    'set record value
    Me.Event_Date = tbxEventDate
'    tbxRecordEventDate = tbxEventDate
    
'    'set the event date
'    tbxDate = tbxRecordEventDate
'
'    'check
    ReadyForSave
'
'    'set focus on button (vs. PseudoEvent)
'    btnCreate.SetFocus
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxEventDate_AfterUpdate[EventAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tglPseudoEvent_AfterUpdate
' Description:  Toggle button after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Private Sub tglPseudoEvent_AfterUpdate()
On Error GoTo Err_Handler

    'display as checkbox
    ToggleCaption tglPseudoEvent, True
    
    'set value for PseudoEvent
    Debug.Print "pse=" & tglPseudoEvent.value
    tbxPseudoEvent.value = CByte(Abs(tglPseudoEvent.value))
    Debug.Print "tbxpse=" & tbxPseudoEvent.value
    
    'set database value
    Me.PseudoEvent = CByte(Abs(tglPseudoEvent.value))
'    tbxRecordPseudoEvent.Value = CByte(Abs(tglPseudoEvent.Value))
    
    'check
    ReadyForSave
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglPseudoEvent_AfterUpdate[EventAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCreate_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cmdEvent_Create > btnCreate
'               BLC  - 10/23/2018 - added PseudoEvent handling
' ---------------------------------
Private Sub btnCreate_Click()
On Error GoTo Err_Handler

    'Save the new event if all of the needed information is provided, and open the Event form

    Dim strDocName As String
    Dim strLinkCriteria As String
    
    If IsNull(Me!cbxLocationID) Then
        MsgBox "You must select a location before you can enter record details!", _
            vbExclamation, "Enter Location First"
        Me!cbxLocationID.SetFocus
    Else
        If IsNull(Me!tbxEventDate) Then
            MsgBox "You must enter a date before you can enter record details!", _
                vbExclamation, "Enter Start Date"
            Me!tbxEventDate.SetFocus
        Else
            
'    'Generate string GUID for Event_ID
'    'If Me.NewRecord = True Then
'        If GetDataType("tbl_Events", "Event_ID") = dbText Then
''            Me!Event_ID = fxnGUIDGen
''            Me.tbxEID = Me!Event_ID
'            Me.tbxEID = fxnGUIDGen
'            Me.tbxEventID = Me.tbxEID
'        End If
'    'End If
            
Debug.Print "Dirty = " & Me.Dirty
Debug.Print "NewRec = " & Me.NewRecord
            
    If Me.Dirty = True Then
        DoCmd.RunCommand acCmdSaveRecord
    Else
        MsgBox "nothing to save"
    End If
'            DoCmd.RunCommand acCmdSaveRecord
        DoCmd.RunCommand acCmdSaveRecord
        
            'retrieve the EventID
'Debug.Print "eid = " & Me.tbxEID 'tbxEventID
Debug.Print "eid = " & Me.Event_ID

            strDocName = "frm_Events"
            strLinkCriteria = "[Event_ID]=" & "'" & Me.Event_ID & "'"
'            strLinkCriteria = "[Event_ID]=" & "'" & Me![tbxEventID] & "'"
Debug.Print strLinkCriteria
 '           DoCmd.OpenForm strDocName, , , strLinkCriteria, , , " (Creating)," & Me.tbxEID
            DoCmd.OpenForm strDocName, , , strLinkCriteria, , , "(Browsing)"
            
            DoCmd.Close acForm, Me.Name '"EventAdd"
'            DoCmd.Close acForm, "EventAdd", acSavePrompt
            'DoCmd.Close acForm, "frm_Event_Add"
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCreate_Click[EventAdd])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCancel_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cmdEvent_Cancel > btnCancel
' ---------------------------------
Private Sub btnCancel_Click()
On Error GoTo Err_Handler

    'Close the Create Event form without creating a record

'    If Me.Dirty Then Me.Undo
'    If Not Me.NewRecord Then
'        DoCmd.RunCommand acCmdDeleteRecord
'    End If
Debug.Print "Dirty = " & Me.Dirty
Debug.Print "NewRec = " & Me.NewRecord

    'remove new record if created
    If Me.Dirty Then Me.Undo
    If Not Me.NewRecord = True Then
        DoCmd.RunCommand acCmdDelete
    End If
    
    DoCmd.Close
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[EventAdd])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          ReadyForSave
' Description:  Check if form values are ready to save
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Public Sub ReadyForSave()
On Error GoTo Err_Handler

    Dim isOK As Boolean

    'default
    isOK = False
    
    If cbxLocationID.value > 0 Then tbxEventDate.Enabled = True
    If IsDate(tbxEventDate.value) Then tglPseudoEvent.Enabled = True

    If Len(Nz(cbxParkCode.value, "")) > 0 _
        And isGUID(cbxLocationID.value) = True _
        And IsDate(tbxEventDate.value) = True Then '_
        
        isOK = True
        
    End If
    
    'enable save button only for new sites (tbxID = 0)
'   If tbxID = 0 Then btnSave.Enabled = isOK
    
'    btnSubstrateCover.Enabled = IIf(tbxID.Value > 0, True, False)
'    btnSetObserverRecorder.Enabled = IIf(tbxID.Value > 0, True, False)
    
    'enable create if data is ok
    btnCreate.Enabled = isOK
    
    'refresh form
'    Me.Requery
   
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ReadyForSave[EventAdd form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          IsGUID
' Description:  Check if value is a valid GUID
' Assumptions:
'               GUID is 32 hex digits grouped into chunks of 8-4-4-4-12
'               Regex is
'                   "^(\{){0,1}[0-9a-fA-F]{8}\-" & _
'                   "[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-" & _
'                   "[0-9a-fA-F]{12}(\}){0,1}$"
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Torbis, January 16, 2007
'   http://www.vbforums.com/showthread.php?447414-Solved-Check-if-string-is-Guid
' Source/date:  Bonnie Campbell, October 23, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/23/2018 - initial version
' ---------------------------------
Public Function isGUID(strInspect As String) As Boolean
On Error GoTo Err_Handler

    Dim strPattern As String
    strPattern = "^(\{){0,1}[0-9a-fA-F]{8}\-" & _
                 "[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-" & _
                 "[0-9a-fA-F]{12}(\}){0,1}$"

    isGUID = IsRegExpMatch(strInspect, strPattern)
   
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsGUID[mod_Validation])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SetPlots
' Description:  filter plots by park code
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2019
' Adapted:      -
' Revisions:
'   BLC - 3/18/2019 - initial version
' ---------------------------------
Public Function SetPlots(ParkCode As String)
On Error GoTo Err_Handler
    
    Me.cbxLocationID.RowSource = "SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, " _
            & "tbl_Locations.Panel, tbl_Locations.Frame, tbl_Locations.Unit_Code " _
            & "FROM tbl_Locations " _
            & "WHERE (((tbl_Locations.Panel) = [Forms]![frm_Switchboard]![Panel]) " _
            & "AND ((tbl_Locations.Unit_Code) = '" & ParkCode & "')) " _
            & "ORDER BY tbl_Locations.Plot_Name;"

    'enable plot
    cbxLocationID.Enabled = True
    
    'set focus on next field
    cbxLocationID.SetFocus
    
    'Me.cbxLocationID = Me.cbxLocationID.ItemData(0) #Error 2448 - can't assign value to this object
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetPlots[EventAdd])"
    End Select
    Resume Exit_Handler
End Function
