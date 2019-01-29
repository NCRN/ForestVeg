Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11355
    DatasheetFontHeight =10
    ItemSuffix =64
    Left =285
    Top =495
    Right =13020
    Bottom =5880
    DatasheetGridlinesColor =12632256
    OrderBy ="Plot_Name"
    RecSrcDt = Begin
        0xd3905458b532e540
    End
    RecordSource ="qFrm_PseudoEvents"
    Caption ="Pseudo Events"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnGotFocus ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =1560
            BackColor =15921906
            Name ="FormHeader"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =3525
                    Top =1260
                    Width =630
                    Height =300
                    FontSize =12
                    Name ="lblUnitCode"
                    Caption ="Unit"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3525
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4155
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6315
                    Top =1260
                    Width =630
                    Height =300
                    FontSize =12
                    Name ="lblPanel"
                    Caption ="Panel"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6315
                    LayoutCachedTop =1260
                    LayoutCachedWidth =6945
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1680
                    Top =1260
                    Width =1665
                    Height =300
                    FontSize =12
                    Name ="lblEventDate"
                    Caption ="Sample Date*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1680
                    LayoutCachedTop =1260
                    LayoutCachedWidth =3345
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =8100
                    Top =1260
                    Width =780
                    Height =300
                    FontSize =12
                    Name ="lblEventYear"
                    Caption ="Year*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8100
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =120
                    Top =1260
                    Width =1485
                    Height =300
                    FontSize =12
                    Name ="lblPlotName"
                    Caption ="Plot Name*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1260
                    LayoutCachedWidth =1605
                    LayoutCachedHeight =1560
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =540
                    Top =720
                    Width =960
                    Height =300
                    ColumnOrder =1
                    FontSize =11
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cbxParkFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Nz([tbl_Locations]![Unit_Code],\"[Null]\") AS Expr1, Nz([Descrip"
                        "tion],\"[Null]\") AS Expr2 FROM tbl_Locations LEFT JOIN tlu_Units ON tbl_Locatio"
                        "ns.Unit_Code = tlu_Units.Unit_Code ORDER BY Nz([tbl_Locations]![Unit_Code],\"[Nu"
                        "ll]\");"
                    ColumnWidths ="864;4896"
                    StatusBarText ="Park code"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =540
                    LayoutCachedTop =720
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =735
                            Width =495
                            Height =285
                            FontSize =11
                            ForeColor =0
                            Name ="lblParkFilter"
                            Caption ="Unit:"
                            FontName ="Calibri"
                            LayoutCachedTop =735
                            LayoutCachedWidth =495
                            LayoutCachedHeight =1020
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =10140
                    Top =600
                    Width =1080
                    Height =240
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =6
                    Name ="tglFilter"
                    AfterUpdate ="[Event Procedure]"
                    Caption ="Filter Is On"
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the filter on or off"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =600
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4200
                    Top =720
                    Width =480
                    Height =300
                    ColumnOrder =2
                    FontSize =11
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cbxPanelFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT nz(Panel,\"[Null]\") FROM qfrm_Data_Gateway ORDER BY nz(Panel,\""
                        "[Null]\"); "
                    StatusBarText ="Panel"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =4200
                    LayoutCachedTop =720
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3480
                            Top =720
                            Width =660
                            Height =285
                            FontSize =11
                            ForeColor =0
                            Name ="lblPanelFilter"
                            Caption ="Panel:"
                            FontName ="Calibri"
                            LayoutCachedLeft =3480
                            LayoutCachedTop =720
                            LayoutCachedWidth =4140
                            LayoutCachedHeight =1005
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7140
                    Top =720
                    Width =840
                    Height =300
                    ColumnOrder =3
                    FontSize =11
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cbxYearFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT nz([Event_Year],\"[Null]\") FROM qfrm_Data_Gateway ORDER BY nz(["
                        "Event_Year],\"[Null]\");"
                    StatusBarText ="Year"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =7140
                    LayoutCachedTop =720
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6540
                            Top =720
                            Width =570
                            Height =285
                            FontSize =11
                            ForeColor =0
                            Name ="lblYearFilter"
                            Caption ="Year:"
                            FontName ="Calibri"
                            LayoutCachedLeft =6540
                            LayoutCachedTop =720
                            LayoutCachedWidth =7110
                            LayoutCachedHeight =1005
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =7013
                    Top =1260
                    Width =960
                    Height =300
                    FontSize =12
                    Name ="lblFrame"
                    Caption ="Frame"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7013
                    LayoutCachedTop =1260
                    LayoutCachedWidth =7973
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =4320
                    Top =1260
                    Width =990
                    Height =300
                    FontSize =12
                    Name ="lblUnitGroup"
                    Caption ="Unit Grp*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4320
                    LayoutCachedTop =1260
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5340
                    Top =1260
                    Width =960
                    Height =300
                    FontSize =12
                    Name ="lblSubunitCode"
                    Caption ="Subunit*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    LayoutCachedLeft =5340
                    LayoutCachedTop =1260
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =1560
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5460
                    Top =720
                    Width =1020
                    Height =300
                    ColumnOrder =5
                    FontSize =11
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cbxFrameFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT nz([Frame],\"[Null]\") AS Expr1 FROM qfrm_Data_Gateway ORDER BY "
                        "nz([Frame],\"[Null]\"); "
                    StatusBarText ="Panel"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =5460
                    LayoutCachedTop =720
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =4680
                            Top =720
                            Width =705
                            Height =285
                            FontSize =11
                            ForeColor =0
                            Name ="lblFrameFilter"
                            Caption ="Frame:"
                            FontName ="Calibri"
                            LayoutCachedLeft =4680
                            LayoutCachedTop =720
                            LayoutCachedWidth =5385
                            LayoutCachedHeight =1005
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =24
                    ListWidth =5616
                    Left =2520
                    Top =720
                    Width =960
                    Height =300
                    ColumnOrder =6
                    FontSize =11
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cbxUnitGroupFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Nz([tbl_Locations]![Unit_Group],\"[Null]\") AS Expr1, Nz([Descri"
                        "ption],\"[Null]\") AS Expr2, tlu_Unit_Group.Sort_Order FROM tbl_Locations LEFT J"
                        "OIN tlu_Unit_Group ON tbl_Locations.Unit_Group = tlu_Unit_Group.Unit_Group ORDER"
                        " BY tlu_Unit_Group.Sort_Order;"
                    ColumnWidths ="864;4752"
                    StatusBarText ="Park code"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =2520
                    LayoutCachedTop =720
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =1500
                            Top =735
                            Width =990
                            Height =285
                            FontSize =11
                            ForeColor =0
                            Name ="lblUnitGroupFilter"
                            Caption ="Unit Grp:"
                            FontName ="Calibri"
                            LayoutCachedLeft =1500
                            LayoutCachedTop =735
                            LayoutCachedWidth =2490
                            LayoutCachedHeight =1020
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =87
                    TextFontCharSet =204
                    Left =10140
                    Top =840
                    Width =1080
                    Height =240
                    FontWeight =700
                    TabIndex =7
                    Name ="btnClearFilter"
                    Caption ="Clear Filter"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Add a new location record"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =10140
                    LayoutCachedTop =840
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =1080
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =7620
                    Top =1260
                    Width =1500
                    Height =300
                    FontSize =12
                    Name ="lblLocationStatus"
                    Caption ="Status"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7620
                    LayoutCachedTop =1260
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =1560
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =8760
                    Top =720
                    Width =1200
                    Height =300
                    ColumnOrder =0
                    FontSize =11
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cbxStatusFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT DISTINCT Nz([Location_Status],\"[Null]\") AS Expr1 FROM qfrm_Data_Gateway"
                        " ORDER BY Nz([Location_Status],\"[Null]\");"
                    StatusBarText ="Year"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =720
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =1020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8040
                            Top =720
                            Width =690
                            Height =285
                            FontSize =11
                            ForeColor =0
                            Name ="lblStatusFilter"
                            Caption ="Status:"
                            FontName ="Calibri"
                            LayoutCachedLeft =8040
                            LayoutCachedTop =720
                            LayoutCachedWidth =8730
                            LayoutCachedHeight =1005
                            ForeThemeColorIndex =0
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =-30
                    Width =11385
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Pseudo Events"
                    FontName ="Calibri"
                    LayoutCachedLeft =-30
                    LayoutCachedWidth =11355
                    LayoutCachedHeight =540
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Top =1200
                    Width =11340
                    Name ="lnSeparatorTop"
                    GridlineColor =10921638
                    LayoutCachedTop =1200
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =1200
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =120
                    Top =30
                    Width =2220
                    Height =480
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    Name ="btnAddEvent"
                    Caption ="Create New Event"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Create a new event..."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =30
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =2
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =9226162
                    HoverThemeColorIndex =7
                    HoverTint =60.0
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
                    OverlapFlags =247
                    Left =10140
                    Top =30
                    Width =1140
                    Height =480
                    FontSize =12
                    FontWeight =700
                    TabIndex =9
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10140
                    LayoutCachedTop =30
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =510
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =2
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =7775995
                    HoverThemeColorIndex =5
                    HoverTint =60.0
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
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7200
                    Top =60
                    Width =900
                    Height =420
                    FontSize =14
                    FontWeight =700
                    TabIndex =10
                    ForeColor =13421823
                    Name ="tbxFilteredRecordCount"
                    ControlSource ="=Nz(Count([Plot_Name]),0)"
                    FontName ="Calibri"

                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =480
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =215
                            TextAlign =2
                            Left =8160
                            Top =180
                            Width =900
                            Height =300
                            FontSize =10
                            ForeColor =13421823
                            Name ="lblRecordsSelected"
                            Caption ="events"
                            FontName ="Calibri"
                            LayoutCachedLeft =8160
                            LayoutCachedTop =180
                            LayoutCachedWidth =9060
                            LayoutCachedHeight =480
                        End
                    End
                End
            End
        End
        Begin Section
            Height =360
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Width =10740
                    Height =360
                    FontSize =13
                    TabIndex =12
                    Name ="tbxPseudoEvent"
                    ControlSource ="PseudoEvent"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000002000000000000000200000001000000 ,
                        0x00000000ffcdcd00000000000200000003000000050000000100000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000300000000000
                    End

                    LayoutCachedWidth =10740
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000ffcdcd00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =5280
                    Width =420
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="txtEvent_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="Name of the location"
                    FontName ="Calibri"

                    LayoutCachedLeft =5280
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Width =1740
                    Height =300
                    ColumnWidth =1710
                    FontSize =13
                    TabIndex =1
                    ForeColor =16711680
                    Name ="tbxEventDate"
                    ControlSource ="Event_Date"
                    Format ="dd-mmm-yyyy"
                    StatusBarText ="Start date of the sampling event"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ShowDatePicker =0

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8100
                    Width =780
                    Height =300
                    ColumnWidth =600
                    FontSize =13
                    TabIndex =6
                    Name ="tbxEventYear"
                    ControlSource ="Event_Year"
                    FontName ="Calibri"

                    LayoutCachedLeft =8100
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3480
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    FontSize =13
                    TabIndex =2
                    Name ="tbxUnitCode"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Unit code"
                    FontName ="Calibri"

                    LayoutCachedLeft =3480
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6300
                    Width =660
                    Height =300
                    ColumnWidth =2310
                    FontSize =13
                    TabIndex =5
                    ForeColor =0
                    Name ="tbxPanel"
                    ControlSource ="Panel"
                    StatusBarText ="Sample location"
                    FontName ="Calibri"

                    LayoutCachedLeft =6300
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =360
                    Width =420
                    FontSize =10
                    TabIndex =7
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Name of the location"
                    FontName ="Calibri"

                    LayoutCachedLeft =360
                    LayoutCachedWidth =780
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Width =1500
                    Height =300
                    FontSize =13
                    ForeColor =16711680
                    Name ="tbxPlotName"
                    ControlSource ="Plot_Name"
                    StatusBarText ="Name of the location"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7020
                    Width =1020
                    Height =300
                    FontSize =13
                    TabIndex =9
                    Name ="tbxFrame"
                    ControlSource ="Frame"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =7020
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4380
                    Width =900
                    Height =300
                    FontSize =13
                    TabIndex =3
                    Name ="tbxUnitGroup"
                    ControlSource ="Unit_Group"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =4380
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5340
                    Width =900
                    Height =300
                    FontSize =13
                    TabIndex =4
                    Name ="tbxSubunitCode"
                    ControlSource ="Subunit_Code"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =5340
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7560
                    Width =1560
                    Height =300
                    FontSize =13
                    TabIndex =10
                    Name ="tbxLocationStatus"
                    ControlSource ="Location_Status"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =7560
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =300
                End
                Begin ToggleButton
                    OverlapFlags =85
                    Left =10860
                    Top =60
                    Width =240
                    Height =240
                    TabIndex =11
                    Name ="tglPseudoEvent"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Toggle pseudoevent"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10860
                    LayoutCachedTop =60
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =2
                    Gradient =12
                    BackColor =14262536
                    BackThemeColorIndex =6
                    BorderColor =14262536
                    BorderThemeColorIndex =6
                    HoverColor =16236067
                    HoverThemeColorIndex =6
                    HoverTint =80.0
                    PressedColor =11436294
                    PressedThemeColorIndex =6
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =25
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =708
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2760
                    Top =180
                    Width =6000
                    Height =528
                    FontSize =10
                    BackColor =16777215
                    ForeColor =0
                    Name ="lblOverview"
                    Caption ="Click on a Plot Name to open location or the Sample Date to open event.\015\012*"
                        "Double-click on the field label to change sort order.  "
                    FontName ="Calibri"
                    ControlTipText ="View mode"
                    LayoutCachedLeft =2760
                    LayoutCachedTop =180
                    LayoutCachedWidth =8760
                    LayoutCachedHeight =708
                End
                Begin Line
                    BorderWidth =1
                    OverlapFlags =85
                    Top =60
                    Width =11340
                    Name ="lnSeparatorBottom"
                    GridlineColor =10921638
                    LayoutCachedTop =60
                    LayoutCachedWidth =11340
                    LayoutCachedHeight =60
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
' FORM:    PseudoEventList form
' Level:        Application module
' Version:      1.05
'
' Description:  form related functions & procedures
'
' Data source:  qPseudoEventList form
' Data access:  view and delete records (delete by cmdDeleteRec)
' Pages:        none
' Functions:    fxnSortRecords, FilterGateway, FilterString, WriteRecordCriteria
' References:   fxnSwitchboardIsOpen
'
' Source/date:  John R. Boetsch, June 7, 2006
' Revisions:    JRB - 6/7/2006 - 1.00 - initial version
'               Simon Kingston, 9/2006 - 1.01 -  added CorrectText calls where strings were being used in criteria
'                                             - updated cmdDeleteRec_Click() event to use appropriate criteria depending on primary key
'               Simon Kingston, 12/2006-1/2007 - 1.02 - added filters to the top of the form and changed toggle button to text caption
'               MEL/GS - unknown - 1.03 - adapted for NCRN
'               BLC - 5/23/2018  - 1.04 - added documentation/error handling
'               BLC - 11/9/2018  - 1.05 - update to handle Pseudoevents
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
Dim strSortField As String    ' Keeps track of current sort settings
Dim strSortOrder As String
Dim strSortFieldLabel As String
Dim strCurrentRecordCriteria As String

' ---------------------------------
'  Properties
' ---------------------------------


' ----------------
'  Events
' ----------------

' ----------------
'  Form
' ----------------
' ---------------------------------
' SUB:          Form_Open
' Description:  form open actions
' Assumptions:  -
' Parameters:   Cancel - whether open action(s) should be cancelled (boolean)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/23/2018 - update documentation, error handling
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim varReturn As Variant

    ' On opening the form, set the initial sort order
    strSortFieldLabel = "lblPlot_Name"
    varReturn = SortRecords("Plot_Name")
    ' Set the filter
    If fxnSwitchboardIsOpen Then
        'Not currently choosing to select default filter
        'Me!cboPanelFilter = Forms!frm_Switchboard!cPanel
        Me.FilterGateway (True)
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_GotFocus
' Description:  form open actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/23/2018 - update documentation, error handling
' ---------------------------------
Private Sub Form_GotFocus()
On Error GoTo Err_Handler

    Dim rst As DAO.Recordset
    
    'return to same record when coming back to Data Gateway from another form
    If Not IsNothing(strCurrentRecordCriteria) Then
        Set rst = Me.RecordsetClone
        rst.FindFirst strCurrentRecordCriteria
        Me.Bookmark = rst.Bookmark
        Set rst = Nothing
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_GotFocus[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click Events
' ----------------

' ----------------
'  Filters
' ----------------

'CODE TO UPDATE FILTER WHEN ON-SCREEN SELECTIONS CHANGE

' ---------------------------------
' SUB:          cbxParkFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub cbxParkFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxParkFilter_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxUnitGroupFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub cbxUnitGroupFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxUnitGroupFilter_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxPanelFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub cbxPanelFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPanelFilter_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxFrameFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub cbxFrameFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxFrameFilter_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxYearFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub cbxYearFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxYearFilter_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxStatusFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub cbxStatusFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxStatusFilter_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglPseudoEvent_AfterUpdate
' Description:  toggle button after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub tglPseudoEvent_AfterUpdate()
On Error GoTo Err_Handler

    'tglPseudoEvent.Value = Not tglPseudoEvent.Value
    
    'set pseudoevent field
    tbxPseudoEvent = IIf(tglPseudoEvent.Value = True, 1, 0)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglPseudoEvent_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lblPlotName_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub lblPlotName_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Plot_Name")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblPlot_Name_DblClick[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

'CODE TO RESORT RECORDS WHEN HEADING IS CLICKED

Private Sub lblPlot_Name_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    'fxnSortRecords ("Plot_Name")
    SortRecords ("Plot_Name")

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          lblUnitGroup_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub lblUnitGroup_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Unit_Group")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblUnitGroup_DblClick[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub lblUnit_Group_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    'fxnSortRecords ("Unit_Group")
    SortRecords ("Unit_Group")

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          lblSubunitCode_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub lblSubunitCode_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Subunit_Code")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblSubunitCode_DblClick[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub lblSubunit_Code_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    'fxnSortRecords ("Subunit_Code")
    SortRecords ("Subunit_Code")

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          lblEventYear_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub lblEventYear_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Event_Year")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblEventYear_DblClick[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub lblEvent_Year_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    'fxnSortRecords ("Event_Year")
    SortRecords ("Event_Year")

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          lblEventDate_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub lblEventDate_DblClick(Cancel As Integer)
On Error GoTo Err_Handler
    
    SortRecords ("Event_Date")
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblEventDate_DblClick[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

'Private Sub lblEvent_Date_DblClick(Cancel As Integer)
'    On Error GoTo Err_Handler
'
'    fxnSortRecords ("Event_Date")
'
'Exit_Procedure:
'    Exit Sub
'Err_Handler:
'    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
'    Resume Exit_Procedure
'End Sub
' ---------------------------------
' SUB:          btnClearFilter_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnClearFilter_Click()
On Error GoTo Err_Handler

    Me!cbxParkFilter = Null
    Me!cbxUnitGroupFilter = Null
    Me!cbxPanelFilter = Null
    Me!cbxFrameFilter = Null
    Me!cbxYearFilter = Null
    Me!cbxStatusFilter = Null
    Me.Filter = ""
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClearFilter_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cmdClearFilter_Click()
    On Error GoTo Err_Handler
   
    Me!cboParkFilter = Null
    Me!cboUnitGroupFilter = Null
    Me!cboPanelFilter = Null
    Me!cboFrameFilter = Null
    Me!cboYearFilter = Null
    Me!cboStatusFilter = Null
    Me.Filter = ""
    
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          btnAddLocation_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnAddLocation_Click()
On Error GoTo Err_Handler
    
    MsgBox "This function is being developed"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddLocation_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cmdNewLocation_Click()
    MsgBox "This function is being developed"
End Sub

' ---------------------------------
' SUB:          btnAddEvent_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub btnAddEvent_Click()
On Error GoTo Err_Handler
    
    'record what the current record is so we can go back to that record on return
    WriteRecordCriteria
    DoCmd.Close acForm, "PseudoEventList form"
    DoCmd.OpenForm "EventAdd"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddEvent_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnClose_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilter_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub tglFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    Me.FilterGateway (Me!tglFilter)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilter_AfterUpdate[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxEventDate_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub tbxEventDate_Click()
On Error GoTo Err_Handler
    
    Dim strCriteriaLoc As String
    Dim strCriteriaEvent As String

    'Record what the current record is so we can go back to that record on return
    WriteRecordCriteria
    
    'NCRN NOTE: For this database, we will not create new events through this mechanism.
    'It is unclear to me how to use this mechanism to create a second event for a location (mel).
    
    'If there is not an event id, add a new data entry record
    'If IsNull(Me!txtEvent_ID) Then
    '            DoCmd.OpenForm "frm_Events", , , , acFormAdd, , "New record"
    '    If Not IsNull(Me!txtLocation_ID) Then
    '        ' Fill in Location
    '        Forms!frm_Events!cboLocation_ID = Me!txtLocation_ID
    '        Forms!frm_Events.Update_Loc_Info
    '    End If
    'if there is an event id, bring up the selected data entry record
    'Else
        'strCriteriaLoc = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        strCriteriaEvent = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
        ' Filter by location and event
        DoCmd.OpenForm "frm_Events", , , strCriteriaEvent, , , "(Browsing)"
    'End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxEventDate_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxPlotName_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub tbxPlotName_Click()
On Error GoTo Err_Handler
    
    Dim strCriteria As String

    'record what the current record is so we can go back to that record on return
    If Not IsNothing(Me!Location_ID) Then
        WriteRecordCriteria
        strCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Locations", , , strCriteria, , , "Filter by location"
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxPlotName_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxReportTrigger_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub tbxReportTrigger_Click()
On Error GoTo Err_Handler
    
    Dim strDocName As String
    Dim strCriteria As String
    
    If IsNothing(Me!Event_ID) Then
        'Trap records that do not contain an event.
        MsgBox ("This Record is not linked to an Event.  Please choose another Record.")
        GoTo Exit_Handler
    Else
        'Record what the current record is so we can go back to that record on return
        WriteRecordCriteria
        strDocName = "rpt_Event_Summary_Unfiltered"
        strCriteria = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
        'DoCmd.OpenReport stDocName, acPreview, "qRpt_Event_Summary_Unfiltered", stCriteria
        DoCmd.OpenReport strDocName, acPreview, , strCriteria
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxReportTrigger_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxViewPhotos_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, November 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 11/9/2018 - initial version
' ---------------------------------
Private Sub tbxViewPhotos_Click()
On Error GoTo Err_Handler
    
    Dim RetVal As Double
    Dim RootFolder As String
    Dim PhotoFolder As String
    
    RootFolder = "T:\I&M"
    PhotoFolder = "T:\I&M\Monitoring\Forest_Vegetation\Photos\"
    If FolderExists(PhotoFolder & Me!txtPlot_Name) Then
        RetVal = Shell("explorer /e,/root, " & PhotoFolder & Me!txtPlot_Name, vbNormalFocus)
        GoTo Exit_Handler
    Else
        If FolderExists(RootFolder) Then
            MsgBox ("Folder for this plot not found....Opening the root of the Photos folder.")
            RetVal = Shell("explorer /e,/root, " & PhotoFolder, vbNormalFocus)
            GoTo Exit_Handler
        Else
            MsgBox ("The network appears to be unavailable. Network access is required to view photos.")
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxViewPhotos_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Methods
' ---------------------------------

' ---------------------------------
' SUB:          FilterString
' Description:  Builds filter string for the Data Gateway form
' Assumptions:  -
' Parameters:   Val - filter control value (variant)
'               FieldName - field being filtered (string)
'               CurrentFilter - current filter value (string)
' Returns:      Filter string or null if no filter built yet
' Throws:       none
' References:   -
' Source/date:  Simon Kingston, 1/17/2007
'               Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   SK - 1/17/2007 - initial version
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Function FilterString(val As Variant, FieldName As String, CurrentFilter As Variant) As Variant
On Error GoTo Err_Handler

    Const cstrNull As String = "[Null]"
    Dim Filter As Variant

    If IsNull(val) Then
        Filter = CurrentFilter
    Else
        Filter = (CurrentFilter + " AND ") & FieldName
        If val = cstrNull Then
            Filter = Filter & " Is Null"
        Else
        If IsNumeric(val) Then
            Filter = Filter & "=" & val & ""
            Else
            If IsDate(val) Then
                Filter = Filter & "=#" & val & "#"
                Else
                    Filter = Filter & "=" & CorrectText(CStr(val))
                End If
            End If
        End If
    End If
    
    FilterString = Filter

    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterString[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          FilterGateway
' Description:  filters gateway form
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  Simon Kingston, 1/17/2007
'               Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   SK - 1/17/2007 - initial version
'   MEL/GS - unknown - NCRN version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Public Sub FilterGateway(FilterOn As Boolean)
On Error GoTo Err_Handler
    
    Dim Filter As Variant

    Filter = Null
    
    Me!tglFilter = FilterOn
    
    If FilterOn Then
        Me!tglFilter.Caption = "Filter Is On"
    
        'add park filter to filter string
        Filter = FilterString(Me!cbxParkFilter, "Unit_Code", Filter)
        'add unit filter to filter string
        Filter = FilterString(Me!cbxUnitGroupFilter, "Unit_Group", Filter)
        'add panel filter to filter string
        Filter = FilterString(Me!cbxPanelFilter, "Panel", Filter)
        'add frame filter to filter string
        Filter = FilterString(Me!cbxFrameFilter, "Frame", Filter)
        'add year filter to filter string
        Filter = FilterString(Me!cbxYearFilter, "Event_Year", Filter)
        'add status filter to filter string
        Filter = FilterString(Me!cbxStatusFilter, "Location_Status", Filter)
        Me.Filter = Nz(Filter)
    Else
        Me!tglFilter.Caption = "Filter Is Off"
    End If
    Me.FilterOn = FilterOn
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterGateway[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          WriteRecordCriteria
' Description:  Records Location & Event IDs of the current record so that it can be made the current record when coming
'               back to the form from another form (=bookmark).
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   GetCriteriaString
' Source/date:  Simon Kingston, 1/17/2007
'               Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   SK - 1/17/2007 - initial version
'   MEL/GS - unknown - initial NCRN version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub WriteRecordCriteria()
On Error GoTo Err_Handler

    If Not IsNothing(Me!Location_ID) Then
        strCurrentRecordCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        If IsNothing(Me!Event_ID) Then
            strCurrentRecordCriteria = strCurrentRecordCriteria & " AND Event_ID Is Null"
        Else
            strCurrentRecordCriteria = strCurrentRecordCriteria & " AND " & GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - WriteRecordCriteria[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     SortRecords
' Description:  sorts records by desired field
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  John R. Boetsch, May 5, 2006
'               Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling,
'                     renamed from fxnSortRecords
' ---------------------------------
Private Function SortRecords(ByVal strFieldName As String, _
    Optional ByVal strField2Name As String)
On Error GoTo Err_Handler
    
    Dim strOrderBy As String

    ' If already sorting in ascending order by this field, sort descending
    If strFieldName = strSortField And strSortOrder = "" Then
        strSortOrder = " DESC"
    Else: strSortOrder = ""
    End If
    
    ' Create the order by string and activate the filter
    strOrderBy = strFieldName & strSortOrder
    If strField2Name <> "" Then
        strOrderBy = strField2Name & " DESC, " & strOrderBy
    End If
    strSortField = strFieldName
    Me.Form.OrderBy = strOrderBy
    Me.Form.OrderByOn = True

    ' Change the label format to indicate the sorted field
    strSortFieldLabel = "lbl" & Replace(strFieldName, "_", "")
    With Me.Controls.Item(strSortFieldLabel)
        .FontItalic = IIf(.FontItalic = False, True, False)
        .FontBold = IIf(.FontBold = False, True, False)
    
'        .FontItalic = False
'        .fontBold = False

'        .FontItalic = True
'        .fontBold = True
    End With
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortRecords[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GoToForm
' Description:  open desired form
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 24, 2018
' Adapted:      -
' Revisions:
'   BLC - 5/24/2018 - initial version
' ---------------------------------
Public Function GoToForm(frm As String)
On Error GoTo Err_Handler
    
    WriteRecordCriteria
    
    If DbObjectExists(frm, "frm") Then _
        DoCmd.OpenForm frm, acNormal
        
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnGoToTags_Click[PseudoEventList form])"
    End Select
    Resume Exit_Handler
End Function
