Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11580
    DatasheetFontHeight =10
    ItemSuffix =68
    Left =15
    Top =645
    Right =11595
    Bottom =6030
    DatasheetGridlinesColor =12632256
    OrderBy ="Plot_Name"
    RecSrcDt = Begin
        0x0f463b98b308e440
    End
    RecordSource ="qfrm_Data_Gateway"
    Caption ="Location and Event Data Gateway"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnGotFocus ="[Event Procedure]"
    AllowDatasheetView =0
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
                    Left =4245
                    Top =1260
                    Width =630
                    Height =300
                    FontSize =12
                    Name ="lblUnitCode"
                    Caption ="Unit"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4245
                    LayoutCachedTop =1260
                    LayoutCachedWidth =4875
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =7035
                    Top =1260
                    Width =630
                    Height =300
                    FontSize =12
                    Name ="lblPanel"
                    Caption ="Panel"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7035
                    LayoutCachedTop =1260
                    LayoutCachedWidth =7665
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2040
                    Top =1260
                    Width =1665
                    Height =300
                    FontSize =12
                    Name ="lblEventDate"
                    Caption ="Sample Date*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =1260
                    LayoutCachedWidth =3705
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =8820
                    Top =1260
                    Width =780
                    Height =300
                    FontSize =12
                    Name ="lblEventYear"
                    Caption ="Year*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8820
                    LayoutCachedTop =1260
                    LayoutCachedWidth =9600
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
                    OverlapFlags =85
                    TextAlign =2
                    Left =7733
                    Top =1260
                    Width =960
                    Height =300
                    FontSize =12
                    Name ="lblFrame"
                    Caption ="Frame"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7733
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8693
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5040
                    Top =1260
                    Width =990
                    Height =300
                    FontSize =12
                    Name ="lblUnitGroup"
                    Caption ="Unit Grp*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5040
                    LayoutCachedTop =1260
                    LayoutCachedWidth =6030
                    LayoutCachedHeight =1560
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =6060
                    Top =1260
                    Width =960
                    Height =300
                    FontSize =12
                    Name ="lblSubunitCode"
                    Caption ="Subunit*"
                    FontName ="Calibri"
                    OnDblClick ="[Event Procedure]"
                    LayoutCachedLeft =6060
                    LayoutCachedTop =1260
                    LayoutCachedWidth =7020
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
                    OverlapFlags =85
                    TextAlign =2
                    Left =9720
                    Top =1260
                    Width =1500
                    Height =300
                    FontSize =12
                    Name ="lblLocationStatus"
                    Caption ="Status"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9720
                    LayoutCachedTop =1260
                    LayoutCachedWidth =11220
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
                    Left =-15
                    Width =11355
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="lblEvent_Form_Header"
                    Caption ="Data Gateway"
                    FontName ="Calibri"
                    LayoutCachedLeft =-15
                    LayoutCachedWidth =11340
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
                Begin Label
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =9840
                    Top =105
                    Width =1320
                    Height =345
                    FontSize =12
                    FontWeight =600
                    BorderColor =52479
                    ForeColor =16776960
                    Name ="lblQCMode"
                    Caption ="QC MODE"
                    LayoutCachedLeft =9840
                    LayoutCachedTop =105
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =450
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
                    Width =11220
                    Height =360
                    FontSize =13
                    TabIndex =14
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

                    LayoutCachedWidth =11220
                    LayoutCachedHeight =360
                    ConditionalFormat14 = Begin
                        0x01000200000000000000020000000100000000000000ffcdcd00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ffffff000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =255
                    Left =3840
                    Width =300
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =12
                    Name ="btnViewReport"
                    Caption ="Browse PLANTS"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddddddddddddddddddddd0000000007ddd ,
                        0xdd0f066666660dddddd7066666660ddddd0f066666660dddddd7066666660ddd ,
                        0xdd0f066666660dddddd7060000060ddddd0f060fff060dddddd7060000060ddd ,
                        0xdd0f066666660dddddd0000000007ddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View the Summary Report for this Event"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =3840
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =2
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
                    OverlapFlags =255
                    Left =1680
                    Width =300
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =11
                    Name ="btnViewPhotos"
                    Caption ="Browse PLANTS"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddd7007ddddddd70000000000007dd07700777700770dd07707777770770d ,
                        0xd07707877770770dd07707e87770770dd0ff00777700ff0dd0fff000000fff0d ,
                        0xd00000000000000ddd00d70ff07d00dddddddd7007dddddddddddddddddddddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview the plot photos"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =300
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =2
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
                    Left =2040
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

                    LayoutCachedLeft =2040
                    LayoutCachedWidth =3780
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
                    Left =8820
                    Width =780
                    Height =300
                    ColumnWidth =600
                    FontSize =13
                    TabIndex =6
                    Name ="tbxEventYear"
                    ControlSource ="Event_Year"
                    FontName ="Calibri"

                    LayoutCachedLeft =8820
                    LayoutCachedWidth =9600
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
                    Left =4200
                    Width =840
                    Height =300
                    ColumnWidth =2310
                    FontSize =13
                    TabIndex =2
                    Name ="tbxUnitCode"
                    ControlSource ="Unit_Code"
                    StatusBarText ="Unit code"
                    FontName ="Calibri"

                    LayoutCachedLeft =4200
                    LayoutCachedWidth =5040
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
                    Left =7020
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

                    LayoutCachedLeft =7020
                    LayoutCachedWidth =7680
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
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7740
                    Width =1020
                    Height =300
                    FontSize =13
                    TabIndex =9
                    Name ="tbxFrame"
                    ControlSource ="Frame"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =7740
                    LayoutCachedWidth =8760
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
                    Left =5100
                    Width =900
                    Height =300
                    FontSize =13
                    TabIndex =3
                    Name ="tbxUnitGroup"
                    ControlSource ="Unit_Group"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =5100
                    LayoutCachedWidth =6000
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
                    Left =6060
                    Width =900
                    Height =300
                    FontSize =13
                    TabIndex =4
                    Name ="tbxSubunitCode"
                    ControlSource ="Subunit_Code"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =6060
                    LayoutCachedWidth =6960
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
                    Left =9660
                    Width =1560
                    Height =300
                    FontSize =13
                    TabIndex =10
                    Name ="tbxLocationStatus"
                    ControlSource ="Location_Status"
                    StatusBarText ="The name or code of the protocol governing the event"
                    FontName ="Calibri"

                    LayoutCachedLeft =9660
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =300
                End
                Begin ToggleButton
                    OverlapFlags =119
                    Left =11280
                    Top =45
                    Width =240
                    Height =240
                    TabIndex =13
                    Name ="tglPseudoEvent"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Toggle pseudoevent"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =11280
                    LayoutCachedTop =45
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =285
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
            Height =960
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =10080
                    Top =180
                    Width =1140
                    Height =600
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10080
                    LayoutCachedTop =180
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =780
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
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1260
                    Top =180
                    Width =1140
                    Height =600
                    FontSize =12
                    FontWeight =700
                    Name ="btnGoToPlants"
                    Caption ="Browse PLANTS"
                    OnClick ="=GoToForm(\"frm_Plants\")"
                    FontName ="Calibri"
                    ControlTipText ="Browse PLANT species"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =1260
                    LayoutCachedTop =180
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =2
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
                    OverlapFlags =87
                    Left =60
                    Top =180
                    Width =1140
                    Height =600
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="btnGoToTags"
                    Caption ="Browse TAGS"
                    OnClick ="=GoToForm(\"frm_Tags\")"
                    FontName ="Calibri"
                    ControlTipText ="Browse TAGs"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =2
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
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9060
                    Top =120
                    Width =900
                    Height =420
                    ColumnOrder =4
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="tbxFilteredRecordCount"
                    ControlSource ="=nz(Count([Plot_Name]),0)"
                    FontName ="Calibri"

                    LayoutCachedLeft =9060
                    LayoutCachedTop =120
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =540
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =93
                            TextAlign =2
                            Left =9060
                            Top =420
                            Width =900
                            Height =420
                            Name ="lblRecordsSelected"
                            Caption ="records selected"
                            FontName ="Calibri"
                            LayoutCachedLeft =9060
                            LayoutCachedTop =420
                            LayoutCachedWidth =9960
                            LayoutCachedHeight =840
                        End
                    End
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
' FORM:    frm_Data_Gateway
' Level:        Application module
' Version:      1.08
'
' Description:  form related functions & procedures
'
' Data source:  qfrm_Data_Gateway
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
'               BLC - 4/17/2019  - 1.06 - update PseudoEvent handling
'               BLC - 4/2/2020   - 1.07 - fit report to window after opening vs. default smaller view
'               BLC - 6/22/2020  - 1.08 - add QC mode
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
'   BLC - 4/17/2019 - hide PseudoEvent toggle IF not DEV_MODE
'   BLC - 6/22/2020 - add QC mode
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'QC mode?
    SetTempVar "QC_MODE", IIf(Me.OpenArgs = "QC_MODE", True, False)
    lblQCMode.visible = Nz(TempVars("QC_MODE"), False)

    Dim varReturn As Variant

    ' On opening the form, set the initial sort order
    strSortFieldLabel = "lblPlot_Name"
    varReturn = fxnSortRecords("Plot_Name")
    ' Set the filter
    If fxnSwitchboardIsOpen Then
        'Not currently choosing to select default filter
        'Me!cboPanelFilter = Forms!frm_Switchboard!cPanel
        Me.FilterGateway (True)
    End If
    
    'hide PseudoEvent toggle when not in DEV_MODE
    Me.tglPseudoEvent.visible = IIf(TempVars("DEV_MODE") = True, False, True)
    
    'temporarily hide toggle
    Me.tglPseudoEvent.visible = False
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Data_Gateway])"
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
            "Error encountered (#" & Err.Number & " - Form_GotFocus[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click Events
' ----------------

' ---------------------------------
' SUB:          btnViewPhotos_Click
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
'   BLC - 5/23/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnViewPhotos_Click()
On Error GoTo Err_Handler

    Dim strCriteria As String

    'record what the current record is so we can go back to that record on return
    If Not IsNothing(Me!Location_ID) Then
        WriteRecordCriteria
        strCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Photos", , , strCriteria, , , "Filter by location"
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewPhotos_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnViewReport_Click
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
'   BLC - 5/23/2018 - update documentation, error handling,
'                     revise to open rpt_Event_Summary_Unfiltered vs.
'                     Copy of rpt_Event_Summary_Unfiltered
'   BLC - 4/2/2020  - fit report to window after opening vs. default smaller view
' ---------------------------------
Private Sub btnViewReport_Click()
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
        
        '10/23/2018 BLC
        'set TempVar for qry_Status_Sapling_Current_Event/qry_Status_Tree_Current_Event
        SetTempVar "EventID", CStr(Me.txtEvent_ID)
        
        strDocName = "rpt_Event_Summary_Unfiltered"
        'strDocName = "Copy of rpt_Event_Summary_Unfiltered"
        strCriteria = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
Debug.Print strCriteria
        'DoCmd.OpenReport stDocName, acPreview, "qRpt_Event_Summary_Unfiltered", stCriteria
        DoCmd.OpenReport strDocName, acPreview, , strCriteria
        
        'set to full size
        DoCmd.Maximize
        DoCmd.RunCommand acCmdZoom100 '100%
        'DoCmd.RunCommand acCmdFitToWindow 'fit window size
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewReport_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

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
            "Error encountered (#" & Err.Number & " - cbxParkFilter_AfterUpdate[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cboParkFilter_AfterUpdate()
    On Error GoTo Err_Handler

    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - cbxUnitGroupFilter_AfterUpdate[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cboUnitGroupFilter_AfterUpdate()
    On Error GoTo Err_Handler

    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - cbxPanelFilter_AfterUpdate[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cboPanelFilter_AfterUpdate()
    On Error GoTo Err_Handler

    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - cbxFrameFilter_AfterUpdate[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cboFrameFilter_AfterUpdate()
    On Error GoTo Err_Handler

    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - cbxYearFilter_AfterUpdate[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cboYearFilter_AfterUpdate()
    On Error GoTo Err_Handler

    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - cbxStatusFilter_AfterUpdate[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cboStatusFilter_AfterUpdate()
    On Error GoTo Err_Handler

    If Me!tglFilter Then
        Me.FilterGateway (True)
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - tglPseudoEvent_AfterUpdate[frm_Data_Gateway])"
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
            "Error encountered (#" & Err.Number & " - lblPlot_Name_DblClick[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

'CODE TO RESORT RECORDS WHEN HEADING IS CLICKED

Private Sub lblPlot_Name_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Plot_Name")

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
            "Error encountered (#" & Err.Number & " - lblUnitGroup_DblClick[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub lblUnit_Group_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Unit_Group")

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
            "Error encountered (#" & Err.Number & " - lblSubunitCode_DblClick[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub lblSubunit_Code_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Subunit_Code")

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
            "Error encountered (#" & Err.Number & " - lblEventYear_DblClick[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub lblEvent_Year_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    fxnSortRecords ("Event_Year")

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
            "Error encountered (#" & Err.Number & " - lblEventDate_DblClick[frm_Data_Gateway])"
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
            "Error encountered (#" & Err.Number & " - btnClearFilter_Click[frm_Data_Gateway])"
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
            "Error encountered (#" & Err.Number & " - btnAddLocation_Click[frm_Data_Gateway])"
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
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/24/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnAddEvent_Click()
On Error GoTo Err_Handler
    
    'record what the current record is so we can go back to that record on return
    WriteRecordCriteria
    DoCmd.Close acForm, "frm_Data_Gateway"
    DoCmd.OpenForm "frm_Event_Add"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddEvent_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cmdNewEvent_Click()
    On Error GoTo Err_Handler

    'record what the current record is so we can go back to that record on return
    WriteRecordCriteria
    DoCmd.Close acForm, "frm_Data_Gateway"
    DoCmd.OpenForm "frm_Event_Add"
        
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          btnClose_Click
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
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          tglFilter_AfterUpdate
' Description:  toggle after update actions
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
Private Sub tglFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    Me.FilterGateway (Me!tglFilter)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilter_AfterUpdate[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub xtglFilter_AfterUpdate()
    Me.FilterGateway (Me!tglFilter)
End Sub

' ---------------------------------
' SUB:          tbxEventDate_Click
' Description:  textbox click actions
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
            "Error encountered (#" & Err.Number & " - tbxEventDate_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub txtEvent_Date_Click()
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
    
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          tbxPlotName_Click
' Description:  textbox click actions
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
            "Error encountered (#" & Err.Number & " - tbxPlotName_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub txtPlot_Name_Click()
    On Error GoTo Err_Handler
    Dim strCriteria As String

    'record what the current record is so we can go back to that record on return
    If Not IsNothing(Me!Location_ID) Then
        WriteRecordCriteria
        strCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Locations", , , strCriteria, , , "Filter by location"
    End If
    
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          tbxReportTrigger_Click
' Description:  textbox click actions
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
'   BLC - 4/2/2020 - fit report to window after opening vs. default smaller view
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
    
        'set to full size
        DoCmd.Maximize
        DoCmd.RunCommand acCmdZoom100 '100%
        'DoCmd.RunCommand acCmdFitToWindow 'fit window size
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxReportTrigger_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub txtReportTrigger_Click()
On Error GoTo Err_Handler
    Dim strDocName As String
    Dim strCriteria As String
    
    If IsNothing(Me!Event_ID) Then
        'Trap records that do not contain an event.
        MsgBox ("This Record is not linked to an Event.  Please choose another Record.")
        GoTo Exit_Procedure
    Else
        'Record what the current record is so we can go back to that record on return
        WriteRecordCriteria
        strDocName = "rpt_Event_Summary_Unfiltered"
        strCriteria = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
        'DoCmd.OpenReport stDocName, acPreview, "qRpt_Event_Summary_Unfiltered", stCriteria
        DoCmd.OpenReport strDocName, acPreview, , strCriteria
        
        'set to full size
        DoCmd.Maximize
        DoCmd.RunCommand acCmdZoom100 '100%
        'DoCmd.RunCommand acCmdFitToWindow 'fit window size
    End If
    
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          tbxViewPhotos_Click
' Description:  textbox click actions
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
Private Sub tbxViewPhotos_Click()
On Error GoTo Err_Handler
    
    Dim RetVal As Double
    Dim RootFolder As String
    Dim PhotoFolder As String
    
    RootFolder = "T:\I&M"
    PhotoFolder = "T:\I&M\Monitoring\Forest_Vegetation\Photos\"
    If FolderExists(PhotoFolder & Me!txtPlot_Name) Then
        RetVal = shell("explorer /e,/root, " & PhotoFolder & Me!txtPlot_Name, vbNormalFocus)
        GoTo Exit_Handler
    Else
        If FolderExists(RootFolder) Then
            MsgBox ("Folder for this plot not found....Opening the root of the Photos folder.")
            RetVal = shell("explorer /e,/root, " & PhotoFolder, vbNormalFocus)
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
            "Error encountered (#" & Err.Number & " - tbxViewPhotos_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub txtViewPhotos_Click()
On Error GoTo Err_Handler

    Dim RetVal As Double
    Dim RootFolder As String
    Dim PhotoFolder As String
    
    RootFolder = "T:\I&M"
    PhotoFolder = "T:\I&M\Monitoring\Forest_Vegetation\Photos\"
    If FolderExists(PhotoFolder & Me!txtPlot_Name) Then
        RetVal = shell("explorer /e,/root, " & PhotoFolder & Me!txtPlot_Name, vbNormalFocus)
        GoTo Exit_Procedure
    Else
        If FolderExists(RootFolder) Then
            MsgBox ("Folder for this plot not found....Opening the root of the Photos folder.")
            RetVal = shell("explorer /e,/root, " & PhotoFolder, vbNormalFocus)
            GoTo Exit_Procedure
        Else
            MsgBox ("The network appears to be unavailable. Network access is required to view photos.")
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          btnGoToTags_Click
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
Private Sub btnGoToTags_Click()
On Error GoTo Err_Handler
    
'    'record what the current record is so we can go back to that record on return
'    WriteRecordCriteria
'    DoCmd.OpenForm "frm_Tags"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnGoToTags_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cmdGoto_Tags_Click()
    On Error GoTo Err_Handler

    'record what the current record is so we can go back to that record on return
    WriteRecordCriteria
    DoCmd.OpenForm "frm_Tags"
        
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          btnGoToPlants
' Description:  Open the plants form
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
Private Sub btnGoToPlants_Click()
On Error GoTo Err_Handler
    
'    'record what the current record is so we can go back to that record on return
'    WriteRecordCriteria
'    DoCmd.OpenForm "frm_Plants"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnGoToPlants[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub cmdGoto_Plants_Click()
    On Error GoTo Err_Handler

    'record what the current record is so we can go back to that record on return
    WriteRecordCriteria
    DoCmd.OpenForm "frm_Plants"
        
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
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
Private Function FilterString(val As Variant, fieldName As String, CurrentFilter As Variant) As Variant
On Error GoTo Err_Handler

    Const cstrNull As String = "[Null]"
    Dim Filter As Variant

    If IsNull(val) Then
        Filter = CurrentFilter
    Else
        Filter = (CurrentFilter + " AND ") & fieldName
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
            "Error encountered (#" & Err.Number & " - FilterString[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Function

Private Function xFilterString(varValue As Variant, strFieldName As String, varCurrentFilter As Variant) As Variant
' Description:  Builds a filter string for the Data Gateway form
' Parameters:   varValue = the value of the filter control
'               strFieldName = the field that is being filtered
'               varCurrentFilter = the current filter value as it is being built up
' Returns:      Filter string or null if no filter built yet
' Throws:       none
' References:   none
' Source/date:  Simon Kingston, 1/17/2007
' Revisions:    <name, date, desc - add lines as you go>

Const cstrNull As String = "[Null]"
Dim varFilter As Variant

On Error GoTo Error_Handler

If IsNull(varValue) Then
    varFilter = varCurrentFilter
Else
    varFilter = (varCurrentFilter + " AND ") & strFieldName
    If varValue = cstrNull Then
        varFilter = varFilter & " Is Null"
    Else
    If IsNumeric(varValue) Then
        varFilter = varFilter & "=" & varValue & ""
        Else
        If IsDate(varValue) Then
            varFilter = varFilter & "=#" & varValue & "#"
            Else
                varFilter = varFilter & "=" & CorrectText(CStr(varValue))
            End If
        End If
    End If
End If

xFilterString = varFilter

Exit_Handler:
    Exit Function
Error_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (FilterString)"
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
            "Error encountered (#" & Err.Number & " - FilterGateway[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Public Sub xFilterGateway(booFilterOn As Boolean)
' Description:  Filters the Data Gateway form
' Parameters:   booFilterOn = true if filter is to be applied, false if filter is to be removed
' Returns:      none
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  Simon Kingston, 1/17/2007
' Revisions:    <name, date, desc - add lines as you go>
Dim varFilter As Variant

On Error GoTo Error_Handler

varFilter = Null

Me!tglFilter = booFilterOn

If booFilterOn Then
    Me!tglFilter.Caption = "Filter Is On"

    'add park filter to filter string
    varFilter = FilterString(Me!cboParkFilter, "Unit_Code", varFilter)
    'add park filter to filter string
    varFilter = FilterString(Me!cboUnitGroupFilter, "Unit_Group", varFilter)
    'add panel filter to filter string
    varFilter = FilterString(Me!cboPanelFilter, "Panel", varFilter)
    'add frame filter to filter string
    varFilter = FilterString(Me!cboFrameFilter, "Frame", varFilter)
    'add year filter to filter string
    varFilter = FilterString(Me!cboYearFilter, "Event_Year", varFilter)
    'add status filter to filter string
    varFilter = FilterString(Me!cboStatusFilter, "Location_Status", varFilter)
    Me.Filter = Nz(varFilter)
Else
    Me!tglFilter.Caption = "Filter Is Off"
End If
Me.FilterOn = booFilterOn

Exit_Handler:
    Exit Sub
Error_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (FilterGateway)"
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
Debug.Print strCurrentRecordCriteria
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - WriteRecordCriteria[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Sub

Private Sub xWriteRecordCriteria()
' Description:  Records the Location ID and Event ID of the current record so that it can be made the current record when coming
'               back to the form from another form (=bookmark).
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   GetCriteriaString
' Source/date:  Simon Kingston, 1/17/2007
' Revisions:    <name, date, desc - add lines as you go>

On Error GoTo Error_Handler

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

Error_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (FilterGateway)"
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
        .fontBold = IIf(.fontBold = False, True, False)
    
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
            "Error encountered (#" & Err.Number & " - SortRecords[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          fxnSortRecords
' Description:  record sorting actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling, add stripping _ from label
' ---------------------------------
Private Function fxnSortRecords(ByVal strFieldName As String, _
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

    'adjust for label name
    strSortFieldLabel = Replace(strSortFieldLabel, "_", "")

    ' Change the label format to indicate the sorted field
    Me.Controls.Item(strSortFieldLabel).FontItalic = False
    Me.Controls.Item(strSortFieldLabel).fontBold = False
    'strSortFieldLabel = "lbl" & strFieldName
    Me.Controls.Item(strSortFieldLabel).FontItalic = True
    Me.Controls.Item(strSortFieldLabel).fontBold = True

Exit_Procedure:
    Exit Function
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (fxnSortRecords)"
    Resume Exit_Procedure
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
            "Error encountered (#" & Err.Number & " - btnGoToTags_Click[frm_Data_Gateway])"
    End Select
    Resume Exit_Handler
End Function
