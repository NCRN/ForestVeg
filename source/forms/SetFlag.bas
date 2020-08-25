Version =21
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    DefaultView =0
    ScrollBars =2
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7560
    DatasheetFontHeight =11
    ItemSuffix =55
    Left =5025
    Top =1455
    Right =12585
    Bottom =5205
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x74b6cfd4997ce540
    End
    RecordSource ="DataFlags"
    Caption ="Set Data Flag"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FetchDefaults =0
    FetchDefaults =0
    FilterOnLoad =255
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
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
        Begin FormHeader
            Height =2460
            BackColor =4144959
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Top =30
                    Width =3480
                    Height =300
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedTop =30
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =330
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =120
                    Top =15
                    Width =7260
                    Height =840
                    BorderColor =8355711
                    ForeColor =16777164
                    Name ="lblDirections"
                    Caption ="Sort records by clicking the header.\015\012Effective date color reflects if dat"
                        "a flag is retired or not."
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =15
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =855
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =1020
                    Top =2100
                    Width =660
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFlagType"
                    Caption ="Type"
                    FontName ="Franklin Gothic Book"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1020
                    LayoutCachedTop =2100
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =660
                    Top =2100
                    Width =270
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblHdrID"
                    Caption ="ID"
                    FontName ="Franklin Gothic Book"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =660
                    LayoutCachedTop =2100
                    LayoutCachedWidth =930
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =1860
                    Top =2100
                    Width =3780
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblFlag"
                    Caption ="Flag"
                    FontName ="Franklin Gothic Book"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =1860
                    LayoutCachedTop =2100
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =223
                    Left =4800
                    Top =2100
                    Width =720
                    Height =315
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblCategory"
                    Caption ="Version"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =4800
                    LayoutCachedTop =2100
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =5790
                    Top =1860
                    Width =900
                    Height =555
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblEffectiveDate"
                    Caption ="Effective Date"
                    FontName ="Franklin Gothic Book"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638
                    LayoutCachedLeft =5790
                    LayoutCachedTop =1860
                    LayoutCachedWidth =6690
                    LayoutCachedHeight =2415
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6480
                    Top =1380
                    Width =720
                    ForeColor =4210752
                    Name ="btnAddNewFlag"
                    Caption ="Add Flag"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00ddddddddddddddddddddddddddddddddddddddddddd777dd ,
                        0xddddddddddd777dddddd666ddddddddddddd666ddddddddddddd666dddd777dd ,
                        0xd666666666d777ddd666666666d777ddd666666666d777dddddd666dddd777dd ,
                        0xdddd666dddd777dddddd666dddd777ddddddddddddd777dddddddddddddddddd ,
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
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add new flag for this data context"
                    GridlineColor =10921638

                    LayoutCachedLeft =6480
                    LayoutCachedTop =1380
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1740
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6000
                    Top =180
                    Width =720
                    TabIndex =1
                    ForeColor =4210752
                    Name ="btnOpenTable"
                    Caption ="Add Record"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open data flags table"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000b0a090ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff0000000000000000b0a090ffe0c8c0ffd0c0b0ffd0b8b0ffd0b8b0ff ,
                        0xc0b0a0ffc0b0a0ffc0b0a0ffc0a8a0ffc0a890ffc0a890ffb0a090ffb0a090ff ,
                        0x604830ff0000000000000000b0a090fffffffffffffffffffff8ffffd0b8b0ff ,
                        0xfff0f0fffff0e0ffffe8e0ffc0a8a0fff0d8d0fff0d8c0fff0d0b0ffb0a090ff ,
                        0x604830ff0000000000000000b0a090ffffffffffffffffffffffffffd0c0b0ff ,
                        0xfff8f0fffff0f0fffff0e0ffc0b0a0ffffe0d0fff0d8d0fff0d8c0ffc0a890ff ,
                        0x604830ff0000000000000000b0a090ffe0d0d0ffd0c8c0ffd0c0c0ffd0c0b0ff ,
                        0xd0c0b0ffd0b8b0ffd0b8b0ffc0b0a0ffc0b0a0ffc0b0a0ffc0a8a0ffc0a890ff ,
                        0x604830ff0000000000000000c0a890ffffffffffffffffffffffffffd0c8c0ff ,
                        0xfffffffffff8fffffff8f0ffd0b8b0fffff0e0ffffe8e0ffffe0d0ffc0a8a0ff ,
                        0x604830ff0000000000000000c0a8a0ffffffffffffffffffffffffffd0c8c0ff ,
                        0xfffffffffffffffffff8ffffd0b8b0fffff0f0fffff0e0ffffe8e0ffc0a8a0ff ,
                        0x604830ff0000000000000000c0b0a0ffe0d8d0ffe0d0c0ffe0d0c0ffe0c8c0ff ,
                        0xd0c8c0ffd0c8c0ffd0c0b0ffd0c0b0ffd0b8b0ffd0b8b0ffc0b0a0ffc0b0a0ff ,
                        0x604830ff0000000000000000d0b0a0ffffffffffffffffffffffffffe0d0c0ff ,
                        0xffffffffffffffffffffffffd0c0b0fffff8fffffff8f0fffff0f0ffc0b0a0ff ,
                        0x604830ff0000000000000000d0b8a0ffffffffffffffffffffffffffe0d0c0ff ,
                        0xffffffffffffffffffffffffd0c8c0fffffffffffff8fffffff8f0ffd0b8b0ff ,
                        0x604830ff0000000000000000f0a890fff0a890fff0a890fff0a880fff0a080ff ,
                        0xe09870ffe09060ffe08850ffe08050ffe07840ffe07040ffe07040ffe07040ff ,
                        0xd06030ff0000000000000000f0a890ffffc0a0ffffc0a0ffffc0a0ffffb890ff ,
                        0xffb890ffffb090ffffa880ffffa880fff0a070fff0a070fff09870fff09860ff ,
                        0xd06830ff0000000000000000f0a890fff0a890fff0a890fff0a890fff0a880ff ,
                        0xf0a080fff09870ffe09870ffe09060ffe08860ffe08050ffe07840ffe07840ff ,
                        0xe07040ff00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =180
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =540
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6780
                    Top =600
                    Width =720
                    TabIndex =2
                    ForeColor =4210752
                    Name ="btnAddFlag"
                    Caption ="Add SOP"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Add new data flag record"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000b09880ff201010ff201010ff201010ff201010ff201010ff ,
                        0x201010ff201010ff201010ff201010ff201010ff201010ff201010ff00000000 ,
                        0x0000000000000000c0a090fffff8f0fffff8f0fffff0f0fffff0e0fff0e8e0ff ,
                        0xf0e8d0fff0e0d0fff0e0d0fff0e0d0fff0d8d0fff0d8d0ff201810ff00000000 ,
                        0x0000000000000000c0a090ffffffffffd07850ffd07840ffd07040ffc07040ff ,
                        0xc06840ffc06840ffc06840ffc07040ffa06040fff0e0d0ff403830ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850fff0b8a0fff0b090fff0a880ff ,
                        0xf0a080fff09870fff09870fff0a880ffc09880fffff0f0ff909090ff00000000 ,
                        0x0000000000000000c0a890ffffffffffd07850ffd07850ffd07840ffd07040ff ,
                        0xc07040ffc07050ffd09070ff70b8c0ff90d8f0ff90f0ffff40c0e0ffa0f0ffff ,
                        0xa0e8ffff90d8f0ffc0a8a0fffffffffffffffffffffffffffffffffffff8f0ff ,
                        0xfff8f0fffff8f0fffff8f0ffb0e8ffff30b8e0ff80e8ffff60c8e0ff90f0ffff ,
                        0x30b8e0ffa0e8ffffc0a8a0ffc0a8a0ffc0a890ffc0a090ffc0a090ffc0a090ff ,
                        0xc09880ffc0a090ffd0c0b0ffa0e8ffff90f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffffa0f0ffff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000020a8e0ff50c0e0ffb0e8f0fff0ffffffb0e8f0ff ,
                        0x50c0e0ff30b8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000080e8ffc090f0ffffc0f8ffffb0e8f0ffc0f8ffff ,
                        0x90f0ffff90d8e0ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000050d8ff8030b8e0ff90f0ffff60c0e0ff90f0ffff ,
                        0x30b8e0ff50d0f080000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000030b0e0a040c8f09080e8ffc020b0e0ff70e8ffc0 ,
                        0x50d8f08030b0e080000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedTop =600
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =960
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1140
                    Top =960
                    Width =2040
                    Height =288
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =2171426
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cbxFlagType"
                    RowSourceType ="Table/Query"
                    RowSource ="s_flag_types"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =960
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =1248
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =180
                            Top =960
                            Width =960
                            Height =345
                            BorderColor =8355711
                            ForeColor =16777215
                            Name ="lblSelectFlagType"
                            Caption ="Flag Type"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =960
                            LayoutCachedWidth =1140
                            LayoutCachedHeight =1305
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =1395
                    Width =2040
                    Height =288
                    ColumnOrder =0
                    FontSize =10
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxTableRef"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =1395
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =1683
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1395
                            Width =945
                            Height =315
                            BorderColor =8355711
                            ForeColor =16777215
                            Name ="lblTableRef"
                            Caption ="Table"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =1395
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =1710
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =1380
                    Width =2040
                    Height =288
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxFieldRef"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedTop =1380
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =1668
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3360
                            Top =1380
                            Width =945
                            Height =315
                            BorderColor =8355711
                            ForeColor =16777215
                            Name ="lblFieldRef"
                            Caption ="Field"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =3360
                            LayoutCachedTop =1380
                            LayoutCachedWidth =4305
                            LayoutCachedHeight =1695
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =1740
                    Width =2040
                    Height =288
                    ColumnOrder =2
                    FontSize =10
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxUserRef"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =1140
                    LayoutCachedTop =1740
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =2028
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Top =1740
                            Width =945
                            Height =315
                            BorderColor =8355711
                            ForeColor =16777215
                            Name ="lblUserRef"
                            Caption ="User"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =180
                            LayoutCachedTop =1740
                            LayoutCachedWidth =1125
                            LayoutCachedHeight =2055
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =4320
                    Top =1815
                    Width =2040
                    Height =288
                    FontSize =10
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxRecordRef"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedTop =1815
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =2103
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3360
                            Top =1815
                            Width =945
                            Height =315
                            BorderColor =8355711
                            ForeColor =16777215
                            Name ="lblRecordRef"
                            Caption ="RecordID"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =3360
                            LayoutCachedTop =1815
                            LayoutCachedWidth =4305
                            LayoutCachedHeight =2130
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =960
                    Width =2040
                    Height =288
                    FontSize =10
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =2171426
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cbxFlags"
                    RowSourceType ="Table/Query"
                    RowSource ="s_flag_types"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4320
                    LayoutCachedTop =960
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =1248
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3360
                            Top =960
                            Width =945
                            Height =315
                            BorderColor =8355711
                            ForeColor =16777215
                            Name ="lblFlags"
                            Caption ="Flag"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =3360
                            LayoutCachedTop =960
                            LayoutCachedWidth =4305
                            LayoutCachedHeight =1275
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1560
            Name ="Detail"
            OnMouseMove ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =2
                    BorderColor =8355711
                    ForeColor =690698
                    Name ="tbxIcon"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =45
                    LayoutCachedWidth =480
                    LayoutCachedHeight =360
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =50.0
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontFamily =2
                    Left =6360
                    Width =720
                    FontSize =14
                    TabIndex =1
                    ForeColor =255
                    Name ="btnDelete"
                    Caption ="í ½í·´"
                    OnClick ="[Event Procedure]"
                    FontName ="Academy Engraved LET"
                    ControlTipText ="Delete this data flag"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedWidth =7080
                    LayoutCachedHeight =360
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =-1
                    BackColor =11710639
                    BorderColor =11710639
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =60
                    Width =3720
                    Height =315
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxFlag"
                    ControlSource ="FlagName"
                    FontName ="Franklin Gothic Book"
                    OnMouseMove ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =600
                    Top =45
                    Width =360
                    Height =315
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxID"
                    ControlSource ="ID"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =45
                    LayoutCachedWidth =960
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4890
                    Top =60
                    Width =420
                    Height =315
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxVersion"
                    ControlSource ="FlagCategory"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4890
                    LayoutCachedTop =60
                    LayoutCachedWidth =5310
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5700
                    Top =60
                    Width =1020
                    Height =315
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxEffectiveDate"
                    ControlSource ="EffectiveDate"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x0100000056010000030000000100000000000000000000001e00000001000000 ,
                        0x22b14c00ffffff0001000000000000001f0000005c00000001000000ed1c2400 ,
                        0xffffff0001000000000000005d0000007a000000010000000000ff00ffffff00 ,
                        0x4900490066002800490073004e0075006c006c0028005b005200650074006900 ,
                        0x7200650044006100740065005d0029002c0031002c0030002900000000004900 ,
                        0x4900660028004e006f0074002000490073004e0075006c006c0028005b005200 ,
                        0x6500740069007200650044006100740065005d0029002c004900490066002800 ,
                        0x5b0052006500740069007200650044006100740065005d003c00440061007400 ,
                        0x6500280029002c0031002c00300029002c003000290000000000490049006600 ,
                        0x28005b0052006500740069007200650044006100740065005d003d0044006100 ,
                        0x74006500280029002c0031002c003000290000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000022b14c00ffffff001d0000004900 ,
                        0x490066002800490073004e0075006c006c0028005b0052006500740069007200 ,
                        0x650044006100740065005d0029002c0031002c00300029000000000000000000 ,
                        0x00000000000000000000000000010000000000000001000000ed1c2400ffffff ,
                        0x003c00000049004900660028004e006f0074002000490073004e0075006c006c ,
                        0x0028005b0052006500740069007200650044006100740065005d0029002c0049 ,
                        0x004900660028005b0052006500740069007200650044006100740065005d003c ,
                        0x004400610074006500280029002c0031002c00300029002c0030002900000000 ,
                        0x0000000000000000000000000000000000000100000000000000010000000000 ,
                        0xff00ffffff001c00000049004900660028005b00520065007400690072006500 ,
                        0x44006100740065005d003d004400610074006500280029002c0031002c003000 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6780
                    Width =720
                    ForeColor =4210752
                    Name ="btnEdit"
                    Caption ="Edit"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Edit data flag info"
                    GridlineColor =10921638
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000303840ff404040ff505050ff504850f080686020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000606060ff909890ffd0d0d0ffa0a8b0ff304850ff ,
                        0xa090905000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000a0a0a0fff0f0f0fff0f8ffffc0e0f0ff5090b0ff ,
                        0x204850ff80686020000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000080787080e0e0e0ffd0f0f0ff90e0f0ff50c0d0ff ,
                        0x4098b0ff204850ff806860200000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000006090a080c0e8f0ffa0f0f0ff70e0f0ff ,
                        0x50c0d0ff4098b0ff204850ff8068602000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000006090a090b0e8f0ffa0f0f0ff ,
                        0x70e0f0ff50c0d0ff4098b0ff204850ff80686020000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000006090a090b0e8f0ff ,
                        0xa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff806860200000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000006090a0a0 ,
                        0xb0e8f0ffa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff8068602000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x6090a0a0b0e8f0ffa0f0f0ff70e0f0ff50c0d0ff4098b0ff204850ff80686020 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xd08060006090a0a0b0e8f0ffa0f0f0ff70e0f0ff50b8d0ff4098b0ff204850ff ,
                        0x8068602000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000d0d8e0006090a0b0b0e8f0ffa0f0f0ff70d0e0ff50a0b0ff808890ff ,
                        0x303870ff80686020000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000d0d8e0006090a0b0c0f0f0ffa0e0e0ffb0b0a0ff5058b0ff ,
                        0x303090ff505880ff000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000d0d8e0006090a0b0a0b8d0ff8088d0ff6070d0ff ,
                        0x303090ff202860ff000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000d0d8e0006070b0b09098d0ff7078d0ff ,
                        0x4050a0ff9098b0ff000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000d0d8e000606090d05060a0ff ,
                        0x9090b0ff00000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =360
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =840
                    Height =315
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxFlagType"
                    ControlSource ="FlagType"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x0100000098000000030000000000000002000000000000000800000001010000 ,
                        0xed1c2400ffffff000000000002000000090000001300000001010000fba30500 ,
                        0xffffff000000000002000000140000001b000000010100000072bc00ffffff00 ,
                        0x22004500720072006f00720022000000000022005700610072006e0069006e00 ,
                        0x6700220000000000220049006e0066006f00220000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =1080
                    LayoutCachedTop =60
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010003000000000000000200000001010000ed1c2400ffffff00070000002200 ,
                        0x4500720072006f00720022000000000000000000000000000000000000000000 ,
                        0x00000000000200000001010000fba30500ffffff000900000022005700610072 ,
                        0x006e0069006e0067002200000000000000000000000000000000000000000000 ,
                        0x0000000002000000010100000072bc00ffffff0006000000220049006e006600 ,
                        0x6f002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =60
                    Width =1200
                    Height =315
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =2171426
                    Name ="tbxLabel"
                    ControlSource ="Label"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =375
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
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
' Form:         SetFlag
' Level:        Application form
' Version:      1.00
' Basis:        Dropdown form
'
' Description:  List form object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, June 22, 2020
' References:   -
' Revisions:    BLC - 6/22/2020  - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String
Private m_ButtonCaption
Private m_SelectedID As Integer
Private m_SelectedValue As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)
Public Event InvalidLabel(Value As String)
Public Event InvalidCaption(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(Value As String)
        m_CallingForm = Value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

Public Property Let ButtonCaption(Value As String)
    If Len(Value) > 0 Then
        m_ButtonCaption = Value

        'set the form button caption
        'Me.btnEdit.Caption = m_ButtonCaption
    Else
        RaiseEvent InvalidCaption(Value)
    End If
End Property

Public Property Get ButtonCaption() As String
    ButtonCaption = m_ButtonCaption
End Property

Public Property Let SelectedID(Value As Integer)
        m_SelectedID = Value
End Property

Public Property Get SelectedID() As Integer
    SelectedID = m_SelectedID
End Property

Public Property Let SelectedValue(Value As String)
        m_SelectedValue = Value
End Property

Public Property Get SelectedValue() As String
    SelectedValue = m_SelectedValue
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 11, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/11/2020 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim oa As Variant
    Dim tblRef As String
    Dim fldRef As String
    Dim RecordID As Variant
Debug.Print Me.OpenArgs

    'default
    Me.CallingForm = "frm_Events"

    If Len(Me.OpenArgs) > 0 Then
        'OpenArgs expect - CallingForm | Table | Field | Record ID
        oa = Split(Me.OpenArgs, "|")
        tblRef = Nz(oa(1), "")
        fldRef = Nz(oa(2), "")
        RecordID = Nz(oa(3), "")
        
        Me.CallingForm = Nz(oa(0), "frm_Switchboard")
        Me.tbxTableRef = tblRef
        Me.tbxFieldRef = fldRef
        Me.tbxRecordRef = RecordID
        Me.tbxUserRef = Nz(oa(4), "")
    End If

    'filter flags
    If Len(tblRef) > 0 And Len(fldRef) > 0 Then FilterFlags tblRef, fldRef

    'minimize calling form
    ToggleForm Me.CallingForm, -1
    
    Me.Caption = "Data Flags"
    lblTitle.Caption = ""
    lblDirections.Caption = "Add the flag as appropriate." _
                            & vbCrLf & ""
    tbxIcon.Value = StringFromCodepoint(uLocked)
    tbxIcon.forecolor = lngDkGreen
    lblDirections.forecolor = lngLtBlue
    
    'set hover
    btnAddNewFlag.hoverColor = lngGreen
    btnOpenTable.hoverColor = lngGreen
    btnAddFlag.hoverColor = lngGreen
    btnDelete.hoverColor = lngGreen
    btnEdit.hoverColor = lngGreen
    
    btnDelete.Caption = StringFromCodepoint(uDelete)
    btnDelete.forecolor = lngRed

    'enable textbox to ensure scrollbar is available for longer text
    tbxFlag.Enabled = True
    
    'cover to avoid data entry

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

    'eliminate NULLs
    If IsNull(Me.OpenArgs) Then GoTo Exit_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxFlag_MouseMove
' Description:  textbox mouse move actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   datAdrenaline, March 29, 2005
'   http://www.utteraccess.com/forum/controltiptext-continuo-t669660.html
' Source/date:  Bonnie Campbell, June 18, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/18/2016 - initial version
' ---------------------------------
Private Sub tbxFlag_MouseMove(Button As Integer, Shift As Integer, _
                                    x As Single, y As Single)
On Error GoTo Err_Handler
       Static lngRecordID As Long 'keeps track of current record to make sure
                                  'code will execute once when mouse enters control
        If lngRecordID <> Me.tbxID Then
            lngRecordID = Me.tbxID
            Me.tbxFlag.ControlTipText = Nz(Me.tbxLabel, "")
        End If
       
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxFlag_MouseMove[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxFlagType_AfterUpdate
' Description:  Open table button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 18, 2020
' Adapted:      -
' Revisions:
'   BLC - 1/18/2020 - initial version
' ---------------------------------
Private Sub cbxFlagType_AfterUpdate()
On Error GoTo Err_Handler
    
    Me.Filter = "FlagType = '" & cbxFlagType & "'"
    Me.FilterOn = Not IsNull(Me.cbxFlagType)

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxFlagType_AfterUpdate[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnOpenTable_Click
' Description:  Open table button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 11, 2020
' Adapted:      -
' Revisions:
'   BLC - 1/10/2017 - initial version
' ---------------------------------
Private Sub btnOpenTable_Click()
On Error GoTo Err_Handler
    
    'minimize SetFlag
    ToggleForm "SetFlag", -1
    
    DoCmd.OpenTable "Flags", acViewNormal ',acAdd

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenTable_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnCreateVersionTable_Click
' Description:  Open table button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Microsoft, January 19, 2017
'   https://docs.microsoft.com/en-us/sql/odbc/microsoft/column-name-limitations
' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/19/2017 - initial version
' ---------------------------------
Private Sub btnCreateVersionTable_Click()
On Error GoTo Err_Handler
    
    'minimize SetFlag
    ToggleForm "SetFlag", -1
    
    're-generate the current SOP Version Table
    
    ' SOP_crosstab
    '
    ' retrieve SOP names
    ' retrieve SOP numbers
    Dim rs As DAO.Recordset
    
    Dim ary() As Variant
    ary = RetrieveTableColumnData("SOP_VersionTable")

    Debug.Print ary(1)

    'prepare top level hdr - SOP Names
    'prepare second level hdr - SOP #s

    Dim ary2() As String
    ary2 = Split(ary(1), ",")
    
    Dim rsColInfo As DAO.Recordset
    Set rsColInfo = ary(0)
    
    Dim i As Integer
    Dim sop As String, sopnum As Integer
    
    Dim aryRecord() As Variant
    
    For i = 0 To UBound(ary2)

'        Debug.Print ary2(i)(0)

'    'create new record
'        rs.AddNew
'
'
'        'aryRecord() = IIF(InStr(ary2(i)Split(ary2(i)(0), "-")
'
'        rs!EffectiveDate = aryRecord(0)
'        rs!ColType = aryRecord(5)
'        rs!IsReqd = IIf(aryRecord(3) = False, 0, 1)
'        rs!Length = aryRecord(2)
'        rs!AllowZLS = IIf(aryRecord(4) = False, 0, 1)
'
'        'add the new record
'        rs.Update


    Next
    
    'create columns - EffectiveDate, 1-XXX SOP names
    
    Dim tdf As New DAO.TableDef
    Dim aryCols() As String
    Dim tbl As String

    aryCols = Split(ary(1), ",")
        
    'generate table name
    tbl = "SOP_Version_" & Format(Now, "YYYYmmdd_hhmmss")

    'remove table if it already exists
    Dim result As Boolean
    If TableExists(tbl) Then _
         result = MsgBox("Version table already exists. Delete existing table: #" & tbl & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Existing SOP Version Table?")

    If result = vbYes Then CurrDb.TableDefs.Delete tbl

    With tdf
        .Name = tbl
        .Fields.Append .CreateField("EffectiveDate", dbDate)
        
        'iterate through the SOPs (skip EffectiveDate = first record, aryCols(0))
        'maximum column name length = 64
        'column w/ any other characters other than letters, #s, or underscores
        'name must be delimited by enclosing it in back quotes (`)
        For i = 1 To UBound(aryCols)
        'Debug.Print aryCols(i) & " " & Len(aryCols(i))
        
            'add only viable fields
            If Len(Trim(aryCols(i))) > 0 Then
                'add the column
                .Fields.Append .CreateField("" & Trim(aryCols(i)) & "", dbDouble)
            End If
        Next
        
        CurrDb.TableDefs.Append tdf
    End With

    'move table to RESULT TABLES group
    SetNavGroup "RESULT TABLES", tbl, "table"

    Dim rsNew As DAO.Recordset
    
    'open a rs from the table
    Set rsNew = CurrDb.OpenRecordset(tbl)
    
    
    'iterate through SOP data
    If Not (rsColInfo.BOF And rsColInfo.EOF) Then
        rsColInfo.MoveFirst
        Do Until rsColInfo.EOF = True

            rs.AddNew
                
                Debug.Print rsColInfo.Fields("Column")
                Debug.Print rsColInfo.Fields("ColType")
                Debug.Print rsColInfo.Fields("IsReqd")
                Debug.Print rsColInfo.Fields("Length")
                Debug.Print rsColInfo.Fields("AllowZLS")
                
                'create columns
                
                'EffectiveDate first
                
            rs.Update
    
        rsColInfo.MoveNext
        Loop
    End If


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCreateVersionTable_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          btnAddFlag_Click
' Description:  Add SOP button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 11, 2020 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/11/2020 - initial version
' ---------------------------------
Private Sub btnAddFlag_Click()
On Error GoTo Err_Handler
    
    'minimize SetFlag
    ToggleForm "SetFlag", -1
    
    DoCmd.OpenTable "Flags", acViewNormal, acAdd
    DoCmd.GoToRecord acDataTable, "Flags", acNewRec

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddFlag_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

'FIX - MISSING ID PROPERTY
'' ---------------------------------
'' Sub:          btnEdit_Click
'' Description:  Enter button click actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 5/31/2016 - initial version
''   BLC - 1/19/2017 - converted to SOPs from SQL Templates
'' ---------------------------------
'Private Sub btnEdit_Click()
'On Error GoTo Err_Handler
'
'    'minimize form
'    ToggleForm "SetFlag", -1
'
'    'open the table for editing
'    DoCmd.OpenTable "SOP", acViewNormal, acEdit
'
'    DoCmd.GoToRecord acDataTable, "SOP", acGoTo, Me.ID
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - btnEdit_Click[SetFlag form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Sub:          btnDelete_Click
' Description:  Delete button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 10/16/2017 - revised to use tbxID vs. ID on delete
' ---------------------------------
Private Sub btnDelete_Click()
On Error GoTo Err_Handler
    
    Dim result As Integer
    
    'identify the record ID
     result = MsgBox("Delete Record this record: #" & tbxID & " ?" _
                        & vbCrLf & "This action cannot be undone.", vbYesNo, "Delete Record?")

    If result = vbYes Then DeleteRecord "SetFlag", tbxID
    
    'clear the deleted record
    Me.Requery

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDelete_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Sorts
'---------------------

' ---------------------------------
' Sub:          lblHdrID_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblHdrID_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblHdrID
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblHdrID_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblFlagType_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/19/2017 - initial version
' ---------------------------------
Private Sub lblFlagType_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblFlagType

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblFlagType_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

'' ---------------------------------
'' Sub:          lblFlagType_Click
'' Description:  lbl click actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:
''   pere_de_chipstic, August 5, 2012
''   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
'' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 1/19/2017 - initial version
'' ---------------------------------
'Private Sub lblFlagType_Click()
'On Error GoTo Err_Handler
'
'    'set the sort
'    SortListForm Me, Me.lblFlagType
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - lblFlagType_Click[SetFlag form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Sub:          lblFlag_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/19/2017 - initial version
' ---------------------------------
Private Sub lblFlag_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblFlag

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblFlag_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          lblEffectiveDate_Click
' Description:  lbl click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub lblEffectiveDate_Click()
On Error GoTo Err_Handler

    'set the sort
    SortListForm Me, Me.lblEffectiveDate

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblEffectiveDate_Click[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          tbxSOP_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Template Name textbox is disabled, so control tips won't display
'               Otherwise this would be tbxSOPName_MouseMove instead & tbxSOP would
'               not be necessary
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   hnaser, March 17, 2013
'   https://www.experts-exchange.com/questions/28067200/MS-Access-tooltip-on-a-disabled-control.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub tbxSOP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler

'    Me.tbxSOP.ControlTipText = Nz(FetchAddlData("SOP", "Remarks", Me.tbxID)(0), "")
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxSOP_MouseMove[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Detail_MouseMove
' Description:  mouse move (hover) actions
' Assumptions:  -
'               Template Name textbox is disabled, so control tips won't display
'               Otherwise this would be tbxTemplateName_MouseMove instead & tbxControlTip would
'               not be necessary
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   hnaser, March 17, 2013
'   https://www.experts-exchange.com/questions/28067200/MS-Access-tooltip-on-a-disabled-control.html
' Source/date:  Bonnie Campbell, September 13, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/13/2016 - initial version
' ---------------------------------
Private Sub Detail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Err_Handler

'    Me.tbxSOP.ControlTipText = ""
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_MouseMove[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Close
' Description:  form closing actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 31, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/31/2016 - initial version
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'close flags table if open
    DoCmd.Close acTable, "Flags", acSavePrompt

    'restore calling form
    ToggleForm Me.CallingForm, 0
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          FilterFlags
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 23, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/23/2020 - initial version
' ---------------------------------
Private Sub FilterFlags(TableRef As String, FieldRef As String)
On Error GoTo Err_Handler

Debug.Print "tblRef: " & TableRef & " fldRef: " & FieldRef

    Me.cbxFlags.RowSource = "SELECT * FROM DataFlags "
    
    'GetTemplate("s_flags_context_filtered","")
    


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterFlags[SetFlag form])"
    End Select
    Resume Exit_Handler
End Sub
