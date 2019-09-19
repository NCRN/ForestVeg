Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14040
    DatasheetFontHeight =10
    ItemSuffix =124
    Left =4830
    Top =4305
    Right =18855
    Bottom =10875
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x382bdd274ff0e240
    End
    RecordSource ="SELECT tbl_Quadrat_Data.* FROM tbl_Quadrat_Data; "
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =255
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =6600
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3240
                    Top =60
                    Width =1680
                    Height =420
                    ColumnOrder =9
                    FontSize =16
                    FontWeight =700
                    TabIndex =1
                    Name ="txtQuadrat_Number"
                    ControlSource ="Quadrat_Number"
                    FontName ="Calibri"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =60
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3540
                    Top =2160
                    Width =10440
                    Height =603
                    ColumnOrder =5
                    FontSize =12
                    TabIndex =10
                    Name ="txtQuadrat_Comments"
                    ControlSource ="Quadrat_Notes"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedTop =2160
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =2763
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3240
                            Top =1860
                            Width =2160
                            Height =300
                            FontSize =12
                            Name ="lblQuadrat_Comments"
                            Caption ="Quadrat Comments:"
                            FontName ="Calibri"
                            LayoutCachedLeft =3240
                            LayoutCachedTop =1860
                            LayoutCachedWidth =5400
                            LayoutCachedHeight =2160
                        End
                    End
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =87
                    Left =60
                    Top =3120
                    Width =7320
                    Height =3420
                    TabIndex =11
                    Name ="fsub_Quad_Seedlings"
                    SourceObject ="Form.fsub_Quad_Seedlings"
                    LinkChildFields ="Quadrat_Data_ID"
                    LinkMasterFields ="Quadrat_Data_ID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =3120
                    LayoutCachedWidth =7380
                    LayoutCachedHeight =6540
                End
                Begin OptionGroup
                    SpecialEffect =1
                    OverlapFlags =85
                    Left =60
                    Top =180
                    Width =3060
                    Height =2580
                    ColumnOrder =10
                    TabIndex =13
                    Name ="grpQuadrat_Selection"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =3120
                    LayoutCachedHeight =2760
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =240
                            Top =60
                            Width =1920
                            Height =300
                            FontSize =12
                            BackColor =15527148
                            ForeColor =255
                            Name ="lblQuadrat_Selection"
                            Caption ="Select a Quadrat"
                            FontName ="Calibri"
                            LayoutCachedLeft =240
                            LayoutCachedTop =60
                            LayoutCachedWidth =2160
                            LayoutCachedHeight =360
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =1140
                            Top =420
                            Width =900
                            Height =360
                            FontSize =10
                            OptionValue =3
                            ForeColor =0
                            Name ="tglQuad_360-13"
                            Caption ="360-13m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_360_13"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =1140
                            LayoutCachedTop =420
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =780
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =1140
                            Top =780
                            Width =900
                            Height =360
                            FontSize =10
                            TabIndex =1
                            OptionValue =2
                            ForeColor =0
                            Name ="tglQuad_360-8"
                            Caption ="360-8m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_360_8"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =1140
                            LayoutCachedTop =780
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =1140
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =300
                            Top =1020
                            Width =839
                            Height =360
                            FontSize =10
                            TabIndex =2
                            OptionValue =12
                            ForeColor =0
                            Name ="tglQuad_300"
                            Caption ="300"
                            FontName ="Calibri"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =300
                            LayoutCachedTop =1020
                            LayoutCachedWidth =1139
                            LayoutCachedHeight =1380
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =2040
                            Top =1020
                            Width =839
                            Height =360
                            FontSize =10
                            TabIndex =3
                            OptionValue =4
                            ForeColor =0
                            Name ="tglQuad_60"
                            Caption ="60"
                            FontName ="Calibri"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =2040
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2879
                            LayoutCachedHeight =1380
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =1140
                            Top =1140
                            Width =900
                            Height =360
                            FontSize =10
                            TabIndex =4
                            OptionValue =1
                            ForeColor =0
                            Name ="tglQuad_360-3"
                            Caption ="360-3m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_360_3"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =1140
                            LayoutCachedTop =1140
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =1500
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =600
                            Top =1500
                            Width =839
                            Height =360
                            FontSize =10
                            TabIndex =5
                            OptionValue =9
                            ForeColor =0
                            Name ="tglQuad_240-3"
                            Caption ="240-3m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_240_3"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =600
                            LayoutCachedTop =1500
                            LayoutCachedWidth =1439
                            LayoutCachedHeight =1860
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =1860
                            Top =1500
                            Width =839
                            Height =360
                            FontSize =10
                            TabIndex =6
                            OptionValue =5
                            ForeColor =0
                            Name ="tglQuad_120-3"
                            Caption ="120-3m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_120_3"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =1860
                            LayoutCachedTop =1500
                            LayoutCachedWidth =2699
                            LayoutCachedHeight =1860
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =360
                            Top =1860
                            Width =839
                            Height =360
                            FontSize =10
                            TabIndex =7
                            OptionValue =10
                            ForeColor =0
                            Name ="tglQuad_240-8"
                            Caption ="240-8m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_240_8"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =360
                            LayoutCachedTop =1860
                            LayoutCachedWidth =1199
                            LayoutCachedHeight =2220
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =95
                            Left =2040
                            Top =1860
                            Width =839
                            Height =360
                            FontSize =10
                            TabIndex =8
                            OptionValue =6
                            ForeColor =0
                            Name ="tglQuad_120-8"
                            Caption ="120-8m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_120_8"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =2040
                            LayoutCachedTop =1860
                            LayoutCachedWidth =2879
                            LayoutCachedHeight =2220
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1200
                            Top =1980
                            Width =840
                            Height =360
                            FontSize =10
                            TabIndex =9
                            OptionValue =8
                            ForeColor =0
                            Name ="tglQuad_180"
                            Caption ="180"
                            FontName ="Calibri"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =1200
                            LayoutCachedTop =1980
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =2340
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =120
                            Top =2220
                            Width =840
                            Height =360
                            FontSize =10
                            TabIndex =10
                            OptionValue =11
                            ForeColor =0
                            Name ="tglQuad_240-13"
                            Caption ="240-13m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_240_13"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =120
                            LayoutCachedTop =2220
                            LayoutCachedWidth =960
                            LayoutCachedHeight =2580
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2160
                            Top =2220
                            Width =840
                            Height =360
                            FontSize =10
                            TabIndex =11
                            OptionValue =7
                            ForeColor =0
                            Name ="tglQuad_120-13"
                            Caption ="120-13m"
                            FontName ="Calibri"
                            EventProcPrefix ="tglQuad_120_13"
                            LeftPadding =60
                            RightPadding =75
                            BottomPadding =120

                            LayoutCachedLeft =2160
                            LayoutCachedTop =2220
                            LayoutCachedWidth =3000
                            LayoutCachedHeight =2580
                            ForeThemeColorIndex =0
                            UseTheme =1
                            Shape =1
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
                            QuickStyle =23
                            QuickStyleMask =-5
                            WebImagePaddingLeft =4
                            WebImagePaddingTop =2
                            WebImagePaddingRight =4
                            WebImagePaddingBottom =7
                            Overlaps =1
                        End
                    End
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7440
                    Top =3120
                    Width =6360
                    Height =3420
                    TabIndex =12
                    Name ="fsub_Quad_Herbaceous"
                    SourceObject ="Form.fsub_Quad_Herbaceous"
                    LinkChildFields ="Quadrat_Data_ID"
                    LinkMasterFields ="Quadrat_Data_ID"

                    LayoutCachedLeft =7440
                    LayoutCachedTop =3120
                    LayoutCachedWidth =13800
                    LayoutCachedHeight =6540
                End
                Begin Label
                    OverlapFlags =93
                    Left =180
                    Top =2820
                    Width =1680
                    Height =300
                    FontSize =12
                    FontWeight =700
                    Name ="lblQuad_Seedlings"
                    Caption ="Seedlings"
                    FontName ="Calibri"
                    LayoutCachedLeft =180
                    LayoutCachedTop =2820
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =3120
                End
                Begin Label
                    OverlapFlags =85
                    Left =7440
                    Top =2820
                    Width =3720
                    Height =270
                    FontSize =12
                    FontWeight =700
                    Name ="lblQuad_Herbaceous"
                    Caption ="Targeted Herbaceous Vegetation"
                    FontName ="Calibri"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =2820
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =3090
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =3240
                    Top =1260
                    Width =10740
                    Height =540
                    Name ="shpVegetation_Cover"
                    LayoutCachedLeft =3240
                    LayoutCachedTop =1260
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =1800
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =3240
                    Top =600
                    Width =10740
                    Height =540
                    Name ="shpFloor_Condition"
                    LayoutCachedLeft =3240
                    LayoutCachedTop =600
                    LayoutCachedWidth =13980
                    LayoutCachedHeight =1140
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =3360
                    Top =660
                    Width =1926
                    Height =366
                    FontSize =12
                    TabIndex =14
                    ForeColor =6108695
                    Name ="cmdOpen_Popup_Floor_Condition"
                    Caption ="Floor Condition %"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =660
                    LayoutCachedWidth =5286
                    LayoutCachedHeight =1026
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =3360
                    Top =1320
                    Width =1506
                    Height =366
                    FontSize =12
                    TabIndex =16
                    ForeColor =6108695
                    Name ="cmdOpen_Popup_Veg_Cover"
                    Caption ="Veg Cover %"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =3360
                    LayoutCachedTop =1320
                    LayoutCachedWidth =4866
                    LayoutCachedHeight =1686
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11040
                    Top =720
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =15
                    Name ="txtPercent_FWD"
                    ControlSource ="Percent_Fine_Woody_Debris"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000009e000000010000000100000000000000000000001e00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004600570044005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =11040
                    LayoutCachedTop =720
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =1020
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001d0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004600570044005d0029003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =10380
                            Top =720
                            Width =600
                            Height =300
                            FontSize =12
                            Name ="Label121"
                            Caption ="FWD"
                            FontName ="Calibri"
                            LayoutCachedLeft =10380
                            LayoutCachedTop =720
                            LayoutCachedWidth =10980
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =720
                    Width =720
                    Height =300
                    FontSize =12
                    Name ="txtPercent_Trees"
                    ControlSource ="Percent_Trees"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a2000000010000000100000000000000000000002000000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f00540072006500650073005d0029003d0054007200750065000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =720
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1020
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001f0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f00540072006500650073005d0029003d00540072007500650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =5340
                            Top =720
                            Width =600
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Trees"
                            Caption ="Trees"
                            FontName ="Calibri"
                            LayoutCachedLeft =5340
                            LayoutCachedTop =720
                            LayoutCachedWidth =5940
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =7740
                    Top =720
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =2
                    Name ="txtPercent_Rock"
                    ControlSource ="Percent_Rock"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a0000000010000000100000000000000000000001f00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f0052006f0063006b005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =7740
                    LayoutCachedTop =720
                    LayoutCachedWidth =8460
                    LayoutCachedHeight =1020
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001e0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f0052006f0063006b005d0029003d005400720075006500000000000000 ,
                        0x000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =7140
                            Top =720
                            Width =540
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Rock"
                            Caption ="Rock"
                            FontName ="Calibri"
                            LayoutCachedLeft =7140
                            LayoutCachedTop =720
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9360
                    Top =720
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =3
                    Name ="txtPercent_CWD"
                    ControlSource ="Percent_Woody_Debris"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000009e000000010000000100000000000000000000001e00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004300570044005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =9360
                    LayoutCachedTop =720
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1020
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001d0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004300570044005d0029003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =8700
                            Top =720
                            Width =600
                            Height =300
                            FontSize =12
                            Name ="lblPercent_CWD"
                            Caption ="CWD"
                            FontName ="Calibri"
                            LayoutCachedLeft =8700
                            LayoutCachedTop =720
                            LayoutCachedWidth =9300
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13200
                    Top =720
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =5
                    Name ="txtPercent_Other"
                    ControlSource ="Percent_Other"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a2000000010000000100000000000000000000002000000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004f0074006800650072005d0029003d0054007200750065000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =13200
                    LayoutCachedTop =720
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =1020
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001f0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004f0074006800650072005d0029003d00540072007500650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =12480
                            Top =720
                            Width =660
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Other"
                            Caption ="Other"
                            FontName ="Calibri"
                            LayoutCachedLeft =12480
                            LayoutCachedTop =720
                            LayoutCachedWidth =13140
                            LayoutCachedHeight =1020
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13200
                    Top =1380
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =4
                    Name ="txtPercent_Bryophytes"
                    ControlSource ="Percent_Bryophytes"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000ac000000010000000100000000000000000000002500000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004200720079006f007000680079007400650073005d0029003d00 ,
                        0x540072007500650000000000
                    End

                    LayoutCachedLeft =13200
                    LayoutCachedTop =1380
                    LayoutCachedWidth =13920
                    LayoutCachedHeight =1680
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500240000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004200720079006f007000680079007400650073005d0029003d005400 ,
                        0x720075006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =12000
                            Top =1380
                            Width =1140
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Bryophytes"
                            Caption ="Bryophytes"
                            FontName ="Calibri"
                            LayoutCachedLeft =12000
                            LayoutCachedTop =1380
                            LayoutCachedWidth =13140
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =11070
                    Top =1380
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =9
                    Name ="txtPercent_Ferns"
                    ControlSource ="Percent_Ferns"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a2000000010000000100000000000000000000002000000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004600650072006e0073005d0029003d0054007200750065000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =11070
                    LayoutCachedTop =1380
                    LayoutCachedWidth =11790
                    LayoutCachedHeight =1680
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001f0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004600650072006e0073005d0029003d00540072007500650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =10380
                            Top =1380
                            Width =600
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Ferns"
                            Caption ="Ferns"
                            FontName ="Calibri"
                            LayoutCachedLeft =10380
                            LayoutCachedTop =1380
                            LayoutCachedWidth =10980
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =9375
                    Top =1380
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =8
                    Name ="txtPercent_Herbs"
                    ControlSource ="Percent_Herbs"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a2000000010000000100000000000000000000002000000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f00480065007200620073005d0029003d0054007200750065000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =9375
                    LayoutCachedTop =1380
                    LayoutCachedWidth =10095
                    LayoutCachedHeight =1680
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001f0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f00480065007200620073005d0029003d00540072007500650000000000 ,
                        0x0000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =8700
                            Top =1380
                            Width =600
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Herbs"
                            Caption ="Herbs"
                            FontName ="Calibri"
                            LayoutCachedLeft =8700
                            LayoutCachedTop =1380
                            LayoutCachedWidth =9300
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =7755
                    Top =1380
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =7
                    Name ="txtPercent_Sedges"
                    ControlSource ="Percent_Sedges"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a4000000010000000100000000000000000000002100000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f005300650064006700650073005d0029003d005400720075006500 ,
                        0x00000000
                    End

                    LayoutCachedLeft =7755
                    LayoutCachedTop =1380
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =1680
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500200000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f005300650064006700650073005d0029003d0054007200750065000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =6900
                            Top =1380
                            Width =780
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Sedges"
                            Caption ="Sedges"
                            FontName ="Calibri"
                            LayoutCachedLeft =6900
                            LayoutCachedTop =1380
                            LayoutCachedWidth =7680
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =6000
                    Top =1380
                    Width =720
                    Height =300
                    FontSize =12
                    TabIndex =6
                    Name ="txtPercent_Grasses"
                    ControlSource ="Percent_Grasses"
                    ValidationRule ="Is Null Or Between 0 And 1"
                    ValidationText ="Enter % cover values between 0 and 100% (inclusive)"
                    DefaultValue ="Null"
                    FontName ="Calibri"
                    OnLostFocus ="=ValidPct([Screen].[ActiveControl],True)"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x01000000a6000000010000000100000000000000000000002200000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f0047007200610073007300650073005d0029003d00540072007500 ,
                        0x650000000000
                    End

                    LayoutCachedLeft =6000
                    LayoutCachedTop =1380
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1680
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500210000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f0047007200610073007300650073005d0029003d005400720075006500 ,
                        0x000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =5100
                            Top =1380
                            Width =840
                            Height =300
                            FontSize =12
                            Name ="lblPercent_Grasses"
                            Caption ="Grasses"
                            FontName ="Calibri"
                            LayoutCachedLeft =5100
                            LayoutCachedTop =1380
                            LayoutCachedWidth =5940
                            LayoutCachedHeight =1680
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =4
                    ListWidth =1800
                    Left =3240
                    Top =2160
                    Width =240
                    Height =600
                    FontSize =12
                    TabIndex =17
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"192\""
                    Name ="cboQuick_Comment"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Quadrat Comments\")) ORDER BY"
                        " tlu_Enumerations.[Sort_Order];"
                    ColumnWidths ="0;1800;0;0"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =3240
                    LayoutCachedTop =2160
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =2760
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
' MODULE:       fsub_Quadrats
' Level:        Application module
' Version:      1.01
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/22/2018 - 1.01 - added documentation, error handling
' =================================

' ---------------------------------
'  Declarations
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
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/22/2018 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler
    
    Dim strQuadrat As String
    Me!txtQuadrat_Number.DefaultValue = ""
    Me!txtQuadrat_Number.Requery
    Me!lblQuadrat_Selection.ForeColor = 255
    
    Me!fsub_Quad_Herbaceous.Visible = False
    Me!fsub_Quad_Seedlings.Visible = False
    Me!txtQuadrat_Number.Visible = False
    Me!txtQuadrat_Comments.Visible = False
    Me!txtPercent_Trees.Visible = False
    Me!txtPercent_Bryophytes.Visible = False
    Me!txtPercent_CWD.Visible = False
    Me!txtPercent_FWD.Visible = False
    Me!txtPercent_Rock.Visible = False
    Me!txtPercent_Other.Visible = False
    Me!txtPercent_Grasses.Visible = False
    Me!txtPercent_Sedges.Visible = False
    Me!txtPercent_Herbs.Visible = False
    Me!txtPercent_Ferns.Visible = False
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  form actions before updates
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
    
    If Me.NewRecord Then
        If GetDataType("tbl_Quadrat_Data", "Quadrat_Data_ID") = dbText Then
            Me!Quadrat_Data_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click
' ----------------

' ---------------------------------
' SUB:          txtPercent_Trees_Click
' Description:  textbox click actions
' Requires:     Keypad Utils module
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Trees_Click()
On Error GoTo Err_Handler
    
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'set keypad form to launch & control on this form to be updated by it
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Trees"
'
'    'launch keypad
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    'launch keypad
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Trees"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Trees_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_Rock_Click
' Description:  textbox click actions
' Requires:     Keypad Utils module
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Rock_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Rock"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    'launch keypad
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Rock"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Rock_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_CWD_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_CWD_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_CWD"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_CWD"
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_CWD_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_FWD_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_FWD_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_FWD"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_FWD"
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_FWD_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_Other_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Other_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Other"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Other"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Other_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_Grasses_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Grasses_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Grasses"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Grasses"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Grasses_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_Sedges_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Sedges_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Sedges"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Sedges"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Sedges_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_Herbs_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Herbs_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Herbs"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Herbs"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Herbs_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_Ferns_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Ferns_Click()
On Error GoTo Err_Handler
    
'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Ferns"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Ferns"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Ferns_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          txtPercent_Bryophytes_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub txtPercent_Bryophytes_Click()
On Error GoTo Err_Handler

'    'This routine requires the presence of the Keypad_Utils module.
'    Dim strKeypadFormName As String
'    Dim strControlToUpdate As String
'    Dim frmFormToUpdate As Form
'
'    'The two lines below should be changed to reflect the name of the keypad to open
'    '    and the name of the control to be updated.
'    strKeypadFormName = "frm_Pad_Percent"
'    strControlToUpdate = "txtPercent_Bryophytes"
'    'The lines below should not usually be edited.
'    Set frmFormToUpdate = Me
'    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)
    
    LaunchKeypad Me, "frm_Pad_Percent", "txtPercent_Bryophytes"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - txtPercent_Bryophytes_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Form Opening
' ----------------
' ---------------------------------
' SUB:          cmdOpen_Popup_Floor_Condition_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub cmdOpen_Popup_Floor_Condition_Click()
On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Floor_and_Cover"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdOpen_Popup_Floor_Condition_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cmdOpen_Popup_Veg_Cover_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub cmdOpen_Popup_Veg_Cover_Click()
On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Floor_and_Cover"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdOpen_Popup_Veg_Cover_Click[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  After Update
' ----------------
' ---------------------------------
' SUB:          cboQuick_Comment_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling
' ---------------------------------
Private Sub cboQuick_Comment_AfterUpdate()
On Error GoTo Err_Handler
    
    Me.txtQuadrat_Comments = LTrim(Me.txtQuadrat_Comments & " " & Me.cboQuick_Comment)
    Me.txtQuadrat_Comments.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cboQuick_Comment_AfterUpdate[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          grpQuadrat_Selection_AfterUpdate
' Description:  group after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC - 4/22/2018 - added documentation, error handling,
'                                 adjusted
' ---------------------------------
Private Sub grpQuadrat_Selection_AfterUpdate()
On Error GoTo Err_Handler
    
    Dim strQuadrat As String
    Dim aryPct() As Variant
    Dim strPct As Variant
    Dim ctlName As String
    
    'percents
    aryPct = Array("Trees", "Bryophytes", "CWD", "FWD", "Rock", "Other", _
                "Grasses", "Sedges", "Herbs", "Ferns")
    
    'determine quadrat
    Select Case Me!grpQuadrat_Selection.Value
        Case 1
            strQuadrat = "360-3m"
        Case 2
            strQuadrat = "360-8m"
        Case 3
            strQuadrat = "360-13m"
        Case 4
            strQuadrat = "60"
        Case 5
            strQuadrat = "120-3m"
        Case 6
            strQuadrat = "120-8m"
        Case 7
            strQuadrat = "120-13m"
        Case 8
            strQuadrat = "180"
        Case 9
            strQuadrat = "240-3m"
        Case 10
            strQuadrat = "240-8m"
        Case 11
            strQuadrat = "240-13m"
        Case 12
            strQuadrat = "300"
    End Select
    
    Me!lblQuadrat_Selection.ForeColor = 0
    Me!fsub_Quad_Herbaceous.Visible = True
    Me!fsub_Quad_Seedlings.Visible = True
    Me!txtQuadrat_Comments.Visible = True
    Me!txtQuadrat_Number.Visible = True
    
'    Me!txtPercent_Trees.Visible = True
'    Me!txtPercent_Bryophytes.Visible = True
'    Me!txtPercent_CWD.Visible = True
'    Me!txtPercent_FWD.Visible = True
'    Me!txtPercent_Rock.Visible = True
'    Me!txtPercent_Other.Visible = True
'    Me!txtPercent_Grasses.Visible = True
'    Me!txtPercent_Sedges.Visible = True
'    Me!txtPercent_Herbs.Visible = True
'    Me!txtPercent_Ferns.Visible = True
    
    For Each strPct In aryPct
        
        ctlName = "txtPercent_" & strPct
    
        With Me.Form.Controls(ctlName)
        
            .Visible = True
            .Locked = False
            .Enabled = True
            
        End With
        
    
    Next
    
    Me!fsub_Quad_Herbaceous.Locked = False
    Me!fsub_Quad_Seedlings.Locked = False
    Me!txtQuadrat_Comments.Locked = False
    
'    Me!txtPercent_Trees.Locked = False
'    Me!txtPercent_Bryophytes.Locked = False
'    Me!txtPercent_CWD.Locked = False
'    Me!txtPercent_FWD.Locked = False
'    Me!txtPercent_Rock.Locked = False
'    Me!txtPercent_Other.Locked = False
'    Me!txtPercent_Grasses.Locked = False
'    Me!txtPercent_Sedges.Locked = False
'    Me!txtPercent_Herbs.Locked = False
'    Me!txtPercent_Ferns.Locked = False
    
    Me!fsub_Quad_Herbaceous.Enabled = True
    Me!fsub_Quad_Seedlings.Enabled = True
    Me!txtQuadrat_Comments.Enabled = True
    
'    Me!txtPercent_Trees.Enabled = True
'    Me!txtPercent_Bryophytes.Enabled = True
'    Me!txtPercent_CWD.Enabled = True
'    Me!txtPercent_FWD.Enabled = True
'    Me!txtPercent_Rock.Enabled = True
'    Me!txtPercent_Other.Enabled = True
'    Me!txtPercent_Grasses.Enabled = True
'    Me!txtPercent_Sedges.Enabled = True
'    Me!txtPercent_Herbs.Enabled = True
'    Me!txtPercent_Ferns.Enabled = True
    
    'strQuadrat = Me!Frame_Quadrat_Selection.Value
    Me.txtQuadrat_Number.DefaultValue = "'" & strQuadrat & "'"
    Me.Filter = "[Quadrat_Number] = """ & strQuadrat & """ "
    Me.FilterOn = True
    'Temporary fix to save Quadrat record before entering subform
    Me!txtQuadrat_Comments.Value = Me!txtQuadrat_Comments.Value & " "
    Me!txtQuadrat_Comments.Value = Left(Me!txtQuadrat_Comments.Value, Len(Me!txtQuadrat_Comments.Value) - 1)
    'DoCmd.RunCommand acCmdSaveRecord
    'Me!txt_Percent_Trees.SetFocus
    Me!fsub_Quad_Herbaceous.Requery
    Me!fsub_Quad_Seedlings.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - grpQuadrat_Selection_AfterUpdate[fsub_Quadrats])"
    End Select
    Resume Exit_Handler
End Sub
