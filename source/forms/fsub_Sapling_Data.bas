Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13920
    DatasheetFontHeight =9
    ItemSuffix =73
    Left =3975
    Top =3675
    Right =17745
    Bottom =9990
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xd0ed4c4b94aee340
    End
    RecordSource ="tbl_Sapling_Data"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            Height =0
            BackColor =16768194
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =7140
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1619
                    Top =5340
                    Width =12239
                    Height =361
                    ColumnWidth =2055
                    FontSize =12
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =0
                    Name ="tbxComments"
                    ControlSource ="Sapling_Notes"
                    StatusBarText ="Notes about this sampling of this tree"
                    OnEnter ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000faf3e800000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4c0065006e0028005400720069006d0028005b0074006200780043006f006d00 ,
                        0x6d0065006e00740073005d00290029003e00300000000000
                    End

                    LayoutCachedLeft =1619
                    LayoutCachedTop =5340
                    LayoutCachedWidth =13858
                    LayoutCachedHeight =5701
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000faf3e8001a0000004c00 ,
                        0x65006e0028005400720069006d0028005b0074006200780043006f006d006d00 ,
                        0x65006e00740073005d00290029003e0030000000000000000000000000000000 ,
                        0x00000000000000
                    End
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =5340
                            Width =1168
                            Height =361
                            FontSize =13
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BackColor =15527148
                            Name ="lblComments"
                            Caption ="Comments"
                            LayoutCachedLeft =60
                            LayoutCachedTop =5340
                            LayoutCachedWidth =1228
                            LayoutCachedHeight =5701
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    BorderWidth =2
                    Left =60
                    Top =435
                    Width =13800
                    Height =944
                    TabIndex =3
                    BorderColor =7633277
                    Name ="fsub_Tag_Sapling"
                    SourceObject ="Form.fsub_Tag_Sapling"
                    LinkChildFields ="Tag_ID"
                    LinkMasterFields ="Tag_ID"

                    LayoutCachedLeft =60
                    LayoutCachedTop =435
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1379
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =20
                    ListWidth =5760
                    Left =2881
                    Top =60
                    Width =240
                    Height =315
                    FontSize =14
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cbxSelectUnsampledTag"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag, tbl_Tags.Tag_Status AS Class,  IIf(IsNull("
                        "[azimuth]),\"\",[Azimuth] & \" / \" & [distance] & \"m\") AS Azi_Dist, tbl_Tags."
                        "Microplot_Number AS MP  FROM (tbl_Tags  LEFT JOIN qry_Status_Sapling_Current_Eve"
                        "nt  ON tbl_Tags.Tag_ID = qry_Status_Sapling_Current_Event.Tag_ID)  LEFT JOIN qry"
                        "_Status_Tree_Current_Event  ON tbl_Tags.Tag_ID = qry_Status_Tree_Current_Event.T"
                        "ag_ID WHERE (((qry_Status_Sapling_Current_Event.Event_ID) Is Null)  AND ((tbl_Ta"
                        "gs.Location_ID)=[Forms]![frm_Events]![Location_ID])  AND ((qry_Status_Tree_Curre"
                        "nt_Event.Event_ID) Is Null)) ORDER BY tbl_Tags.Tag_Status, tbl_Tags.Tag;"
                    ColumnWidths ="0;1080;2520;1440;720"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    LayoutCachedLeft =2881
                    LayoutCachedTop =60
                    LayoutCachedWidth =3121
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =60
                            Width =2805
                            Height =315
                            FontSize =12
                            FontWeight =700
                            Name ="lblSelectTag"
                            Caption ="Select an unsampled tag ->"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =2865
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =1319
                    Top =2880
                    Width =2819
                    Height =2220
                    TabIndex =5
                    Name ="fsub_Sapling_DBH"
                    SourceObject ="Form.fsub_Sapling_DBH"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"
                    OnExit ="[Event Procedure]"

                    LayoutCachedLeft =1319
                    LayoutCachedTop =2880
                    LayoutCachedWidth =4138
                    LayoutCachedHeight =5100
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =2880
                            Width =1199
                            Height =360
                            FontSize =13
                            Name ="fsub_Tree_DBH Label"
                            Caption ="Stems (cm)"
                            EventProcPrefix ="fsub_Tree_DBH_Label"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2880
                            LayoutCachedWidth =1259
                            LayoutCachedHeight =3240
                        End
                    End
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1318
                    Top =1559
                    Width =2819
                    Height =359
                    ColumnWidth =1875
                    FontSize =13
                    TabIndex =1
                    BorderColor =0
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0063006200780054007200650065005300 ,
                        0x740061007400750073005d0029003d00540072007500650000000000
                    End
                    Name ="cbxSaplingStatus"
                    ControlSource ="Sapling_Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Tree Status\")) ORDER BY tlu_Enumer"
                        "ations.Sort_Order; "
                    ColumnWidths ="1440"
                    StatusBarText ="Health status of this specimen"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1318
                    LayoutCachedTop =1559
                    LayoutCachedWidth =4137
                    LayoutCachedHeight =1918
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001c0000004900 ,
                        0x73004e0075006c006c0028005b00630062007800540072006500650053007400 ,
                        0x61007400750073005d0029003d00540072007500650000000000000000000000 ,
                        0x0000000000000000000000
                    End
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1559
                            Width =1199
                            Height =360
                            FontSize =13
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BackColor =15527148
                            Name ="lblStatus"
                            Caption ="Status"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1559
                            LayoutCachedWidth =1259
                            LayoutCachedHeight =1919
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =1379
                    Top =5820
                    Width =12539
                    Height =1275
                    TabIndex =6
                    Name ="fsub_Tags_History_Summary"
                    SourceObject ="Form.fsub_Tags_History_Summary"
                    LinkChildFields ="Tag_ID"
                    LinkMasterFields ="Tag_ID"

                    LayoutCachedLeft =1379
                    LayoutCachedTop =5820
                    LayoutCachedWidth =13918
                    LayoutCachedHeight =7095
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =5820
                            Width =1195
                            Height =599
                            FontSize =13
                            Name ="fsub_Tags_History_Summary Label"
                            Caption ="Tag History"
                            EventProcPrefix ="fsub_Tags_History_Summary_Label"
                            LayoutCachedLeft =120
                            LayoutCachedTop =5820
                            LayoutCachedWidth =1315
                            LayoutCachedHeight =6419
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7319
                    Top =60
                    Width =2399
                    Height =300
                    FontSize =12
                    TabIndex =7
                    ForeColor =0
                    Name ="btnTagNewSpecimen"
                    Caption ="Tag New Specimen"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7319
                    LayoutCachedTop =60
                    LayoutCachedWidth =9718
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =16236067
                    HoverThemeColorIndex =6
                    HoverTint =80.0
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
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =6900
                    Top =60
                    Width =270
                    Height =285
                    FontWeight =700
                    ForeColor =3751056
                    Name ="lblOr2"
                    Caption ="or"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =60
                    LayoutCachedWidth =7170
                    LayoutCachedHeight =345
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1319
                    Top =1980
                    Width =2819
                    Height =359
                    FontSize =13
                    TabIndex =8
                    BoundColumn =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"192\""
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0063006200780048006100620069007400 ,
                        0x5d0029003d00540072007500650000000000
                    End
                    Name ="cbxHabit"
                    ControlSource ="Habit"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Description, tlu_Enumerations.Enum_Code FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Habit\")) ORDER BY tlu_Enumer"
                        "ations.Sort_Order;"
                    ColumnWidths ="1080;0;0;0"
                    StatusBarText ="Growth Habit (Shrub or Tree) as sampled"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1319
                    LayoutCachedTop =1980
                    LayoutCachedWidth =4138
                    LayoutCachedHeight =2339
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500170000004900 ,
                        0x73004e0075006c006c0028005b00630062007800480061006200690074005d00 ,
                        0x29003d0054007200750065000000000000000000000000000000000000000000 ,
                        0x00
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =1980
                            Width =1199
                            Height =360
                            FontSize =13
                            Name ="lblHabit"
                            Caption ="Habit"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1980
                            LayoutCachedWidth =1259
                            LayoutCachedHeight =2340
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =8220
                    Top =1500
                    Width =900
                    Height =360
                    FontSize =8
                    BackColor =14276557
                    Name ="lblDeerBrowse"
                    Caption ="Deer Browse"
                    LayoutCachedLeft =8220
                    LayoutCachedTop =1500
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =1860
                    BackThemeColorIndex =3
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =20
                    ListWidth =6840
                    Left =6570
                    Top =60
                    Width =240
                    Height =315
                    FontSize =14
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cbxSelectSampledTag"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag, tbl_Tags.Microplot_Number AS MP, qry_Statu"
                        "s_Sapling_Current_Event.Sapling_Status, tlu_Plants.Latin_Name FROM (tbl_Tags LEF"
                        "T JOIN tlu_Plants ON tbl_Tags.TSN = tlu_Plants.TSN) INNER JOIN qry_Status_Saplin"
                        "g_Current_Event ON tbl_Tags.Tag_ID = qry_Status_Sapling_Current_Event.Tag_ID WHE"
                        "RE (((tbl_Tags.Location_ID)=[Forms]![frm_Events]![Location_ID])) ORDER BY tbl_Ta"
                        "gs.Tag;"
                    ColumnWidths ="0;1080;720;2160;2880"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    LayoutCachedLeft =6570
                    LayoutCachedTop =60
                    LayoutCachedWidth =6810
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =3600
                            Top =60
                            Width =2940
                            Height =315
                            FontSize =12
                            FontWeight =700
                            Name ="lblSelectSample"
                            Caption ="Select an existing sample ->"
                            LayoutCachedLeft =3600
                            LayoutCachedTop =60
                            LayoutCachedWidth =6540
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =3240
                    Top =60
                    Width =270
                    Height =285
                    FontWeight =700
                    ForeColor =3751056
                    Name ="lblOr1"
                    Caption ="or"
                    LayoutCachedLeft =3240
                    LayoutCachedTop =60
                    LayoutCachedWidth =3510
                    LayoutCachedHeight =345
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =13440
                    Top =1440
                    Width =426
                    Height =396
                    TabIndex =11
                    ForeColor =0
                    Name ="btnDeleteSample"
                    Caption ="Command73"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddddddddddddddddd ,
                        0xdddddddddddddddddddd177ddddd77dd1ddd1177dddd17dd11dd7117ddd71ddd ,
                        0x111dd1177d117ddd1111d7117711dddd11111d11111ddddd1111dd71117ddddd ,
                        0x111d77111177dddd11d711dd71177ddd1dddddddd71177ddddddddddddd11ddd ,
                        0xdddddddddddddddd
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete This Sample"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13440
                    LayoutCachedTop =1440
                    LayoutCachedWidth =13866
                    LayoutCachedHeight =1836
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =4819194
                    HoverThemeColorIndex =5
                    HoverTint =80.0
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =10259
                    Top =60
                    Width =360
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =12
                    ForeColor =0
                    Name ="btnOpenFormTagTransitions"
                    Caption ="?"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10259
                    LayoutCachedTop =60
                    LayoutCachedWidth =10619
                    LayoutCachedHeight =360
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =16236067
                    HoverThemeColorIndex =6
                    HoverTint =80.0
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
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =9840
                    Top =60
                    Width =270
                    Height =285
                    FontWeight =700
                    ForeColor =3751056
                    Name ="lblOr3"
                    Caption ="or"
                    LayoutCachedLeft =9840
                    LayoutCachedTop =60
                    LayoutCachedWidth =10110
                    LayoutCachedHeight =345
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BorderWidth =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =25
                    ListWidth =4320
                    Left =12420
                    Top =1500
                    Width =240
                    Height =360
                    FontSize =12
                    TabIndex =13
                    BackColor =-2147483643
                    BorderColor =5026082
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbxBrowsePick"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Browse_Status\")) ORDER BY tl"
                        "u_Enumerations.Sort_Order;"
                    ColumnWidths ="1080;3240;0;0;0"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="Quick Find"
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =12420
                    LayoutCachedTop =1500
                    LayoutCachedWidth =12660
                    LayoutCachedHeight =1860
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =11040
                    Top =1500
                    Width =600
                    Height =358
                    FontSize =13
                    TabIndex =10
                    Name ="tbxBrowsable"
                    ControlSource ="Browsable"
                    StatusBarText ="This sapling was browseable (leaves below 2 meters)"
                    OnEnter ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00420072006f0077007300610062006c00 ,
                        0x65005d0029003d00540072007500650000000000
                    End

                    LayoutCachedLeft =11040
                    LayoutCachedTop =1500
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =1858
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500180000004900 ,
                        0x73004e0075006c006c0028005b00420072006f0077007300610062006c006500 ,
                        0x5d0029003d005400720075006500000000000000000000000000000000000000 ,
                        0x000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =2
                            Left =9120
                            Top =1560
                            Width =1860
                            Height =300
                            FontSize =10
                            Name ="lblBrowse"
                            Caption =" Browsable / Browsed"
                            LayoutCachedLeft =9120
                            LayoutCachedTop =1560
                            LayoutCachedWidth =10980
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =11820
                    Top =1500
                    Width =540
                    Height =358
                    FontSize =13
                    TabIndex =9
                    Name ="tbxBrowsed"
                    ControlSource ="Browsed"
                    StatusBarText ="Deer browse was noticable on this sapling"
                    OnEnter ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00420072006f0077007300650064005d00 ,
                        0x29003d00540072007500650000000000
                    End

                    LayoutCachedLeft =11820
                    LayoutCachedTop =1500
                    LayoutCachedWidth =12360
                    LayoutCachedHeight =1858
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500160000004900 ,
                        0x73004e0075006c006c0028005b00420072006f0077007300650064005d002900 ,
                        0x3d005400720075006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =11640
                            Top =1500
                            Width =180
                            Height =360
                            FontSize =13
                            Name ="lblSlash"
                            Caption ="/"
                            LayoutCachedLeft =11640
                            LayoutCachedTop =1500
                            LayoutCachedWidth =11820
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =7680
                    Top =1980
                    Width =210
                    Height =209
                    TabIndex =14
                    Name ="chkConditionsChecked"
                    ControlSource ="Conditions_Checked"
                    StatusBarText ="This tree was checked for disease/damage conditions"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7890
                    LayoutCachedHeight =2189
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =7680
                    Top =1980
                    Width =210
                    Height =209
                    TabIndex =15
                    Name ="chkFoliageConditionsChecked"
                    ControlSource ="Foliage_Conditions_Checked"
                    StatusBarText ="This tree was checked for foliage conditions"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7890
                    LayoutCachedHeight =2189
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =4380
                    Top =1440
                    Width =1200
                    Height =480
                    TabIndex =16
                    BackColor =15527148
                    BorderColor =0
                    Name ="tbxHighlightVines"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b00560069006e006500730043006800650063006b0065006400 ,
                        0x5d003c003e00540072007500650000000000
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1440
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500170000005b00 ,
                        0x630068006b00560069006e006500730043006800650063006b00650064005d00 ,
                        0x3c003e0054007200750065000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =223
                    Left =4410
                    Top =2220
                    Width =3539
                    Height =2400
                    TabIndex =17
                    Name ="fsub_Sapling_Foliage_Conditions"
                    SourceObject ="Form.fsub_Sapling_Foliage_Conditions"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =4410
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7949
                    LayoutCachedHeight =4620
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =4440
                            Top =1980
                            Width =1620
                            Height =306
                            FontSize =10
                            Name ="lblFoliageConditions"
                            Caption ="Foliage Conditions"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =1980
                            LayoutCachedWidth =6060
                            LayoutCachedHeight =2286
                        End
                    End
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =204
                    Left =4440
                    Top =1920
                    Width =2106
                    Height =306
                    FontSize =10
                    TabIndex =18
                    ForeColor =6108695
                    Name ="btnOpenFormConditionsAndPests"
                    Caption ="Conditions and Pests"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =4440
                    LayoutCachedTop =1920
                    LayoutCachedWidth =6546
                    LayoutCachedHeight =2226
                    Alignment =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =7680
                    Top =1980
                    Width =210
                    Height =209
                    TabIndex =19
                    Name ="chkVinesChecked"
                    ControlSource ="Vines_Checked"
                    StatusBarText ="This tree was checked for vines"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7890
                    LayoutCachedHeight =2189
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =255
                    Left =4410
                    Top =2220
                    Width =3539
                    Height =2400
                    TabIndex =20
                    Name ="fsub_Sapling_Vines"
                    SourceObject ="Form.fsub_Sapling_Vines"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =4410
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7949
                    LayoutCachedHeight =4620
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =223
                            Left =4440
                            Top =1986
                            Width =600
                            Height =306
                            FontSize =10
                            Name ="lblVines"
                            Caption ="Vines"
                            LayoutCachedLeft =4440
                            LayoutCachedTop =1986
                            LayoutCachedWidth =5040
                            LayoutCachedHeight =2292
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =4410
                    Top =2220
                    Width =3539
                    Height =2400
                    TabIndex =21
                    Name ="fsub_Sapling_Conditions"
                    SourceObject ="Form.fsub_Sapling_Conditions"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =4410
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7949
                    LayoutCachedHeight =4620
                End
                Begin TextBox
                    Locked = NotDefault
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Top =1920
                    Width =1035
                    Height =285
                    FontSize =10
                    TabIndex =22
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    BorderColor =0
                    Name ="lblCompleted"
                    ControlSource ="=\"Completed\""

                    LayoutCachedLeft =6600
                    LayoutCachedTop =1920
                    LayoutCachedWidth =7635
                    LayoutCachedHeight =2205
                End
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =4
                    BorderWidth =3
                    Left =8160
                    Top =2220
                    Width =5700
                    Height =2400
                    TabIndex =23
                    Name ="fsub_Conditions_Summary"
                    SourceObject ="Form.fsub_Sapling_All_Conditions"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =8160
                    LayoutCachedTop =2220
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =4620
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =8160
                            Top =1920
                            Width =5100
                            Height =300
                            FontSize =13
                            FontWeight =700
                            Name ="lblTree_All_Conditions"
                            Caption ="Summary of all vines and conditions"
                            LayoutCachedLeft =8160
                            LayoutCachedTop =1920
                            LayoutCachedWidth =13260
                            LayoutCachedHeight =2220
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4440
                    Top =1500
                    Width =1080
                    FontSize =12
                    TabIndex =24
                    ForeColor =0
                    Name ="btnShowVines"
                    Caption ="Vines"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =4440
                    LayoutCachedTop =1500
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =16236067
                    HoverThemeColorIndex =6
                    HoverTint =80.0
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
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =5580
                    Top =1440
                    Width =1200
                    Height =480
                    TabIndex =25
                    BackColor =15527148
                    BorderColor =0
                    Name ="tbxHighlightCondition"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b0043006f006e0064006900740069006f006e00730043006800 ,
                        0x650063006b00650064005d003c003e00540072007500650000000000
                    End

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1440
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001c0000005b00 ,
                        0x630068006b0043006f006e0064006900740069006f006e007300430068006500 ,
                        0x63006b00650064005d003c003e00540072007500650000000000000000000000 ,
                        0x0000000000000000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =5640
                    Top =1500
                    Width =1080
                    FontSize =12
                    TabIndex =26
                    ForeColor =0
                    Name ="btnShowCondition"
                    Caption ="Condition"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =5640
                    LayoutCachedTop =1500
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =16236067
                    HoverThemeColorIndex =6
                    HoverTint =80.0
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
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =95
                    IMESentenceMode =3
                    Left =6780
                    Top =1440
                    Width =1200
                    Height =480
                    TabIndex =27
                    BackColor =15527148
                    BorderColor =0
                    Name ="tbxHighlightFoliage"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x01000000aa000000010000000100000000000000000000002400000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b0046006f006c00690061006700650043006f006e0064006900 ,
                        0x740069006f006e00730043006800650063006b00650064005d003c003e005400 ,
                        0x72007500650000000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedTop =1440
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500230000005b00 ,
                        0x630068006b0046006f006c00690061006700650043006f006e00640069007400 ,
                        0x69006f006e00730043006800650063006b00650064005d003c003e0054007200 ,
                        0x75006500000000000000000000000000000000000000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =6840
                    Top =1500
                    Width =1080
                    FontSize =12
                    TabIndex =28
                    ForeColor =0
                    Name ="btnShowFoliage"
                    Caption ="Foliage"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6840
                    LayoutCachedTop =1500
                    LayoutCachedWidth =7920
                    LayoutCachedHeight =1860
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =16236067
                    HoverThemeColorIndex =6
                    HoverTint =80.0
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
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =5760
                    Left =1320
                    Top =5340
                    Width =240
                    Height =360
                    FontSize =12
                    TabIndex =29
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxQuickComment"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Tree Comments\")) ORDER BY tlu_Enumerations.Sort_Order;"
                    ColumnWidths ="5760"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =5340
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =5700
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1319
                    Top =2400
                    Width =2819
                    Height =359
                    FontSize =13
                    TabIndex =30
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0063006200780048006100620069007400 ,
                        0x5d0029003d00540072007500650000000000
                    End
                    Name ="cbxSapVigor"
                    ControlSource ="SaplingVigor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tluTreeVigor.TreeVigorCode, tluTreeVigor.TreeVigorClass FROM tluTreeVigor"
                        ";"
                    ColumnWidths ="360;2160"
                    StatusBarText ="Growth Habit (Shrub or Tree) as sampled"
                    AllowValueListEdits =0

                    LayoutCachedLeft =1319
                    LayoutCachedTop =2400
                    LayoutCachedWidth =4138
                    LayoutCachedHeight =2759
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500170000004900 ,
                        0x73004e0075006c006c0028005b00630062007800480061006200690074005d00 ,
                        0x29003d0054007200750065000000000000000000000000000000000000000000 ,
                        0x00
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2400
                            Width =1199
                            Height =360
                            FontSize =13
                            Name ="lblVigor"
                            Caption ="Vigor"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2400
                            LayoutCachedWidth =1259
                            LayoutCachedHeight =2760
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1080
                    Top =3420
                    Width =210
                    Height =209
                    TabIndex =31
                    Name ="chkDBHCheck"
                    StatusBarText ="Check if DBH was double checked"
                    DefaultValue ="0"
                    ControlTipText ="Check if DBH was double checked"

                    LayoutCachedLeft =1080
                    LayoutCachedTop =3420
                    LayoutCachedWidth =1290
                    LayoutCachedHeight =3629
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =180
                    Top =3360
                    Width =855
                    Height =780
                    FontSize =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="lblDBHCheck"
                    Caption ="DBH Double Checked?"
                    ControlTipText ="Was DBH double checked?"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3360
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =4140
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
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
' FORM:         fsub_Sapling_Data
' Level:        Application report
' Version:      1.02
'
' Description:  Form related functions & procedures for application
' Requires:     Keypad Utils module
'
' Source/date:  Bonnie Campbell, April 3, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/3/2018 - 1.01 - added documentation, error handling
'               BLC   - 4/9/2018 - 1.02 - updated TreeStatus > SaplingStatus
'                                         updated checkbox naming (removed _)
'                                         added tag vs. sapling status check
' =================================

' ---------------------------------
'  Properties
' ---------------------------------
Public SaplingStatus As String

' ---------------------------------
'  Events
' ---------------------------------

' ---------------------------------
' SUB:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'   BLC - 4/9/2018 - check tag status vs sapling status
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler
    
'    If Me!cbxHabit = "Tree" Then
'        Me!fsub_Sapling_DBH.Visible = True
'    Else
'        Me!fsub_Sapling_DBH.Visible = False
'    End If

    'compare status
    CheckTagStatus "Sapling"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
    
    If Me.NewRecord Then
        If GetDataType("tbl_Sapling_Data", "Sapling_Data_ID") = dbText Then
            Me!Sapling_Data_ID = fxnGUIDGen
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click Events
' ----------------

' ---------------------------------
' SUB:          btnOpenFormTagTransitions_Click
' Description:  open form tag transitions button actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation,
'                    renamed cmdOpen_Form_Tag_Transitions > btnOpenFormTagTransitions
' ---------------------------------
Private Sub btnOpenFormTagTransitions_Click()
On Error GoTo Err_Handler

    Dim strDocName As String
    Dim strLinkCriteria As String
    
    strDocName = "frm_Popup_Tag_Transitions"
    DoCmd.OpenForm strDocName, , , strLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenFormTagTransitions_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Show/Open Form Events
' ----------------

' ---------------------------------
' SUB:          btnShowVines_Click
' Description:  show vines button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cmdShow_Vines > btnShowVines
'   BLC - 4/9/2018 - updated checkbox naming (removed _)
' ---------------------------------
Private Sub btnShowVines_Click()
On Error GoTo Err_Handler
    
    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, True
    DoCmd.SetProperty "chkVinesChecked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Sapling_Vines", acPropertyVisible, True
    DoCmd.SetProperty "btnOpenFormConditionsAndPests", acPropertyVisible, "0"
    DoCmd.SetProperty "chkConditionsChecked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Sapling_Conditions", acPropertyVisible, "0"
    DoCmd.SetProperty "lblFoliageConditions", acPropertyVisible, "0"
    DoCmd.SetProperty "chkFoliageConditionsChecked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Sapling_Foliage_Conditions", acPropertyVisible, "0"
    DoCmd.RunCommand acCmdRefresh
   
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnShowVines_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnShowCondition_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cmdShow_Condition > btnShowCondition
'   BLC - 4/9/2018 - updated checkbox naming (removed _)
' ---------------------------------
Private Sub btnShowCondition_Click()
On Error GoTo Err_Handler

    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, False
    DoCmd.SetProperty "chkVinesChecked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Sapling_Vines", acPropertyVisible, False
    DoCmd.SetProperty "btnOpenFormConditionsAndPests", acPropertyVisible, True
    DoCmd.SetProperty "chkConditionsChecked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Sapling_Conditions", acPropertyVisible, True
    DoCmd.SetProperty "lblFoliageConditions", acPropertyVisible, "0"
    DoCmd.SetProperty "chkFoliageConditionsChecked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Sapling_Foliage_Conditions", acPropertyVisible, "0"
    DoCmd.RunCommand acCmdRefresh
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnShowCondition_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnShowFoliage_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cmdShow_Foliage > btnShowFoliage
'   BLC - 4/9/2018 - updated checkbox naming (removed _)
' ---------------------------------
Private Sub btnShowFoliage_Click()
On Error GoTo Err_Handler
    
    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, False
    DoCmd.SetProperty "chkVinesChecked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Sapling_Vines", acPropertyVisible, False
    DoCmd.SetProperty "btnOpenFormConditionsAndPests", acPropertyVisible, False
    DoCmd.SetProperty "chkConditionsChecked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Sapling_Conditions", acPropertyVisible, False
    DoCmd.SetProperty "lblFoliageConditions", acPropertyVisible, True
    DoCmd.SetProperty "chkFoliageConditionsChecked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Sapling_Foliage_Conditions", acPropertyVisible, True
    DoCmd.RunCommand acCmdRefresh
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnShowFoliage_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenFormConditionsAndPests_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cmdOpen_Form_Conditions_and_Pests > btnOpenFormConditionsAndPests
' ---------------------------------
Private Sub btnOpenFormConditionsAndPests_Click()
On Error GoTo Err_Handler
    
    Dim strDocName As String
    Dim strLinkCriteria As String

    strDocName = "frm_Popup_Conditions_and_Pests"
    DoCmd.OpenForm strDocName, , , strLinkCriteria
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenFormConditionsAndPests_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenFormCrownClass_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    rename cmdOpen_Form_Crown_Class > btnOpenFormCrownClass
' ---------------------------------
Private Sub btnOpenFormCrownClass_Click(Cancel As Integer)
On Error GoTo Err_Handler
    Dim strDocName As String
    Dim strLinkCriteria As String
    
    strDocName = "frm_Popup_Crown_Classes"
    DoCmd.OpenForm strDocName, , , strLinkCriteria
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenFormCrownClass_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDeleteSample_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    rename cmdDelete_Sample > btnDeleteSample
'   BLC - 4/9/2018 - updated TreeStatus > SaplingStatus
' ---------------------------------
Private Sub btnDeleteSample_Click()
On Error GoTo Err_Handler
    
    If MsgBox("You are about to DELETE all data for this sapling for this sampling event only." _
            & vbNewLine & vbNewLine & "Is this OK?", vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then _
            GoTo Exit_Handler
            
    With CodeContextObject
        On Error Resume Next
        'DoCmd.GoToControl Screen.PreviousControl.Name
        DoCmd.GoToControl cbxSaplingStatus
        Err.Clear
        If (Not .Form.NewRecord) Then
            DoCmd.RunCommand acCmdDeleteRecord
        End If
        If (.Form.NewRecord And Not .Form.Dirty) Then
            Beep
        End If
        If (.Form.NewRecord And .Form.Dirty) Then
            DoCmd.RunCommand acCmdUndo
        End If
    End With

    Me.Parent.Refresh
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDeleteSample_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTagNewSpecimen_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    rename cmdTag_New_Specimen > btnTagNewSpecimen
' ---------------------------------
Private Sub btnTagNewSpecimen_Click()
On Error GoTo Err_Handler
    
    Dim strCriteria As String
    
    strCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Parent.Name, "txtLocation_ID")
    DoCmd.OpenForm "frm_Locations", , , strCriteria, , , "Filter by location"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTagNewSpecimen_Click[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Enter Events
' ----------------

' ---------------------------------
' SUB:          cbxSaplingStatus_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboTree_Status > cbxTreeStatus
'   BLC - 4/9/2018 - renamed SaplingStatus vs. TreeStatus
' ---------------------------------
Private Sub cbxSaplingStatus_Enter()
On Error GoTo Err_Handler

    ValidateSaplingSubform
    
    SaplingStatus = Me!cbxSaplingStatus
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 94
            Resume Next
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSaplingStatus_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxHabit_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboHabit > cbxHabit
' ---------------------------------
Private Sub cbxHabit_Enter()
On Error GoTo Err_Handler
    
    ValidateSaplingSubform
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxHabit_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxBrowsable_Enter
' Description:  textbox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed Browsable > tbxBrowsable
' ---------------------------------
Private Sub tbxBrowsable_Enter()
On Error GoTo Err_Handler

    ValidateSaplingSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxBrowsable_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxBrowsed_Enter
' Description:  textbox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed Browsed > tbxBrowsed
' ---------------------------------
Private Sub tbxBrowsed_Enter()
On Error GoTo Err_Handler
    
    ValidateSaplingSubform
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxBrowsed_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxComments_Enter
' Description:  textbox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed txtComments > tbxComments
' ---------------------------------
Private Sub tbxComments_Enter()
On Error GoTo Err_Handler

    ValidateSaplingSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComments_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSelectSampledTag_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboSelect_SampledTag > cbxSelectSampledTag
' ---------------------------------
Private Sub cbxSelectSampledTag_Enter()
On Error GoTo Err_Handler

    Me!cbxSelectSampledTag.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectSampledTag_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTagFinder_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboTag_Finder > cbxTagFinder
' ---------------------------------
Private Sub cbxTagFinder_Enter()
On Error GoTo Err_Handler
    
    Me!cboTag_Finder.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTagFinder_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSelectUnsampledTag_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboSelect_UnsampledTag > cbxSelectUnsampledTag
' ---------------------------------
Private Sub cbxSelectUnsampledTag_Enter()
On Error GoTo Err_Handler

    Me!cbxSelectUnsampledTag.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectUnsampledTag_Enter[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Change Events
' ----------------

' ---------------------------------
' SUB:          cbxSaplingStatus_Change
' Description:  combobox change actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 9, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/9/2018 - initial version
' ---------------------------------
Private Sub cbxSaplingStatus_Change()
On Error GoTo Err_Handler

    CheckTagStatus "Sapling"

'    'sapling status = Dead* ?
'    ' --> trigger tag status = RIO (Retired (In Office))
'    If Left(cbxSaplingStatus, 4) = "Dead" Then
'
''Debug.Print "tag status = " & Me.fsub_Tag_Sapling.Controls("cbxTagStatus")
'
'        Select Case fsub_Tag_Sapling.Controls("cbxTagStatus")
'         Case Is <> "Retired (In Office)"
'            Me.fsub_Tag_Sapling.Controls("cbxTagStatus").BackColor = lngYellow
'
'         Case Is = Null
'Debug.Print "tag status = NULL " & Me.fsub_Tag_Sapling.Controls("cbxTagStatus")
'                'set the value
'                fsub_Tag_Sapling.Controls("cbxTagStatus") = "Retired (In Office)"
'         Case Else
'            'do nothing
'        End Select
'
'    Else
'
'        Me.fsub_Tag_Sapling.Controls("cbxTagStatus").BackColor = lngWhite
'
'    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSaplingStatus_Change[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  After Update Events
' ----------------

' ---------------------------------
' SUB:          cbxBrowsePick_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboBrowsePick > cbxBrowsePick
' ---------------------------------
Private Sub cbxBrowsePick_AfterUpdate()
On Error GoTo Err_Handler

    Select Case Me!cboBrowsePick.Column(0)
        Case "Yes / Yes"
            Me!txtBrowsable.Value = "Yes"
            Me!txtBrowsed.Value = "Yes"
        Case "Yes / No"
            Me!txtBrowsable.Value = "Yes"
            Me!txtBrowsed.Value = "No"
        Case "No / No"
            Me!txtBrowsable.Value = "No"
            Me!txtBrowsed.Value = "No"
    End Select

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxBrowsePick_AfterUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxQuickComment_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboQuick_Comment > cbxQuickComment
' ---------------------------------
Private Sub cbxQuickComment_AfterUpdate()
On Error GoTo Err_Handler

    Me.tbxComments = LTrim(Me.tbxComments & " " & Me.cbxQuickComment)
    Me.tbxComments.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxQuickComment_AfterUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSaplingStatus_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboTree_Status > cbxTreeStatus
'   BLC - 4/9/2018 - renamed SaplingStatus vs. TreeStatus
' ---------------------------------
Private Sub cbxSaplingStatus_AfterUpdate()
On Error GoTo Err_Handler
    
    Dim Response As String
    
    If Left(SaplingStatus, 4) = "Dead" And Left(Me!cbxSaplingStatus, 5) = "Alive" Then
        Response = MsgBox("You have changed the status of this sapling from dead to alive", vbOKCancel, "NCRN Forest Vegetation Monitoring")
            If Response = vbOK Then
                MsgBox "Changes approved"
            Else
                MsgBox "Changes rejected"
                Me!cbxSaplingStatus = SaplingStatus
                
            End If
               
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSaplingStatus_AfterUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxHabit_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboHabit > cbxHabit
' ---------------------------------
Private Sub cbxHabit_AfterUpdate()
On Error GoTo Err_Handler
    
    ToggleDBH
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxHabit_AfterUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSelectSampledTag_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboSelect_SampledTag > cbxSelectSampledTag
'                    revised error 3200 issue to use rstClone vs. rst
' ---------------------------------
Private Sub cbxSelectSampledTag_AfterUpdate()
On Error GoTo Err_Handler
    
    ' Find the record that matches the control, if record doesn't exist, create it.
    
    Dim rstClone As DAO.Recordset
    Dim strFind As String
    Dim strSearchField As String
    
    strFind = Me!cbxSelectSampledTag.Column(0)
    strSearchField = "Tag_ID"
    
    'Search for a matching record
    Set rstClone = Me.Recordset.Clone
    
    Do Until rstClone.EOF
        If rstClone(strSearchField) = strFind Then
            'Goto matching record and exit subroutine
            Me.Bookmark = rstClone.Bookmark
            GoTo Exit_Handler
        End If
        rstClone.MoveNext
    Loop
'    'If we haven't found record and exited by now, create new record.
'    DoCmd.GoToRecord , , acNewRec
'    Tag_ID.Value = strFind
'    DoCmd.RunCommand acCmdSaveRecord
'    Me!fsub_Tag_Sapling.Requery
'    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tag_Sapling]!txtTag_Status = "Sapling"
'    Me!fsub_Tag_Sapling.Requery
'    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tag_Sapling]!cmdShow_Species.Visible = True
'    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tags_History_Summary].Requery

Exit_Handler:
    ToggleDBH
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            'rst.CancelUpdate 'I hope this is the correct fix.
            rstClone.CancelUpdate
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - cbxSelectSampledTag_AfterUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSelectUnsampledTag_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboSelect_UnsampledTag > btnSelectUnsampledTag
'                    revised error 3200 issue to use rstClone vs. rst
'   BLC - 4/17/2018 - revosed tbxTagStatus > cbxTagStatus
' ---------------------------------
Private Sub cbxSelectUnsampledTag_AfterUpdate()
On Error GoTo Err_Handler
    
    ' Find the record that matches the control, if record doesn't exist, create it.
    
    Dim rstClone As DAO.Recordset
    Dim strFind As String
    Dim strSearchField As String
    
    strFind = Me!cbxSelectUnsampledTag.Column(0)
    strSearchField = "Tag_ID"
    
    If Me!cbxSelectUnsampledTag.Column(2) = "Tree" Then
        If MsgBox("You are downgrading a TREE to a SAPLING.  Is this OK?", vbOKCancel) = vbCancel Then GoTo Exit_Handler
    End If
        
    'Search for a matching record
    Set rstClone = Me.Recordset.Clone
    
    Do Until rstClone.EOF
        If rstClone(strSearchField) = strFind Then
            'Goto matching record and exit subroutine
            Me.Bookmark = rstClone.Bookmark
            GoTo Exit_Handler
        End If
        rstClone.MoveNext
    Loop
    'If we haven't found record and exited by now, create new record.
    DoCmd.GoToRecord , , acNewRec
    Tag_ID.Value = strFind
    DoCmd.RunCommand acCmdSaveRecord
    Me!fsub_Tag_Sapling.Requery
    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tag_Sapling]!cbxTagStatus = "Sapling"
    Me!fsub_Tag_Sapling.Requery
    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tags_History_Summary].Requery
        
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            'rst.CancelUpdate 'I hope this is the correct fix.
            rstClone.CancelUpdate
        Case 3021 'record not found .... Mel says DOUBLE CHECK
            MsgBox ("Case 3021 error cbxSelectUnsampledTag code")
            DoCmd.GoToRecord , , acNewRec
'FIX            txtTag_ID.Value = Me!cbxSelectUnsampledTag.Column(0)
            DoCmd.RunCommand acCmdSaveRecord
            Me!fsub_Sapling_Data.Requery
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - cbxSelectUnsampledTag_AfterUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTagFinder_AfterUpdate
' Description:  button after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed cboTag_Finder > btnTagFinder
'                    revised error 3200 issue to use rstClone vs. rst
' ---------------------------------
Private Sub btnTagFinder_AfterUpdate()
On Error GoTo Err_Handler
    
    ' Find the record that matches the control, if record doesn't exist, create it.
       
    Dim rstClone As DAO.Recordset
    Dim strFind As String
    Dim strSearchField As String
    
    strFind = Me!cboTag_Finder.Column(0)
    strSearchField = "Tag_ID"
    
    If Me!cboTag_Finder.Column(2) = "Tree" Then
        If MsgBox("You are downgrading a TREE to a SAPLING.  Is this OK?", _
            vbOKCancel) = vbCancel Then GoTo Exit_Handler
        
    End If
        
    'Search for a matching record
    Set rstClone = Me.Recordset.Clone
    
    Do Until rstClone.EOF
        If rstClone(strSearchField) = strFind Then
            'Goto matching record and exit subroutine
            Me.Bookmark = rstClone.Bookmark
            GoTo Exit_Handler
        End If
        rstClone.MoveNext
    Loop
    'If we haven't found record and exited by now, create new record.
    DoCmd.GoToRecord , , acNewRec
    Tag_ID.Value = strFind
    DoCmd.RunCommand acCmdSaveRecord
    Me!fsub_Tag_Sapling.Requery
    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tag_Sapling]!tbxTagStatus = "Sapling"
    Me!fsub_Tag_Sapling.Requery
    Forms![frm_Events]![fsub_Sapling_Data]![fsub_Tags_History_Summary].Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            'rst.CancelUpdate 'I hope this is the correct fix.
            rstClone.CancelUpdate
        Case 3021 'record not found .... Mel says DOUBLE CHECK
            MsgBox ("Case 3021 error cbxTagFinder code")
            DoCmd.GoToRecord , , acNewRec
'FIX            txtTag_ID.Value = Me!cbxTag_Finder.Column(0)
            DoCmd.RunCommand acCmdSaveRecord
            Me!fsub_Sapling_Data.Requery
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnTagFinder_AfterUpdate[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Exit Events
' ----------------

' ---------------------------------
' SUB:          fsub_Sapling_DBH_Exit
' Description:  subreport exit actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub fsub_Sapling_DBH_Exit(Cancel As Integer)
On Error GoTo Err_Handler

    Me.Refresh
    
    'check for +/-4cm or < 1cm sapling DBH
    Select Case ValidDBH
        Case True
            tbxComments.BackColor = lngWhite
            'hide DBH double check
            lblDBHCheck.BackColor = lngWhite
            lblDBHCheck.Visible = False
            chkDBHCheck.Visible = False
        Case False
            tbxComments.BackColor = lngYellow
            'expose DBH double check
            lblDBHCheck.BackColor = lngYellow
            lblDBHCheck.Visible = True
            chkDBHCheck.Visible = True
            MsgBox "Warning...change in DBH exceeds threshold. Please check value.", vbExclamation, "NCRN Vegetation Monitoring"
    End Select
    
'    Dim db As DAO.Database
'    Set db = CurrentDb
'
'    'Check to see if the temporary query exists and if it does delete it.
'
'    If fxnQueryExists("_qCOMPARE_DBH") Then
'        db.QueryDefs.Delete ("_qCOMPARE_DBH")
'    End If
'
'    Dim strLocID As String
'    strLocID = Forms!frm_Events!txtLocation_ID
'
'    Dim intTag As Integer
'    intTag = Forms!frm_Events!fsub_Sapling_Data!fsub_Tag_Sapling!txtTag
'
'    Dim varDBH_Current As Variant
'    Dim varDBH_Past As Variant
'
'    Dim strSQL As String
'    strSQL = "SELECT tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag, " _
'            & "Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2,6) AS EquivDBH " _
'            & "FROM ((tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " _
'            & "INNER JOIN (tbl_Sapling_Data INNER JOIN tbl_Tags ON tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID) ON tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID) " _
'            & "INNER JOIN tbl_Sapling_DBH ON tbl_Sapling_Data.Sapling_Data_ID = tbl_Sapling_DBH.Sapling_Data_ID " _
'            & "GROUP BY tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag " _
'            & "HAVING (((tbl_Locations.Location_ID) = """ & strLocID & """) And ((tbl_Tags.Tag) = " & intTag & ")) " _
'            & "ORDER BY tbl_Events.Event_Date;"
'
'    Dim qDef As DAO.QueryDef
'    Set qDef = db.CreateQueryDef("_qCOMPARE_DBH", strSQL)
'
'    Dim rs As DAO.Recordset
'    Set rs = db.OpenRecordset("_qCOMPARE_DBH")
'
'    rs.MoveLast
'    If rs.RecordCount > 1 Then
'
'        varDBH_Current = rs![EquivDBH]
'        rs.MovePrevious
'        varDBH_Past = rs![EquivDBH]
'
'        If varDBH_Current - varDBH_Past >= 4 Or varDBH_Current - varDBH_Past <= -4 Then
'            MsgBox "Warning...change in DBH exceeds threshold. Please check value.", vbExclamation, "NCRN Vegetation Monitoring"
'        End If
'    End If
'
'
'    If Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!txtEquivDBH < 1 Then
'        MsgBox "Saplings must have a minimum DBH of 1.0. Please address the issue"
'        Forms!frm_Events!fsub_Sapling_Data!fsub_Sapling_DBH!txtDBH.SetFocus
'    End If
'
'    DoCmd.DeleteObject acQuery, "_qCOMPARE_DBH"
'    Set varDBH_Current = Nothing
'    Set varDBH_Past = Nothing
'    Set rs = Nothing
'    Set qDef = Nothing
'    Set db = Nothing

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fsub_Sapling_DBH_Exit[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Methods
' ---------------------------------

' ---------------------------------
' SUB:          ValidateSaplingSubform
' Description:  form validation actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
' ---------------------------------
Private Sub ValidateSaplingSubform()
On Error GoTo Err_Handler
    
    ' confirm a Tag has been selected
    If IsNull(Me!fsub_Tag_Sapling!tbxTag) Then
        MsgBox "You must SELECT A TAG before you can enter record details!", vbExclamation, "Enter Tag First"
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ValidateSaplingSubform[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
'  Functions
' ---------------------------------

' ---------------------------------
' FUNCTION:     ToggleDBH
' Description:  DBH show/hide actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  ML/GS, unknown
' Adapted:      Bonnie Campbell, April 3, 2018
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 4/3/2018 - added error handling, documentation
'                    renamed Show_Hide_DBH > ToggleDBH
' ---------------------------------
Private Function ToggleDBH()
On Error GoTo Err_Handler
    
    Select Case Me!cbxHabit.Value
        Case "Tree"
            Me!fsub_Sapling_DBH.Visible = True
        Case "Shrub"
            Me!fsub_Sapling_DBH.Visible = False
        Case Else
            Me!fsub_Sapling_DBH.Visible = True
    End Select
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleDBH[fsub_Sapling_Data])"
    End Select
    Resume Exit_Handler
End Function
