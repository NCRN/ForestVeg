Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14040
    DatasheetFontHeight =9
    ItemSuffix =79
    Left =1095
    Top =2190
    Right =15120
    Bottom =7455
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x015274d28119e540
    End
    RecordSource ="tbl_Tree_Data"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =255
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
        Begin FormHeader
            Height =0
            BackColor =16768194
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =6855
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =3255
                    Width =1260
                    Height =855
                    TabIndex =31
                    BackColor =14745599
                    Name ="tbxHighlightChk"

                    LayoutCachedLeft =60
                    LayoutCachedTop =3255
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =4110
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =7680
                    Top =1980
                    Width =210
                    Height =209
                    TabIndex =5
                    Name ="chkConditionsChecked"
                    ControlSource ="Conditions_Checked"
                    StatusBarText ="This tree was checked for disease/damage conditions"
                    AfterUpdate ="[Event Procedure]"

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
                    TabIndex =6
                    Name ="chkFoliageConditionsChecked"
                    ControlSource ="Foliage_Conditions_Checked"
                    StatusBarText ="This tree was checked for foliage conditions"
                    AfterUpdate ="[Event Procedure]"

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
                    TabIndex =20
                    BackColor =15527148
                    BorderColor =0
                    Name ="tbxVinesHighlight"
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
                    TabIndex =12
                    Name ="fsub_Tree_Foliage_Conditions"
                    SourceObject ="Form.fsub_Tree_Foliage_Conditions"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"
                    OnEnter ="[Event Procedure]"

                    LayoutCachedLeft =4410
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7949
                    LayoutCachedHeight =4620
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextFontCharSet =204
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
                    Visible = NotDefault
                    FontUnderline = NotDefault
                    OverlapFlags =223
                    TextFontCharSet =204
                    Left =4440
                    Top =1920
                    Width =2106
                    Height =306
                    FontSize =10
                    TabIndex =14
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
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    SpecialEffect =2
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1559
                    Top =5160
                    Width =12240
                    Height =361
                    ColumnWidth =2055
                    FontSize =12
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BorderColor =0
                    Name ="tbxComments"
                    ControlSource ="Tree_Notes"
                    StatusBarText ="Notes about this sampling of this tree"
                    OnEnter ="[Event Procedure]"

                    LayoutCachedLeft =1559
                    LayoutCachedTop =5160
                    LayoutCachedWidth =13799
                    LayoutCachedHeight =5521
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextAlign =3
                            Top =5160
                            Width =1169
                            Height =361
                            FontSize =13
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            BackColor =15527148
                            Name ="lblComments"
                            Caption ="Comments"
                            LayoutCachedTop =5160
                            LayoutCachedWidth =1169
                            LayoutCachedHeight =5521
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =7680
                    Top =1980
                    Width =210
                    Height =209
                    TabIndex =4
                    Name ="chkVinesChecked"
                    ControlSource ="Vines_Checked"
                    StatusBarText ="This tree was checked for vines"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =1980
                    LayoutCachedWidth =7890
                    LayoutCachedHeight =2189
                End
                Begin Subform
                    OverlapFlags =85
                    BorderWidth =2
                    Top =435
                    Width =13860
                    Height =945
                    TabIndex =7
                    BorderColor =7633277
                    Name ="fsub_Tag_Tree"
                    SourceObject ="Form.fsub_Tag_Tree"
                    LinkChildFields ="Tag_ID"
                    LinkMasterFields ="Tag_ID"

                    LayoutCachedTop =435
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1380
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =20
                    ListWidth =5760
                    Left =2880
                    Top =60
                    Width =240
                    Height =315
                    FontSize =14
                    TabIndex =8
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cbxSelectUnsampledTag"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag, tbl_Tags.Tag_Status AS Class, IIf(IsNull(["
                        "azimuth]),\"\",[Azimuth] & \" / \" & [distance] & \"m\") AS Azi_Dist, tbl_Tags.M"
                        "icroplot_Number AS MP FROM (tbl_Tags LEFT JOIN qry_Status_Tree_Current_Event ON "
                        "tbl_Tags.Tag_ID = qry_Status_Tree_Current_Event.Tag_ID) LEFT JOIN qry_Status_Sap"
                        "ling_Current_Event ON tbl_Tags.Tag_ID = qry_Status_Sapling_Current_Event.Tag_ID "
                        "WHERE (((qry_Status_Sapling_Current_Event.Event_ID) Is Null) AND ((qry_Status_Tr"
                        "ee_Current_Event.Event_ID) Is Null) AND ((tbl_Tags.Location_ID)=[Forms]![frm_Eve"
                        "nts]![Location_ID])) ORDER BY tbl_Tags.Tag_Status DESC , tbl_Tags.Tag;"
                    ColumnWidths ="0;1080;2520;1440;720"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3120
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
                            Name ="lblSelect_Tag"
                            Caption ="Select an unsampled tag ->"
                            LayoutCachedLeft =60
                            LayoutCachedTop =60
                            LayoutCachedWidth =2865
                            LayoutCachedHeight =375
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =119
                    SpecialEffect =2
                    Left =1320
                    Top =2880
                    Width =2820
                    Height =2220
                    TabIndex =9
                    Name ="fsub_Tree_DBH"
                    SourceObject ="Form.fsub_Tree_DBH"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =2880
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =5100
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2880
                            Width =1200
                            Height =360
                            FontSize =13
                            Name ="lblfsubTreeDBH"
                            Caption ="Stems (cm)"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2880
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =3240
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =255
                    Left =4410
                    Top =2220
                    Width =3539
                    Height =2400
                    TabIndex =10
                    Name ="fsub_Tree_Vines"
                    SourceObject ="Form.fsub_Tree_Vines"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =4410
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7949
                    LayoutCachedHeight =4620
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextFontCharSet =204
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
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =4410
                    Top =2220
                    Width =3539
                    Height =2400
                    TabIndex =11
                    Name ="fsub_Tree_Conditions"
                    SourceObject ="Form.fsub_Tree_Conditions"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"
                    OnEnter ="[Event Procedure]"

                    LayoutCachedLeft =4410
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7949
                    LayoutCachedHeight =4620
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1320
                    Top =1980
                    Width =2820
                    Height =359
                    FontSize =13
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"192\""
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00430072006f0077006e005f0043006c00 ,
                        0x6100730073005d0029003d00540072007500650000000000
                    End
                    Name ="cbxCrownClass"
                    ControlSource ="Crown_Class"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description, tlu_Enumer"
                        "ations.Enum_Group FROM tlu_Enumerations WHERE (((tlu_Enumerations.Enum_Group)=\""
                        "Crown Class\")) ORDER BY tlu_Enumerations.Sort_Order; "
                    ColumnWidths ="0;1440"
                    StatusBarText ="Options: (1)open-grown (2)Dominant (3)Co-dominant (4)Intermediate (5)Overtopped"
                    OnEnter ="[Event Procedure]"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1320
                    LayoutCachedTop =1980
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =2339
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b00430072006f0077006e005f0043006c006100 ,
                        0x730073005d0029003d0054007200750065000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin ComboBox
                    SpecialEffect =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1319
                    Top =1560
                    Width =2820
                    Height =359
                    ColumnWidth =1875
                    FontSize =13
                    TabIndex =2
                    BorderColor =0
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0063006200780054007200650065005300 ,
                        0x740061007400750073005d0029003d00540072007500650000000000
                    End
                    Name ="cbxTreeStatus"
                    ControlSource ="Tree_Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Tree Status\")) ORDER BY tlu_Enumer"
                        "ations.Sort_Order; "
                    ColumnWidths ="3168"
                    StatusBarText ="Health status of this specimen"
                    OnEnter ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1319
                    LayoutCachedTop =1560
                    LayoutCachedWidth =4139
                    LayoutCachedHeight =1919
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
                            Top =1560
                            Width =1200
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
                            LayoutCachedTop =1560
                            LayoutCachedWidth =1260
                            LayoutCachedHeight =1920
                        End
                    End
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    Left =60
                    Top =1980
                    Width =1200
                    Height =420
                    FontSize =13
                    TabIndex =13
                    ForeColor =6108695
                    Name ="btnOpenFormCrownClass"
                    Caption ="Crown"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =1980
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =2400
                    Alignment =3
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Subform
                    OverlapFlags =85
                    Left =1320
                    Top =5580
                    Width =12540
                    Height =1275
                    TabIndex =15
                    Name ="fsub_Tags_History_Summary"
                    SourceObject ="Form.fsub_Tags_History_Summary"
                    LinkChildFields ="Tag_ID"
                    LinkMasterFields ="Tag_ID"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =5580
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =6855
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =60
                            Top =5580
                            Width =1195
                            Height =600
                            FontSize =13
                            Name ="fsub_Tags_History_Summary Label"
                            Caption ="Tag History"
                            EventProcPrefix ="fsub_Tags_History_Summary_Label"
                            LayoutCachedLeft =60
                            LayoutCachedTop =5580
                            LayoutCachedWidth =1255
                            LayoutCachedHeight =6180
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7320
                    Top =60
                    Width =2400
                    Height =300
                    FontSize =12
                    TabIndex =16
                    ForeColor =0
                    Name ="btnTagNewSpecimen"
                    Caption ="Tag New Specimen"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Tag a new tree (Do not use this to replace a lost tag)."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7320
                    LayoutCachedTop =60
                    LayoutCachedWidth =9720
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
                    Overlaps =1
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
                    TabIndex =17
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
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =5
                    ListRows =20
                    ListWidth =7920
                    Left =6570
                    Top =60
                    Width =240
                    Height =315
                    FontSize =14
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cbxSelectSampledTag"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag, IIf(IsNull([azimuth]),\"\",[Azimuth] & \" "
                        "/ \" & [distance] & \"m\") AS Azi_Dist, qry_Status_Tree_Current_Event.Tree_Statu"
                        "s, tlu_Plants.Latin_Name FROM (tbl_Tags INNER JOIN qry_Status_Tree_Current_Event"
                        " ON tbl_Tags.Tag_ID = qry_Status_Tree_Current_Event.Tag_ID) LEFT JOIN tlu_Plants"
                        " ON tbl_Tags.TSN = tlu_Plants.TSN WHERE (((tbl_Tags.Location_ID)=[Forms]![frm_Ev"
                        "ents]![Location_ID])) ORDER BY tbl_Tags.Tag;"
                    ColumnWidths ="0;1080;1800;2160;2880"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnChange ="[Event Procedure]"
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
                            Name ="lblSelect_Sample"
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
                Begin Subform
                    OverlapFlags =87
                    SpecialEffect =4
                    BorderWidth =3
                    Left =8160
                    Top =1860
                    Width =5700
                    Height =2760
                    TabIndex =18
                    Name ="fsub_Conditions_Summary"
                    SourceObject ="Form.fsub_Tree_All_Conditions"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =8160
                    LayoutCachedTop =1860
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =4620
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =8160
                            Top =1500
                            Width =5100
                            Height =360
                            FontSize =13
                            FontWeight =700
                            Name ="lblTree_All_Conditions"
                            Caption ="Summary of all vines and conditions"
                            LayoutCachedLeft =8160
                            LayoutCachedTop =1500
                            LayoutCachedWidth =13260
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =4440
                    Top =1500
                    Width =1080
                    FontSize =12
                    TabIndex =19
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
                    TabIndex =21
                    BackColor =15527148
                    BorderColor =0
                    Name ="tbxConditionHighlight"
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
                    TabIndex =22
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
                    TabIndex =23
                    BackColor =15527148
                    BorderColor =0
                    Name ="tbxFoliageHighlight"
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
                    TabIndex =24
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =13440
                    Top =1440
                    Width =426
                    Height =396
                    TabIndex =25
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =10260
                    Top =60
                    Width =360
                    Height =300
                    FontSize =12
                    FontWeight =700
                    TabIndex =26
                    ForeColor =0
                    Name ="btnOpenFormTagTransitions"
                    Caption ="?"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Tag a new tree (Do not use this to replace a lost tag)."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10260
                    LayoutCachedTop =60
                    LayoutCachedWidth =10620
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
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ListWidth =5760
                    Left =1320
                    Top =5160
                    Width =240
                    Height =360
                    FontSize =12
                    TabIndex =27
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxQuickComment"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Tree Comments\")) ORDER BY tlu_Enumerations.Sort_Order;"
                    ColumnWidths ="5760"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =5160
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =5520
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1320
                    Top =2400
                    Width =2820
                    Height =359
                    FontSize =13
                    TabIndex =28
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"3\";\"2\""
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062007800430072006f0077006e00 ,
                        0x43006c006100730073005d0029003d00540072007500650000000000
                    End
                    Name ="cbxTreeVigor"
                    ControlSource ="TreeVigor"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tluTreeVigor.TreeVigorCode, tluTreeVigor.TreeVigorClass FROM tluTreeVigor"
                        ";"
                    ColumnWidths ="360;1440"
                    StatusBarText ="Options: (1)open-grown (2)Dominant (3)Co-dominant (4)Intermediate (5)Overtopped"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1320
                    LayoutCachedTop =2400
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =2759
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001c0000004900 ,
                        0x73004e0075006c006c0028005b00630062007800430072006f0077006e004300 ,
                        0x6c006100730073005d0029003d00540072007500650000000000000000000000 ,
                        0x0000000000000000000000
                    End
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =87
                    Left =60
                    Top =2400
                    Width =1200
                    FontSize =13
                    TabIndex =29
                    ForeColor =6108695
                    Name ="btnTreeVigorDesc"
                    Caption ="Vigor"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x00000000
                    End
                    BackStyle =0

                    LayoutCachedLeft =60
                    LayoutCachedTop =2400
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =2760
                    Alignment =3
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CheckBox
                    SpecialEffect =0
                    OverlapFlags =247
                    OldBorderStyle =0
                    Left =1020
                    Top =3360
                    Width =210
                    Height =209
                    TabIndex =30
                    BorderColor =255
                    Name ="chkDBHCheck"
                    StatusBarText ="Check if DBH was double checked"
                    DefaultValue ="0"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Check if DBH was double checked"
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1

                    LayoutCachedLeft =1020
                    LayoutCachedTop =3360
                    LayoutCachedWidth =1230
                    LayoutCachedHeight =3569
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =247
                            TextAlign =3
                            Left =120
                            Top =3300
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
                            LayoutCachedLeft =120
                            LayoutCachedTop =3300
                            LayoutCachedWidth =975
                            LayoutCachedHeight =4080
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
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
' MODULE:       fsub_Tree_Data
' Level:        Application module
' Version:      1.06
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/5/2018 - 1.01 - added documentation, error handling
'               BLC   - 4/9/2018 - 1.02 - added tag vs. sapling status check
'               BLC   - 4/19/2018 - 1.03 - added Form_Open, chkDBHCheck_Click events
'                                          update ValidDBH w/ Habit
'               BLC   - 4/21/2018 - 1.04 - set record's DBH_Check value, code cleanup
'               BLC - 4/22/2018   - 1.05 - added change events for tags (sampled/unsampled),
'                                          CheckDBH
'               BLC - 4/30/2018   - 1.06 - add DBH validation on exit (shift from DBH subform events)
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
' Source/date:  Bonnie Campbell, April 19, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/19/2018 - initial version
'   BLC - 4/21/2018 - set DBH check from db, check DBH
'   BLC - 4/22/2018 - revised to use CheckDBH
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'hide double check unless necessary
    'lblDBHCheck.Visible = False
    'chkDBHCheck.Visible = False
    tbxHighlightChk.Visible = False
    
'    'set default comment bgd color
'    tbxComments.BackColor = lngWhite
'
'    'fetch DBH_Check value from db (convert 1 -> -1 for Access logic)
'    chkDBHCheck = IIf(Me!DBH_Check = 1, -1, 0)
'
'    'check for +/-4cm or < 1cm sapling DBH
'    ValidDBH "Tree"
'
'    'set text color if checked
'    If Me!DBH_Check = 1 Then Me.lblDBHCheck.ForeColor = lngBlue
    
    CheckDBH
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Current
' Description:  form current actions
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
Private Sub Form_Current()
On Error GoTo Err_Handler
       
    'compare status
    CheckTagStatus "Tree"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Tree_Data", "Tree_Data_ID") = dbText Then
            Me!Tree_Data_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Subforms
' ----------------
' ---------------------------------
' SUB:          fsub_Tree_Conditions_Enter
' Description:  subform enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub fsub_Tree_Conditions_Enter()
On Error GoTo Err_Handler

    ValidateTreeSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fsub_Tree_Conditions_Enter[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          fsub_Tree_Foliage_Conditions_Enter
' Description:  subform enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub fsub_Tree_Foliage_Conditions_Enter()
On Error GoTo Err_Handler

    ValidateTreeSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fsub_Tree_Foliage_Conditions_Enter[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          fsub_Tree_DBH_Enter
' Description:  subform enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub fsub_Tree_DBH_Enter()
On Error GoTo Err_Handler

    ValidateTreeSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fsub_Tree_DBH_Enter[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          fsub_Tree_DBH_Exit
' Description:  subform exit actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
'   BLC - 4/19/2018 - update ValidDBH w/ Habit
'   BLC - 4/21/2018 - code cleanup
' ---------------------------------
Private Sub fsub_Tree_DBH_Exit(Cancel As Integer)
On Error GoTo Err_Handler
   
    Me.Refresh
    
    'check for +/-4cm or < 1cm sapling DBH
    ValidDBH "Tree"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fsub_Tree_DBH_Exit[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Click
' ----------------

' ---------------------------------
' SUB:          cbxCrownClass_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub cbxCrownClass_Enter()
On Error GoTo Err_Handler

    ValidateTreeSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxCrownClass_Enter[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkDBHCheck_Click
' Description:  checkbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 19, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/19/2018 - initial version
'   BLC - 4/21/2018 - set DBH value
' ---------------------------------
Private Sub chkDBHCheck_Click()
On Error GoTo Err_Handler
    
    'Toggle check label color based on if checked or not
    lblDBHCheck.ForeColor = IIf(chkDBHCheck, lngBlue, lngRed)
    
    'update the record's value (since DBH_Check is 0/1 vs. 0/-1)
    SetDBHCheck Me.Tree_Data_ID, "Tree", chkDBHCheck
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkDBHCheck_Click[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
            "Error encountered (#" & Err.Number & " - btnOpenFormConditionsAndPests_Click[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
            "Error encountered (#" & Err.Number & " - btnTagNewSpecimen_Click[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenFormTagTransitions_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
            "Error encountered (#" & Err.Number & " - btnOpenFormTagTransitions_Click[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnOpenFormCrownClass_Click()
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
            "Error encountered (#" & Err.Number & " - btnOpenFormCrownClass_Click[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTreeVigorDesc_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnTreeVigorDesc_Click()
On Error GoTo Err_Handler

    Dim strDocName As String
    Dim strLinkCriteria As String

    strDocName = "frm_Popup_Vigor_Classes"
    DoCmd.OpenForm strDocName, , , strLinkCriteria
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTreeVigorDesc_Click[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnShowVines_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnShowVines_Click()
On Error GoTo Err_Handler

    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, True
    DoCmd.SetProperty "chkVinesChecked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Tree_Vines", acPropertyVisible, True
    DoCmd.SetProperty "btnOpenFormConditionsAndPests", acPropertyVisible, "0"
    DoCmd.SetProperty "chkConditionsChecked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Tree_Conditions", acPropertyVisible, "0"
    DoCmd.SetProperty "lblFoliageConditions", acPropertyVisible, "0"
    DoCmd.SetProperty "chkFoliageConditionsChecked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Tree_Foliage_Conditions", acPropertyVisible, "0"
    DoCmd.RunCommand acCmdRefresh

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnShowVines_Click[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnShowCondition_Click()
On Error GoTo Err_Handler

    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, False
    DoCmd.SetProperty "chkVinesChecked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Tree_Vines", acPropertyVisible, False
    DoCmd.SetProperty "btnOpenFormConditionsAndPests", acPropertyVisible, True
    DoCmd.SetProperty "chkConditionsChecked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Tree_Conditions", acPropertyVisible, True
    DoCmd.SetProperty "lblFoliageConditions", acPropertyVisible, "0"
    DoCmd.SetProperty "chkFoliageConditionsChecked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Tree_Foliage_Conditions", acPropertyVisible, "0"
    DoCmd.RunCommand acCmdRefresh

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnShowCondition_Click[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnShowFoliage_Click()
On Error GoTo Err_Handler

    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, False
    DoCmd.SetProperty "chkVinesChecked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Tree_Vines", acPropertyVisible, False
    DoCmd.SetProperty "btnOpenFormConditionsAndPests", acPropertyVisible, False
    DoCmd.SetProperty "chkConditionsChecked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Tree_Conditions", acPropertyVisible, False
    DoCmd.SetProperty "lblFoliageConditions", acPropertyVisible, True
    DoCmd.SetProperty "chkFoliageConditionsChecked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Tree_Foliage_Conditions", acPropertyVisible, True
    DoCmd.RunCommand acCmdRefresh

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnShowFoliage_Click[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnDeleteSample_Click()
On Error GoTo Err_Handler

    If MsgBox("You are about to DELETE all data for this tree for this " _
        & "sampling event only." & vbNewLine & vbNewLine & "Is this OK?", _
        vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then GoTo Exit_Handler
    
    With CodeContextObject
        On Error Resume Next
        'DoCmd.GoToControl Screen.PreviousControl.Name
        DoCmd.GoToControl cbxTreeStatus
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
            "Error encountered (#" & Err.Number & " - btnDeleteSample_Click[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Enter
' ----------------

' ---------------------------------
' SUB:          cbxSelectUnsampledTag_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
            "Error encountered (#" & Err.Number & " - cbxSelectUnsampledTag_Enter[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
            "Error encountered (#" & Err.Number & " - cbxSelectSampledTag_Enter[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTreeStatus_Enter
' Description:  combobox enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub cbxTreeStatus_Enter()
On Error GoTo Err_Handler

    ValidateTreeSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTreeStatus_Enter[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub tbxComments_Enter()
On Error GoTo Err_Handler

    ValidateTreeSubform

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxComments_Enter[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Change Events
' ----------------
' ---------------------------------
' SUB:          cbxSelectUnsampledTag_Change
' Description:  combobox change actions
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
Private Sub cbxSelectUnsampledTag_Change()
On Error GoTo Err_Handler

'    'fetch DBH_Check value from db (convert 1 -> -1 for Access logic)
'    chkDBHCheck = IIf(Me!DBH_Check = 1, -1, 0)

    CheckDBH

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - cbxSelectUnsampledTag_Change[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSelectSampledTag_Change
' Description:  combobox change actions
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
Private Sub cbxSelectSampledTag_Change()
On Error GoTo Err_Handler

'    'fetch DBH_Check value from db (convert 1 -> -1 for Access logic)
'    chkDBHCheck = IIf(Me!DBH_Check = 1, -1, 0)

    CheckDBH
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - cbxSelectSampledTag_Change[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxTreeStatus_Change
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
'   BLC - 4/21/2018 - code cleanup
' ---------------------------------
Private Sub cbxTreeStatus_Change()
On Error GoTo Err_Handler

    CheckTagStatus "Tree"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxTreeStatus_Change[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  After Update
' ----------------
' ---------------------------------
' SUB:          cbxSelectUnsampledTag_AfterUpdate
' Description:  combobox after udpate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub cbxSelectUnsampledTag_AfterUpdate()
On Error GoTo Err_Handler

' Find the record that matches the control, if record doesn't exist, create it.
    
    Dim rstClone As DAO.Recordset
    Dim strFind As String
    Dim strSearchField As String
    
    strFind = Me!cbxSelectUnsampledTag.Column(0)
    strSearchField = "Tag_ID"
    
    If Me!cbxSelectUnsampledTag.Column(2) = "Sapling" Then
        If MsgBox("You are upgrading a SAPLING to a TREE.  Is this OK?", vbOKCancel) = vbCancel Then GoTo Exit_Handler
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
    Me!fsub_Tag_Tree.Requery
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tag_Tree]!cbxTagStatus = "Tree"
    Me!fsub_Tag_Tree.Requery
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tags_History_Summary].Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            'rst.CancelUpdate 'I hope this is the correct fix.
            rstClone.CancelUpdate
        Case 3021 'record not found .... Mel says DOUBLE CHECK
            MsgBox ("Case 3021 error cboTagFinder code")
            DoCmd.GoToRecord , , acNewRec
            'FIX? txtTag_ID control not found
            'txtTag_ID.Value = Me!cbxSelectUnsampledTag.Column(0)
            DoCmd.RunCommand acCmdSaveRecord
            Me!fsub_Tree_Data.Requery
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - cbxSelectUnsampledTag_AfterUpdate[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
    'If we haven't found record and exited by now, create new record.
'    DoCmd.GoToRecord , , acNewRec
'    Tag_ID.Value = strFind
'    DoCmd.RunCommand acCmdSaveRecord
'    Me!fsub_Tag_Tree.Requery
'    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tag_Tree]!txtTag_Status = "Tree"
'    Me!fsub_Tag_Tree.Requery
'    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tag_Tree]!cmdShow_Species.Visible = True
'    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tags_History_Summary].Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            'rst.CancelUpdate 'I hope this is the correct fix.
            rstClone.CancelUpdate
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - cbxSelectSampledTag_AfterUpdate[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkVinesChecked_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub chkVinesChecked_AfterUpdate()
On Error GoTo Err_Handler

    tbxVinesHighlight.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkVinesChecked_AfterUpdate[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkConditionsChecked_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub chkConditionsChecked_AfterUpdate()
On Error GoTo Err_Handler

    tbxConditionHighlight.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkConditionsChecked_AfterUpdate[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkFoliageConditionsChecked_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub chkFoliageConditionsChecked_AfterUpdate()
On Error GoTo Err_Handler

    tbxFoliageHighlight.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkFoliageConditionsChecked_AfterUpdate[fsub_Tree_Data])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
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
            "Error encountered (#" & Err.Number & " - cbxQuickComment_AfterUpdate[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Methods
' ----------------
' ---------------------------------
' SUB:          ValidateTreeSubform
' Description:  tree subform validation actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 9, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/9/2018 - added documentation, error handling
' ---------------------------------
Private Sub ValidateTreeSubform()
On Error GoTo Err_Handler

    ' Description:  Confirms that a Tag has been selected
    If IsNull(Me!fsub_Tag_Tree!tbxTag) Then
        MsgBox "You must SELECT A TAG before you can enter record details!", vbExclamation, "Enter Tag First"
        'Me!cboLocation_ID.SetFocus
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ValidateTreeSubform[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          CheckDBH
' Description:  form validation actions
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
Private Sub CheckDBH()
On Error GoTo Err_Handler
    
    'set default comment bgd color
    tbxComments.backcolor = lngWhite
    
    'fetch DBH_Check value from db (convert 1 -> -1 for Access logic)
    chkDBHCheck = IIf(Me!DBH_Check = 1, -1, 0)

    'DBH records?
    If Me.Form.Controls("fsub_Tree_DBH").Form.Recordset.RecordCount > 0 Then
        
        'check for +/-4cm or < 1cm sapling DBH
        ValidDBH "Tree"

    End If

    'set text color if checked
    If Me!DBH_Check = 1 Then Me.lblDBHCheck.ForeColor = lngBlue
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CheckDBH[fsub_Tree_Data])"
    End Select
    Resume Exit_Handler
End Sub
