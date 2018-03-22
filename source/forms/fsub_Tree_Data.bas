﻿Version =20
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
    Left =1500
    Top =3150
    Right =15270
    Bottom =9960
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x4d5502714caae340
    End
    RecordSource ="tbl_Tree_Data"
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
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =7680
                    Top =1980
                    Width =210
                    Height =209
                    TabIndex =5
                    Name ="chkConditions_Checked"
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
                    Name ="chkFoliage_Conditions_Checked"
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
                    Name ="txtVines_Highlight"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x0100000094000000010000000100000000000000000000001900000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b00560069006e00650073005f0043006800650063006b006500 ,
                        0x64005d003c003e00540072007500650000000000
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =1440
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500180000005b00 ,
                        0x630068006b00560069006e00650073005f0043006800650063006b0065006400 ,
                        0x5d003c003e005400720075006500000000000000000000000000000000000000 ,
                        0x000000
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
                            Name ="lblFoliage_Conditions"
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
                    Name ="cmdOpen_Form_Conditions_and_Pests"
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
                    Name ="txtComments"
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
                    Name ="chkVines_Checked"
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
                    Name ="cboSelect_UnsampledTag"
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
                    OverlapFlags =85
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
                            Name ="fsub_Tree_DBH Label"
                            Caption ="Stems (cm)"
                            EventProcPrefix ="fsub_Tree_DBH_Label"
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
                    Name ="Crown_Class"
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
                        0x010000009e000000010000000100000000000000000000001e00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0054007200650065005f00 ,
                        0x5300740061007400750073005d0029003d00540072007500650000000000
                    End
                    Name ="cboTree_Status"
                    ControlSource ="Tree_Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Tree Status\")) ORDER BY tlu_Enumer"
                        "ations.Sort_Order; "
                    ColumnWidths ="3168"
                    StatusBarText ="Health status of this specimen"
                    OnEnter ="[Event Procedure]"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1319
                    LayoutCachedTop =1560
                    LayoutCachedWidth =4139
                    LayoutCachedHeight =1919
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001d0000004900 ,
                        0x73004e0075006c006c0028005b00630062006f0054007200650065005f005300 ,
                        0x740061007400750073005d0029003d0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
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
                    OverlapFlags =85
                    Left =60
                    Top =1980
                    Width =1200
                    FontSize =13
                    TabIndex =13
                    ForeColor =6108695
                    Name ="cmdOpen_Form_Crown_Class"
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
                    LayoutCachedHeight =2340
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
                    Name ="cmdTag_New_Specimen"
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
                    Name ="cboSelect_SampledTag"
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
                    Name ="lblOR1"
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
                    Name ="cmdShow_Vines"
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
                    Name ="txtCondition_Highlight"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x010000009e000000010000000100000000000000000000001e00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b0043006f006e0064006900740069006f006e0073005f004300 ,
                        0x6800650063006b00650064005d003c003e00540072007500650000000000
                    End

                    LayoutCachedLeft =5580
                    LayoutCachedTop =1440
                    LayoutCachedWidth =6780
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001d0000005b00 ,
                        0x630068006b0043006f006e0064006900740069006f006e0073005f0043006800 ,
                        0x650063006b00650064005d003c003e0054007200750065000000000000000000 ,
                        0x00000000000000000000000000
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
                    Name ="cmdShow_Condition"
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
                    Name ="txtFoliage_Highlight"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x01000000ae000000010000000100000000000000000000002600000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b0046006f006c0069006100670065005f0043006f006e006400 ,
                        0x6900740069006f006e0073005f0043006800650063006b00650064005d003c00 ,
                        0x3e00540072007500650000000000
                    End

                    LayoutCachedLeft =6780
                    LayoutCachedTop =1440
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =1920
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500250000005b00 ,
                        0x630068006b0046006f006c0069006100670065005f0043006f006e0064006900 ,
                        0x740069006f006e0073005f0043006800650063006b00650064005d003c003e00 ,
                        0x5400720075006500000000000000000000000000000000000000000000
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
                    Name ="cmdShow_Foliage"
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
                    Name ="cmdDelete_Sample"
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
                    Name ="Label74"
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
                    Name ="cmdOpen_Form_Tag_Transitions"
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
                    Name ="cboQuick_Comment"
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
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00430072006f0077006e005f0043006c00 ,
                        0x6100730073005d0029003d00540072007500650000000000
                    End
                    Name ="cboTreeVigor"
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
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b00430072006f0077006e005f0043006c006100 ,
                        0x730073005d0029003d0054007200750065000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin CommandButton
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =2400
                    Width =1200
                    FontSize =13
                    TabIndex =29
                    ForeColor =6108695
                    Name ="cmdTreeVigorDesc"
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

Private Sub ValidateTreeSubform()
' Description:  Confirms that a Tag has been selected
If IsNull(Me!fsub_Tag_Tree!txtTag) Then
    MsgBox "You must SELECT A TAG before you can enter record details!", vbExclamation, "Enter Tag First"
    'Me!cboLocation_ID.SetFocus
End If
End Sub

Private Sub cboQuick_Comment_AfterUpdate()
    Me.txtComments = LTrim(Me.txtComments & " " & Me.cboQuick_Comment)
    Me.txtComments.Requery
End Sub

Private Sub cboTree_Status_Enter()
    ValidateTreeSubform
End Sub

Private Sub cmdTreeVigorDesc_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Vigor_Classes"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdTreeVigorDesc_Click:
    Exit Sub
Err_cmdTreeVigorDesc_Click:
    MsgBox Err.Description
    Resume Exit_cmdTreeVigorDesc_Click
End Sub

Private Sub Crown_Class_Enter()
    ValidateTreeSubform
End Sub

Private Sub cboSelect_SampledTag_Enter()
    Me!cboSelect_SampledTag.Requery
End Sub

Private Sub fsub_Tree_Conditions_Enter()
    ValidateTreeSubform
End Sub

Private Sub fsub_Tree_DBH_Enter()
    ValidateTreeSubform
End Sub

Private Sub fsub_Tree_DBH_Exit(Cancel As Integer)
Dim db As DAO.Database
Set db = CurrentDb

'Check to see if the temporary query exists and if it does delete it.

If fxnQueryExists("_qCOMPARE_DBH") Then
    db.QueryDefs.Delete ("_qCOMPARE_DBH")
End If

Dim strLocID As String
strLocID = Forms!frm_Events!txtLocation_ID

Dim intTag As Integer
intTag = Forms!frm_Events!fsub_Tree_Data!fsub_Tag_Tree!txtTag

'dbh variables for current and previous sampling events.

Dim varDBH_Current As Variant
Dim varDBH_Past As Variant

'This code creates a temporary query that will pulls the dbh from the previous sampling event as well as the dbh that was entered for the current event.

Dim strSQL As String
strSQL = "SELECT tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag, " _
        & "Round((((Sum(3.1415*((IIf([Live]=True,[DBH],0))/2)^2))*(1/3.1415))^0.5)*2,6) AS EquivDBH " _
        & "FROM ((tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) " _
        & "INNER JOIN (tbl_Tree_Data INNER JOIN tbl_Tags ON tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID) ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) " _
        & "INNER JOIN tbl_Tree_DBH ON tbl_Tree_Data.Tree_Data_ID = tbl_Tree_DBH.Tree_Data_ID " _
        & "GROUP BY tbl_Locations.Location_ID, tbl_Events.Event_ID, tbl_Locations.Admin_Unit_Code, tbl_Locations.Subunit_Code, tbl_Events.Event_Date, tbl_Tags.Tag " _
        & "HAVING (((tbl_Locations.Location_ID) = """ & strLocID & """) And ((tbl_Tags.Tag) = " & intTag & ")) " _
        & "ORDER BY tbl_Events.Event_Date;"

Dim qDef As DAO.QueryDef
Set qDef = db.CreateQueryDef("_qCOMPARE_DBH", strSQL)

Dim rs As DAO.Recordset
Set rs = db.OpenRecordset("_qCOMPARE_DBH")

rs.MoveLast
If rs.RecordCount <= 1 Then
    Exit Sub
Else
    varDBH_Current = rs![EquivDBH]
        rs.MovePrevious
    varDBH_Past = rs![EquivDBH]
End If

If varDBH_Current - varDBH_Past >= 4 Or varDBH_Current - varDBH_Past <= -4 Then
    MsgBox "Warning!!!!! change in DBH exceeds threshold. Please check value.", vbExclamation, "NCRN Vegetation Monitoring"
End If

DoCmd.DeleteObject acQuery, "_qCOMPARE_DBH"
Set varDBH_Current = Nothing
Set varDBH_Past = Nothing
Set rs = Nothing
Set qDef = Nothing
Set db = Nothing

End Sub

Private Sub fsub_Tree_Foliage_Conditions_Enter()
    ValidateTreeSubform
End Sub

Private Sub txtComments_Enter()
    ValidateTreeSubform
End Sub

Private Sub cboSelect_SampledTag_AfterUpdate()
    ' Find the record that matches the control, if record doesn't exist, create it.
    
    On Error GoTo HandleErrors
    
    Dim rstClone As DAO.Recordset
    Dim strFind As String
    Dim strSearchField As String
    
    strFind = Me!cboSelect_SampledTag.Column(0)
    strSearchField = "Tag_ID"
    
    'Search for a matching record
    Set rstClone = Me.Recordset.Clone
    
    Do Until rstClone.EOF
        If rstClone(strSearchField) = strFind Then
            'Goto matching record and exit subroutine
            Me.Bookmark = rstClone.Bookmark
            GoTo ExitHere
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
    
ExitHere:
    Exit Sub
HandleErrors:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            rst.CancelUpdate 'I hope this is the correct fix.
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error encountered in procedure" & strProcName
            Exit Sub
    End Select
End Sub

Private Sub cboSelect_UnsampledTag_Enter()
    Me!cboSelect_UnsampledTag.Requery
End Sub

Private Sub cboSelect_UnsampledTag_AfterUpdate()
' Find the record that matches the control, if record doesn't exist, create it.
    
    On Error GoTo HandleErrors
    
    Dim rstClone As DAO.Recordset
    Dim strFind As String
    Dim strSearchField As String
    
    strFind = Me!cboSelect_UnsampledTag.Column(0)
    strSearchField = "Tag_ID"
    
    If Me!cboSelect_UnsampledTag.Column(2) = "Sapling" Then
        If MsgBox("You are upgrading a SAPLING to a TREE.  Is this OK?", vbOKCancel) = vbCancel Then GoTo ExitHere
    End If
        
    'Search for a matching record
    Set rstClone = Me.Recordset.Clone
    
    Do Until rstClone.EOF
        If rstClone(strSearchField) = strFind Then
            'Goto matching record and exit subroutine
            Me.Bookmark = rstClone.Bookmark
            GoTo ExitHere
        End If
        rstClone.MoveNext
    Loop
    'If we haven't found record and exited by now, create new record.
    DoCmd.GoToRecord , , acNewRec
    Tag_ID.Value = strFind
    DoCmd.RunCommand acCmdSaveRecord
    Me!fsub_Tag_Tree.Requery
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tag_Tree]!cboTag_Status = "Tree"
    Me!fsub_Tag_Tree.Requery
    Forms![frm_Events]![fsub_Tree_Data]![fsub_Tags_History_Summary].Requery
    
ExitHere:
    Exit Sub
HandleErrors:
    Select Case Err.Number
        Case 3200 'Record cannot be edited or saved because it has related records?
            MsgBox "Could not move to the requested record, because it would adversely affect related records.", vbOKOnly
            rst.CancelUpdate 'I hope this is the correct fix.
        Case 3021 'record not found .... Mel says DOUBLE CHECK
            MsgBox ("Case 3021 error cboTagFinder code")
            DoCmd.GoToRecord , , acNewRec
            txtTag_ID.Value = Me!cboSelect_UnsampledTag.Column(0)
            DoCmd.RunCommand acCmdSaveRecord
            Me!fsub_Tree_Data.Requery
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error encountered in procedure" & strProcName
            Exit Sub
    End Select


End Sub

Private Sub chkConditions_Checked_AfterUpdate()
    txtCondition_Highlight.Requery
End Sub

Private Sub chkFoliage_Conditions_Checked_AfterUpdate()
    txtFoliage_Highlight.Requery
End Sub

Private Sub chkVines_Checked_AfterUpdate()
    txtVines_Highlight.Requery
End Sub

Private Sub cmdOpen_Form_Conditions_and_Pests_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Conditions_and_Pests"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpen_Popup_Click:
    Exit Sub
Err_cmdOpen_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Popup_Click
End Sub

Private Sub cmdOpen_Form_Crown_Class_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Crown_Classes"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpen_Popup_Click:
    Exit Sub
Err_cmdOpen_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Popup_Click
End Sub

Private Sub cmdOpen_Form_Tag_Transitions_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Tag_Transitions"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpen_Popup_Click:
    Exit Sub
Err_cmdOpen_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Popup_Click
End Sub

Private Sub cmdTag_New_Specimen_Click()
On Error GoTo Err_cmdTag_New_Specimen_Click
    Dim strCriteria As String

    strCriteria = GetCriteriaString("[Location_ID]=", "tbl_Locations", "Location_ID", Me.Parent.Name, "txtLocation_ID")
    DoCmd.OpenForm "frm_Locations", , , strCriteria, , , "Filter by location"

Exit_cmdTag_New_Specimen_Click:
    Exit Sub
Err_cmdTag_New_Specimen_Click:
    MsgBox Err.Description
    Resume Exit_cmdTag_New_Specimen_Click
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    If Me.NewRecord Then
        If GetDataType("tbl_Tree_Data", "Tree_Data_ID") = dbText Then
            Me!Tree_Data_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub cmdShow_Vines_Click()
On Error GoTo Err_Handler

    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, True
    DoCmd.SetProperty "chkVines_Checked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Tree_Vines", acPropertyVisible, True
    DoCmd.SetProperty "cmdOpen_Form_Conditions_and_Pests", acPropertyVisible, "0"
    DoCmd.SetProperty "chkConditions_Checked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Tree_Conditions", acPropertyVisible, "0"
    DoCmd.SetProperty "lblFoliage_Conditions", acPropertyVisible, "0"
    DoCmd.SetProperty "chkFoliage_Conditions_Checked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Tree_Foliage_Conditions", acPropertyVisible, "0"
    DoCmd.RunCommand acCmdRefresh

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Error$
    Resume Exit_Procedure
End Sub


Private Sub cmdShow_Condition_Click()
On Error GoTo Err_Handler

    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, False
    DoCmd.SetProperty "chkVines_Checked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Tree_Vines", acPropertyVisible, False
    DoCmd.SetProperty "cmdOpen_Form_Conditions_and_Pests", acPropertyVisible, True
    DoCmd.SetProperty "chkConditions_Checked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Tree_Conditions", acPropertyVisible, True
    DoCmd.SetProperty "lblFoliage_Conditions", acPropertyVisible, "0"
    DoCmd.SetProperty "chkFoliage_Conditions_Checked", acPropertyVisible, "0"
    DoCmd.SetProperty "fsub_Tree_Foliage_Conditions", acPropertyVisible, "0"
    DoCmd.RunCommand acCmdRefresh

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Error$
    Resume Exit_Procedure
End Sub

Private Sub cmdShow_Foliage_Click()
On Error GoTo Err_Handler

    DoCmd.SetProperty "lblCompleted", acPropertyVisible, True
    DoCmd.SetProperty "lblVines", acPropertyVisible, False
    DoCmd.SetProperty "chkVines_Checked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Tree_Vines", acPropertyVisible, False
    DoCmd.SetProperty "cmdOpen_Form_Conditions_and_Pests", acPropertyVisible, False
    DoCmd.SetProperty "chkConditions_Checked", acPropertyVisible, False
    DoCmd.SetProperty "fsub_Tree_Conditions", acPropertyVisible, False
    DoCmd.SetProperty "lblFoliage_Conditions", acPropertyVisible, True
    DoCmd.SetProperty "chkFoliage_Conditions_Checked", acPropertyVisible, True
    DoCmd.SetProperty "fsub_Tree_Foliage_Conditions", acPropertyVisible, True
    DoCmd.RunCommand acCmdRefresh

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Error$
    Resume Exit_Procedure
End Sub

Private Sub cmdDelete_Sample_Click()
On Error GoTo Err_Handler

    If MsgBox("You are about to DELETE all data for this tree for this sampling event only." & vbNewLine & vbNewLine & "Is this OK?", vbOKCancel + vbDefaultButton2, "Warning") = vbCancel Then GoTo Exit_Procedure
    With CodeContextObject
        On Error Resume Next
        'DoCmd.GoToControl Screen.PreviousControl.Name
        DoCmd.GoToControl cboTree_Status
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

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Error$
    Resume Exit_Procedure
End Sub
