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
    FilterOn = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =10
    ItemSuffix =154
    Left =1020
    Right =15420
    Bottom =9195
    DatasheetGridlinesColor =12632256
    Filter ="[Event_ID]='{5DD03496-502A-462F-AF9C-34C036D06379}'"
    RecSrcDt = Begin
        0x58c05212730ae440
    End
    RecordSource ="qfrm_Events"
    Caption ="NCRN Sampling Event - (Browsing) - (Browsing) - (Browsing) - (Browsing) - (Brows"
        "ing) - (Browsing) - (Browsing) - (Browsing) - (Browsing) - (Browsing) - (Browsin"
        "g)"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =9210
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =14400
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="lblEvent_Form_Header"
                    Caption ="Vegetation Sampling Events"
                    FontName ="Calibri"
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =13320
                    Top =120
                    Width =900
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13320
                    LayoutCachedTop =120
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =450
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Top =480
                    Width =2160
                    Height =420
                    FontSize =18
                    FontWeight =700
                    TabIndex =2
                    Name ="txtStart_Date"
                    ControlSource ="Event_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    FontName ="Calibri"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =480
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =600
                    Width =4560
                    Height =300
                    FontSize =12
                    TabIndex =3
                    Name ="txtXY"
                    FontName ="Calibri"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =600
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =900
                End
                Begin Tab
                    MultiRow = NotDefault
                    OverlapFlags =85
                    Top =1200
                    Width =14250
                    Height =8010
                    FontSize =12
                    TabIndex =4
                    Name ="tabctlData"
                    FontName ="Calibri"

                    LayoutCachedTop =1200
                    LayoutCachedWidth =14250
                    LayoutCachedHeight =9210
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =135
                            Top =1695
                            Width =13980
                            Height =7385
                            Name ="pagIntro"
                            Caption ="Intro"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1695
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9080
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =240
                                    Top =2055
                                    Width =5520
                                    Height =2100
                                    Name ="subObservers"
                                    SourceObject ="Form.fsub_Observers"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =2055
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =4155
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    BorderWidth =3
                                    Left =6060
                                    Top =2115
                                    Width =7980
                                    Height =6885
                                    TabIndex =1
                                    Name ="fsub_Note_History"
                                    SourceObject ="Form.fsub_Note_History"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                    LayoutCachedLeft =6060
                                    LayoutCachedTop =2115
                                    LayoutCachedWidth =14040
                                    LayoutCachedHeight =9000
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =6060
                                            Top =1815
                                            Width =2760
                                            Height =300
                                            FontSize =12
                                            FontWeight =700
                                            Name ="fsub_Note_History Label"
                                            Caption ="Event History"
                                            FontName ="Calibri"
                                            EventProcPrefix ="fsub_Note_History_Label"
                                            LayoutCachedLeft =6060
                                            LayoutCachedTop =1815
                                            LayoutCachedWidth =8820
                                            LayoutCachedHeight =2115
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =240
                                    Top =4680
                                    Width =5460
                                    TabIndex =2
                                    Name ="subPlot_Floor_Conditions"
                                    SourceObject ="Form.fsub_Plot_Floor_Condition_Data"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =4680
                                    LayoutCachedWidth =5700
                                    LayoutCachedHeight =6120
                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =240
                                            Top =4380
                                            Width =3480
                                            Height =300
                                            FontSize =14
                                            FontWeight =700
                                            Name ="lblPlot Floor Conditions"
                                            Caption ="Plot Floor Conditions"
                                            FontName ="Calibri"
                                            EventProcPrefix ="lblPlot_Floor_Conditions"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =4380
                                            LayoutCachedWidth =3720
                                            LayoutCachedHeight =4680
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =12060
                                    Top =1695
                                    Width =1980
                                    Height =300
                                    FontSize =10
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="cmdAdd_Event_Note"
                                    Caption ="Add/Edit Event Note"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    ControlTipText ="Add a new contact record"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =12060
                                    LayoutCachedTop =1695
                                    LayoutCachedWidth =14040
                                    LayoutCachedHeight =1995
                                    ForeThemeColorIndex =0
                                    UseTheme =255
                                    Shape =1
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
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    Left =3360
                                    Top =6360
                                    Height =210
                                    TabIndex =4
                                    BorderColor =2366701
                                    Name ="chkPictures_Taken"
                                    ControlSource ="Pictures_Taken"
                                    AfterUpdate ="[Event Procedure]"

                                    LayoutCachedLeft =3360
                                    LayoutCachedTop =6360
                                    LayoutCachedWidth =3620
                                    LayoutCachedHeight =6570
                                End
                                Begin Rectangle
                                    SpecialEffect =4
                                    BorderWidth =3
                                    OverlapFlags =223
                                    Left =240
                                    Top =7710
                                    Width =5520
                                    Height =1080
                                    Name ="boxMetadata"
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =7710
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =8790
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =240
                                    Top =7470
                                    Width =1260
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    Name ="lblMetadata_Box"
                                    Caption ="Metadata"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =7470
                                    LayoutCachedWidth =1500
                                    LayoutCachedHeight =7710
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    Left =1440
                                    Top =7830
                                    Width =1200
                                    FontSize =10
                                    TabIndex =5
                                    Name ="txtMeta_Updated_Date"
                                    ControlSource ="Updated_Date"
                                    Format ="Short Date"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =7830
                                    LayoutCachedWidth =2640
                                    LayoutCachedHeight =8070
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =204
                                            TextAlign =3
                                            Left =240
                                            Top =7830
                                            Width =1080
                                            Height =240
                                            FontSize =10
                                            Name ="lblMeta_Updated"
                                            Caption ="Updated"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =7830
                                            LayoutCachedWidth =1320
                                            LayoutCachedHeight =8070
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =2880
                                    Left =2760
                                    Top =7830
                                    Width =2823
                                    Height =252
                                    FontSize =10
                                    TabIndex =6
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                                    Name ="cboMeta_Updated_Contact_ID"
                                    ControlSource ="Updated_By"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                                        "Last_Name, tlu_Contacts.First_Name; "
                                    ColumnWidths ="0;2880"
                                    StatusBarText ="Observer identifier"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =2760
                                    LayoutCachedTop =7830
                                    LayoutCachedWidth =5583
                                    LayoutCachedHeight =8082
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    Left =1440
                                    Top =8130
                                    Width =1200
                                    FontSize =10
                                    TabIndex =7
                                    Name ="txtMeta_Verified_Date"
                                    ControlSource ="Verified_Date"
                                    Format ="Short Date"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =8130
                                    LayoutCachedWidth =2640
                                    LayoutCachedHeight =8370
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =204
                                            TextAlign =3
                                            Left =240
                                            Top =8130
                                            Width =1080
                                            Height =240
                                            FontSize =10
                                            Name ="lblMeta_Verified"
                                            Caption ="Verified"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =8130
                                            LayoutCachedWidth =1320
                                            LayoutCachedHeight =8370
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =2880
                                    Left =2760
                                    Top =8130
                                    Width =2823
                                    Height =252
                                    FontSize =10
                                    TabIndex =8
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                                    Name ="cboMeta_Verified_Contact_ID"
                                    ControlSource ="Verified_By"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                                        "Last_Name, tlu_Contacts.First_Name; "
                                    ColumnWidths ="0;2880"
                                    StatusBarText ="Observer identifier"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =2760
                                    LayoutCachedTop =8130
                                    LayoutCachedWidth =5583
                                    LayoutCachedHeight =8382
                                End
                                Begin TextBox
                                    OverlapFlags =215
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    Left =1440
                                    Top =8430
                                    Width =1200
                                    FontSize =10
                                    TabIndex =9
                                    Name ="txtMeta_Certified_Date"
                                    ControlSource ="Certified_Date"
                                    Format ="Short Date"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =8430
                                    LayoutCachedWidth =2640
                                    LayoutCachedHeight =8670
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextFontCharSet =204
                                            TextAlign =3
                                            Left =240
                                            Top =8430
                                            Width =1080
                                            Height =240
                                            FontSize =10
                                            Name ="lblMeta_Certified"
                                            Caption ="Certified"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =8430
                                            LayoutCachedWidth =1320
                                            LayoutCachedHeight =8670
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =215
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =2880
                                    Left =2760
                                    Top =8430
                                    Width =2823
                                    Height =252
                                    FontSize =10
                                    TabIndex =10
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                                    Name ="cboMeta_Certified_Contact_ID"
                                    ControlSource ="Certified_By"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                                        "Last_Name, tlu_Contacts.First_Name; "
                                    ColumnWidths ="0;2880"
                                    StatusBarText ="Observer identifier"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =2760
                                    LayoutCachedTop =8430
                                    LayoutCachedWidth =5583
                                    LayoutCachedHeight =8682
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3600
                                    Top =6240
                                    Width =2220
                                    Height =360
                                    FontSize =14
                                    FontWeight =700
                                    TabIndex =11
                                    BackColor =15527148
                                    ForeColor =-2147483630
                                    Name ="lblPictures_Taken"
                                    ControlSource ="=\"Pictures Taken\""
                                    FontName ="Calibri"
                                    ConditionalFormat = Begin
                                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x5b00500069006300740075007200650073005f00540061006b0065006e005d00 ,
                                        0x3c003e00540072007500650000000000
                                    End

                                    LayoutCachedLeft =3600
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =5820
                                    LayoutCachedHeight =6600
                                    ConditionalFormat14 = Begin
                                        0x01000100000001000000000000000100000000000000dfa7a500160000005b00 ,
                                        0x500069006300740075007200650073005f00540061006b0065006e005d003c00 ,
                                        0x3e005400720075006500000000000000000000000000000000000000000000
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =223
                                    Left =4020
                                    Top =2415
                                    Width =300
                                    Height =300
                                    FontSize =12
                                    FontWeight =700
                                    TabIndex =12
                                    Name ="cmdNewUser"
                                    Caption ="+"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dddddddddddddddddddddddddddddddddd0000dddddddddd ,
                                        0xdd00ff07dddddddddd0ff7f07ddddddddd0fff7b07ddddddddd0fbb7b07ddddd ,
                                        0xdddd0bbb7b07ddddddddd0bbb0707ddddddddd0b077707ddddddddd07870007d ,
                                        0xdddddddd07001117ddddddddd009111ddddddddddd0191ddddddddddddd11ddd ,
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
                                    ControlTipText ="Add a new contact record"
                                    LeftPadding =15
                                    TopPadding =15
                                    RightPadding =15
                                    BottomPadding =15
                                    ImageData = Begin
                                        0x00000000
                                    End

                                    LayoutCachedLeft =4020
                                    LayoutCachedTop =2415
                                    LayoutCachedWidth =4320
                                    LayoutCachedHeight =2715
                                    WebImagePaddingLeft =1
                                    WebImagePaddingTop =1
                                    Overlaps =1
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =87
                                    Left =240
                                    Top =7410
                                    Width =5520
                                    Name ="Line137"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =7410
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =7410
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =87
                                    Left =240
                                    Top =6180
                                    Width =5520
                                    Name ="Line142"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =6180
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =6180
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =87
                                    Left =240
                                    Top =4260
                                    Width =5520
                                    Name ="Line143"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =4260
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =4260
                                End
                                Begin Label
                                    OverlapFlags =223
                                    TextAlign =1
                                    Left =240
                                    Top =1755
                                    Width =3480
                                    Height =311
                                    FontSize =14
                                    FontWeight =700
                                    Name ="lblContact_ID"
                                    Caption ="Participants and Roles"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =1755
                                    LayoutCachedWidth =3720
                                    LayoutCachedHeight =2066
                                End
                                Begin ComboBox
                                    OverlapFlags =215
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =3888
                                    Left =1680
                                    Top =6240
                                    Width =720
                                    Height =359
                                    FontSize =13
                                    TabIndex =13
                                    BackColor =16777215
                                    ForeColor =0
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    ConditionalFormat = Begin
                                        0x010000009e000000010000000100000000000000000000001e00000001000000 ,
                                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x490073004e0075006c006c0028005b00630062006f0054007200650065005f00 ,
                                        0x5300740061007400750073005d0029003d00540072007500650000000000
                                    End
                                    Name ="cboTree_Status"
                                    ControlSource ="Deer_Impact"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Deer Impact\")) ORDER BY tlu_"
                                        "Enumerations.Sort_Order;"
                                    ColumnWidths ="720;3168"
                                    StatusBarText ="Health status of this specimen"
                                    FontName ="Calibri"
                                    AllowValueListEdits =1
                                    InheritValueList =1
                                    LeftMargin =22
                                    TopMargin =22
                                    RightMargin =22
                                    BottomMargin =22

                                    LayoutCachedLeft =1680
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =2400
                                    LayoutCachedHeight =6599
                                    ConditionalFormat14 = Begin
                                        0x01000100000001000000000000000100000000000000dfa7a5001d0000004900 ,
                                        0x73004e0075006c006c0028005b00630062006f0054007200650065005f005300 ,
                                        0x740061007400750073005d0029003d0054007200750065000000000000000000 ,
                                        0x00000000000000000000000000
                                    End
                                End
                                Begin CommandButton
                                    FontUnderline = NotDefault
                                    OverlapFlags =215
                                    Left =240
                                    Top =6240
                                    Width =1380
                                    FontSize =13
                                    TabIndex =14
                                    ForeColor =6108695
                                    Name ="cmdOpen_Form_Deer_Impact"
                                    Caption ="Deer Impact"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    ControlTipText ="Open Form"
                                    ImageData = Begin
                                        0x00000000
                                    End
                                    BackStyle =0

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =6240
                                    LayoutCachedWidth =1620
                                    LayoutCachedHeight =6600
                                    Alignment =3
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =87
                                    Left =240
                                    Top =6660
                                    Width =5520
                                    Name ="Line147"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =6660
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =6660
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =300
                                    Top =6750
                                    Width =240
                                    TabIndex =15
                                    Name ="chk_Early_Detect"
                                    ControlSource ="Early_Detect"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =6750
                                    LayoutCachedWidth =540
                                    LayoutCachedHeight =6990
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            Left =540
                                            Top =6720
                                            Width =2100
                                            Height =240
                                            FontWeight =700
                                            Name ="Label149"
                                            Caption ="Early Detection Species"
                                            LayoutCachedLeft =540
                                            LayoutCachedTop =6720
                                            LayoutCachedWidth =2640
                                            LayoutCachedHeight =6960
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =300
                                    Top =7110
                                    Width =240
                                    Height =180
                                    TabIndex =16
                                    Name ="chk_Rare_Spp"
                                    ControlSource ="Rare_Spp"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =7110
                                    LayoutCachedWidth =540
                                    LayoutCachedHeight =7290
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =525
                                            Top =7080
                                            Width =1215
                                            Height =240
                                            FontWeight =700
                                            Name ="Label151"
                                            Caption ="Rare Species "
                                            LayoutCachedLeft =525
                                            LayoutCachedTop =7080
                                            LayoutCachedWidth =1740
                                            LayoutCachedHeight =7320
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =215
                                    Left =2820
                                    Top =6750
                                    Width =240
                                    TabIndex =17
                                    Name ="chk_Plot_Maint"
                                    ControlSource ="Plot_Maint"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =6750
                                    LayoutCachedWidth =3060
                                    LayoutCachedHeight =6990
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =3045
                                            Top =6720
                                            Width =1575
                                            Height =240
                                            FontWeight =700
                                            Name ="Label153"
                                            Caption ="Plot Maintenance"
                                            LayoutCachedLeft =3045
                                            LayoutCachedTop =6720
                                            LayoutCachedWidth =4620
                                            LayoutCachedHeight =6960
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =1695
                            Width =13980
                            Height =7380
                            Name ="pagTransects"
                            Caption ="Transect"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1695
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9075
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin OptionGroup
                                    SpecialEffect =1
                                    OverlapFlags =247
                                    Left =360
                                    Top =2715
                                    Width =1680
                                    Height =1200
                                    Name ="grpTransect_Selection"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="1"

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =2715
                                    LayoutCachedWidth =2040
                                    LayoutCachedHeight =3915
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextAlign =2
                                            Left =480
                                            Top =2595
                                            Width =1440
                                            Height =240
                                            FontSize =10
                                            BackColor =15527148
                                            ForeColor =0
                                            Name ="lblTransect_Selection"
                                            Caption ="Select a Transect"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =480
                                            LayoutCachedTop =2595
                                            LayoutCachedWidth =1920
                                            LayoutCachedHeight =2835
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =247
                                            Left =840
                                            Top =2955
                                            Height =390
                                            FontSize =14
                                            FontWeight =700
                                            OptionValue =360
                                            Name ="tglTransect360"
                                            Caption ="360"
                                            FontName ="Calibri"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120
                                            ImageData = Begin
                                                0x00000000
                                            End

                                            LayoutCachedLeft =840
                                            LayoutCachedTop =2955
                                            LayoutCachedWidth =1560
                                            LayoutCachedHeight =3345
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
                                        Begin ToggleButton
                                            OverlapFlags =247
                                            Left =480
                                            Top =3435
                                            Height =390
                                            FontSize =14
                                            FontWeight =700
                                            TabIndex =1
                                            OptionValue =240
                                            Name ="tglTransect240"
                                            Caption ="240"
                                            FontName ="Calibri"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120
                                            ImageData = Begin
                                                0x00000000
                                            End

                                            LayoutCachedLeft =480
                                            LayoutCachedTop =3435
                                            LayoutCachedWidth =1200
                                            LayoutCachedHeight =3825
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
                                        Begin ToggleButton
                                            OverlapFlags =247
                                            Left =1260
                                            Top =3435
                                            Height =390
                                            FontSize =14
                                            FontWeight =700
                                            TabIndex =2
                                            OptionValue =120
                                            Name ="tglTransect120"
                                            Caption ="120"
                                            FontName ="Calibri"
                                            LeftPadding =60
                                            RightPadding =75
                                            BottomPadding =120
                                            ImageData = Begin
                                                0x00000000
                                            End

                                            LayoutCachedLeft =1260
                                            LayoutCachedTop =3435
                                            LayoutCachedWidth =1980
                                            LayoutCachedHeight =3825
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
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =360
                                    Top =1935
                                    Width =1680
                                    Height =540
                                    FontSize =22
                                    FontWeight =700
                                    TabIndex =1
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtTransect_Selection"
                                    DefaultValue ="'360'"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =1935
                                    LayoutCachedWidth =2040
                                    LayoutCachedHeight =2475
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =795
                                    Top =4515
                                    Width =335
                                    Height =285
                                    TabIndex =2
                                    Name ="chkTransectChecked_360"
                                    ControlSource ="CWD_Check_360"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"

                                    LayoutCachedLeft =795
                                    LayoutCachedTop =4515
                                    LayoutCachedWidth =1130
                                    LayoutCachedHeight =4800
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =780
                                    Top =4995
                                    Width =335
                                    Height =285
                                    TabIndex =3
                                    Name ="chkTransectChecked_120"
                                    ControlSource ="CWD_Check_120"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"

                                    LayoutCachedLeft =780
                                    LayoutCachedTop =4995
                                    LayoutCachedWidth =1115
                                    LayoutCachedHeight =5280
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =780
                                    Top =5475
                                    Width =335
                                    Height =285
                                    TabIndex =4
                                    Name ="chkTransectChecked_240"
                                    ControlSource ="CWD_Check_240"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"

                                    LayoutCachedLeft =780
                                    LayoutCachedTop =5475
                                    LayoutCachedWidth =1115
                                    LayoutCachedHeight =5760
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =360
                                    Top =4200
                                    Width =1679
                                    Height =1650
                                    Name ="shpTransect_Checked"
                                    LayoutCachedLeft =360
                                    LayoutCachedTop =4200
                                    LayoutCachedWidth =2039
                                    LayoutCachedHeight =5850
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =420
                                    Top =4035
                                    Width =1515
                                    Height =240
                                    FontSize =10
                                    BackColor =15527148
                                    Name ="lblTransectChecked"
                                    Caption ="Transect Checked"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =420
                                    LayoutCachedTop =4035
                                    LayoutCachedWidth =1935
                                    LayoutCachedHeight =4275
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =2520
                                    Top =1905
                                    Width =10065
                                    Height =6435
                                    TabIndex =5
                                    Name ="fsub_Transects"
                                    SourceObject ="Form.fsub_Transects"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =2520
                                    LayoutCachedTop =1905
                                    LayoutCachedWidth =12585
                                    LayoutCachedHeight =8340
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1040
                                    Top =4395
                                    Width =705
                                    Height =375
                                    FontSize =16
                                    FontWeight =700
                                    TabIndex =6
                                    BackColor =15527148
                                    ForeColor =-2147483630
                                    Name ="lblTransectChecked_360"
                                    ControlSource ="=\"360\""
                                    FontName ="Calibri"
                                    ConditionalFormat = Begin
                                        0x010000008e000000010000000100000000000000000000001600000001010000 ,
                                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x5b004300570044005f0043006800650063006b005f003300360030005d003c00 ,
                                        0x3e00540072007500650000000000
                                    End

                                    LayoutCachedLeft =1040
                                    LayoutCachedTop =4395
                                    LayoutCachedWidth =1745
                                    LayoutCachedHeight =4770
                                    ConditionalFormat14 = Begin
                                        0x01000100000001000000000000000101000000000000dfa7a500150000005b00 ,
                                        0x4300570044005f0043006800650063006b005f003300360030005d003c003e00 ,
                                        0x5400720075006500000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1025
                                    Top =4875
                                    Width =705
                                    Height =375
                                    FontSize =16
                                    FontWeight =700
                                    TabIndex =7
                                    BackColor =15527148
                                    ForeColor =-2147483630
                                    Name ="lblTransectChecked_120"
                                    ControlSource ="=\"120\""
                                    FontName ="Calibri"
                                    ConditionalFormat = Begin
                                        0x010000008e000000010000000100000000000000000000001600000001010000 ,
                                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x5b004300570044005f0043006800650063006b005f003100320030005d003c00 ,
                                        0x3e00540072007500650000000000
                                    End

                                    LayoutCachedLeft =1025
                                    LayoutCachedTop =4875
                                    LayoutCachedWidth =1730
                                    LayoutCachedHeight =5250
                                    ConditionalFormat14 = Begin
                                        0x01000100000001000000000000000101000000000000dfa7a500150000005b00 ,
                                        0x4300570044005f0043006800650063006b005f003100320030005d003c003e00 ,
                                        0x5400720075006500000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1005
                                    Top =5355
                                    Width =705
                                    Height =375
                                    FontSize =16
                                    FontWeight =700
                                    TabIndex =8
                                    BackColor =15527148
                                    ForeColor =-2147483630
                                    Name ="lblTransectChecked_240"
                                    ControlSource ="=\"240\""
                                    FontName ="Calibri"
                                    ConditionalFormat = Begin
                                        0x010000008e000000010000000100000000000000000000001600000001010000 ,
                                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x5b004300570044005f0043006800650063006b005f003200340030005d003c00 ,
                                        0x3e00540072007500650000000000
                                    End

                                    LayoutCachedLeft =1005
                                    LayoutCachedTop =5355
                                    LayoutCachedWidth =1710
                                    LayoutCachedHeight =5730
                                    ConditionalFormat14 = Begin
                                        0x01000100000001000000000000000101000000000000dfa7a500150000005b00 ,
                                        0x4300570044005f0043006800650063006b005f003200340030005d003c003e00 ,
                                        0x5400720075006500000000000000000000000000000000000000000000
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =1695
                            Width =14055
                            Height =7380
                            Name ="pagTrees"
                            Caption ="Trees"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1695
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9075
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =60
                                    Top =1724
                                    Width =14054
                                    Height =7094
                                    Name ="fsub_Tree_Data"
                                    SourceObject ="Form.fsub_Tree_Data"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =1724
                                    LayoutCachedWidth =14114
                                    LayoutCachedHeight =8818
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =1695
                            Width =14055
                            Height =7380
                            Name ="pagSaplings"
                            Caption ="Saplings"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1695
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9075
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =60
                                    Top =1724
                                    Width =14054
                                    Height =6599
                                    Name ="fsub_Sapling_Data"
                                    SourceObject ="Form.fsub_Sapling_Data"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"
                                    OnEnter ="[Event Procedure]"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =1724
                                    LayoutCachedWidth =14114
                                    LayoutCachedHeight =8323
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =1695
                            Width =14055
                            Height =7380
                            Name ="pagQuadrats"
                            Caption ="Quadrats"
                            ImageData = Begin
                                0x00000000
                            End
                            LayoutCachedLeft =60
                            LayoutCachedTop =1695
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9075
                            BorderThemeColorIndex =-1
                            BorderShade =100.0
                            GridlineThemeColorIndex =-1
                            GridlineShade =100.0
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    SpecialEffect =3
                                    Left =60
                                    Top =1724
                                    Width =14054
                                    Height =6599
                                    Name ="fsub_Quadrats"
                                    SourceObject ="Form.fsub_Quadrats"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =1724
                                    LayoutCachedWidth =14114
                                    LayoutCachedHeight =8323
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =12300
                    Top =600
                    Width =900
                    Height =660
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdEditLocation"
                    Caption ="Edit Location"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Edit the current location record."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12300
                    LayoutCachedTop =600
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =1260
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
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
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =223
                    Left =240
                    Top =120
                    Width =1080
                    Height =330
                    FontWeight =700
                    TabIndex =6
                    Name ="tglBrowse_Edit"
                    Caption ="Editing OFF"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Toggle between browse and edit modes"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =1320
                    LayoutCachedHeight =450
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
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
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =480
                    Width =2583
                    Height =420
                    ColumnWidth =1440
                    FontSize =18
                    FontWeight =700
                    Name ="txtPlot_Name"
                    StatusBarText ="Unique identifier for each sample location"
                    FontName ="Calibri"

                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =2643
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =10200
                    Top =600
                    Width =1005
                    TabIndex =7
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =600
                    LayoutCachedWidth =11205
                    LayoutCachedHeight =840
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5040
                    Top =600
                    Width =1620
                    Height =300
                    FontSize =12
                    TabIndex =8
                    Name ="txtProtocol_Name"
                    ControlSource ="=\"Protocol: \" & [Protocol_Name]"
                    FontName ="Calibri"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =600
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =3000
                    Top =960
                    Width =3540
                    TabIndex =9
                    Name ="txtEvent_ID"
                    ControlSource ="Event_ID"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =960
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1200
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =7080
                    Top =1260
                    Width =1980
                    Height =300
                    FontSize =10
                    ForeColor =1279872587
                    Name ="lblLink_to_Google_Maps"
                    Caption ="Show on Google Maps"
                    FontName ="Calibri"
                    HyperlinkAddress ="http://maps.google.com/maps?q=ANTI-0045@39.483246,-77.743994&iwloc=A&t=h"
                    LayoutCachedLeft =7080
                    LayoutCachedTop =1260
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =1560
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    Left =9180
                    Top =1260
                    Width =1800
                    Height =300
                    FontSize =10
                    ForeColor =1279872587
                    Name ="lblLink_To_Plot_Photos"
                    Caption ="Explore Plot Photos"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =9180
                    LayoutCachedTop =1260
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =1560
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =13320
                    Top =600
                    Width =900
                    Height =660
                    FontWeight =700
                    TabIndex =10
                    Name ="cmdTriggerReport"
                    Caption ="Event Report"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Display a Summary Report of this Event."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13320
                    LayoutCachedTop =600
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =1260
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
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
                    OverlapFlags =255
                    Left =11280
                    Top =600
                    Width =900
                    Height =660
                    FontWeight =700
                    TabIndex =11
                    Name ="cmdPlot_Chart"
                    Caption ="Plot Chart"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Display the current location plot chart."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =11280
                    LayoutCachedTop =600
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =1260
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
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
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6720
                    Top =900
                    Width =4560
                    Height =300
                    FontSize =12
                    TabIndex =12
                    Name ="txtSlope_Aspect"
                    FontName ="Calibri"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =900
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =1200
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

Private Sub chkPictures_Taken_AfterUpdate()
    lblPictures_Taken.Requery
End Sub

Private Sub chkTransectChecked_120_AfterUpdate()
    lblTransectChecked_120.Requery
End Sub

Private Sub chkTransectChecked_240_AfterUpdate()
    lblTransectChecked_240.Requery
End Sub

Private Sub chkTransectChecked_360_AfterUpdate()
    lblTransectChecked_360.Requery
End Sub

Private Sub cmdOpen_Form_Deer_Impact_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Deer_Impact"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdOpen_Popup_Click:
    Exit Sub
Err_cmdOpen_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpen_Popup_Click
End Sub

Private Sub cmdPlot_Chart_Click()
Dim strOpenargs As String
Dim strCriteria As String
    If Not IsNothing(Me!txtLocation_ID) Then
        strOpenargs = XML_Tag("FormFrom", Me.Name)
        strOpenargs = strOpenargs & XML_Tag("ControlFrom", "txtLocation_ID")
        strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Plot_Chart", , , strCriteria, acFormEdit, acWindowNormal, strOpenargs
    End If
End Sub

Private Sub cmdTriggerReport_Click()
On Error GoTo Err_Handler
    Dim strDocName As String
    Dim strCriteria As String
    
    strDocName = "rpt_Event_Summary_Unfiltered"
    strCriteria = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
    DoCmd.OpenReport strDocName, acPreview, , strCriteria
    
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub

' =================================
' FORM NAME:    frm_Data_Entry
' Description:  Primary field data entry form
' Data source:  tbl_Locations
' Data access:  edit; allow additions off except for new records
' Pages:        none
' Functions:    Update_Loc_Info, ValidateForm
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Source/date:  John R. Boetsch, June 2006
' Revisions:    Simon Kingston, October - January 2006
'                   - extensive updates, adding GUID generation code, new controls
' =================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strCaptionSuffix As String
    Dim booEditOn As Boolean

    ' Set the opening parameters depending on the arguments passed from the previous form
    If Me.OpenArgs = "(Browsing)" Then
        strCaptionSuffix = " - " & Me.OpenArgs
        booEditOn = False
    ElseIf Me.OpenArgs = "(Creating)" Then
        strCaptionSuffix = " - " & Me.OpenArgs
        booEditOn = True
    ElseIf Me.OpenArgs <> "" Then
        strCaptionSuffix = " - " & "No arguments"
        booEditOn = False
    End If
    
    'TO DO
    'Insert code here to update Plot Status in the Location table if this is the first sampling of this plot.
        
    Me.Caption = Me.Caption & strCaptionSuffix
    SetEditMode (booEditOn)

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Current()
    'Update fields in header from Locations table
    Update_Loc_Info
    'Enable edit location function if there is an active location
    Me!cmdEditLocation.Enabled = Not IsNull(Me!txtLocation_ID)
    'Event groups not implemented in this database
    'Me!cmdEditEventGroup.Enabled = Not IsNull(Me!cboEvent_Group_ID)
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Create the GUID primary key value if needed for a string GUID
    If IsNull(Me!Event_ID) Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me!Event_ID = fxnGUIDGen
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub ValidateForm()
' Description:  Confirms that a Location and Start Date are entered
' References:   none
' Source/date:  Simon Kingston, Dec. 2006
' Revisions:    <name, date, desc - add lines as you go>

If IsNull(Me!txtLocation_ID) Then
    MsgBox "You must select a location before you can enter record details!", vbExclamation, "Enter Location First"
    'Me!cboLocation_ID.SetFocus
Else
    If IsNull(Me!txtStart_date) Then
        MsgBox "You must enter a start date before you can enter record details!", vbExclamation, "Enter Start Date"
        'Me!txtStart_date.SetFocus
    End If
End If
End Sub

Private Sub grpTransect_Selection_AfterUpdate()
Dim strTransect As String
    strTransect = Me!grpTransect_Selection.Value
    Me.txtTransect_Selection.Value = "'" & strTransect & "'"
    Forms![frm_Events]![fsub_Transects]!txtTransect_Azimuth.DefaultValue = "'" & strTransect & "'"
    Forms![frm_Events]![fsub_Transects].Form.Filter = "[Transect_Azimuth] = """ & strTransect & """ "
    Forms![frm_Events]![fsub_Transects].Form.FilterOn = True
End Sub

Private Sub lblLink_To_Plot_Photos_Click()
On Error GoTo Err_Handler

    Dim RetVal As Double
    Dim RootFolder As String
    Dim PhotoFolder As String
    
    RootFolder = "T:\I&M"
    PhotoFolder = "T:\I&M\Monitoring\Forest_Vegetation\Photos\"
    If FolderExists(PhotoFolder & Me!txtPlot_Name) Then
        RetVal = Shell("explorer /e,/root, " & PhotoFolder & Me!txtPlot_Name, vbNormalFocus)
        GoTo Exit_Procedure
    Else
        If FolderExists(RootFolder) Then
            MsgBox ("Folder for this plot not found....Opening the root of the Photos folder.")
            RetVal = Shell("explorer /e,/root, " & PhotoFolder, vbNormalFocus)
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

Private Sub subObservers_Enter()
    ValidateForm
End Sub

Private Sub cboEvent_Group_ID_AfterUpdate()
    Me!cmdEditEventGroup.Enabled = Not IsNothing(Me!cboEvent_Group_ID)
End Sub

Private Sub cmdEditLocation_Click()
Dim strOpenargs As String
Dim strCriteria As String

    If Not IsNothing(Me!txtLocation_ID) Then
        strOpenargs = XML_Tag("FormFrom", Me.Name)
        strOpenargs = strOpenargs & XML_Tag("ControlFrom", "txtLocation_ID")
        strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Locations", , , strCriteria, acFormEdit, acWindowNormal, strOpenargs
    End If
End Sub

Private Sub cmdNewUser_Click()
    DoCmd.OpenForm "frm_Contacts", , , , acFormAdd, , "new"
End Sub

Public Sub Update_Loc_Info()
' Description:  Updates associated location information when Location_ID is updated
' References:   GetCriteriaString
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:    <name, date, desc - add lines as you go>

Dim strXY As Variant
Dim strSlopeAspect As String

Dim strCriteria As String

If IsNull(Me!txtLocation_ID) Then
    Me!txtXY = Null
    Me!txtUnit_Code = Null
    Me!txtSlope_Aspect = Null
    
    lblLink_to_Google_Maps.HyperlinkAddress = "http://maps.google.com"
    'lblLink_To_Plot_Photos.Tag = "T:\I&M\Monitoring\Forest_Vegetation\Photos"
Else
    strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
    strXY = "UTM 18N NAD83 E: " & Nz(DLookup("X_Coord", "tbl_Locations", strCriteria), "")
    strXY = strXY & "  N: " & Nz(DLookup("Y_Coord", "tbl_Locations", strCriteria), "")
    Me!txtXY = strXY
    strSlopeAspect = "Slope: " & Nz(DLookup("Slope", "tbl_Locations", strCriteria), "")
    strSlopeAspect = strSlopeAspect & "; Aspect: " & Nz(DLookup("Aspect", "tbl_Locations", strCriteria), "")
    Me!txtSlope_Aspect = strSlopeAspect
    
    Me!txtPlot_Name = DLookup("Plot_Name", "tbl_Locations", strCriteria)
    lblLink_to_Google_Maps.HyperlinkAddress = "http://maps.google.com/maps?q=" & Me!txtPlot_Name & "@" & DLookup("Lat_WGS84", "tbl_Locations", strCriteria) & "," & DLookup("Lon_WGS84", "tbl_Locations", strCriteria) & "&iwloc=A&t=h"
    'lblLink_To_Plot_Photos.Tag = "T:\I&M\Monitoring\Forest_Vegetation\Photos\" & Me!txtPlot_Name
End If
End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Close()
If IsLoaded("frm_Data_Gateway") Then
    Forms("frm_Data_Gateway").Requery
End If
End Sub

Private Sub tglBrowse_Edit_Click()
    'Call the SetEditMode subroutine with the current status of the Browse/Edit toggle
    Me.SetEditMode (Me!tglBrowse_Edit)
End Sub

Public Sub SetEditMode(booEditOn As Boolean)
' Description:  Toggles the form between browse and edit mode
' Parameters:   booFilterOn = true if edit mode, false if browse mode
' Returns:      none
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  Simon Kingston, 1/17/2007
' Revisions:    Mark Lehman 3/15/2010 Repurposed version of FilterGateway by Kingston

On Error GoTo Error_Handler

Me!tglBrowse_Edit = booEditOn

If booEditOn Then
    Me!tglBrowse_Edit.Caption = "Editing ON"
    Me!lblEvent_Form_Header.BackColor = RGB(128, 0, 0)
Else
    Me!tglBrowse_Edit.Caption = "Editing OFF"
    Me!lblEvent_Form_Header.BackColor = vbBlack
End If

'Me.FilterOn = booEditOn

Exit_Handler:
    Exit Sub
Error_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
        "Error encountered (SetEditMode)"
    Resume Exit_Handler
End Sub
Private Sub cmdAdd_Edit_Event_Note_Click()
On Error GoTo Err_cmdAdd_Edit_Event_Note_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Event_Add_Note"
    
    stLinkCriteria = "[Event_ID]=" & "'" & Me![txtEvent_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdAdd_Edit_Event_Note_Click:
    Exit Sub

Err_cmdAdd_Edit_Event_Note_Click:
    MsgBox Err.Description
    Resume Exit_cmdAdd_Edit_Event_Note_Click
    
End Sub

Private Sub cmdAdd_Event_Note_Click()
On Error GoTo Err_cmdAdd_Event_Note_Click

Me.Requery
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Event_Add_Note"
    
    stLinkCriteria = "[Event_ID]=" & "'" & Me![txtEvent_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_cmdAdd_Event_Note_Click:
    Exit Sub

Err_cmdAdd_Event_Note_Click:
    MsgBox Err.Description
    Resume Exit_cmdAdd_Event_Note_Click
End Sub
