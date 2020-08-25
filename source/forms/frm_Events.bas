Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
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
    ItemSuffix =167
    Left =735
    Top =45
    Right =15135
    Bottom =10395
    DatasheetGridlinesColor =12632256
    Filter ="[Event_ID]='{445DF397-13FD-4344-9209-5889A362FB7C}'"
    RecSrcDt = Begin
        0x47be11900a4be540
    End
    RecordSource ="qfrm_Events"
    Caption ="NCRN Sampling Event - (Browsing)"
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
        Begin FormHeader
            Height =540
            BackColor =0
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
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
                    OverlapFlags =223
                    Left =13440
                    Top =90
                    Width =900
                    Height =330
                    FontSize =10
                    FontWeight =700
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13440
                    LayoutCachedTop =90
                    LayoutCachedWidth =14340
                    LayoutCachedHeight =420
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
                Begin ToggleButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =120
                    Top =90
                    Width =1080
                    Height =330
                    ColumnOrder =0
                    FontWeight =700
                    TabIndex =1
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

                    LayoutCachedLeft =120
                    LayoutCachedTop =90
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =420
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
                    Visible = NotDefault
                    OverlapFlags =215
                    Left =10620
                    Top =75
                    Width =2766
                    Height =366
                    FontWeight =600
                    TabIndex =2
                    Name ="btnConvertPseudoEvent"
                    Caption ="  Convert to Regular Event"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddd00d00d00dddddddd00d00d00dddddddd ,
                        0xdddddddddddddddddddddddddddddddddddddddddddddd00ddddd7ddddddd000 ,
                        0xd7dd7c7ddddd000d7c7dd7ddddd000ddd7dddddddd000dddddddddddd000dddd ,
                        0xddd7dddd000ddddddd7c7dd000ddddddddd7dd0b0dddddddddddd0b0dddddddd ,
                        0xddddd70ddddddddd000000000000000000000000000000000000000000000000 ,
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
                    ControlTipText ="Convert pseudoevent to normal event (NOTE - you cannot convert it back so make s"
                        "ure!)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =10620
                    LayoutCachedTop =75
                    LayoutCachedWidth =13386
                    LayoutCachedHeight =441
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =1796857
                    BackThemeColorIndex =5
                    BorderColor =1796857
                    BorderThemeColorIndex =5
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =413911
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =24
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =1320
                    Top =60
                    Width =3120
                    Height =450
                    FontSize =18
                    FontWeight =700
                    BackColor =-2147483643
                    ForeColor =16777215
                    Name ="lblPseudoEventFlag"
                    Caption ="* PSEUDO EVENT *"
                    FontName ="Berlin Sans FB Demi"
                    LayoutCachedLeft =1320
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =510
                End
                Begin Label
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =9840
                    Top =60
                    Width =600
                    Height =420
                    FontWeight =600
                    BorderColor =52479
                    ForeColor =16776960
                    Name ="lblQCMode"
                    Caption ="QC MODE"
                    LayoutCachedLeft =9840
                    LayoutCachedTop =60
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9465
            BackColor =15527148
            Name ="Detail"
            Begin
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
                    Top =120
                    Width =4560
                    Height =300
                    FontSize =12
                    TabIndex =11
                    Name ="txtXY"
                    FontName ="Calibri"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =120
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2700
                    Width =2160
                    Height =420
                    FontSize =18
                    FontWeight =700
                    TabIndex =13
                    Name ="txtStart_Date"
                    ControlSource ="Event_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    FontName ="Calibri"

                    LayoutCachedLeft =2700
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =420
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =255
                    Width =14400
                    Height =1140
                    Name ="rctPseudoEvent"
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =1140
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =255
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Width =2583
                    Height =420
                    ColumnWidth =1440
                    FontSize =18
                    FontWeight =700
                    TabIndex =12
                    Name ="txtPlot_Name"
                    StatusBarText ="Unique identifier for each sample location"
                    FontName ="Calibri"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2643
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =10200
                    Top =120
                    Width =1005
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                    LayoutCachedLeft =10200
                    LayoutCachedTop =120
                    LayoutCachedWidth =11205
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5040
                    Top =120
                    Width =1620
                    Height =300
                    FontSize =12
                    TabIndex =1
                    Name ="txtProtocol_Name"
                    ControlSource ="=\"Protocol: \" & [Protocol_Name]"
                    FontName ="Calibri"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =120
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =420
                End
                Begin Tab
                    MultiRow = NotDefault
                    OverlapFlags =119
                    Top =1140
                    Width =14250
                    Height =8325
                    FontSize =12
                    TabIndex =2
                    Name ="tabctlData"
                    FontName ="Calibri"

                    LayoutCachedTop =1140
                    LayoutCachedWidth =14250
                    LayoutCachedHeight =9465
                    Begin
                        Begin Page
                            OverlapFlags =119
                            Left =135
                            Top =1635
                            Width =13980
                            Height =7700
                            Name ="pagIntro"
                            Caption ="Intro"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1635
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9335
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
                                    Top =2310
                                    Width =5520
                                    Height =2100
                                    Name ="subObservers"
                                    SourceObject ="Form.fsub_Observers"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =2310
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =4410
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    BorderWidth =3
                                    Left =6060
                                    Top =2370
                                    Width =7980
                                    Height =6885
                                    TabIndex =1
                                    Name ="fsub_Note_History"
                                    SourceObject ="Form.fsub_Note_History"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                    LayoutCachedLeft =6060
                                    LayoutCachedTop =2370
                                    LayoutCachedWidth =14040
                                    LayoutCachedHeight =9255
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =6060
                                            Top =2070
                                            Width =2760
                                            Height =300
                                            FontSize =12
                                            FontWeight =700
                                            Name ="fsub_Note_History Label"
                                            Caption ="Event History"
                                            FontName ="Calibri"
                                            EventProcPrefix ="fsub_Note_History_Label"
                                            LayoutCachedLeft =6060
                                            LayoutCachedTop =2070
                                            LayoutCachedWidth =8820
                                            LayoutCachedHeight =2370
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    OldBorderStyle =0
                                    SpecialEffect =0
                                    Left =240
                                    Top =4935
                                    Width =5460
                                    TabIndex =2
                                    Name ="subPlot_Floor_Conditions"
                                    SourceObject ="Form.fsub_Plot_Floor_Condition_Data"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =240
                                    LayoutCachedTop =4935
                                    LayoutCachedWidth =5700
                                    LayoutCachedHeight =6375
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =240
                                            Top =4635
                                            Width =3480
                                            Height =300
                                            FontSize =14
                                            FontWeight =700
                                            Name ="lblPlot Floor Conditions"
                                            Caption ="Plot Floor Conditions"
                                            FontName ="Calibri"
                                            EventProcPrefix ="lblPlot_Floor_Conditions"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =4635
                                            LayoutCachedWidth =3720
                                            LayoutCachedHeight =4935
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =12060
                                    Top =1950
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
                                    LayoutCachedTop =1950
                                    LayoutCachedWidth =14040
                                    LayoutCachedHeight =2250
                                    ForeThemeColorIndex =0
                                    UseTheme =255
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
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                    Overlaps =1
                                End
                                Begin CheckBox
                                    SpecialEffect =0
                                    OverlapFlags =255
                                    Left =3360
                                    Top =6615
                                    Height =210
                                    TabIndex =4
                                    BorderColor =2366701
                                    Name ="chkPictures_Taken"
                                    ControlSource ="Pictures_Taken"
                                    AfterUpdate ="[Event Procedure]"

                                    LayoutCachedLeft =3360
                                    LayoutCachedTop =6615
                                    LayoutCachedWidth =3620
                                    LayoutCachedHeight =6825
                                End
                                Begin Rectangle
                                    SpecialEffect =4
                                    BorderWidth =3
                                    OverlapFlags =255
                                    Left =240
                                    Top =7965
                                    Width =5520
                                    Height =1080
                                    Name ="boxMetadata"
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =7965
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =9045
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =240
                                    Top =7725
                                    Width =1260
                                    Height =240
                                    FontSize =10
                                    FontWeight =700
                                    Name ="lblMetadata_Box"
                                    Caption ="Metadata"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =7725
                                    LayoutCachedWidth =1500
                                    LayoutCachedHeight =7965
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    Left =1440
                                    Top =8085
                                    Width =1200
                                    FontSize =10
                                    TabIndex =5
                                    Name ="txtMeta_Updated_Date"
                                    ControlSource ="Updated_Date"
                                    Format ="Short Date"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =8085
                                    LayoutCachedWidth =2640
                                    LayoutCachedHeight =8325
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =204
                                            TextAlign =3
                                            Left =240
                                            Top =8085
                                            Width =1080
                                            Height =240
                                            FontSize =10
                                            Name ="lblMeta_Updated"
                                            Caption ="Updated"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =8085
                                            LayoutCachedWidth =1320
                                            LayoutCachedHeight =8325
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =2880
                                    Left =2760
                                    Top =8085
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
                                    LayoutCachedTop =8085
                                    LayoutCachedWidth =5583
                                    LayoutCachedHeight =8337
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    Left =1440
                                    Top =8385
                                    Width =1200
                                    FontSize =10
                                    TabIndex =7
                                    Name ="txtMeta_Verified_Date"
                                    ControlSource ="Verified_Date"
                                    Format ="Short Date"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =8385
                                    LayoutCachedWidth =2640
                                    LayoutCachedHeight =8625
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =204
                                            TextAlign =3
                                            Left =240
                                            Top =8385
                                            Width =1080
                                            Height =240
                                            FontSize =10
                                            Name ="lblMeta_Verified"
                                            Caption ="Verified"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =8385
                                            LayoutCachedWidth =1320
                                            LayoutCachedHeight =8625
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =2880
                                    Left =2760
                                    Top =8385
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
                                    LayoutCachedTop =8385
                                    LayoutCachedWidth =5583
                                    LayoutCachedHeight =8637
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    Left =1440
                                    Top =8685
                                    Width =1200
                                    FontSize =10
                                    TabIndex =9
                                    Name ="txtMeta_Certified_Date"
                                    ControlSource ="Certified_Date"
                                    Format ="Short Date"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1440
                                    LayoutCachedTop =8685
                                    LayoutCachedWidth =2640
                                    LayoutCachedHeight =8925
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextFontCharSet =204
                                            TextAlign =3
                                            Left =240
                                            Top =8685
                                            Width =1080
                                            Height =240
                                            FontSize =10
                                            Name ="lblMeta_Certified"
                                            Caption ="Certified"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =8685
                                            LayoutCachedWidth =1320
                                            LayoutCachedHeight =8925
                                        End
                                    End
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    OverlapFlags =247
                                    TextFontCharSet =204
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =2880
                                    Left =2760
                                    Top =8685
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
                                    LayoutCachedTop =8685
                                    LayoutCachedWidth =5583
                                    LayoutCachedHeight =8937
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3600
                                    Top =6495
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
                                    LayoutCachedTop =6495
                                    LayoutCachedWidth =5820
                                    LayoutCachedHeight =6855
                                    ConditionalFormat14 = Begin
                                        0x01000100000001000000000000000100000000000000dfa7a500160000005b00 ,
                                        0x500069006300740075007200650073005f00540061006b0065006e005d003c00 ,
                                        0x3e005400720075006500000000000000000000000000000000000000000000
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =255
                                    Left =4020
                                    Top =2670
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
                                    LayoutCachedTop =2670
                                    LayoutCachedWidth =4320
                                    LayoutCachedHeight =2970
                                    WebImagePaddingLeft =1
                                    WebImagePaddingTop =1
                                    Overlaps =1
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =119
                                    Left =240
                                    Top =7665
                                    Width =5520
                                    Name ="lnMetadata"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =7665
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =7665
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =119
                                    Left =240
                                    Top =6435
                                    Width =5520
                                    Name ="lnDeerImpact"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =6435
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =6435
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =119
                                    Left =240
                                    Top =4515
                                    Width =5520
                                    Name ="lnPlotFloor"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =4515
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =4515
                                End
                                Begin Label
                                    OverlapFlags =255
                                    TextAlign =1
                                    Left =240
                                    Top =2010
                                    Width =3480
                                    Height =311
                                    FontSize =14
                                    FontWeight =700
                                    Name ="lblContact_ID"
                                    Caption ="Participants and Roles"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =2010
                                    LayoutCachedWidth =3720
                                    LayoutCachedHeight =2321
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =1
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =3888
                                    Left =1680
                                    Top =6495
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
                                    LayoutCachedTop =6495
                                    LayoutCachedWidth =2400
                                    LayoutCachedHeight =6854
                                    ConditionalFormat14 = Begin
                                        0x01000100000001000000000000000100000000000000dfa7a5001d0000004900 ,
                                        0x73004e0075006c006c0028005b00630062006f0054007200650065005f005300 ,
                                        0x740061007400750073005d0029003d0054007200750065000000000000000000 ,
                                        0x00000000000000000000000000
                                    End
                                End
                                Begin CommandButton
                                    FontUnderline = NotDefault
                                    OverlapFlags =247
                                    Left =240
                                    Top =6495
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
                                    LayoutCachedTop =6495
                                    LayoutCachedWidth =1620
                                    LayoutCachedHeight =6855
                                    Alignment =3
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Line
                                    LineSlant = NotDefault
                                    OverlapFlags =119
                                    Left =240
                                    Top =6915
                                    Width =5520
                                    Name ="lnCheckboxes"
                                    GridlineColor =10921638
                                    LayoutCachedLeft =240
                                    LayoutCachedTop =6915
                                    LayoutCachedWidth =5760
                                    LayoutCachedHeight =6915
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =300
                                    Top =7005
                                    Width =240
                                    TabIndex =15
                                    Name ="chk_Early_Detect"
                                    ControlSource ="Early_Detect"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =7005
                                    LayoutCachedWidth =540
                                    LayoutCachedHeight =7245
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =540
                                            Top =6975
                                            Width =2100
                                            Height =240
                                            FontWeight =700
                                            Name ="lblEarlyDetectSpecies"
                                            Caption ="Early Detection Species"
                                            LayoutCachedLeft =540
                                            LayoutCachedTop =6975
                                            LayoutCachedWidth =2640
                                            LayoutCachedHeight =7215
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =300
                                    Top =7365
                                    Width =240
                                    Height =180
                                    TabIndex =16
                                    Name ="chk_Rare_Spp"
                                    ControlSource ="Rare_Spp"

                                    LayoutCachedLeft =300
                                    LayoutCachedTop =7365
                                    LayoutCachedWidth =540
                                    LayoutCachedHeight =7545
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =525
                                            Top =7335
                                            Width =1215
                                            Height =240
                                            FontWeight =700
                                            Name ="lblRareSpecies"
                                            Caption ="Rare Species "
                                            LayoutCachedLeft =525
                                            LayoutCachedTop =7335
                                            LayoutCachedWidth =1740
                                            LayoutCachedHeight =7575
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2820
                                    Top =7005
                                    Width =240
                                    TabIndex =17
                                    Name ="chk_Plot_Maint"
                                    ControlSource ="Plot_Maint"

                                    LayoutCachedLeft =2820
                                    LayoutCachedTop =7005
                                    LayoutCachedWidth =3060
                                    LayoutCachedHeight =7245
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =3045
                                            Top =6975
                                            Width =1575
                                            Height =240
                                            FontWeight =700
                                            Name ="lblPlotMaintenance"
                                            Caption ="Plot Maintenance"
                                            LayoutCachedLeft =3045
                                            LayoutCachedTop =6975
                                            LayoutCachedWidth =4620
                                            LayoutCachedHeight =7215
                                        End
                                    End
                                End
                                Begin CommandButton
                                    Visible = NotDefault
                                    OverlapFlags =255
                                    Left =4140
                                    Top =1680
                                    Width =576
                                    Height =576
                                    TabIndex =18
                                    Name ="btnFlag"
                                    Caption ="btnFlag"
                                    PictureData = Begin
                                        0x2800000020000000200000000100180000000000000c0000c40e0000c40e0000 ,
                                        0x0000000000000000b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7 ,
                                        0xb8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9 ,
                                        0xb7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8 ,
                                        0xb9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7 ,
                                        0xb8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9 ,
                                        0xb7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8 ,
                                        0xb9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7 ,
                                        0xb8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9 ,
                                        0xb7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8 ,
                                        0xb9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7000000000000b8b9b7 ,
                                        0xb8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b6b8b9b6b8b9b7b8b9 ,
                                        0xb7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8b9b7b8 ,
                                        0xb9b7b8b9b7b8b9b7b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000696aed ,
                                        0xaaacc4b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b60000000000006c6deb ,
                                        0x1415ff7a7be5b7b9b7b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000a5a7c8 ,
                                        0x0004ff0004ff2c2dfc8889ddb8b9b6b8b9b6b8b9b6b8b9b69798d45d5ef12c2d ,
                                        0xfc0f0fff0004ff2a2bfd7d7ee3b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0x2829fd0004ff0004ff0004ff2122fe5354f45758f32829fd0004ff0004ff0004 ,
                                        0xff0004ff0004ff0004ff0004ff4b4cf6b4b6bab8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0x5a5bf20004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004 ,
                                        0xff0004ff0004ff0004ff0004ff0004ff4041f9b6b7b9b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0x7475e80004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004 ,
                                        0xff0004ff0004ff0004ff0004ff0004ff0004ff5556f4b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0x7475e80004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004 ,
                                        0xff0004ff0004ff0004ff0004ff0004ff0004ff0004ff7c7de4b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6000000000000b8b9b6 ,
                                        0x3e3efa0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004 ,
                                        0xff0004ff0004ff0004ff2c2cfc4243f93a3afa0f0fff0f0fffa4a6c9b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b60000000000008b8ddb ,
                                        0x0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004 ,
                                        0xff0004ff494af7a1a2ccb8b9b6b8b9b6b8b9b6b3b5bb7f80e26465efb8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b60000000000002e2efc ,
                                        0x0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff0004ff2c2c ,
                                        0xfc9294d6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b60000000000009d9fcf ,
                                        0x5c5df11819fe0004ff0004ff0004ff0004ff0004ff0004ff0f0fff7273e9b6b8 ,
                                        0xb8b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b6b7b88f91d8595af22627fd0004ff1c1cfe6061f0adafc1b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6 ,
                                        0xb8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9 ,
                                        0xb6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8b9b6b8 ,
                                        0xb9b6b8b9b6b8b9b6
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Find Next"
                                    Picture ="flag_red.bmp"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =4140
                                    LayoutCachedTop =1680
                                    LayoutCachedWidth =4716
                                    LayoutCachedHeight =2256
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Shape =1
                                    BorderColor =8289145
                                    BorderThemeColorIndex =4
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-49
                                    WebImagePaddingTop =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =1635
                            Width =13980
                            Height =7695
                            Name ="pagTransects"
                            Caption ="Transect"
                            LayoutCachedLeft =135
                            LayoutCachedTop =1635
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9330
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
                                    Top =2970
                                    Width =1680
                                    Height =1200
                                    Name ="grpTransect_Selection"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="1"

                                    LayoutCachedLeft =360
                                    LayoutCachedTop =2970
                                    LayoutCachedWidth =2040
                                    LayoutCachedHeight =4170
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =247
                                            TextAlign =2
                                            Left =480
                                            Top =2850
                                            Width =1440
                                            Height =240
                                            FontSize =10
                                            BackColor =15527148
                                            ForeColor =0
                                            Name ="lblTransect_Selection"
                                            Caption ="Select a Transect"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =480
                                            LayoutCachedTop =2850
                                            LayoutCachedWidth =1920
                                            LayoutCachedHeight =3090
                                        End
                                        Begin ToggleButton
                                            OverlapFlags =247
                                            Left =840
                                            Top =3210
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
                                            LayoutCachedTop =3210
                                            LayoutCachedWidth =1560
                                            LayoutCachedHeight =3600
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
                                            OverlapFlags =247
                                            Left =480
                                            Top =3690
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
                                            LayoutCachedTop =3690
                                            LayoutCachedWidth =1200
                                            LayoutCachedHeight =4080
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
                                            OverlapFlags =247
                                            Left =1260
                                            Top =3690
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
                                            LayoutCachedTop =3690
                                            LayoutCachedWidth =1980
                                            LayoutCachedHeight =4080
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
                                Begin TextBox
                                    Locked = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =360
                                    Top =2190
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
                                    LayoutCachedTop =2190
                                    LayoutCachedWidth =2040
                                    LayoutCachedHeight =2730
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =795
                                    Top =4770
                                    Width =335
                                    Height =285
                                    TabIndex =2
                                    Name ="chkTransectChecked_360"
                                    ControlSource ="CWD_Check_360"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"

                                    LayoutCachedLeft =795
                                    LayoutCachedTop =4770
                                    LayoutCachedWidth =1130
                                    LayoutCachedHeight =5055
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =780
                                    Top =5250
                                    Width =335
                                    Height =285
                                    TabIndex =3
                                    Name ="chkTransectChecked_120"
                                    ControlSource ="CWD_Check_120"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"

                                    LayoutCachedLeft =780
                                    LayoutCachedTop =5250
                                    LayoutCachedWidth =1115
                                    LayoutCachedHeight =5535
                                End
                                Begin CheckBox
                                    OverlapFlags =255
                                    Left =780
                                    Top =5730
                                    Width =335
                                    Height =285
                                    TabIndex =4
                                    Name ="chkTransectChecked_240"
                                    ControlSource ="CWD_Check_240"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"

                                    LayoutCachedLeft =780
                                    LayoutCachedTop =5730
                                    LayoutCachedWidth =1115
                                    LayoutCachedHeight =6015
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =360
                                    Top =4455
                                    Width =1679
                                    Height =1650
                                    Name ="shpTransect_Checked"
                                    LayoutCachedLeft =360
                                    LayoutCachedTop =4455
                                    LayoutCachedWidth =2039
                                    LayoutCachedHeight =6105
                                End
                                Begin Label
                                    BackStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =420
                                    Top =4290
                                    Width =1515
                                    Height =240
                                    FontSize =10
                                    BackColor =15527148
                                    Name ="lblTransectChecked"
                                    Caption ="Transect Checked"
                                    FontName ="Calibri"
                                    LayoutCachedLeft =420
                                    LayoutCachedTop =4290
                                    LayoutCachedWidth =1935
                                    LayoutCachedHeight =4530
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =2520
                                    Top =2160
                                    Width =10065
                                    Height =6435
                                    TabIndex =5
                                    Name ="fsub_Transects"
                                    SourceObject ="Form.fsub_Transects"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =2520
                                    LayoutCachedTop =2160
                                    LayoutCachedWidth =12585
                                    LayoutCachedHeight =8595
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1040
                                    Top =4650
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
                                    LayoutCachedTop =4650
                                    LayoutCachedWidth =1745
                                    LayoutCachedHeight =5025
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
                                    Top =5130
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
                                    LayoutCachedTop =5130
                                    LayoutCachedWidth =1730
                                    LayoutCachedHeight =5505
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
                                    Top =5610
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
                                    LayoutCachedTop =5610
                                    LayoutCachedWidth =1710
                                    LayoutCachedHeight =5985
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
                            Top =1635
                            Width =14055
                            Height =7695
                            Name ="pagTrees"
                            Caption ="Trees"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1635
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9330
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
                                    Top =1979
                                    Width =14054
                                    Height =7094
                                    Name ="fsub_Tree_Data"
                                    SourceObject ="Form.fsub_Tree_Data"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =1979
                                    LayoutCachedWidth =14114
                                    LayoutCachedHeight =9073
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =1635
                            Width =14055
                            Height =7695
                            Name ="pagSaplings"
                            Caption ="Saplings"
                            LayoutCachedLeft =60
                            LayoutCachedTop =1635
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9330
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
                                    Top =1979
                                    Width =14054
                                    Height =6599
                                    Name ="fsub_Sapling_Data"
                                    SourceObject ="Form.fsub_Sapling_Data"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"
                                    OnEnter ="[Event Procedure]"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =1979
                                    LayoutCachedWidth =14114
                                    LayoutCachedHeight =8578
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =60
                            Top =1635
                            Width =14055
                            Height =7695
                            Name ="pagQuadrats"
                            Caption ="Quadrats"
                            ImageData = Begin
                                0x00000000
                            End
                            LayoutCachedLeft =60
                            LayoutCachedTop =1635
                            LayoutCachedWidth =14115
                            LayoutCachedHeight =9330
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
                                    Top =1979
                                    Width =14054
                                    Height =6599
                                    Name ="fsub_Quadrats"
                                    SourceObject ="Form.fsub_Quadrats"
                                    LinkChildFields ="Event_ID"
                                    LinkMasterFields ="Event_ID"

                                    LayoutCachedLeft =60
                                    LayoutCachedTop =1979
                                    LayoutCachedWidth =14114
                                    LayoutCachedHeight =8578
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =255
                    Left =12300
                    Top =120
                    Width =900
                    Height =660
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdEditLocation"
                    Caption ="Edit Location"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Edit the current location record."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =12300
                    LayoutCachedTop =120
                    LayoutCachedWidth =13200
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =255
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
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =3000
                    Top =480
                    Width =3540
                    TabIndex =4
                    Name ="txtEvent_ID"
                    ControlSource ="Event_ID"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =480
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =720
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =247
                    Left =7080
                    Top =780
                    Width =1980
                    Height =300
                    FontSize =10
                    ForeColor =1279872587
                    Name ="lblLink_to_Google_Maps"
                    Caption ="Show on Google Maps"
                    FontName ="Calibri"
                    HyperlinkAddress ="http://maps.google.com/maps?q=ANTI-0092@39.4746557,-77.7262205&iwloc=A&t=h"
                    LayoutCachedLeft =7080
                    LayoutCachedTop =780
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =1080
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =247
                    Left =9180
                    Top =780
                    Width =1800
                    Height =300
                    FontSize =10
                    ForeColor =1279872587
                    Name ="lblLink_To_Plot_Photos"
                    Caption ="Explore Plot Photos"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =9180
                    LayoutCachedTop =780
                    LayoutCachedWidth =10980
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =13320
                    Top =120
                    Width =900
                    Height =660
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdTriggerReport"
                    Caption ="Event Report"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Display a Summary Report of this Event."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13320
                    LayoutCachedTop =120
                    LayoutCachedWidth =14220
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =255
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
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2460
                    Top =60
                    Width =483
                    Height =360
                    FontSize =18
                    FontWeight =700
                    TabIndex =6
                    Name ="tbxPseudoEvent"
                    ControlSource ="PseudoEvent"
                    StatusBarText ="Unique identifier for each sample location"
                    FontName ="Calibri"

                    LayoutCachedLeft =2460
                    LayoutCachedTop =60
                    LayoutCachedWidth =2943
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =255
                    Left =11280
                    Top =120
                    Width =900
                    Height =660
                    FontWeight =700
                    TabIndex =7
                    Name ="cmdPlot_Chart"
                    Caption ="Plot Chart"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Display the current location plot chart."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =11280
                    LayoutCachedTop =120
                    LayoutCachedWidth =12180
                    LayoutCachedHeight =780
                    ForeThemeColorIndex =0
                    UseTheme =255
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
                    Top =420
                    Width =4560
                    Height =300
                    FontSize =12
                    TabIndex =8
                    Name ="txtSlope_Aspect"
                    FontName ="Calibri"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =420
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =720
                End
                Begin CommandButton
                    OverlapFlags =255
                    Left =60
                    Top =480
                    Width =2766
                    FontWeight =600
                    TabIndex =9
                    Name ="btnEditEventDate"
                    Caption ="  Change Event Date"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddd00d00d00dddddddd00d00d00dddddddd ,
                        0xdddddddddddddddddddddddddddddddddddddddddddddd00ddddd7ddddddd000 ,
                        0xd7dd7c7ddddd000d7c7dd7ddddd000ddd7dddddddd000dddddddddddd000dddd ,
                        0xddd7dddd000ddddddd7c7dd000ddddddddd7dd0b0dddddddddddd0b0dddddddd ,
                        0xddddd70ddddddddd000000000000000000000000000000000000000000000000 ,
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
                    ControlTipText ="Change event date"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =60
                    LayoutCachedTop =480
                    LayoutCachedWidth =2826
                    LayoutCachedHeight =840
                    PictureCaptionArrangement =5
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =1796857
                    BackThemeColorIndex =5
                    BorderColor =1796857
                    BorderThemeColorIndex =5
                    HoverColor =65280
                    HoverTint =80.0
                    PressedColor =413911
                    PressedThemeColorIndex =5
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeThemeColorIndex =0
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    Shadow =-1
                    QuickStyle =24
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2760
                    Top =660
                    Width =2160
                    Height =420
                    FontSize =18
                    FontWeight =700
                    TabIndex =10
                    Name ="tbxEventDate"
                    ControlSource ="Event_Date"
                    Format ="Short Date"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    FontName ="Calibri"

                    LayoutCachedLeft =2760
                    LayoutCachedTop =660
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =1080
                End
                Begin Image
                    Visible = NotDefault
                    PictureType =2
                    Left =5520
                    Top =1320
                    Width =540
                    Height =540
                    Name ="imgFlag"
                    OnClick ="[Event Procedure]"
                    Picture ="flag_red"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =1860
                    TabIndex =14
                End
                Begin Image
                    Visible = NotDefault
                    PictureType =2
                    Left =6120
                    Top =1320
                    Width =540
                    Height =540
                    Name ="Image165"
                    OnClick ="[Event Procedure]"
                    Picture ="flag_yellow"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =1860
                    TabIndex =15
                End
                Begin Image
                    Visible = NotDefault
                    PictureType =2
                    Left =6720
                    Top =1320
                    Width =540
                    Height =540
                    Name ="Image166"
                    OnClick ="[Event Procedure]"
                    Picture ="flag_blue"

                    LayoutCachedLeft =6720
                    LayoutCachedTop =1320
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =1860
                    TabIndex =16
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =15527148
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
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
' FORM:         frm_Events
' Level:        Form module
' Version:      1.06
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 24, 2018
' Revisions:    ML/GS - unknown   - 1.00 - initial version
'               BLC   - 5/24/2018 - 1.01 - added documentation, error handling
'               BLC   - 11/9/2018 - 1.02 - added pseudoevent functionality
'               BLC   - 4/17/2018 - 1.03 - updated convert pseudoevent to regular event
'               BLC   - 5/3/2019  - 1.04 - set plot & event ID temp vars
'               BLC - 4/2/2020    - 1.05 - fit report to window after opening vs. default smaller view
'               BLC - 6/22/2020   - 1.06 - add QC mode for flagging
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 4/24/2018 -
' ---------------------------------
Private Sub xForm_Open(Cancel As Integer)
On Error GoTo Err_Handler
   
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' FORM NAME:    frm_Data_Entry
' Description:  Primary field data entry form
' Version:      1.03
' Data source:  tbl_Locations
' Data access:  edit; allow additions off except for new records
' Pages:        none
' Functions:    Update_Loc_Info, ValidateForm
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Source/date:  John R. Boetsch, June 2006
' Revisions:    Simon Kingston, October - January 2006
'                   - extensive updates, adding GUID generation code, new controls
'               BLC - 4/2/2019 - 1.03 - added psuedoevent handling
' =================================

' ---------------------------------
' SUB:          Form_Open
' Description:  form open actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
'   BLC - 4/2/2019 - added pseudoevent handling
'   BLC - 5/3/2019 - set plot & event ID temp vars
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim strCaptionSuffix As String
    Dim booEditOn As Boolean
'SetAppIcon "ncrn_forestveg.ico"

    'minimize utilities
    ToggleForm "frm_Data_Gateway", -1
    
Debug.Print "QCmode = " & TempVars("QC_MODE")
    'QC mode?
    lblQCMode.visible = Nz(TempVars("QC_MODE"), False)
    btnFlag.visible = Nz(TempVars("QC_MODE"), False)
    imgFlag.visible = Nz(TempVars("QC_MODE"), False)

'    ' Set the opening parameters depending on the arguments passed from the previous form
'    If Me.OpenArgs = "(Browsing)" Then
'        strCaptionSuffix = " - " & Me.OpenArgs
'        booEditOn = False
'    ElseIf Me.OpenArgs = "(Creating)" Then
'        strCaptionSuffix = " - " & Me.OpenArgs
'        booEditOn = True
'    ElseIf Me.OpenArgs <> "" Then
'        strCaptionSuffix = " - " & "No arguments"
'        booEditOn = False
'    End If
    
    'split out Caption, EventID
    Dim ary() As String
    ary = IIf(Len(Me.OpenArgs) > 0, Split(Me.OpenArgs, ","), "")
    
    'default
    
    'cleanup caption (prevents (Browsing)-(Browsing)-(Browsing)...)
    Me.Caption = IIf(CountInString(Me.Caption, "-") > 1, _
                      Replace(Replace(Me.Caption, ary(0), ""), "-", ""), _
                      Me.Caption)
                      
    booEditOn = False
    strCaptionSuffix = " - "
    
    Select Case ary(0)
        Case "(Browsing)"
            Add2Self strCaptionSuffix, ary(0)
        Case "(Creating)"
            booEditOn = True
            Add2Self strCaptionSuffix, ary(0)
        Case ""
            Add2Self strCaptionSuffix, "No arguments"
        Case Else
    End Select
    
    'update form title
    Me.Caption = Add2Self(Trim(Me.Caption), strCaptionSuffix)
    
    'TO DO
    'Insert code here to update Plot Status in the Location table if this is the first sampling of this plot.
    
    SetEditMode (booEditOn)

    'check for PseudoEvents
    SetTempVar "IsPseudoEvent", Nz(Me.tbxPseudoEvent.Value, 0) 'tbxPseudoEvent.Value
    Dim bgdColor As Long, txtColor As Long

    'defaults
    txtColor = lngWhite
    bgdColor = HTMLConvert("#ECECEC")
    btnConvertPseudoEvent.hoverColor = lngGreen
    btnConvertPseudoEvent.visible = False
    lblPseudoEventFlag.visible = False
    rctPseudoEvent.visible = False

    If TempVars("IsPseudoEvent") = 1 Then
        'bgdColor = lngLtPink
        txtColor = lngLtPink
        lblPseudoEventFlag.forecolor = txtColor
        lblPseudoEventFlag.visible = True
        rctPseudoEvent.backcolor = txtColor
        rctPseudoEvent.visible = True

        'expose conversion button ONLY in edit mode
        If Not Right(Me.tglBrowse_Edit.Caption, 4) = "EDIT" Then
            btnConvertPseudoEvent.visible = True
        End If
    End If
    
    'Me.Detail.BackColor = bgdColor
    'lblEvent_Form_Header.ForeColor = txtColor
    'lblPseudoEventFlag.ForeColor = lngBlack
    'lblPseudoEventFlag.BackColor = txtColor
    
    'set globals
    SetTempVar "plot", Me.Location_ID.Value
    SetTempVar "eventID", Me.Event_ID.Value

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Events])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub Form_Current() 'Cancel As Integer)
On Error GoTo Err_Handler
   
    'Update fields in header from Locations table
    Update_Loc_Info
    'Enable edit location function if there is an active location
    Me!cmdEditLocation.Enabled = Not IsNull(Me!txtLocation_ID)
    'Event groups not implemented in this database
    'Me!cmdEditEventGroup.Enabled = Not IsNull(Me!cboEvent_Group_ID)
   
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_BeforeInsert
' Description:  form before insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
On Error GoTo Err_Handler

    ' Create the GUID primary key value if needed for a string GUID
    If IsNull(Me!Event_ID) Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me!Event_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnConvertPseudoEvent_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, February, 2019
' Adapted:      -
' Revisions:
'   BLC - 2/20/2019 - initial version
'   BLC - 4/17/2019 - add conversion code
' ---------------------------------
Private Sub btnConvertPseudoEvent_Click()
On Error GoTo Err_Handler

    Dim RetVal As Boolean
    
    RetVal = MsgBox("Click OK to confirm you want to convert this event to a regular event." _
                    & vbCrLf & vbCrLf & "NOTE:" & vbCrLf & vbCrLf _
                    & "You cannot revert back to a pseudoevent, so be sure you want to do this!", _
                     vbOKCancel, "Confirm Convert from PseudoEvent to Regular Event")
    
    'convert if desired
    If RetVal = True Then
        lblPseudoEventFlag.visible = False
        btnConvertPseudoEvent.Enabled = False
        btnConvertPseudoEvent.visible = False
        rctPseudoEvent.visible = False

        'convert to regular event
        Me.PseudoEvent = 0

    Else
        lblPseudoEventFlag.visible = True
        btnConvertPseudoEvent.Enabled = True
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnConvertPseudoEvent_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnEditEventDate_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 4/2/2020 - initial version, adapted from Location edit
' ---------------------------------
Private Sub btnEditEventDate_Click()
On Error GoTo Err_Handler

    DisplayMessage "notready"
    GoTo Exit_Handler

    Dim strOpenArgs As String
    Dim strCriteria As String

    If Not IsNothing(Me.txtStart_date) Then
        strOpenArgs = XML_Tag("FormFrom", Me.Name)
        Add2Self strOpenArgs, XML_Tag("ControlFrom", "txtStart_date")
        Add2Self strOpenArgs, XML_Tag("ControlValue", Me.txtStart_date)
        Add2Self strOpenArgs, XML_Tag("EditTable", "tbl_Events")
        Add2Self strOpenArgs, XML_Tag("EditID", Me.txtEvent_ID)
        Add2Self strOpenArgs, XML_Tag("EditField", "Event_Date")
        Add2Self strOpenArgs, XML_Tag("EditIDField", "Event_ID")
        If IsNull(TempVars("UserID")) Then DoCmd.OpenForm "frm_Select_User", acNormal, , , acFormEdit, acWindowNormal, "frm_Pad_Date"
        Add2Self strOpenArgs, XML_Tag("UpdateByID", TempVars("UserID"))
        'strCriteria = GetCriteriaString("Event_Date=", "tbl_Events", "Event_Date", Me.Name, "txtStart_date")
        strCriteria = ""
        DoCmd.OpenForm "frm_Pad_Date", , , strCriteria, acFormEdit, acWindowNormal, strOpenArgs
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEditEventDate_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkPicturesTaken_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub chkPictures_Taken_AfterUpdate()
On Error GoTo Err_Handler
    
    lblPictures_Taken.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkPicturesTaken_AfterUpdate[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkTransect120Check_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub chkTransectChecked_120_AfterUpdate()
On Error GoTo Err_Handler

    lblTransectChecked_120.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkTransect120Check_AfterUpdate[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkTransect240Check_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub chkTransectChecked_240_AfterUpdate()
On Error GoTo Err_Handler

    lblTransectChecked_240.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkTransect240Check_AfterUpdate[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          chkTransect360Check_AfterUpdate
' Description:  checkbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub chkTransectChecked_360_AfterUpdate()
On Error GoTo Err_Handler

    lblTransectChecked_360.Requery

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - chkTransect360Check_AfterUpdate[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenDeerImpactForm_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cmdOpen_Form_Deer_Impact_Click()
On Error GoTo Err_Handler
    
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Popup_Deer_Impact"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenDeerImpactForm_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPlotChart_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cmdPlot_Chart_Click()
On Error GoTo Err_Handler
    
    Dim strOpenArgs As String
    Dim strCriteria As String
    If Not IsNothing(Me!txtLocation_ID) Then
        strOpenArgs = XML_Tag("FormFrom", Me.Name)
        strOpenArgs = strOpenArgs & XML_Tag("ControlFrom", "txtLocation_ID")
        strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Plot_Chart", , , strCriteria, acFormEdit, acWindowNormal, strOpenArgs
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnPlotChart_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnReport_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
'   BLC - 4/2/2020  - fit report to window after opening vs. default smaller view
' ---------------------------------
Private Sub cmdTriggerReport_Click()
On Error GoTo Err_Handler
    Dim strDocName As String
    Dim strCriteria As String
    
    '10/23/2018 BLC
    'set TempVar for qry_Status_Sapling_Current_Event/qry_Status_Tree_Current_Event
    SetTempVar "EventID", CStr(Me.txtEvent_ID)
    
    strDocName = "rpt_Event_Summary_Unfiltered"
    strCriteria = GetCriteriaString("[Event_ID]=", "tbl_Events", "Event_ID", Me.Name, "txtEvent_ID")
    DoCmd.OpenReport strDocName, acPreview, , strCriteria
    
       'set to full size
    DoCmd.Maximize
    DoCmd.RunCommand acCmdZoom100 '100%
    'DoCmd.RunCommand acCmdFitToWindow 'fit window size
    
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub

' ---------------------------------
' SUB:          ValidateForm
' Description:  form validation actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Simon Kingston, Dec. 2006
' Adapted:  Mark Lehman/Geoffrey Sanders, unknown
' Revisions:
'   SK      - 12/2006 - initial version
'   MEL/GS  - unknown - adapted version
'   BLC     - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub ValidateForm()
On Error GoTo Err_Handler

    If IsNull(Me!txtLocation_ID) Then
        MsgBox "You must select a location before you can enter record details!", vbExclamation, "Enter Location First"
        'Me!cboLocation_ID.SetFocus
    Else
        If IsNull(Me!txtStart_date) Then
            MsgBox "You must enter a start date before you can enter record details!", vbExclamation, "Enter Start Date"
            'Me!txtStart_date.SetFocus
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ValidateForm[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          grpTransectSelection_AfterUpdate
' Description:  group after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub grpTransect_Selection_AfterUpdate()
On Error GoTo Err_Handler
    
    Dim strTransect As String
    
    strTransect = Me!grpTransect_Selection.Value
    Me.txtTransect_Selection.Value = "'" & strTransect & "'"
    Forms![frm_Events]![fsub_Transects]!txtTransect_Azimuth.DefaultValue = "'" & strTransect & "'"
    Forms![frm_Events]![fsub_Transects].Form.Filter = "[Transect_Azimuth] = """ & strTransect & """ "
    Forms![frm_Events]![fsub_Transects].Form.FilterOn = True

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - grpTransectSelection_AfterUpdate[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lblPhotosLink_Click
' Description:  label click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
'   BLC - 12/10/2018 - revised error handling
'   BLC - 4/5/2020 - revised handling of root & photo folder locations
' ---------------------------------
Private Sub lblLink_To_Plot_Photos_Click()
On Error GoTo Err_Handler

    Dim RetVal As Double
    Dim RootFolder As String
    Dim PhotoFolder As String
    
    RootFolder = TempVars("Root") '"T:\I&M"
    PhotoFolder = TempVars("Photo") '"T:\I&M\Monitoring\Forest_Vegetation\Photos\"
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
            "Error encountered (#" & Err.Number & " - lblPhotosLink_Click[frm_Events])"
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' SUB:          subObservers_Enter
' Description:  subform enter actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub subObservers_Enter()
On Error GoTo Err_Handler

    ValidateForm

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - subObservers_Enter[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxEventGroupID_AfterUpdate
' Description:  combobox after udpate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cboEvent_Group_ID_AfterUpdate()
On Error GoTo Err_Handler

    Me!cmdEditEventGroup.Enabled = Not IsNothing(Me!cboEvent_Group_ID)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxEventGroupID_AfterUpdate[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnEditLocation_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cmdEditLocation_Click()
On Error GoTo Err_Handler

    Dim strOpenArgs As String
    Dim strCriteria As String
    
    If Not IsNothing(Me!txtLocation_ID) Then
        strOpenArgs = XML_Tag("FormFrom", Me.Name)
        strOpenArgs = strOpenArgs & XML_Tag("ControlFrom", "txtLocation_ID")
        strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Locations", , , strCriteria, acFormEdit, acWindowNormal, strOpenArgs
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEditLocation_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnNewUser_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cmdNewUser_Click()
On Error GoTo Err_Handler

    DoCmd.OpenForm "frm_Contacts", , , , acFormAdd, , "new"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnNewUser_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          UpdateLocInfo
' Description:  Updates associates location information when Location_ID is updated
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Public Sub Update_Loc_Info()
' Description:  Updates associated location information when Location_ID is updated
' References:   GetCriteriaString
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:    <name, date, desc - add lines as you go>
On Error GoTo Err_Handler

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

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateLocInfo[frm_Events])"
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
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cmdClose_Click()
On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Close
' Description:  form close actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
'   BLC - 6/22/20202 - restore data gateway
' ---------------------------------
Private Sub Form_Close()
On Error GoTo Err_Handler

    'restore data gateway
    ToggleForm "frm_Data_Gateway", 0
    
    If IsLoaded("frm_Data_Gateway") Then
        Forms("frm_Data_Gateway").Requery
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Close[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglBrowseEdit_Click
' Description:  toggle click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub tglBrowse_Edit_Click()
On Error GoTo Err_Handler

    'Call the SetEditMode subroutine with the current status of the Browse/Edit toggle
    Me.SetEditMode (Me!tglBrowse_Edit)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglBrowseEdit_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetEditMode
' Description:  sets form edit mode
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
'   BLC - 12/10/2018 - revised error handling
'   BLC - 4/17/2019 - added PseudoEvent effects (hide conversion button)
' ---------------------------------
Public Sub SetEditMode(booEditOn As Boolean)
' Description:  Toggles the form between browse and edit mode
' Parameters:   booFilterOn = true if edit mode, false if browse mode
' Returns:      none
' Throws:       none
' References:   strFieldName, strSortOrder, strSortFieldLabel
'               (form-level variables)
' Source/date:  Simon Kingston, 1/17/2007
' Revisions:    Mark Lehman 3/15/2010 Repurposed version of FilterGateway by Kingston

On Error GoTo Err_Handler

    Me!tglBrowse_Edit = booEditOn
    
    If booEditOn Then
        Me!tglBrowse_Edit.Caption = "Editing ON"
        Me!lblEvent_Form_Header.backcolor = RGB(128, 0, 0)
        Me.btnConvertPseudoEvent.visible = False
    Else
        Me!tglBrowse_Edit.Caption = "Editing OFF"
        Me!lblEvent_Form_Header.backcolor = vbBlack
        Me.btnConvertPseudoEvent.visible = True
    End If
    
    'Me.FilterOn = booEditOn

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetEditMode[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnAddEditEventNote_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cmdAdd_Edit_Event_Note_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Event_Add_Note"
    
    stLinkCriteria = "[Event_ID]=" & "'" & Me![txtEvent_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddEditEventNote_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnAddEventNote_Click
' Description:   click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 11/9/2018 - add documentation, error handling
' ---------------------------------
Private Sub cmdAdd_Event_Note_Click()
On Error GoTo Err_Handler

    Me.Requery
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Event_Add_Note"
    
    stLinkCriteria = "[Event_ID]=" & "'" & Me![txtEvent_ID] & "'"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddEventNote[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          imgFlag_Click
' Description:  image click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2020
' Adapted:      -
' Revisions:
'   BLC - 6/22/2020 - initial version
' ---------------------------------
Private Sub imgFlag_Click()
On Error GoTo Err_Handler
    
    Dim strOpenArgs As String
    
    strOpenArgs = Me.Name & "|" & Me.RecordSource & "|" & "Event" & "|" & Me.txtEvent_ID & "|" & "|" & Nz(TempVars("UserID"), "") & "|"
    
    DoCmd.OpenForm "SetFlag", , , , acFormEdit, acWindowNormal, strOpenArgs

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - imgFlag_Click[frm_Events])"
    End Select
    Resume Exit_Handler
End Sub
