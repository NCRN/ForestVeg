Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7380
    DatasheetFontHeight =11
    ItemSuffix =33
    Left =495
    Top =270
    Right =7620
    Bottom =7260
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x2e1f8472d703e440
    End
    Caption ="Utilities"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-16777216
            FontName ="Calibri"
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackColor =-8488071
            BorderLineStyle =0
            BorderColor =-8488071
            ThemeFontIndex =1
            HoverColor =-8488071
            HoverTint =80.0
            PressedColor =-8488071
            PressedShade =80.0
            HoverForeColor =-16777216
            PressedForeColor =-16777216
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BoundObjectFrame
            AddColon = NotDefault
            SizeMode =3
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =-1800
            BackThemeColorIndex =1
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
                    TextAlign =1
                    Width =7320
                    Height =480
                    FontSize =18
                    FontWeight =700
                    LeftMargin =58
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblUtilities_Header"
                    Caption ="Utilities, Configuration && Admin Tools"
                    GridlineColor =10921638
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =480
                    ThemeFontIndex =-1
                    BackThemeColorIndex =0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =6240
                    Top =60
                    Width =1020
                    FontSize =14
                    FontWeight =700
                    ForeColor =0
                    Name ="btnClose"
                    Caption ="Close"
                    FontName ="Franklin Gothic Book"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =6240
                    LayoutCachedTop =60
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =420
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
            End
        End
        Begin Section
            Height =6720
            BackColor =15921906
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =3420
                    Top =2100
                    Width =3600
                    Height =4500
                    TabIndex =14
                    BackColor =13166064
                    BorderColor =10921638
                    Name ="frmAdmin"
                    GridlineColor =10921638

                    LayoutCachedLeft =3420
                    LayoutCachedTop =2100
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =6600
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =223
                            TextAlign =2
                            Left =3829
                            Top =1980
                            Width =735
                            Height =345
                            FontWeight =600
                            BorderColor =8355711
                            Name ="lblAdmin"
                            Caption ="Admin"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =3829
                            LayoutCachedTop =1980
                            LayoutCachedWidth =4564
                            LayoutCachedHeight =2325
                            BackTint =95.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =360
                    Top =2100
                    Width =3000
                    Height =3180
                    TabIndex =13
                    BackColor =16765439
                    BorderColor =10921638
                    Name ="frmQC"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =2100
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =5280
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =829
                            Top =1980
                            Width =375
                            Height =345
                            FontWeight =600
                            BorderColor =8355711
                            ForeColor =16711935
                            Name ="lblQC"
                            Caption ="QC"
                            FontName ="Franklin Gothic Book"
                            ControlTipText ="Quality Control - post data collection quality processes"
                            GridlineColor =10921638
                            LayoutCachedLeft =829
                            LayoutCachedTop =1980
                            LayoutCachedWidth =1204
                            LayoutCachedHeight =2325
                            BackTint =95.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =360
                    Top =180
                    Width =6660
                    Height =1740
                    TabIndex =11
                    BackColor =15924699
                    BorderColor =10921638
                    Name ="frmUtilities"
                    GridlineColor =10921638

                    LayoutCachedLeft =360
                    LayoutCachedTop =180
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =1920
                    BackThemeColorIndex =-1
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =2
                            Left =480
                            Top =60
                            Width =870
                            Height =315
                            FontWeight =600
                            BorderColor =8355711
                            ForeColor =16711680
                            Name ="lblUtilities"
                            Caption ="Utilities"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =480
                            LayoutCachedTop =60
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =375
                            BackTint =95.0
                            ForeThemeColorIndex =-1
                            ForeTint =100.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =223
                    Left =1960
                    Top =2400
                    Width =1260
                    Height =720
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="btnDataQC"
                    Caption ="QA"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the QC Summary Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Data_QA"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =1960
                    LayoutCachedTop =2400
                    LayoutCachedWidth =3220
                    LayoutCachedHeight =3120
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    OverlapFlags =215
                    Left =3760
                    Top =480
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="btnAppend"
                    Caption =" Append Data"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the Append Data Switchboard ti Import Field Data"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Append_Select_Import_File"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"btnAppend\" Event=\"OnClick\" xmlns=\"http://schemas.microso"
                                "ft.com/office/accessservices/2009/11/application\"><Statements><Action Name=\"Op"
                                "enForm\"><Argument Name=\"FormName"
                        End
                        Begin
                            Comment ="_AXL:\">frm_Append_Select_Import_File</Argument></Action></Statements></UserInte"
                                "rfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =3760
                    LayoutCachedTop =480
                    LayoutCachedWidth =5020
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    OverlapFlags =215
                    Left =2180
                    Top =480
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    ForeColor =0
                    Name ="btnBackupBE"
                    Caption ="Create Backup"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Create a Backup of the Backend Database"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =2180
                    LayoutCachedTop =480
                    LayoutCachedWidth =3440
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    OverlapFlags =215
                    Left =5340
                    Top =480
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    ForeColor =0
                    Name ="btnLookups"
                    Caption ="Lookups"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the QA/QC Summary Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Lookups"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"cmdLookups\" Event=\"OnClick\" xmlns=\"http://schemas.micros"
                                "oft.com/office/accessservices/2009/11/application\" xmlns:a=\"http://schemas.mic"
                                "rosoft.com/office/accessservices"
                        End
                        Begin
                            Comment ="_AXL:/2009/11/forms\"><Statements><Action Name=\"OpenForm\"><Argument Name=\"For"
                                "mName\">frm_Lookups</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =5340
                    LayoutCachedTop =480
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =79
                    Left =1980
                    Top =3240
                    Width =1260
                    Height =864
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    ForeColor =0
                    Name ="btnPostSeasonChecks"
                    Caption ="P&ost-Season Checks"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open a form to check that RIO tags are actually in the office"
                    UnicodeAccessKey =111
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1980
                    LayoutCachedTop =3240
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =4104
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    TabStop = NotDefault
                    OverlapFlags =223
                    Left =540
                    Top =3240
                    Width =1260
                    Height =864
                    FontSize =12
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    Name ="btnFlagReports"
                    Caption ="R Flag Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Triggers R QA/QC reports (R/R Studio installs required)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =540
                    LayoutCachedTop =3240
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =4104
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =540
                    Top =2400
                    Width =1260
                    Height =720
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    Name ="btnQCReports"
                    Caption ="QC Reports"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open a form to select which QA/QC Report to run (Access)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =540
                    LayoutCachedTop =2400
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =3120
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    OverlapFlags =215
                    Left =600
                    Top =480
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =12
                    ForeColor =0
                    Name ="btnRelinkTables"
                    Caption =" Relink Tables"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Reset the link to the backend database"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Connect_Tables"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =600
                    LayoutCachedTop =480
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =1740
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    TabStop = NotDefault
                    OverlapFlags =223
                    Left =4260
                    Top =5220
                    Width =1800
                    Height =405
                    TabIndex =15
                    ForeColor =4210752
                    Name ="btn3"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =4260
                    LayoutCachedTop =5220
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =5625
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                    GridlineThemeColorIndex =1
                    BackColor =11710639
                    BackThemeColorIndex =4
                    BackTint =60.0
                    BorderColor =11710639
                    BorderThemeColorIndex =4
                    BorderTint =60.0
                    HoverColor =13355721
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    PressedColor =6249563
                    PressedThemeColorIndex =4
                    PressedShade =75.0
                    HoverForeColor =4210752
                    HoverForeThemeColorIndex =0
                    HoverForeTint =75.0
                    PressedForeColor =4210752
                    PressedForeThemeColorIndex =0
                    PressedForeTint =75.0
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =223
                    Left =3540
                    Top =4740
                    Width =3360
                    Height =1680
                    BackColor =4342595
                    BorderColor =10921638
                    Name ="rctAdminDb"
                    GridlineColor =10921638
                    LayoutCachedLeft =3540
                    LayoutCachedTop =4740
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =6420
                    BackThemeColorIndex =2
                End
                Begin Label
                    OverlapFlags =223
                    TextAlign =2
                    Left =3600
                    Top =4860
                    Width =1860
                    Height =255
                    FontSize =9
                    FontWeight =700
                    BackColor =11056034
                    ForeColor =16777215
                    Name ="lblMgmtTools"
                    Caption ="Management Tools"
                    FontName ="Arial"
                    LayoutCachedLeft =3600
                    LayoutCachedTop =4860
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =5115
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =5400
                    Top =4980
                    Name ="lnUI"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =4980
                    LayoutCachedWidth =6840
                    LayoutCachedHeight =4980
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =82
                    Left =3780
                    Top =5280
                    Width =1260
                    Height =720
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="btnPreSeasonPrep"
                    Caption ="P&re-Season Prep"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Create BE backup and purge annual field data from tables"
                    UnicodeAccessKey =114
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3780
                    LayoutCachedTop =5280
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =6000
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    Shape =2
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5340
                    Top =5280
                    Width =1260
                    Height =720
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    ForeColor =0
                    Name ="btnDPLs"
                    Caption ="Set DPLs"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Set DPLs for all events for a given year. Single event DPLs must be adjusted via"
                        " the event."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5340
                    LayoutCachedTop =5280
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =6000
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    Shape =2
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                Begin Rectangle
                    BackStyle =1
                    OverlapFlags =223
                    Left =3540
                    Top =2340
                    Width =3360
                    Height =2340
                    BackColor =6108695
                    Name ="rctProtocolTools"
                    GridlineColor =10921638
                    LayoutCachedLeft =3540
                    LayoutCachedTop =2340
                    LayoutCachedWidth =6900
                    LayoutCachedHeight =4680
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3600
                    Top =2460
                    Width =1335
                    Height =270
                    FontSize =9
                    FontWeight =700
                    BackColor =11056034
                    ForeColor =16777215
                    Name ="lblProtocolTools"
                    Caption ="Protocol Tools"
                    FontName ="Arial"
                    LayoutCachedLeft =3600
                    LayoutCachedTop =2460
                    LayoutCachedWidth =4935
                    LayoutCachedHeight =2730
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =223
                    Left =3780
                    Top =2820
                    Width =1260
                    Height =720
                    FontSize =12
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    Name ="btnSOPVersions"
                    Caption ="SOPs"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Opens SOP versioning tool"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3780
                    LayoutCachedTop =2820
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =3540
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =3780
                    Top =3660
                    Width =1260
                    Height =720
                    FontWeight =700
                    TabIndex =9
                    ForeColor =0
                    Name ="btnDataFlags"
                    Caption ="Data Flags"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="View and enter data flags"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3780
                    LayoutCachedTop =3660
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =4380
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    Enabled = NotDefault
                    OverlapFlags =223
                    Left =5340
                    Top =2820
                    Width =1260
                    Height =720
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    Name ="btnTaxaReferences"
                    Caption ="Taxonomic Refs"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =5340
                    LayoutCachedTop =2820
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =3540
                    ForeThemeColorIndex =0
                    GridlineThemeColorIndex =1
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =87
                    SpecialEffect =5
                    Left =5040
                    Top =2580
                    Name ="Line28"
                    LayoutCachedLeft =5040
                    LayoutCachedTop =2580
                    LayoutCachedWidth =6480
                    LayoutCachedHeight =2580
                    BorderThemeColorIndex =-1
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin CommandButton
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =5340
                    Top =3660
                    Width =1260
                    Height =720
                    FontWeight =700
                    TabIndex =17
                    ForeColor =0
                    Name ="btnTargetLists"
                    Caption ="Target Lists"
                    StatusBarText ="Shrubs, QuadratHerbaceous, Vines, etc."
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="View and edit species lists for drop-downs (target lists)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =5340
                    LayoutCachedTop =3660
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =4380
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =480
                    Top =5400
                    Width =1260
                    Height =720
                    FontWeight =700
                    TabIndex =18
                    ForeColor =0
                    Name ="btnTBD"
                    Caption ="TBD"
                    StatusBarText ="Shrubs, QuadratHerbaceous, Vines, etc."
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="View and edit species lists for drop-downs (target lists)"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =480
                    LayoutCachedTop =5400
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =6120
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
                    Enabled = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =540
                    Top =4260
                    Width =1260
                    Height =864
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    ForeColor =0
                    Name ="btnAddFlags"
                    Caption ="Add Data Flags"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open event browse form in QC mode w/ flagging functionality"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =540
                    LayoutCachedTop =4260
                    LayoutCachedWidth =1800
                    LayoutCachedHeight =5124
                    ForeThemeColorIndex =0
                    GridlineShade =100.0
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =65280
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedThemeColorIndex =4
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
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15921906
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
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
' MODULE:       frm_Utilities
' Level:        Application form module
' Version:      1.04
' Description:  Standard module - main form for various database functions
' Data source:  -
' Data access:  -
' Pages:        -
' Functions:    none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      Bonnie Campbell, August 16, 2019
' Revisions:
'               ML/GS - unknown   - 1.00 - initial version
'               BLC   - 8/16/2019 - 1.01 - documentation, error handling,
'                                          added Pre-Season Prep for purging db field data tables
'                                          renamed cmdXX to btnXX
'               BLC   - 9/26/2019 - 1.02 - added Post-Season Prep for RIO tag check
'               BLC   - 6/8/2020  - 1.03 - re-organize buttons, add R Flag Report, SOP versions buttons
'               BLC   - 6/22/2020 - 1.04 - add Add Data Flags button
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Methods
' ---------------------------------
' ---------------------------------
' SUB:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 16, 2019
' Adapted:      -
' Revisions:
'   BLC   - 8/16/2019 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'check for DbAdmin functionality (app level DB_ADMIN set in the db_Module)
    Me.btnPreSeasonPrep.visible = IIf(Nz(DB_ADMIN, False), True, False)

    Dim strCaption As String

    ' Set the application font to more closely match the forms.
    ' Useful in cases where the subforms use tables directly
    Application.SetOption "Default Font Name", "Arial"
    Application.SetOption "Default Font Size", 9

'    ' Set the table-driven caption of the switchboard
'    strCaption = Nz(DLookup("[Database_title]", "tsys_App_Releases", "[Release_ID] = '" _
'        & Me!Release_ID & "'"), "")
'    Me.Caption = strCaption

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_App_Releases) (#" & Err.Number & " - Form_Open[frm_Utilities])"
        Case 2001   ' Field name in DLookup improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_App_Releases) (#" & Err.Number & " - Form_Open[frm_Utilities])"
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_App_Releases) (#" & Err.Number & " - Form_Open[frm_Utilities])"
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - Form_Open[frm_Utilities])"
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' SUB:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 16, 2019
' Adapted:      -
' Revisions:
'   BLC   - 8/16/2019 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - Form_Load[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 16, 2019
' Adapted:      -
' Revisions:
'   BLC   - 8/16/2019 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - Form_Current[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------
'   Utilities
' ---------------
' ---------------------------------
' SUB:          btnBackupBE_Click
' Description:  make a backup database backend file
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      Bonnie Campbell, August 16, 2019
' Revisions:
'   ML/GS - unknown   - initial version
'   BLC   - 8/16/2019 - added documentation & error handling
' ---------------------------------
Private Sub btnBackupBE_Click()
On Error GoTo Err_Handler

    fxnMakeBackup

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnBackupBE_Click[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------
'   QA/QC
' ---------------

' ---------------------------------
' SUB:          btnFlagReports_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 8, 2020
' Revisions:
'   BLC   - 6/8/2020 - initial version
' ---------------------------------
Private Sub btnFlagReports_Click()
On Error GoTo Err_Handler

    'check if R & R Studio are installed on machine
    DisplayMessage "notready"
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnFlagReports_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPostSeasonChecks_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 26, 2019
' Revisions:
'   BLC   - 9/26/2019 - initial version
' ---------------------------------
Private Sub btnPostSeasonChecks_Click()
On Error GoTo Err_Handler

    DoCmd.OpenForm "RIOCheck", acNormal, , , acFormAdd, acDialog

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnPostSeasonChecks_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnAddFlags_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 22, 2020
' Revisions:
'   BLC   - 6/22/2020 - initial version
' ---------------------------------
Private Sub btnAddFlags_Click()
On Error GoTo Err_Handler

    'open data gateway in "QC mode"
    DoCmd.OpenForm "frm_Data_Gateway", acNormal, , , acFormAdd, acDialog, "QC_MODE"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnAddFlags_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------
'   Admin Tools
' ---------------
' ---------------------------------
' SUB:          btnSOPVersions_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 8, 2020
' Revisions:
'   BLC   - 6/8/2020 - initial version
' ---------------------------------
Private Sub btnSOPVersions_Click()
On Error GoTo Err_Handler

    DoCmd.OpenForm "SOPVersion", acNormal, , , acFormAdd, acDialog

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnPostSeasonChecks_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPreSeasonPrep_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 16, 2019
' Revisions:
'   BLC   - 7/16/2019 - initial version
' ---------------------------------
Private Sub btnPreSeasonPrep_Click()
On Error GoTo Err_Handler

    'copy BE db
    BackupDbBE
    
    'copy & purge tables
    PurgeAnnualData
    
' shift msg to PurgeAnnualData to display only when purging is selected
'    'update
'    MsgBox "Pre-season backup & annual data purge is complete." & vbCrLf _
'           & "Review APBU_* data tables before deleting them.", _
'           vbOKOnly + vbInformation, "Pre-Season Backup & Annual Db Prep Complete"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnPreSeasonPrep_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDataFlags_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 11, 2020
' Revisions:
'   BLC   - 6/11/2020 - initial version
' ---------------------------------
Private Sub btnDataFlags_Click()
On Error GoTo Err_Handler

    'open form for management of data flags
    DoCmd.OpenForm "DataFlags", acNormal, , , acFormAdd, acDialog

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnDataFlags_Click[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTaxaReferences_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 11, 2020
' Revisions:
'   BLC   - 6/11/2020 - initial version
' ---------------------------------
Private Sub btnTaxaReferences_Click()
On Error GoTo Err_Handler

    'open form for management of data flags
    DoCmd.OpenForm "TaxonomicReferences", acNormal, , , acFormAdd, acDialog

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnTaxaReferences_Click[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTargetLists_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 15, 2020
' Revisions:
'   BLC   - 6/15/2020 - initial version
' ---------------------------------
Private Sub btnTargetLists_Click()
On Error GoTo Err_Handler

    'open form for management of data flags
    DoCmd.OpenForm "TargetLists", acNormal, , , acFormAdd, acDialog

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnTargetLists_Click[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDPLs_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 15, 2020
' Revisions:
'   BLC   - 6/15/2020 - initial version
' ---------------------------------
Private Sub btnDPLs_Click()
On Error GoTo Err_Handler

    'open form for management of data flags
    DoCmd.OpenForm "SetDPL", acNormal, , , acFormAdd, acDialog

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnDPLs_Click[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub
Sub RunRScript()
'runs an external R code through Shell
'The location of the RScript is 'C:\R_code'
'The script name is 'hello.R'
Dim shell As Object
Set shell = VBA.CreateObject("WScript.Shell")
Dim waitTillComplete As Boolean: waitTillComplete = True
Dim style As Integer: style = 1
Dim errorCode As Integer
Dim var1, var2 As Double
var1 = 1.2
var2 = 3.4

Dim Path As String
Path = "RScript C:\R_code\hello.R " & var1 & " " & var2

errorCode = shell.Run(Path, style, waitTillComplete)

End Sub
Public Function TestSubTime()
Debug.Print "testsubtime"
End Function

Public Sub CopyServerFileToLocal(ServerFileFullPath As String, LocalPath As String)

Debug.Print "hello"
'    Dim fso As Object
'    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
'    Dim FileName As String
'FileName = CStr(fso.GetFileName(ServerFileFullPath))
'    If FileExists(ServerFileFullPath) Then
'        If Not FolderExists(LocalPath) Then CreateFolder (LocalPath)
'        'copy file to local directory
'        CopyFile ServerFileFullPath, LocalPath, FileName 'fso.GetFileName(ServerFileFullPath)
'        'Split(ServerFileFullPath,"/")
'    Else
'        MsgBox "Sorry, " & ServerFileFullPath & " does not exist or is unreachable." _
'        & vbCrLf & "Please contact your database administrator to resolve this issue.", _
'        vbExclamation, "Server File Missing or Not Available"
'    End If

End Sub

Public Sub CopyFile(Source As String, DestinationDir As String, NewFileName As String)

    Dim FSO As Object
    Set FSO = VBA.CreateObject("Scripting.FileSystemObject")
    
    Dim Destination As String
    
    Destination = DestinationDir & "\" & NewFileName
    
    Call FSO.CopyFile(Source, Destination)

End Sub
