Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =15005
    DatasheetFontHeight =11
    ItemSuffix =60
    Top =516
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xb8d0f89fb3abe440
    End
    RecordSource ="qry_Tgt_Species_List_Park_Summary"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6a01000068010000680100006d010000000000009d3a00001c02000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    RibbonName ="Export"
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            KeepTogether =1
            ControlSource ="Family"
        End
        Begin BreakLevel
            ControlSource ="Family"
        End
        Begin BreakLevel
            ControlSource ="utah_species"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =780
            BackColor =15849926
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    Left =60
                    Top =60
                    Width =6735
                    Height =525
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblReportHdr"
                    Caption ="INVASIVES SPECIES LIST"
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =585
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10260
                    Width =4680
                    Height =528
                    ColumnOrder =0
                    FontSize =20
                    BorderColor =10921638
                    ForeColor =8355711
                    Name ="tbxPark"
                    ControlSource ="=TempVars(\"Park\") & \" SUMMARY\""
                    StatusBarText ="Park for list"
                    GridlineColor =10921638

                    LayoutCachedLeft =10260
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =528
                    ForeTint =50.0
                End
            End
        End
        Begin PageHeader
            Height =1380
            Name ="PageHeaderSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Rectangle
                    OldBorderStyle =0
                    Width =15005
                    Height =480
                    BackColor =15849926
                    BorderColor =10921638
                    Name ="rectPageHdr"
                    GridlineColor =10921638
                    LayoutCachedWidth =15005
                    LayoutCachedHeight =480
                    BackThemeColorIndex =2
                    BackTint =20.0
                End
                Begin Line
                    BorderWidth =2
                    Top =1320
                    Width =15005
                    Name ="lnHeader"
                    GridlineColor =10921638
                    LayoutCachedTop =1320
                    LayoutCachedWidth =15005
                    LayoutCachedHeight =1320
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13740
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =7
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear8"
                    ControlSource ="=[MinYear]+7"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =13740
                    LayoutCachedTop =600
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =6
                    TabIndex =1
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear7"
                    ControlSource ="=[MinYear]+6"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =600
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12480
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =5
                    TabIndex =2
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear6"
                    ControlSource ="=[MinYear]+5"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =12480
                    LayoutCachedTop =600
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11820
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =4
                    TabIndex =3
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear5"
                    ControlSource ="=[MinYear]+4"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =11820
                    LayoutCachedTop =600
                    LayoutCachedWidth =12120
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11208
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =3
                    TabIndex =4
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear4"
                    ControlSource ="=[MinYear]+3"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =11208
                    LayoutCachedTop =600
                    LayoutCachedWidth =11508
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =2
                    TabIndex =5
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear3"
                    ControlSource ="=[MinYear]+2"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =10620
                    LayoutCachedTop =600
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =600
                    Width =300
                    Height =660
                    ColumnOrder =1
                    TabIndex =6
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear2"
                    ControlSource ="=[MinYear]+1"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =600
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9420
                    Top =600
                    Width =299
                    Height =660
                    ColumnOrder =0
                    TabIndex =7
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblYear1"
                    ControlSource ="=[MinYear]"
                    GridlineStyleBottom =1
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedTop =600
                    LayoutCachedWidth =9719
                    LayoutCachedHeight =1260
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeTint =50.0
                End
                Begin Label
                    TextAlign =2
                    Left =14160
                    Top =660
                    Width =840
                    Height =540
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPri1Parks"
                    Caption ="# Priority 1 Years"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =14160
                    LayoutCachedTop =660
                    LayoutCachedWidth =15000
                    LayoutCachedHeight =1200
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1440
                    Top =900
                    Width =4320
                    Name ="lnSpecies"
                    GridlineColor =10921638
                    LayoutCachedLeft =1440
                    LayoutCachedTop =900
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =900
                End
                Begin Label
                    TextAlign =2
                    Left =1380
                    Top =600
                    Width =4380
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNames"
                    Caption ="Species Names"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1380
                    LayoutCachedTop =600
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =900
                End
                Begin Label
                    TextAlign =1
                    Left =7380
                    Top =960
                    Width =1680
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblCommonName"
                    Caption ="Common Name"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7380
                    LayoutCachedTop =960
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =60
                    Top =960
                    Width =1200
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblFamily"
                    Caption ="Family"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =60
                    LayoutCachedTop =960
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =1
                    Left =6120
                    Top =960
                    Width =1380
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblPlantCode"
                    Caption ="Plant Code"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6120
                    LayoutCachedTop =960
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =3720
                    Top =960
                    Width =1980
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameCO"
                    Caption ="CO"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3720
                    LayoutCachedTop =960
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =1260
                End
                Begin Label
                    TextAlign =2
                    Left =1440
                    Top =960
                    Width =1800
                    Height =300
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblSpeciesNameUT"
                    Caption ="UT"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1440
                    LayoutCachedTop =960
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1260
                End
                Begin Label
                    Left =120
                    Top =60
                    Width =1320
                    Height =300
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="lblPrinted"
                    Caption ="Printed:"
                    GridlineColor =10921638
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1440
                    LayoutCachedHeight =360
                    ForeTint =75.0
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1440
                    Top =60
                    Width =3300
                    Height =315
                    ColumnOrder =10
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDate"
                    ControlSource ="=Format(Now(),\"mmmm d\"\", \"\"yyyy h:nn ampm\")"
                    Format ="Medium Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =1440
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6360
                    Width =2880
                    Height =312
                    ColumnOrder =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxListName"
                    ControlSource ="=IIf([Page]>1,\"Invasives List for \" & TempVars(\"Park\"),\"\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =6360
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =312
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9900
                    Top =60
                    Width =5040
                    Height =312
                    ColumnOrder =8
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxPage"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =9900
                    LayoutCachedTop =60
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =372
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =540
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    IMESentenceMode =3
                    Width =15005
                    Height =490
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxDetail"
                    GridlineColor =10921638

                    LayoutCachedWidth =15005
                    LayoutCachedHeight =490
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1500
                    Top =75
                    Width =2580
                    Height =312
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpeciesUT"
                    ControlSource ="utah_species"
                    GridlineColor =10921638

                    LayoutCachedLeft =1500
                    LayoutCachedTop =75
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =387
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7260
                    Top =75
                    Width =2400
                    Height =312
                    FontSize =8
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxCommon"
                    ControlSource ="Master_Common_Name"
                    GridlineColor =10921638

                    LayoutCachedLeft =7260
                    LayoutCachedTop =75
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =387
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4080
                    Top =75
                    Width =2040
                    Height =312
                    FontSize =8
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSpeciesCO"
                    ControlSource ="Co_Species"
                    GridlineColor =10921638

                    LayoutCachedLeft =4080
                    LayoutCachedTop =75
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =387
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6180
                    Top =75
                    Width =840
                    Height =312
                    FontSize =8
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLUCode"
                    ControlSource ="LU_Code"
                    GridlineColor =10921638

                    LayoutCachedLeft =6180
                    LayoutCachedTop =75
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =387
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =75
                    Width =1500
                    Height =312
                    FontSize =8
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxFamily"
                    ControlSource ="Family"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =75
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =387
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5820
                    Top =180
                    Width =5280
                    Height =300
                    FontSize =9
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxAll"
                    ControlSource ="ParkYearPriorities"
                    GridlineColor =10921638

                    LayoutCachedLeft =5820
                    LayoutCachedTop =180
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =720
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear1"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear1] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =720
                    LayoutCachedTop =180
                    LayoutCachedWidth =960
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1020
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear2"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear2] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =1020
                    LayoutCachedTop =180
                    LayoutCachedWidth =1260
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear3"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear3] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =1320
                    LayoutCachedTop =180
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1620
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear4"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear4] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =1620
                    LayoutCachedTop =180
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1920
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear5"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear5] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =1920
                    LayoutCachedTop =180
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2220
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear6"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear6] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2220
                    LayoutCachedTop =180
                    LayoutCachedWidth =2460
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2520
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear7"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear7] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2520
                    LayoutCachedTop =180
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2820
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear8"
                    ControlSource ="=CountInString([ParkYearPriorities],[TempVars]![Park] & \"-\" & [lblYear8] & \"-"
                        "1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =2820
                    LayoutCachedTop =180
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =480
                    BackThemeColorIndex =-1
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =180
                    Width =660
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxRunSumPri1"
                    ControlSource ="=CountInString([ParkYearPriorities],\"-1\")"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =720
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3360
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year1"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear] & \""
                        "-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =3360
                    LayoutCachedTop =180
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3660
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year2"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear]+1 & "
                        "\"-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =3660
                    LayoutCachedTop =180
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year3"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear]+2 & "
                        "\"-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =3960
                    LayoutCachedTop =180
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4260
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year4"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear]+3 & "
                        "\"-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =4260
                    LayoutCachedTop =180
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4560
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year5"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear]+4 & "
                        "\"-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =4560
                    LayoutCachedTop =180
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4860
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year6"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear]+5 & "
                        "\"-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =4860
                    LayoutCachedTop =180
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5160
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year7"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear]+6 & "
                        "\"-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =5160
                    LayoutCachedTop =180
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5460
                    Top =180
                    Width =240
                    Height =300
                    FontSize =9
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniquePri1Year8"
                    ControlSource ="=IIf(Len(Replace([ParkYearPriorities],[TempVars]![Park] & \"-\" & [MinYear]+7 & "
                        "\"-1\",\"\"))=0,1,0)"
                    GridlineColor =10921638

                    LayoutCachedLeft =5460
                    LayoutCachedTop =180
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9240
                    Top =15
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear1Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9240
                    LayoutCachedTop =15
                    LayoutCachedWidth =9917
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9792
                    Top =15
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =25
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear2Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]+1))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9792
                    LayoutCachedTop =15
                    LayoutCachedWidth =10469
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10392
                    Top =15
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =26
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear3Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]+2))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10392
                    LayoutCachedTop =15
                    LayoutCachedWidth =11069
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10980
                    Top =15
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =27
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear4Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]+3))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10980
                    LayoutCachedTop =15
                    LayoutCachedWidth =11657
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11640
                    Top =15
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =28
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear5Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]+4))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11640
                    LayoutCachedTop =15
                    LayoutCachedWidth =12317
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12300
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =29
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear6Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]+5))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12300
                    LayoutCachedWidth =12977
                    LayoutCachedHeight =420
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12902
                    Top =15
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =30
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear7Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]+6))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12902
                    LayoutCachedTop =15
                    LayoutCachedWidth =13579
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13503
                    Top =15
                    Width =677
                    Height =420
                    FontSize =8
                    TabIndex =31
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxYear8Priority"
                    ControlSource ="=PopulateSpeciesPriorities([TempVars]![Park],[tbxAll],CInt([MinYear]+7))"
                    ControlTipText ="Park Priority"
                    ConditionalFormat = Begin
                        0x0100000090000000020000000000000003000000000000000200000001000000 ,
                        0x00000000ffffff00000000000200000003000000170000000100000000000000 ,
                        0xccffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000005b00740062007800530075006d00530070006500630069006500 ,
                        0x730050007200690031005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13503
                    LayoutCachedTop =15
                    LayoutCachedWidth =14180
                    LayoutCachedHeight =435
                    BackThemeColorIndex =-1
                    ConditionalFormat14 = Begin
                        0x01000200000000000000030000000100000000000000ffffff00010000003100 ,
                        0x0000000000000000000000000000000000000000000000000002000000010000 ,
                        0x0000000000ccffff00130000005b00740062007800530075006d005300700065 ,
                        0x00630069006500730050007200690031005d0000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =14280
                    Top =15
                    Width =660
                    Height =420
                    FontSize =9
                    TabIndex =32
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumSpeciesPri1"
                    ControlSource ="=CountInString([ParkYearPriorities],\"-1\")"
                    StatusBarText ="Park priority"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000002000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x310000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =14280
                    LayoutCachedTop =15
                    LayoutCachedWidth =14940
                    LayoutCachedHeight =435
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ccffff00010000003100 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =4200
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =3
                    Left =5760
                    Width =3480
                    Height =324
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblParkYearPriorities"
                    Caption ="Total # Priority 1 Species by Year =>"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =5760
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =324
                End
                Begin Label
                    BackStyle =1
                    TextAlign =3
                    Left =6852
                    Top =360
                    Width =2388
                    Height =288
                    FontSize =10
                    BackColor =16777164
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblUniquePri1"
                    Caption ="Unique Priority 1 Species =>"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6852
                    LayoutCachedTop =360
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =648
                    BackThemeColorIndex =-1
                End
                Begin Label
                    TextAlign =3
                    Left =7260
                    Top =900
                    Width =2700
                    Height =324
                    FontSize =12
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblTotalNum"
                    Caption ="Total # Priority 1 Species ="
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7260
                    LayoutCachedTop =900
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =1224
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9420
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear1"
                    ControlSource ="=getListLastModifiedDate([lblYear1],TempVars(\"Park\"))"
                    StatusBarText ="=\"List Last Modification Date (\"& [lblYear1] &\")\""
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =4032
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear2"
                    ControlSource ="=getListLastModifiedDate([lblYear2],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (CARE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =4032
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear3"
                    ControlSource ="=getListLastModifiedDate([lblYear3],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (COLM)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10620
                    LayoutCachedTop =1440
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =4032
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11220
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear4"
                    ControlSource ="=getListLastModifiedDate([lblYear4],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (CURE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11220
                    LayoutCachedTop =1440
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =4032
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11820
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear5"
                    ControlSource ="=getListLastModifiedDate([lblYear5],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (DINO)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11820
                    LayoutCachedTop =1440
                    LayoutCachedWidth =12120
                    LayoutCachedHeight =4032
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12420
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear6"
                    ControlSource ="=getListLastModifiedDate([lblYear6],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (FOBU)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12420
                    LayoutCachedTop =1440
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =4032
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear7"
                    ControlSource ="=getListLastModifiedDate([lblYear6],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (GOSP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =1440
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =4032
                End
                Begin TextBox
                    TabStop = NotDefault
                    Vertical = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13740
                    Top =1440
                    Width =300
                    Height =2592
                    FontSize =8
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxLastModYear8"
                    ControlSource ="=getListLastModifiedDate([lblYear8],TempVars(\"Park\"))"
                    StatusBarText ="List Last Modification Date (ZION)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13740
                    LayoutCachedTop =1440
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =4032
                End
                Begin Label
                    TextAlign =3
                    Left =7860
                    Top =1440
                    Width =1260
                    Height =960
                    FontSize =10
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="lblLastModDate"
                    Caption ="Last      Modified  =>\015\012Date      "
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =7860
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =2400
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear1"
                    ControlSource ="=[tbxYear1]"
                    StatusBarText ="=\"Total # priority 1 (\"&[lblYear1]&\")\""
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10020
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear2"
                    ControlSource ="=[tbxYear2]"
                    StatusBarText ="Total # priority 1 (CARE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10620
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear3"
                    ControlSource ="=[tbxYear3]"
                    StatusBarText ="Total # priority 1 (COLM)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10620
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11220
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear4"
                    ControlSource ="=[tbxYear4]"
                    StatusBarText ="Total # priority 1 (CURE)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11220
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =11820
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =12
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear5"
                    ControlSource ="=[tbxyear5]"
                    StatusBarText ="Total # priority 1 (DINO)"
                    GridlineColor =10921638

                    LayoutCachedLeft =11820
                    LayoutCachedWidth =12120
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =12420
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =13
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear6"
                    ControlSource ="=[tbxyear6]"
                    StatusBarText ="Total # priority 1 (FOBU)"
                    GridlineColor =10921638

                    LayoutCachedLeft =12420
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13080
                    Width =300
                    Height =300
                    FontSize =9
                    TabIndex =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear7"
                    ControlSource ="=[tbxYear7]"
                    StatusBarText ="Total # priority 1 (GOSP)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =13740
                    Width =300
                    Height =270
                    FontSize =9
                    TabIndex =15
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumYear8"
                    ControlSource ="=[tbxYear8]"
                    StatusBarText ="Total # priority 1 (ZION)"
                    GridlineColor =10921638

                    LayoutCachedLeft =13740
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9420
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =16
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear1"
                    ControlSource ="=[tbxUniquePri1Year1]"
                    StatusBarText ="=\"Total # priority 1 (\"&[lblYear1]&\")\""
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =9420
                    LayoutCachedTop =360
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10020
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =17
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear2"
                    ControlSource ="=[tbxUniquePri1Year2]"
                    StatusBarText ="Total # priority 1"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10020
                    LayoutCachedTop =360
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10620
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =18
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear3"
                    ControlSource ="=[tbxUniquePri1Year3]"
                    StatusBarText ="Total # priority 1"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =10620
                    LayoutCachedTop =360
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11220
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =19
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear4"
                    ControlSource ="=[tbxUniquePri1Year4]"
                    StatusBarText ="Total # priority 1 (CURE)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11220
                    LayoutCachedTop =360
                    LayoutCachedWidth =11520
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11820
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =20
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear5"
                    ControlSource ="=[tbxUniquePri1Year5]"
                    StatusBarText ="Total # priority 1 (DINO)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =11820
                    LayoutCachedTop =360
                    LayoutCachedWidth =12120
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12420
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =21
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear6"
                    ControlSource ="=[tbxUniquePri1Year6]"
                    StatusBarText ="Total # priority 1 (FOBU)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =12420
                    LayoutCachedTop =360
                    LayoutCachedWidth =12720
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13080
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =22
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear7"
                    ControlSource ="=[tbxUniquePri1Year7]"
                    StatusBarText ="Total # priority 1 (GOSP)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13080
                    LayoutCachedTop =360
                    LayoutCachedWidth =13380
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13740
                    Top =360
                    Width =300
                    Height =252
                    FontSize =9
                    TabIndex =23
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxUniqueSumYear8"
                    ControlSource ="=[tbxUniquePri1Year8]"
                    StatusBarText ="Total # priority 1 (ZION)"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =13740
                    LayoutCachedTop =360
                    LayoutCachedWidth =14040
                    LayoutCachedHeight =612
                    ConditionalFormat14 = Begin
                        0x01000100000000000000040000000100000000000000ccffff00010000003000 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10080
                    Top =900
                    Width =1140
                    Height =312
                    FontSize =12
                    TabIndex =24
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxSumPriority1"
                    ControlSource ="=[tbxRunSumPri1]"
                    StatusBarText ="Standard park code (CANY, FOBU, etc.)"
                    GridlineColor =10921638

                    LayoutCachedLeft =10080
                    LayoutCachedTop =900
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =1212
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
' MODULE:       Report_rpt_Tgt_Species_List_Park_Summary
' Description:  Load species list to target species list functions and routines
'port
' Source/date:  Bonnie Campbell, 9/21/2015
' Revisions:    BLC - 9/21/2015 - initial version
' =================================

' ---------------------------------
' SUB:          Report_Open
' Description:  Actions for when report opens
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Notes:
'   Consider references for performance improvements/user cues that report is still being generated
'   http://stackoverflow.com/questions/11477297/giving-an-alias-to-a-subquery-containing-a-join-in-access
' Source/date:
' Adapted:      Bonnie Campbell, September 21, 2015 - for NCPN tools
' Revisions:
'   BLC - 9/21/2015 - initial version
'   BLC - 9/30/2015 - set report data source SQL to update to TempVars("Park")
'   BLC - 11/27/2015 - cleanup commented out code
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)

On Error GoTo Err_Handler

    'get report data source & alter it using target year to reduce query time?
    Dim i As Integer
    
    Screen.MousePointer = 11 'Hour Glass

    DoCmd.OpenForm "frm_Progress_Bar", acNormal
    
    For i = 1 To 10
        
        Forms("frm_Progress_Bar").Increment i * 10, "Preparing report..."
    
    Next

    'update data source
    Dim strSQL As String
    Dim qdf As DAO.QueryDef
    
'  qry_Tgt_Species_List_Park_Summary_Data SQL:
'    SELECT DISTINCT Master_Plant_Code_FK, LU_Code, Family, Species_Name, utah_species, Co_Species,
'    Wy_Species, Master_Common_Name,
'    ConcatRelated("ParkYearPriority","qry_Annual_Complete_Tgt_Species_Lists","Park = 'DINO'
'    AND Species_Name='"+Species_Name+"'",'',"|") AS ParkYearPriorities,
'    (SELECT Min(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = 'DINO') AS MinYear,
'    (SELECT Max(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = 'DINO') AS MaxYear
'    FROM (SELECT * FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = 'DINO')  AS [%$##@_Alias]
'    GROUP BY Park, Master_Plant_Code_FK, LU_Code, Family, Species_Name, Priority, Transect_Only,
'    Target_Area_ID, Tgt_Area, utah_species, Co_Species, Wy_Species, Master_Common_Name, PriorityTarget,
'    SpeciesYear;
    
    Set qdf = CurrentDb.QueryDefs("qry_Tgt_Species_List_Park_Summary_Data") '("qry_Tgt_Species_List_Park_Summary")
    
    strSQL = "SELECT DISTINCT Master_Plant_Code_FK, LU_Code, Family, Species_Name, utah_species, " _
            & "Co_Species, Wy_Species, Master_Common_Name, " _
            & "ConcatRelated(""ParkYearPriority"", ""qry_Annual_Complete_Tgt_Species_Lists"",""Park= 'PARKNAME' " _
            & "AND Species_Name='""+Species_Name+""'"",'',""|"") AS ParkYearPriorities, " _
            & "(SELECT Min(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = 'PARKNAME') AS MinYear, " _
            & "(SELECT Max(TgtYear) FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = 'PARKNAME') AS MaxYear " _
            & "FROM (SELECT * FROM qry_Annual_Complete_Tgt_Species_Lists WHERE Park = 'PARKNAME') AS [%$##@_Alias] " _
            & "GROUP BY Park, Master_Plant_Code_FK, LU_Code, Family, Species_Name, Priority, Transect_Only, " _
            & "Target_Area_ID, Tgt_Area, utah_species, Co_Species, Wy_Species, Master_Common_Name, " _
            & "PriorityTarget, SpeciesYear;"
            
    strSQL = Replace(strSQL, "PARKNAME", TempVars!Park)
Debug.Print strSQL
    qdf.SQL = strSQL
    

If ReportIsLoaded("rpt_Tgt_Species_List_Park_Summary") Then
     DoEvents
     Pause (5) 'was 15
     DoCmd.Close acForm, "frm_Progress_Bar"
     DoEvents
    
    Pause (10) 'was 30
    ' clear statusbar note running report
    SysCmd acSysCmdSetStatus, "Calculations complete! Fetching report..."
End If

Screen.MousePointer = 1 'Standard Cursor

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Report_rpt_Tgt_Species_List_Park_Summary])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          Report_Load
' Description:  Actions for when report is loaded
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, September 21, 2015 - for NCPN tools
' Revisions:
'   BLC - 9/21/2015 - initial version
' ---------------------------------
Private Sub Report_Load()
On Error GoTo Err_Handler
    



Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Load[Report_rpt_Tgt_Species_List_Park_Summary])"
    End Select
    Resume Exit_Sub
End Sub
