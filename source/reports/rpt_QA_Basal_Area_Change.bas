Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =178
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =9
    ItemSuffix =55
    Left =1380
    Top =-330
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xad5ca9d40f46e440
    End
    RecordSource ="qSum_Tagged_Items_w_Basal_Area_Prev-Curr_Condensed_Filtered"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000403800006801000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin Line
            BorderLineStyle =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin Subform
            BorderLineStyle =0
        End
        Begin EmptyCell
            Height =240
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Plot_Name"
        End
        Begin BreakLevel
            ControlSource ="Plot_Name"
        End
        Begin BreakLevel
            ControlSource ="Tag"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =705
            BackColor =11525325
            Name ="ReportHeader"
            AutoHeight =1
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =60
                    Width =8550
                    Height =645
                    FontSize =22
                    FontWeight =700
                    ForeColor =5054976
                    Name ="Auto_Title0"
                    Caption ="Basal Area Change Summary(2012-2016)"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =8550
                    LayoutCachedHeight =705
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8700
                    Top =60
                    Width =2010
                    Height =315
                    ColumnOrder =0
                    Name ="Text0"
                    ControlSource ="=Date()"
                    Format ="Short Date"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =60
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =375
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8700
                    Top =360
                    Width =2010
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text1"
                    ControlSource ="=Time()"
                    Format ="Long Time"

                    LayoutCachedLeft =8700
                    LayoutCachedTop =360
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =675
                End
            End
        End
        Begin PageHeader
            Height =720
            Name ="PageHeaderSection"
            BackThemeColorIndex =1
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =180
                    Top =60
                    Width =720
                    Height =555
                    FontWeight =700
                    Name ="Label21"
                    Caption ="Tag"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =615
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =960
                    Top =60
                    Width =2280
                    Height =555
                    FontWeight =700
                    Name ="Label22"
                    Caption ="Status \015\012Previous/Current"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =960
                    LayoutCachedTop =60
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =615
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3300
                    Top =60
                    Width =1215
                    Height =555
                    FontWeight =700
                    Name ="Label23"
                    Caption ="Sampled As \015\012Prev/Curr"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =3300
                    LayoutCachedTop =60
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =615
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =4575
                    Top =60
                    Width =2220
                    Height =555
                    FontWeight =700
                    Name ="Label24"
                    Caption ="Stems \015\012Prev/Curr"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4575
                    LayoutCachedTop =60
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =615
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =9780
                    Top =60
                    Width =1620
                    Height =555
                    FontWeight =700
                    Name ="Label25"
                    Caption ="Species"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =9780
                    LayoutCachedTop =60
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =615
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =11460
                    Top =60
                    Width =2040
                    Height =555
                    FontWeight =700
                    Name ="Label27"
                    Caption ="Note"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =11460
                    LayoutCachedTop =60
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =615
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =13560
                    Top =60
                    Width =810
                    Height =555
                    FontWeight =700
                    Name ="Label28"
                    Caption ="BA % Change"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =13560
                    LayoutCachedTop =60
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =615
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =6855
                    Top =60
                    Width =2865
                    Height =555
                    FontWeight =700
                    Name ="Label48"
                    Caption ="Crown Class Prev/Curr"
                    GroupTable =1
                    GridlineColor =10921638
                    LayoutCachedLeft =6855
                    LayoutCachedTop =60
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =615
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =480
            BackColor =14211288
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =60
                    Width =1560
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Plot_Name"
                    ControlSource ="Plot_Name"
                    StatusBarText ="M. Name of the location (Plot_Name)"

                    LayoutCachedTop =60
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =420
                    BackThemeColorIndex =1
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =960
                    Top =30
                    Width =2280
                    Height =255
                    ColumnWidth =2535
                    FontSize =8
                    TabIndex =1
                    Name ="Status_07-11"
                    ControlSource ="Status_Prev-Curr"
                    EventProcPrefix ="Status_07_11"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =960
                    LayoutCachedTop =30
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =180
                    Top =30
                    Width =720
                    Height =255
                    ColumnWidth =735
                    FontSize =10
                    Name ="Tag"
                    ControlSource ="Tag"
                    StatusBarText ="Number of physical tag attached to tree"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =180
                    LayoutCachedTop =30
                    LayoutCachedWidth =900
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3300
                    Top =30
                    Width =1215
                    Height =255
                    ColumnWidth =2070
                    FontSize =8
                    TabIndex =2
                    Name ="Sampled_As_07-11"
                    ControlSource ="Sampled_As_Prev-Curr"
                    EventProcPrefix ="Sampled_As_07_11"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =3300
                    LayoutCachedTop =30
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4575
                    Top =30
                    Width =2220
                    Height =255
                    ColumnWidth =1545
                    FontSize =8
                    TabIndex =3
                    Name ="Stems_07-11"
                    ControlSource ="=[StemList_Previous] & \" /\" & [StemList_Current]"
                    EventProcPrefix ="Stems_07_11"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =4575
                    LayoutCachedTop =30
                    LayoutCachedWidth =6795
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9780
                    Top =30
                    Width =1620
                    Height =255
                    ColumnWidth =1635
                    FontSize =8
                    TabIndex =5
                    Name ="BA_cm2_2007"
                    ControlSource ="Latin_Name"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =9780
                    LayoutCachedTop =30
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =5
                    ColumnEnd =5
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =11460
                    Top =30
                    Width =2040
                    Height =255
                    ColumnWidth =1875
                    FontSize =7
                    TabIndex =6
                    Name ="BA_cm2_Change"
                    ControlSource ="Notes"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =11460
                    LayoutCachedTop =30
                    LayoutCachedWidth =13500
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =6
                    ColumnEnd =6
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =13560
                    Top =30
                    Width =810
                    Height =255
                    ColumnWidth =2130
                    FontSize =10
                    TabIndex =7
                    Name ="BA_cm2_PctChange"
                    ControlSource ="BA_cm2_PctChange"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =13560
                    LayoutCachedTop =30
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =7
                    ColumnEnd =7
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =6855
                    Top =30
                    Width =2865
                    Height =255
                    FontSize =8
                    TabIndex =4
                    Name ="CrownClass_Prev_Curr"
                    ControlSource ="CrownClass_Prev_Curr"
                    GroupTable =1
                    GridlineColor =10921638

                    LayoutCachedLeft =6855
                    LayoutCachedTop =30
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =285
                    RowStart =1
                    RowEnd =1
                    ColumnStart =4
                    ColumnEnd =4
                    LayoutGroup =1
                    GridlineThemeColorIndex =1
                    GridlineShade =65.0
                    GroupTable =1
                End
            End
        End
        Begin PageFooter
            Height =495
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =180
                    Width =3600
                    Height =315
                    Name ="Text2"
                    ControlSource ="=\"Page \" & [Page]"

                    LayoutCachedLeft =3960
                    LayoutCachedTop =180
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =495
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
