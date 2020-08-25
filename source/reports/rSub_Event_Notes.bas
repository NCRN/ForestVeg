Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =162
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =9
    ItemSuffix =5
    Left =570
    Top =1440
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xfb284ba4f4b1e340
    End
    RecordSource ="qFsub_Note_History"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6a01000068010000680100006801000000000000302a00003b01000001000000 ,
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =360
            Name ="ReportHeader"
            AutoHeight =1
            Begin
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Width =1620
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label0"
                    Caption ="Notes"
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =360
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =330
            Name ="Detail"
            AutoHeight =255
            Begin
                Begin TextBox
                    TextFontCharSet =204
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =1200
                    Height =270
                    FontSize =10
                    Name ="Note_Type"
                    ControlSource ="Note_Type"

                    LayoutCachedWidth =1200
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    TextFontCharSet =204
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1200
                    Height =270
                    ColumnWidth =2070
                    FontSize =10
                    TabIndex =1
                    Name ="Note_Date"
                    ControlSource ="Note_Date"
                    Format ="Short Date"

                    LayoutCachedLeft =1200
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextFontCharSet =204
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2640
                    Width =8100
                    Height =270
                    ColumnWidth =13215
                    FontSize =10
                    TabIndex =2
                    Name ="Notes"
                    ControlSource ="Notes"

                    LayoutCachedLeft =2640
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =270
                End
                Begin Line
                    Top =315
                    Width =10800
                    BorderColor =7633277
                    Name ="Line4"
                    LayoutCachedTop =315
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =315
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
