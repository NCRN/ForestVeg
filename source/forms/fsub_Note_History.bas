Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7500
    DatasheetFontHeight =9
    ItemSuffix =8
    Left =7155
    Top =3645
    Right =14850
    Bottom =9525
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x0faa2269f6a6e340
    End
    RecordSource ="qfsub_Note_History"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
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
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin FormHeader
            Height =0
            BackColor =15527148
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =1680
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =1320
                    Height =301
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Note_Type"
                    ControlSource ="Note_Type"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =301
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1410
                    Width =1770
                    Height =361
                    ColumnWidth =2070
                    FontWeight =700
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="Note_Date"
                    ControlSource ="Note_Date"
                    Format ="dd-mmm-yyyy"

                    LayoutCachedLeft =1410
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =361
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Top =300
                    Width =7440
                    Height =1261
                    ColumnWidth =13215
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =15527148
                    Name ="txtNotes"
                    ControlSource ="=IIf(IsNull([Notes]),\"-- None --\",[Notes])"

                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =7500
                    LayoutCachedHeight =1561
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
