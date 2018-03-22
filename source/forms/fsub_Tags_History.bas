Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11610
    DatasheetFontHeight =9
    ItemSuffix =19
    Left =825
    Top =6105
    Right =12435
    Bottom =8340
    OrderBy ="Change_Date"
    RecSrcDt = Begin
        0x70f54ef075aae340
    End
    RecordSource ="qfsub_Tag_History"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
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
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            SizeMode =3
            PictureAlignment =2
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin ComboBox
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin FormHeader
            Height =0
            BackColor =16768194
            Name ="FormHeader"
            AutoHeight =1
        End
        Begin Section
            Height =1455
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =974
                    Top =795
                    Width =9345
                    Height =481
                    ColumnWidth =2370
                    FontSize =10
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtSpecies_History_Notes"
                    ControlSource ="Change_Notes"
                    StatusBarText ="Comments about this identification change"

                    LayoutCachedLeft =974
                    LayoutCachedTop =795
                    LayoutCachedWidth =10319
                    LayoutCachedHeight =1276
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =75
                            Top =795
                            Width =839
                            Height =316
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblSpecies_History_Notes"
                            Caption ="Notes"
                            LayoutCachedLeft =75
                            LayoutCachedTop =795
                            LayoutCachedWidth =914
                            LayoutCachedHeight =1111
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7904
                    Top =75
                    Width =1320
                    Height =299
                    FontSize =10
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="cboNetwork_User_Name"
                    ControlSource ="Network_User_Name"
                    StatusBarText ="The network user name of the person making the change"

                    LayoutCachedLeft =7904
                    LayoutCachedTop =75
                    LayoutCachedWidth =9224
                    LayoutCachedHeight =374
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7365
                            Top =75
                            Width =479
                            Height =299
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblNetwork_User_Name"
                            Caption ="User"
                            LayoutCachedLeft =7365
                            LayoutCachedTop =75
                            LayoutCachedWidth =7844
                            LayoutCachedHeight =374
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2955
                    Top =75
                    Width =1620
                    Height =299
                    ColumnWidth =2760
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtChange_Date"
                    ControlSource ="Change_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that species identification was changed for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =2955
                    LayoutCachedTop =75
                    LayoutCachedWidth =4575
                    LayoutCachedHeight =374
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2445
                            Top =75
                            Width =464
                            Height =299
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblChange_Date"
                            Caption ="Date"
                            LayoutCachedLeft =2445
                            LayoutCachedTop =75
                            LayoutCachedWidth =2909
                            LayoutCachedHeight =374
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =5729
                    Top =75
                    Width =1590
                    Height =299
                    FontSize =10
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cboContact_ID"
                    ControlSource ="Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                        "Last_Name, tlu_Contacts.First_Name; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="M. Contact identifier (Contact_ID)"
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =5729
                    LayoutCachedTop =75
                    LayoutCachedWidth =7319
                    LayoutCachedHeight =374
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4620
                            Top =75
                            Width =1049
                            Height =299
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblContact_ID"
                            Caption ="Changed by"
                            LayoutCachedLeft =4620
                            LayoutCachedTop =75
                            LayoutCachedWidth =5669
                            LayoutCachedHeight =374
                        End
                    End
                End
                Begin Line
                    OverlapFlags =85
                    Left =30
                    Top =1395
                    Width =10305
                    Name ="Line16"
                    LeftPadding =30
                    TopPadding =30
                    RightPadding =30
                    BottomPadding =30
                    GridlineStyleLeft =0
                    GridlineStyleTop =0
                    GridlineStyleRight =0
                    GridlineStyleBottom =0
                    GridlineWidthLeft =1
                    GridlineWidthTop =1
                    GridlineWidthRight =1
                    GridlineWidthBottom =1
                    LayoutCachedLeft =30
                    LayoutCachedTop =1395
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =1395
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =959
                    Top =420
                    Width =3810
                    Height =299
                    FontSize =10
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="cboTSN_New"
                    ControlSource ="New_Value"
                    StatusBarText ="New TSN of Specimen"

                    LayoutCachedLeft =959
                    LayoutCachedTop =420
                    LayoutCachedWidth =4769
                    LayoutCachedHeight =719
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =75
                            Top =420
                            Width =824
                            Height =299
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTSN_New"
                            Caption ="New ID"
                            LayoutCachedLeft =75
                            LayoutCachedTop =420
                            LayoutCachedWidth =899
                            LayoutCachedHeight =719
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6509
                    Top =420
                    Width =3810
                    Height =299
                    FontSize =10
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="cboTSN_Previous"
                    ControlSource ="Old_Value"
                    StatusBarText ="Previous TSN of Specimen"

                    LayoutCachedLeft =6509
                    LayoutCachedTop =420
                    LayoutCachedWidth =10319
                    LayoutCachedHeight =719
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5625
                            Top =437
                            Width =810
                            Height =270
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTSN_Old"
                            Caption ="Old ID"
                            LayoutCachedLeft =5625
                            LayoutCachedTop =437
                            LayoutCachedWidth =6435
                            LayoutCachedHeight =707
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =960
                    Top =75
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="Change_Type"
                    ControlSource ="Change_Type"

                    LayoutCachedLeft =960
                    LayoutCachedTop =75
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =374
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =75
                            Width =900
                            Height =299
                            FontSize =10
                            Name ="Label18"
                            Caption ="Change To"
                            LayoutCachedTop =75
                            LayoutCachedWidth =900
                            LayoutCachedHeight =374
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
