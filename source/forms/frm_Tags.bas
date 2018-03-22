Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14280
    DatasheetFontHeight =9
    ItemSuffix =58
    Left =1515
    Right =16080
    Bottom =6450
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x9807af787caee340
    End
    RecordSource ="qfrm_Tags"
    BeforeUpdate ="[Event Procedure]"
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
    SplitFormDatasheet =1
    SplitFormSize =3255
    SplitFormPrinting =1
    SplitFormDatasheet =1
    SplitFormSize =3255
    SplitFormPrinting =1
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
        Begin Line
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
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
        Begin FormHeader
            Height =540
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =-15
                    Width =14295
                    Height =540
                    FontSize =18
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="Label36"
                    Caption ="Vegetation Tag Summary"
                    FontName ="Tahoma"
                    LayoutCachedLeft =-15
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =540
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5895
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1560
                    Top =765
                    Height =389
                    FontSize =14
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtTag"
                    ControlSource ="Tag"
                    StatusBarText ="Number of physical tag attached to tree"
                    GroupTable =14
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1560
                    LayoutCachedTop =765
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1154
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =14
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =315
                            Top =765
                            Width =1184
                            Height =389
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTag"
                            Caption ="Tag"
                            GroupTable =14
                            BottomPadding =38
                            LayoutCachedLeft =315
                            LayoutCachedTop =765
                            LayoutCachedWidth =1499
                            LayoutCachedHeight =1154
                            LayoutGroup =1
                            GroupTable =14
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7424
                    Top =870
                    Width =1200
                    Height =284
                    FontSize =10
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtAzimuth"
                    ControlSource ="Azimuth"
                    StatusBarText ="Azimuth from plot center to specimen (true north)"

                    LayoutCachedLeft =7424
                    LayoutCachedTop =870
                    LayoutCachedWidth =8624
                    LayoutCachedHeight =1154
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6180
                            Top =870
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblAzimuth"
                            Caption ="Azimuth"
                            LayoutCachedLeft =6180
                            LayoutCachedTop =870
                            LayoutCachedWidth =7364
                            LayoutCachedHeight =1154
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10364
                    Top =840
                    Width =1200
                    Height =284
                    FontSize =10
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtDistance"
                    ControlSource ="Distance"
                    StatusBarText ="Distance (m) from plot center to near EDGE of tree"

                    LayoutCachedLeft =10364
                    LayoutCachedTop =840
                    LayoutCachedWidth =11564
                    LayoutCachedHeight =1124
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9120
                            Top =840
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblDistance"
                            Caption ="Distance"
                            LayoutCachedLeft =9120
                            LayoutCachedTop =840
                            LayoutCachedWidth =10304
                            LayoutCachedHeight =1124
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4649
                    Top =870
                    Width =1200
                    Height =284
                    FontSize =10
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtMicroplot_Number"
                    ControlSource ="Microplot_Number"
                    StatusBarText ="The Microplot location of specimen"

                    LayoutCachedLeft =4649
                    LayoutCachedTop =870
                    LayoutCachedWidth =5849
                    LayoutCachedHeight =1154
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3405
                            Top =870
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblMicroplot"
                            Caption ="Microplot"
                            LayoutCachedLeft =3405
                            LayoutCachedTop =870
                            LayoutCachedWidth =4589
                            LayoutCachedHeight =1154
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1544
                    Top =1980
                    Width =10365
                    Height =736
                    FontSize =10
                    TabIndex =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtTag_Notes"
                    ControlSource ="Tag_Notes"
                    StatusBarText ="Comments about this specimen"

                    LayoutCachedLeft =1544
                    LayoutCachedTop =1980
                    LayoutCachedWidth =11909
                    LayoutCachedHeight =2716
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =1980
                            Width =1184
                            Height =736
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTag_Notes"
                            Caption ="Tag Notes"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1980
                            LayoutCachedWidth =1484
                            LayoutCachedHeight =2716
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10364
                    Top =1260
                    Width =1200
                    Height =284
                    ColumnWidth =2415
                    FontSize =10
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtStart_Date"
                    ControlSource ="Start_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that tracking began on this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =10364
                    LayoutCachedTop =1260
                    LayoutCachedWidth =11564
                    LayoutCachedHeight =1544
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9120
                            Top =1260
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label21lblStart_Date"
                            Caption ="Start Date"
                            LayoutCachedLeft =9120
                            LayoutCachedTop =1260
                            LayoutCachedWidth =10304
                            LayoutCachedHeight =1544
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =10364
                    Top =1620
                    Width =1200
                    Height =284
                    FontSize =10
                    TabIndex =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtStop_Date"
                    ControlSource ="Stop_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that tracking ended for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =10364
                    LayoutCachedTop =1620
                    LayoutCachedWidth =11564
                    LayoutCachedHeight =1904
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =9120
                            Top =1620
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblStop_Date"
                            Caption ="Stop Date"
                            LayoutCachedLeft =9120
                            LayoutCachedTop =1620
                            LayoutCachedWidth =10304
                            LayoutCachedHeight =1904
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1560
                    Top =1230
                    Height =285
                    FontWeight =700
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtPlot_Name"
                    ControlSource ="Plot_Name"
                    StatusBarText ="M. Name of the location (Plot_Name)"
                    GroupTable =14
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1230
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1515
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =14
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =315
                            Top =1230
                            Width =1184
                            Height =285
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblPlot_Name"
                            Caption ="Plot_Name"
                            GroupTable =14
                            BottomPadding =38
                            LayoutCachedLeft =315
                            LayoutCachedTop =1230
                            LayoutCachedWidth =1499
                            LayoutCachedHeight =1515
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =14
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7424
                    Top =1260
                    Width =1215
                    Height =284
                    FontSize =10
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtFrame"
                    ControlSource ="Frame"
                    StatusBarText ="Sampling Frame (Regional or Park)"

                    LayoutCachedLeft =7424
                    LayoutCachedTop =1260
                    LayoutCachedWidth =8639
                    LayoutCachedHeight =1544
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6180
                            Top =1260
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblFrame"
                            Caption ="Frame"
                            LayoutCachedLeft =6180
                            LayoutCachedTop =1260
                            LayoutCachedWidth =7364
                            LayoutCachedHeight =1544
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =85
                    Left =330
                    Top =555
                    Width =11640
                    Name ="Line38"
                    LayoutCachedLeft =330
                    LayoutCachedTop =555
                    LayoutCachedWidth =11970
                    LayoutCachedHeight =555
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =7
                    ListWidth =5850
                    Left =1590
                    Top =105
                    Height =315
                    TabIndex =11
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cboTagFinder"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qfrm_Tags.Tag_ID, qfrm_Tags.Tag, qfrm_Tags.Plot_Name, qfrm_Tags.Tag_Statu"
                        "s, qfrm_Tags.Distance, qfrm_Tags.Azimuth, qfrm_Tags.Microplot_Number FROM qfrm_T"
                        "ags ORDER BY qfrm_Tags.Tag; "
                    ColumnWidths ="0;810;1185;960;795;750;1350"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1590
                    LayoutCachedTop =105
                    LayoutCachedWidth =3030
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =105
                            Width =1470
                            Height =320
                            Name ="lblTagFinder"
                            Caption ="Choose a Tag >"
                            LayoutCachedLeft =60
                            LayoutCachedTop =105
                            LayoutCachedWidth =1530
                            LayoutCachedHeight =425
                        End
                    End
                End
                Begin ComboBox
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =1560
                    Top =2790
                    Width =3555
                    Height =314
                    FontSize =10
                    TabIndex =7
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cboTSN"
                    ControlSource ="TSN"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.TSN_Accepted, tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plant"
                        "s.Tree FROM tlu_Plants WHERE (((tlu_Plants.Tree)=True)) ORDER BY tlu_Plants.Lati"
                        "n_Name; "
                    ColumnWidths ="0;0;3600;0"
                    StatusBarText ="TSN of Specimen"
                    GroupTable =15
                    RightPadding =38
                    BottomPadding =38
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1560
                    LayoutCachedTop =2790
                    LayoutCachedWidth =5115
                    LayoutCachedHeight =3104
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =2
                    GroupTable =15
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =315
                            Top =2790
                            Width =1184
                            Height =314
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTSN"
                            Caption ="Species ID"
                            GroupTable =15
                            BottomPadding =38
                            LayoutCachedLeft =315
                            LayoutCachedTop =2790
                            LayoutCachedWidth =1499
                            LayoutCachedHeight =3104
                            LayoutGroup =2
                            GroupTable =15
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1544
                    Top =1620
                    Width =2475
                    Height =314
                    FontSize =10
                    TabIndex =6
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cboTag_Status"
                    ControlSource ="Tag_Status"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Tag Status\")) ORDER BY tlu_Enumerations.Sort_Order; "
                    StatusBarText ="Last sampled as tree or sapling?"
                    AllowValueListEdits =0
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =1544
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4019
                    LayoutCachedHeight =1934
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =1620
                            Width =1184
                            Height =299
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTag_Status"
                            Caption ="Tag Status"
                            LayoutCachedLeft =300
                            LayoutCachedTop =1620
                            LayoutCachedWidth =1484
                            LayoutCachedHeight =1919
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =315
                    Top =3405
                    Width =13560
                    Height =2130
                    TabIndex =12
                    Name ="fsub_Tags_History_Summary"
                    SourceObject ="Form.fsub_Tags_History_Summary"
                    LinkChildFields ="Tag_ID"
                    LinkMasterFields ="Tag_ID"

                    LayoutCachedLeft =315
                    LayoutCachedTop =3405
                    LayoutCachedWidth =13875
                    LayoutCachedHeight =5535
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =315
                            Top =3165
                            Width =1755
                            Height =315
                            FontSize =10
                            Name ="lblfsub_Tags_History"
                            Caption ="Tag History"
                            LayoutCachedLeft =315
                            LayoutCachedTop =3165
                            LayoutCachedWidth =2070
                            LayoutCachedHeight =3480
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4664
                    Top =1260
                    Width =1215
                    Height =284
                    FontSize =10
                    TabIndex =13
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Text53"
                    ControlSource ="Panel"
                    StatusBarText ="Sampling Frame (Regional or Park)"

                    LayoutCachedLeft =4664
                    LayoutCachedTop =1260
                    LayoutCachedWidth =5879
                    LayoutCachedHeight =1544
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3420
                            Top =1260
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label54"
                            Caption ="Panel"
                            LayoutCachedLeft =3420
                            LayoutCachedTop =1260
                            LayoutCachedWidth =4604
                            LayoutCachedHeight =1544
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =12644
                    Top =3000
                    Width =1200
                    Height =284
                    FontSize =10
                    TabIndex =14
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="txtUpdate"
                    ControlSource ="Updated_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that tracking ended for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =12644
                    LayoutCachedTop =3000
                    LayoutCachedWidth =13844
                    LayoutCachedHeight =3284
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11400
                            Top =3000
                            Width =1184
                            Height =284
                            FontSize =10
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="Label56"
                            Caption ="Last Update"
                            LayoutCachedLeft =11400
                            LayoutCachedTop =3000
                            LayoutCachedWidth =12584
                            LayoutCachedHeight =3284
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4440
                    Top =1650
                    TabIndex =15
                    BorderColor =10921638
                    Name ="RFS"
                    ControlSource ="RFS"
                    StatusBarText ="Removed from study"
                    GridlineColor =10921638

                    LayoutCachedLeft =4440
                    LayoutCachedTop =1650
                    LayoutCachedWidth =4700
                    LayoutCachedHeight =1890
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4670
                            Top =1620
                            Width =405
                            Height =315
                            Name ="Label57"
                            Caption ="RFS"
                            LayoutCachedLeft =4670
                            LayoutCachedTop =1620
                            LayoutCachedWidth =5075
                            LayoutCachedHeight =1935
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

Private Sub cboTagFinder_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[Tag_ID] = '" & Me![cboTagFinder] & "'"
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    cboTagFinder = ""
End Sub

'Private Sub Form_AfterUpdate()
'    Me!Updated_Date = Now()
'End Sub

'Private Sub Form_BeforeUpdate()
'    Me!Updated_Date = Now()
'End Sub
'Change the Updated Date to the current time
'
'           Me!Updated_Date.Value = Now()
    
'End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
    Me!Updated_Date = Now()
End Sub
