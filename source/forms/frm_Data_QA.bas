Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =2
    GridX =24
    GridY =24
    Width =13275
    DatasheetFontHeight =10
    ItemSuffix =679
    Left =7380
    Top =945
    Right =20655
    Bottom =11025
    DatasheetGridlinesColor =12632256
    Filter ="[Query_name]=\"qQA_C_Sapling_GT_10cm_DBH\" And [Time_frame]=\"2010\""
    RecSrcDt = Begin
        0xdef19da9b06be340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="tbl_QA_Results"
    Caption =" Data Validation and Quality Review Tool"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            OldBorderStyle =1
            TextAlign =1
            FontWeight =700
            BackColor =8388608
            BorderColor =8388608
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderWidth =2
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            BackStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            BorderColor =16776960
        End
        Begin CommandButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin CheckBox
            SpecialEffect =4
            BorderWidth =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderColor =8388608
        End
        Begin OptionGroup
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
        End
        Begin BoundObjectFrame
            BorderLineStyle =0
            BackStyle =0
            BorderColor =16776960
        End
        Begin TextBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin ListBox
            BorderLineStyle =0
            BackColor =8421376
            ForeColor =16777215
            BorderColor =16776960
            FontName ="Arial"
        End
        Begin ComboBox
            BorderLineStyle =0
            BackColor =8421376
            BorderColor =16776960
            ForeColor =16777215
            FontName ="Arial"
        End
        Begin Subform
            BorderLineStyle =0
            BorderColor =16776960
        End
        Begin UnboundObjectFrame
            BackStyle =0
            OldBorderStyle =1
            BorderColor =16776960
        End
        Begin ToggleButton
            FontItalic = NotDefault
            FontSize =8
            ForeColor =-2147483630
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Tab
            FontItalic = NotDefault
            BackStyle =0
            FontWeight =700
            FontName ="Arial"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =10095
            BackColor =13025979
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Top =495
                    Width =13275
                    Height =9600
                    Name ="PageTabs"
                    OnChange ="[Event Procedure]"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =120
                            Top =900
                            Width =13020
                            Height =9063
                            Name ="pgResults"
                            Caption =" Results summary"
                            LayoutCachedLeft =120
                            LayoutCachedTop =900
                            LayoutCachedWidth =13140
                            LayoutCachedHeight =9963
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =215
                                    TextAlign =0
                                    Left =120
                                    Top =900
                                    Width =3300
                                    Height =423
                                    FontWeight =400
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labOverview"
                                    Caption ="* Double-click on the label to change sort order.  Click on a query name to open"
                                        "."
                                    ControlTipText ="View mode"
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =215
                                    Left =9120
                                    Top =960
                                    Width =780
                                    Height =300
                                    Name ="cmdRefresh"
                                    Caption ="Refresh"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Run the validation queries and refresh the results summary"

                                    LayoutCachedLeft =9120
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =9900
                                    LayoutCachedHeight =1260
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =215
                                    Left =10020
                                    Top =960
                                    Width =1500
                                    Height =300
                                    TabIndex =1
                                    Name ="cmdViewReport"
                                    Caption ="Summary Report"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="View the quality review results as a report"

                                    LayoutCachedLeft =10020
                                    LayoutCachedTop =960
                                    LayoutCachedWidth =11520
                                    LayoutCachedHeight =1260
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin Subform
                                    CanShrink = NotDefault
                                    OverlapFlags =247
                                    Left =120
                                    Top =1350
                                    Width =13020
                                    Height =8613
                                    TabIndex =2
                                    BorderColor =0
                                    Name ="subResults"
                                    SourceObject ="Form.fsub_QA_Results"
                                    LinkChildFields ="Time_frame;Data_scope"
                                    LinkMasterFields ="cmbTimeframe;optgScope"

                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    SpecialEffect =2
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    Left =4830
                                    Top =997
                                    Width =1170
                                    TabIndex =3
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    Name ="cmbTypeFilter"
                                    RowSourceType ="Value List"
                                    RowSource ="C;Critical;W;Warning;I;Information"
                                    ColumnWidths ="0;2160"
                                    StatusBarText ="Filter by query type"
                                    AfterUpdate ="[Event Procedure]"
                                    ControlTipText ="Filter by query type"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =3600
                                            Top =990
                                            Width =1110
                                            Height =240
                                            FontWeight =400
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labTypeFilter"
                                            Caption ="Query type:"
                                        End
                                    End
                                End
                                Begin ToggleButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =6120
                                    Top =960
                                    Width =480
                                    Height =300
                                    FontWeight =400
                                    TabIndex =4
                                    ForeColor =0
                                    Name ="togFilterByType"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    Caption ="Filter on"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                                        0xadadadadadadadad
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Turn the type filter on or off"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    RowSourceTypeInt =1
                                    SpecialEffect =2
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =7560
                                    Top =997
                                    Width =900
                                    TabIndex =5
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    Name ="cmbDoneFilter"
                                    RowSourceType ="Value List"
                                    RowSource ="True;False"
                                    StatusBarText ="Filter by the 'Done' flag"
                                    AfterUpdate ="[Event Procedure]"
                                    ControlTipText ="Filter by the 'Done' flag"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =6840
                                            Top =997
                                            Width =600
                                            Height =228
                                            FontWeight =400
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labDoneFilter"
                                            Caption ="Done:"
                                        End
                                    End
                                End
                                Begin ToggleButton
                                    FontItalic = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =8580
                                    Top =960
                                    Width =480
                                    Height =300
                                    FontWeight =400
                                    TabIndex =6
                                    ForeColor =0
                                    Name ="togFilterByDone"
                                    AfterUpdate ="[Event Procedure]"
                                    DefaultValue ="0"
                                    Caption ="Filter on"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                                        0xadadadadadadadad
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Turn the 'Done' filter on or off"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =900
                            Width =13020
                            Height =9060
                            Name ="pgQueryViews"
                            Caption =" View and fix query results"
                            LayoutCachedLeft =120
                            LayoutCachedTop =900
                            LayoutCachedWidth =13140
                            LayoutCachedHeight =9960
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =8040
                                    Top =1035
                                    Width =1320
                                    Height =317
                                    Name ="cmdDesignView"
                                    Caption ="Design view"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the selected query in design view"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =1050
                                    Width =6660
                                    Height =252
                                    TabIndex =1
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                                    Name ="selObject"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT MSysObjects.Name AS Query_name FROM MSysObjects WHERE (((MSysObjects.Name"
                                        ") Like \"qQA_*\") AND ((MSysObjects.Type)=5)) ORDER BY MSysObjects.Name; "
                                    AfterUpdate ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =1050
                                            Width =1110
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labObject"
                                            Caption ="Query name"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =10080
                                    Top =1050
                                    Width =1980
                                    Height =252
                                    TabIndex =2
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtUser"
                                    ControlSource ="QA_user"
                                    OnDirty ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =9480
                                            Top =1050
                                            Width =570
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labUser"
                                            Caption ="QA by"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =12120
                                    Top =1050
                                    Width =1020
                                    Height =252
                                    TabIndex =3
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtRemedy_date"
                                    ControlSource ="Remedy_date"
                                    Format ="mm/dd/yy"

                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    ScrollBars =2
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =1410
                                    Width =11820
                                    Height =660
                                    TabIndex =4
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtQueryDesc"
                                    ControlSource ="Query_description"
                                    StatusBarText ="Description of the query"
                                    OnDirty ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =1410
                                            Width =1035
                                            Height =495
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labQueryDesc"
                                            Caption ="Query description"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    EnterKeyBehavior = NotDefault
                                    AllowAutoCorrect = NotDefault
                                    ScrollBars =2
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =1260
                                    Top =2190
                                    Width =11820
                                    Height =690
                                    TabIndex =5
                                    BackColor =16777215
                                    ForeColor =0
                                    Name ="txtRemedy"
                                    ControlSource ="Remedy_desc"
                                    StatusBarText ="Details about actions taken and/or not taken to resolve errors"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =120
                                            Top =2190
                                            Width =810
                                            Height =495
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labRemedy"
                                            Caption ="Remedy details"
                                        End
                                    End
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =3495
                                    Width =12960
                                    Height =6465
                                    TabIndex =6
                                    BorderColor =0
                                    Name ="subQueryResults"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =255
                                            TextAlign =0
                                            Left =120
                                            Top =3255
                                            Width =1212
                                            Height =252
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labQueryResults"
                                            Caption ="Query results"
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    FELineBreak = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3270
                                    Top =3120
                                    Width =606
                                    Height =255
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =7
                                    BackColor =8454143
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="txtEditQuery"
                                    FontName ="Tahoma"
                                    AsianLineBreak =255

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1440
                                            Top =3120
                                            Width =1770
                                            Height =255
                                            FontSize =9
                                            FontWeight =400
                                            BackColor =16777215
                                            BorderColor =0
                                            ForeColor =0
                                            Name ="labEditQuery"
                                            Caption ="Edit results directly?"
                                            FontName ="Tahoma"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    Enabled = NotDefault
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =5040
                                    Top =3060
                                    Width =1080
                                    Height =317
                                    TabIndex =8
                                    ForeColor =0
                                    Name ="cmdAutoFix"
                                    Caption ="Auto-fix"
                                    StatusBarText ="Run a pre-built query to automatically fix all the records"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Run a pre-built query to automatically fix all the records"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =6660
                                    Top =3060
                                    Width =2040
                                    Height =317
                                    TabIndex =9
                                    ForeColor =0
                                    Name ="cmdOpenRecord"
                                    Caption ="Open selected record"
                                    StatusBarText ="Open the form / query / table specified in the query to the record selected in t"
                                        "he subform"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the form / query / table specified in the query to the record selected in t"
                                        "he subform"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =8880
                                    Top =3060
                                    Height =317
                                    TabIndex =10
                                    ForeColor =0
                                    Name ="cmdOpenBrowser"
                                    Caption ="Data browser"
                                    StatusBarText ="Open the project data browser"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Open the project data browser"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =10500
                                    Top =3060
                                    Height =317
                                    TabIndex =11
                                    ForeColor =0
                                    Name ="cmdExport"
                                    Caption ="Export to Excel"
                                    StatusBarText ="Export the results of the selected query to Excel"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Export the results of the selected query to Excel"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    Left =12060
                                    Top =3060
                                    Width =1020
                                    Height =317
                                    TabIndex =12
                                    ForeColor =0
                                    Name ="cmdRequery"
                                    Caption ="Requery"
                                    OnClick ="[Event Procedure]"
                                    ControlTipText ="Requery the results set for the selected query"

                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =120
                            Top =900
                            Width =13020
                            Height =9063
                            Name ="pgDataTables"
                            Caption =" Browse data tables"
                            LayoutCachedLeft =120
                            LayoutCachedTop =900
                            LayoutCachedWidth =13140
                            LayoutCachedHeight =9963
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin ComboBox
                                    LimitToList = NotDefault
                                    SpecialEffect =2
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListRows =20
                                    ListWidth =11520
                                    Left =840
                                    Top =1050
                                    Width =4320
                                    Height =252
                                    BackColor =-2147483643
                                    BorderColor =0
                                    ForeColor =-2147483640
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"200\""
                                    Name ="selTable"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tsys_Link_Tables.Link_table, tsys_Link_Tables.Description_text FROM tsys_"
                                        "Link_Tables WHERE (((tsys_Link_Tables.Link_table) Like \"tbl_*\" And (tsys_Link_"
                                        "Tables.Link_table)<>\"tbl_QA_Results\")) OR (((tsys_Link_Tables.Link_table)=\"tl"
                                        "u_Project_Crew\")) OR (((tsys_Link_Tables.Link_table)=\"tlu_Project_Taxa\")) OR "
                                        "(((tsys_Link_Tables.Link_table)=\"tlu_Park_Taxa\")); "
                                    ColumnWidths ="4320;7200"
                                    AfterUpdate ="[Event Procedure]"
                                    OnEnter ="[Event Procedure]"

                                    Begin
                                        Begin Label
                                            FontItalic = NotDefault
                                            BackStyle =0
                                            OldBorderStyle =0
                                            OverlapFlags =247
                                            TextAlign =0
                                            Left =180
                                            Top =1050
                                            Width =585
                                            Height =270
                                            BackColor =-2147483633
                                            BorderColor =0
                                            ForeColor =-2147483630
                                            Name ="labTable"
                                            Caption ="Table:"
                                        End
                                    End
                                End
                                Begin Subform
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    SpecialEffect =2
                                    Left =120
                                    Top =1698
                                    Width =12960
                                    Height =8265
                                    TabIndex =1
                                    BorderColor =0
                                    Name ="subDataTables"

                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =0
                                    Left =5340
                                    Top =900
                                    Width =7716
                                    Height =699
                                    FontWeight =400
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labEditWarning"
                                    Caption =" Warning:  This is a last resort!  If possible, open the records needing fixes w"
                                        "ithin the data entry form.  Also, when making manual edits in data tables, pleas"
                                        "e be sure to update the updated_date and updated_by fields if they are present i"
                                        "n the table."
                                    ControlTipText ="View mode"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =12180
                    Top =60
                    Width =720
                    Height =354
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Close the form"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =9720
                    Top =60
                    Width =1914
                    Height =355
                    TabIndex =2
                    BackColor =16777215
                    BorderColor =0
                    Name ="optgMode"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="Change the form mode"

                    Begin
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =10800
                            Top =144
                            OptionValue =1
                            BorderColor =0
                            Name ="optEditMode"

                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =11034
                                    Top =120
                                    Width =390
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labEditMode"
                                    Caption ="Edit"
                                    ControlTipText ="Edit mode"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =9840
                            Top =150
                            OptionValue =0
                            BorderColor =0
                            Name ="optViewMode"

                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =10074
                                    Top =120
                                    Width =495
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labViewMode"
                                    Caption ="View"
                                    ControlTipText ="View mode"
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =3
                    OverlapFlags =85
                    Left =4740
                    Top =60
                    Width =4800
                    Height =355
                    TabIndex =4
                    BackColor =16777215
                    BorderColor =0
                    Name ="optgScope"
                    DefaultValue ="0"
                    ControlTipText ="Scope of the data included in the validation queries: uncertified events, certif"
                        "ied events, or both?"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =0
                            OldBorderStyle =0
                            OverlapFlags =215
                            TextAlign =0
                            Left =4800
                            Top =120
                            Width =945
                            Height =255
                            FontWeight =400
                            BackColor =13025979
                            BorderColor =0
                            ForeColor =0
                            Name ="labIncludeCertified"
                            Caption ="Data scope:"
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =5880
                            Top =144
                            OptionValue =0
                            BorderColor =0
                            Name ="optUncertOnly"

                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =6120
                                    Top =120
                                    Width =1050
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labUncertOnly"
                                    Caption ="Uncert. only"
                                    ControlTipText ="Run queries only on uncertified events"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =7380
                            Top =150
                            OptionValue =1
                            BorderColor =0
                            Name ="optBoth"

                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =7620
                                    Top =120
                                    Width =480
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labBoth"
                                    Caption ="Both"
                                End
                            End
                        End
                        Begin OptionButton
                            SpecialEffect =2
                            OverlapFlags =87
                            BorderWidth =0
                            Left =8340
                            Top =150
                            OptionValue =2
                            BorderColor =0
                            Name ="optCertOnly"

                            Begin
                                Begin Label
                                    FontItalic = NotDefault
                                    BackStyle =0
                                    OldBorderStyle =0
                                    OverlapFlags =119
                                    TextAlign =0
                                    Left =8580
                                    Top =120
                                    Width =870
                                    Height =270
                                    BackColor =16777215
                                    BorderColor =0
                                    ForeColor =0
                                    Name ="labCertOnly"
                                    Caption ="Cert. only"
                                    ControlTipText ="Run queries only on certified events"
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2820
                    Top =120
                    Width =1620
                    TabIndex =3
                    BackColor =16777215
                    BorderColor =0
                    ForeColor =0
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbTimeframe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Forms]![frm_Switchboard]![cTimeframe] AS Timeframe FROM tbl_QA_Results  "
                        "UNION SELECT tbl_QA_Results.Time_frame FROM tbl_QA_Results GROUP BY tbl_QA_Resul"
                        "ts.Time_frame ORDER BY Timeframe DESC;"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =0
                            OldBorderStyle =0
                            OverlapFlags =85
                            TextAlign =0
                            Left =180
                            Top =120
                            Width =2520
                            Height =255
                            FontWeight =400
                            BackColor =13025979
                            BorderColor =0
                            ForeColor =0
                            Name ="labTime_frame"
                            Caption ="Time frame of data being certified:"
                        End
                    End
                End
                Begin CommandButton
                    FontItalic = NotDefault
                    OverlapFlags =247
                    Left =11580
                    Top =960
                    Width =1500
                    Height =300
                    TabIndex =5
                    Name ="cmdDetailedReport"
                    Caption ="Detailed Report"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="View the quality review results as a report"

                    LayoutCachedLeft =11580
                    LayoutCachedTop =960
                    LayoutCachedWidth =13080
                    LayoutCachedHeight =1260
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
' FORM NAME:    frm_Data_QA
' Description:  Standard form for data quality review and validation
' Data source:  tbl_QA_Results
' Data access:  edit only, no deletions; opens to allow additions until a query is
'               selected, at which time additions are disallowed (see code in the subform)
' Pages:        pgResults, pgQueryViews, pgDataTables
' Functions:    fxnUpdateQAResults, fxnFilterRecords, fxnSetQueryFlag
' References:   fxnChangeDelimiter, fxnSaveFile, fxnSwitchboardIsOpen, fxnTableExists
' Source/date:  John R. Boetsch, Jan 2006
' Revisions:    JRB, May 16, 2006 - updated to use a subform for results, added conditional
'                   formatting and sort capability, and improved documentation
'               JRB, June 20, 2006 - added a button on pgResults to open the selected record
'                   in the data entry forms to maximize quality control during record fixes
'               JRB, 8/2/2006 - added additional error trapping to cmdOpenRecord
'               JRB, 10/5/2006 - fixed a problem with the refresh button giving a copy/save
'                   error message by saving the current record and turning off the form filter;
'                   added timeframe to fxnUpdateQAResults, and updated to save record before
'                   running the qa report
'               JRB, 11/14/2007 - revised the description and code in fxnUpdateQAResults
'               JRB, 12/17/2007 - added selTable_Enter to restore table pick list functionality
'                   regardless of back end in Access or SQL Server; added PageTabs change
'                   code to update and bookmark the last-selected subform record upon
'                   moving back to the first page; added code to handle multiple possible
'                   data time frames by adding an unbound ctl and linking the subform to this;
'                   added code to the results set report to also filter on data time frame;
'                   also added code to allow the user to flag records using the Is_done field
'               JRB, May 2008 - updated documentation
'               JRB, 6/18/2008 - updated Form_Open to check switchboard and enable/disable
'                   functionality based on application mode
'               JRB, 7/1/2008 - updated by adding blnRunQueries; added filter capability for
'                   Is_done and query type; added fxnFilterRecords
'               JRB, 9/17/2008 - added ref to frm_Progress_Meter (progress meter popup) in
'                   fxnUpdateQAResults
'               JRB, 9/19/2008 - added optgScope; changed txtTime_frame to cmbTimeframe;
'                   updated fxnUpdateQAResults to reflect both changes; updated call to
'                   rpt_QA_Results
'               JRB, 11/21/2008 - added txtEditQuery and fxnSetQueryFlag; updated to lock
'                   subQueryResults except when the query is named in a way that indicates
'                   its results are editable; updated cmdOpenRecord; updated cmdViewReport;
'                   added error traps to selObject and cmdDesignView; fixed a bug with opening
'                   the report and changing the filter values
'               JRB, 1/13/2009 - added save record to PageTabs_Change (copy/edit error)
'               JRB, 2/23/2009 - added cmdOpenBrowser; fixed a bug in selObject_AfterUpdate and
'                   updated fxnUpdateQAResults
'               JRB, 3/27/2009 - added cmdExport to allow quick results export to Excel
'               JRB, 5/1/2009 - updated cmdOpenBrowser to turn browser filters off by default;
'                   updated cmdExport_Click to default to current application path
'               JRB, 5/22/2009 - updated fxnFilterRecords
'               JRB, 6/10/2009 - updated cmdViewReport, cmdExport, fxnUpdateQAResults
'               JRB, 7/9/2009 - updated selTable to rely on tsys_Link_Tables, if present
'               JRB, 11/3/2009 - added cmdAutoFix and fxnEnableAutoFix
'               JRB, 2/8/2010 - updated fxnSetQueryFlag
' =================================

Dim blnRunQueries As Boolean  ' flag to indicate whether to run the queries upon opening

Private Sub cmdDetailedReport_Click()
On Error GoTo Err_Handler
    Dim strDocName As String
    Dim strCriteria As String
    

        strDocName = "rpt_QA_Detailed Summary"
        strCriteria = ""
        'DoCmd.OpenReport stDocName, acPreview, "qRpt_Event_Summary_Unfiltered", stCriteria
        DoCmd.OpenReport strDocName, acPreview
        
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub

Private Sub Command678_Click()

End Sub

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Close the form if the switchboard is not open
    If fxnSwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' Set the time frame for the form to the switchboard time frame
    Me.cmbTimeframe = Forms!frm_Switchboard!cTimeframe

    ' Change form settings depending on application mode
    'Select Case Forms!frm_Switchboard!cAppMode
    '  Case "admin", "power user"
        Me.pgDataTables.visible = True
        Me.cmdDesignView.Enabled = True
        Me.optgScope.Enabled = True
        Me.cmbTimeframe.Enabled = True
        Me.optgMode.Enabled = True
        Me.cmdRefresh.Enabled = True
        Me.cmdRequery.Enabled = True
        Me.cmdOpenRecord.Enabled = True
        Me.cmdAutoFix.Enabled = True
        Me.selObject.Enabled = True
        ' Run the queries if the user selects Yes
        If MsgBox("Would you like to run the QA queries now?" & vbCrLf & _
            "'No' opens the form without running queries ...", _
            vbYesNo, "Quality Assurance Data Checks") = vbYes Then
            blnRunQueries = True
            fxnUpdateQAResults
        End If

     ' Case "data entry"    ' can only view/update for the current year
     '   Me.pgDataTables.Visible = False
     '   Me.cmdDesignView.Enabled = False
     '   Me.optgScope.Enabled = False
     '   Me.cmbTimeframe.Enabled = True
     '   Me.optgMode.Enabled = False
     '   Me.cmdRefresh.Enabled = True
     '   Me.cmdRequery.Enabled = True
     '   Me.cmdOpenRecord.Enabled = True
     '   Me.cmdAutoFix.Enabled = False
     '   Me.selObject.Enabled = False
        ' Run the queries if the user selects Yes
      '  If MsgBox("Would you like to run the QA queries now?" & vbCrLf & _
      '      "'No' opens the form without running queries ...", _
      '      vbYesNo, "Quality Assurance Data Checks") = vbYes Then
      '      blnRunQueries = True
      '      fxnUpdateQAResults
      '  End If

      'Case Else ' read-only mode
      '  Me.pgDataTables.Visible = False
      '  Me.cmdDesignView.Enabled = False
      '  Me.optgScope.Enabled = False
      '  Me.cmbTimeframe.Enabled = True
      '  Me.optgMode.Enabled = False
      '  Me.cmdRefresh.Enabled = False
      '  Me.cmdRequery.Enabled = False
      '  Me.cmdOpenRecord.Enabled = False
      '  Me.cmdAutoFix.Enabled = False
      '  Me.selObject.Enabled = False
    'End Select

    Me.cmbDoneFilter = "False"
    Me.togFilterByDone = True
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_Load()
    On Error GoTo Err_Handler

    ' Requery the results subform to reflect updates if the user chose to run upon opening
    If blnRunQueries Then Me.subResults.Requery
    ' Turn off the form filter and move to a blank record so that no query record is visible
    Me.Filter = ""
    DoCmd.GoToRecord , , acNewRec

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2105   ' Someone saved the form as not allowing new records
        MsgBox "The form has been saved in a manner that does not permit new" & _
            vbCrLf & "records to be added. Contact the database administrator.", _
            vbOKOnly, "Form saved in wrong mode (QA Tool Load Error)"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmbTimeframe_AfterUpdate()
    On Error GoTo Err_Handler

    If Me.cmbTimeframe <> Forms!frm_Switchboard!cTimeframe Then
        Me.cmdRefresh.Enabled = False
        Me.optgMode.Enabled = False
    Else
        'Select Case Forms!frm_Switchboard!cAppMode
          'Case "admin", "power user"
            Me.cmdRefresh.Enabled = True
            Me.optgMode.Enabled = True
          'Case "data entry"
          '  Me.cmdRefresh.Enabled = True
          '  Me.optgMode.Enabled = False
          'Case Else
            ' leave them as is
        'End Select
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub PageTabs_Change()
    On Error GoTo Err_Handler

    Dim rst As DAO.Recordset
    Dim strCriteria As String
    Dim varReturn As Variant

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.cmdRefresh.Enabled = False Then GoTo Exit_Procedure

    ' If moving to the first page, and if a specific query record has been selected
    '   move the subform bookmark to the currently-selected record
    If Me.PageTabs = 0 And IsNull(Me.selObject) = False Then
        ' Save the current record, reset the form filter and query selector, reset the form
        '   to allow additions, and move to a blank record
        If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

        ' Run the function to update the current QA query record
        varReturn = fxnUpdateQAResults(False, Me.selObject)
        Me.Requery
        strCriteria = "[Query_name] = """ & Me.selObject.Value & _
            """ AND [Time_frame] = """ & Me.cmbTimeframe & _
            """ AND [Data_scope] = " & Me.optgScope

        Set rst = Me.subResults.Form.RecordsetClone
        rst.FindFirst strCriteria
        If rst.NoMatch Then
            'MsgBox "No entry found.", vbInformation
        Else
            Me.subResults.Form.Bookmark = rst.Bookmark
        End If
    ElseIf Me.PageTabs = 1 And IsNull(Me.selObject) = False Then
        ' Call the function to update the query flag
        fxnSetQueryFlag
        fxnEnableAutoFix
    End If

Exit_Procedure:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub optgMode_AfterUpdate()
    On Error GoTo Err_Handler

    ' Change the subform data mode depending on the user choice
    If Me.optgMode = 0 Then
    ' View mode
        Me.subQueryResults.Locked = True
        Me.txtUser.Locked = True
        Me.txtQueryDesc.Locked = True
        Me.txtRemedy.Locked = True
        Me.subDataTables.Locked = True
        Me.Detail.BackColor = 13025979 ' steel blue (default)
    Else
    ' Edit mode
        ' Unlock the subform if an editable query
        If Me.txtEditQuery = "OK" Then Me.subQueryResults.Locked = False _
            Else Me.subQueryResults.Locked = True
        Me.txtUser.Locked = False
        Me.txtQueryDesc.Locked = False
        Me.txtRemedy.Locked = False
        Me.subDataTables.Locked = False
        Me.Detail.BackColor = 12574431 ' haystack
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbTypeFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByType = Not IsNull(Me.cmbTypeFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub togFilterByType_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbTypeFilter) = False Then fxnFilterRecords Else Me.togFilterByType = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmbDoneFilter_AfterUpdate()
    On Error GoTo Err_Handler

    Me.togFilterByDone = Not IsNull(Me.cmbDoneFilter)
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub togFilterByDone_AfterUpdate()
    On Error GoTo Err_Handler

    If IsNull(Me.cmbDoneFilter) = False Then fxnFilterRecords Else Me.togFilterByDone = False

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub Form_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Note: this event is ignored on inserting a new record if BeforeInsert code exists

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    If Me.cmdRefresh.Enabled = False Then GoTo Exit_Procedure

    ' Bail out if no object record is selected - keeps from adding bogus new records
    If IsNull(Me.selObject) Then
        DoCmd.CancelEvent
        GoTo Exit_Procedure
    End If

    ' Once a user starts to make edits in the record, update the user field
    '   on the results summary page
    If fxnSwitchboardIsOpen Then Me.txtUser = Forms![frm_Switchboard].[cUser]
    Me.txtRemedy_date = Now()

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    QA Results Summary Page (pgResults)
' Description:  shows an overview of validation query results
' Unbound ctls: none
' Subforms:     subResults - subform for showing the results summaries
' =================================

Private Sub cmdRefresh_Click()
    On Error GoTo Err_Handler

    ' Save the current record, reset the form filter and query selector, reset the form
    '   to allow additions, and move to a blank record
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord
    Me.Filter = ""
    Me.FilterOn = False
    Me.selObject = Null
    Me.subQueryResults.SourceObject = ""
    Me.AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec

    ' Set the form to view mode and call the event procedure for the form mode ctl
    Me.optgMode = 0
    optgMode_AfterUpdate
    Me.Repaint

    ' Refresh the validation query results (filtering requeries the subform)
    fxnUpdateQAResults
    fxnFilterRecords

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdViewReport_Click()
    On Error GoTo Err_Handler

    ' Generate the QA report
    Dim strRptName As String
    Dim strMsg As String
    Dim strFilter As String
    Dim strTimeframe As String
    Dim strScope As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim varResponse As VbMsgBoxResult

    strRptName = "rpt_QA_Results"

    strMsg = "This will open the quality assurance report ..." & vbCrLf & vbCrLf & _
        "Would you like to limit report results to " & Me.cmbTimeframe & "?"
    varResponse = MsgBox(strMsg, vbYesNoCancel, "Quality assurance report")

    Select Case varResponse
      Case vbCancel
        GoTo Exit_Procedure
      Case vbYes
        strTimeframe = Me.cmbTimeframe
        strFilter = "[Time_frame]=""" & strTimeframe & """"
      Case Else
        strTimeframe = Trim(InputBox("Enter the time frame to filter by" & vbCrLf & _
            "(or leave blank to show all):", "Filter by data time frame", _
            Me.cmbTimeframe))
        If strTimeframe <> "" Then
            strFilter = "[Time_frame]=""" & strTimeframe & """"
        Else
            strFilter = ""
        End If
    End Select

    ' Save the current record so that all changes are reflected in the report
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    Select Case Me.optgScope
      Case 0
        strScope = "Uncertifed event data only"
      Case 1
        strScope = "Both certified and uncertified events"
      Case 2
        strScope = "Certified event data only"
    End Select

    If MsgBox("Would you like to filter by the current data scope?" & _
        vbCrLf & vbCrLf & "   " & strScope, vbYesNo, "Filter by data scope?") = vbYes Then
        If strFilter <> "" Then strFilter = strFilter & " AND "
        strFilter = strFilter & "[Data_scope]=" & Me.optgScope
    End If

    ' Open the formatted report output, filtering on time frame
    DoCmd.OpenReport "rpt_QA_Results", acViewPreview, , strFilter
    If MsgBox("Would you like to save this report?", vbYesNo + vbDefaultButton2, _
        "Save report to a file?") = vbYes Then
        If strTimeframe <> "" Then
            ' Add timeframe to file name
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                strTimeframe & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".snp"
        Else
            strInitFile = Application.CurrentProject.Path & "\" & strRptName & "_" & _
                CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".snp"
        End If
        ' Open the save file dialog and update to the actual name given by the user
        strSaveFile = fxnSaveFile(strInitFile, "Snapshot Viewer (*.snp)", "*.snp")
        DoCmd.OutputTo acOutputReport, strRptName, acFormatSNP, strSaveFile, True
        MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Query Results Page (pgQueryViews)
' Description:  shows records returned by individual QA queries, provides the
'               user the opportunity to fix these
' Unbound ctls: selObject - combo box for selecting the query object by name
' Subforms:     subQueryResults - subform showing results of the selected query
' =================================

Private Sub selObject_AfterUpdate()
    On Error GoTo Err_Handler

    Dim strCriteria As String
    Dim varReturn As Variant

    ' Bail out if the refresh button is disabled (app mode or if selected timeframe <>
    '   db timeframe)
    'If Me.cmdRefresh.Enabled = False Then GoTo Exit_Procedure

    ' Exit if no query selected
    If IsNull(Me.selObject) Then
        MsgBox "Please pick from the list", vbOKOnly, "No Query Selected"
        Me.AllowAdditions = True
        DoCmd.GoToRecord , , acNewRec
        Me.txtEditQuery = ""
        Me.txtEditQuery.ForeColor = 0          'black
        Me.txtEditQuery.BackColor = 8454143    'yellow
        GoTo Exit_Procedure
    End If
    
    ' Bind the subform to the selected query
    Me.subQueryResults.SourceObject = "Query." & Me.selObject.Value
    ' Build the filter string and see if a record already exists
    strCriteria = "[Query_name] = """ & Me.selObject.Value & _
        """ AND [Time_frame] = """ & Me.cmbTimeframe & _
        """ AND [Data_scope] = " & Me.optgScope
    If DCount("*", "tbl_QA_Results", strCriteria) = 0 Then
        ' Run the function to update the current QA query record
        varReturn = fxnUpdateQAResults(False, Me.selObject, True)
    End If
    ' Set the form to the selected record
    Me.Form.Filter = strCriteria
    Me.Form.FilterOn = True

    ' Call the function to update the query flag
    fxnSetQueryFlag
    fxnEnableAutoFix

    Dim qdf As DAO.QueryDef
    Dim qdfs As DAO.QueryDefs
    Set qdfs = DBEngine(0)(0).QueryDefs

    On Error Resume Next
    For Each qdf In qdfs
        If qdf.Name = Me.selObject.Value Then
            MsgBox ("This query returns (" & DCount("*", qdf.Name) & _
                ") records that meet the following criteria: " & _
                vbCrLf & vbCrLf & qdf.Properties("Description"))
        End If
    Next qdf

Exit_Procedure:
    On Error Resume Next
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is no longer available in the application." & _
            vbCrLf & """" & Me.selObject & """", , "Query not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdDesignView_Click()
    On Error GoTo Err_Handler

    ' Open the selected query in design view after checking that a query is selected
    If IsNull(Me.selObject) = False Then _
        DoCmd.OpenQuery Me.selObject.Value, acViewDesign, acReadOnly

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is no longer available in the application." & _
            vbCrLf & """" & Me.selObject & """", , "Query not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdAutoFix_Click()
    On Error GoTo Err_Handler

    Dim ctlAutoFix As Control
    Dim varAutoFix As Variant

    varAutoFix = Null

    On Error Resume Next
    Set ctlAutoFix = Forms!frm_Data_QA.subQueryResults!varAutoFix
    varAutoFix = ctlAutoFix.Value
    On Error GoTo Err_Handler

    If IsNull(varAutoFix) Then
        MsgBox "There are no records selected, or no query is specified to fix the results."
    ElseIf Left(varAutoFix, 1) = "t" Then
    ' Object is a table - open in the next tab
        MsgBox "Object is not labeled as a query:" & vbCrLf & vbCrLf & _
            "  " & varAutoFix, , "No action taken"
    ElseIf Left(varAutoFix, 1) = "q" Then
    ' Object is a query - open on its own
        Dim qdf As DAO.QueryDef
        Dim qdfs As DAO.QueryDefs
        Set qdfs = DBEngine(0)(0).QueryDefs
        On Error Resume Next
        For Each qdf In qdfs
            If qdf.Name = varAutoFix Then
                If MsgBox("This will open/run the following query:" & vbCrLf & vbCrLf & _
                    """" & varAutoFix & """" & vbCrLf & vbCrLf & qdf.Properties("Description"), _
                    vbOKCancel, "Open or run query ...") = vbCancel Then
                    GoTo Exit_Procedure
                End If
            End If
        Next qdf
        DoCmd.OpenQuery varAutoFix
        Me.subQueryResults.Requery
    End If

Exit_Procedure:
    On Error Resume Next
    Set ctlAutoFix = Nothing
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2427   ' No records in the subform
        ' Do nothing ...
      Case 2465   ' Needed field is not present in the record set
        MsgBox "No form is specified for fixing these results", , "Missing query field"
      Case 2467   ' No subform recordset
        MsgBox "No query result set"
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdOpenRecord_Click()
    On Error GoTo Err_Handler

    ' Opens the selected subform record in the object specified in the query
    '   to make use of quality control features of the front end during edits

    Dim ctlObject As Control
    Dim ctlFilter As Control
    Dim ctlArgs As Control
    Dim varObject As Variant
    Dim varFilter As Variant
    Dim varArgs As Variant

    varObject = Null
    varFilter = Null
    varArgs = Null
    
    On Error Resume Next
    Set ctlObject = Forms!frm_Data_QA.subQueryResults!varObject
    varObject = ctlObject.Value
    Set ctlFilter = Forms!frm_Data_QA.subQueryResults!varFilter
    varFilter = ctlFilter.Value
    Set ctlArgs = Forms!frm_Data_QA.subQueryResults!varArgs
    varArgs = ctlArgs.Value
    On Error GoTo Err_Handler

    If IsNull(varObject) Then
        MsgBox "There are no records selected, or no form is specified."
    ElseIf Left(varObject, 1) = "t" Then
    ' Object is a table - open in the next tab
        Me.subDataTables.SourceObject = "Table." & varObject
        Me.selTable = varObject
        Me.pgDataTables.SetFocus
    ElseIf Left(varObject, 1) = "q" Then
    ' Object is a query - open on its own
        Dim qdf As DAO.QueryDef
        Dim qdfs As DAO.QueryDefs
        Set qdfs = DBEngine(0)(0).QueryDefs
        On Error Resume Next
        For Each qdf In qdfs
            If qdf.Name = varObject Then
                If MsgBox("This will open/run the following query:" & vbCrLf & vbCrLf & _
                    """" & varObject & """" & vbCrLf & vbCrLf & qdf.Properties("Description"), _
                    vbOKCancel, "Open or run query ...") = vbCancel Then
                    GoTo Exit_Procedure
                End If
            End If
        Next qdf
        DoCmd.OpenQuery varObject
        Me.subQueryResults.Requery
    ElseIf IsNull(varFilter) Then
    ' Filter by form alone if no filter
        Select Case varObject
          Case "frm_Contacts"
            Set gvarRefContactCtl = Me.subQueryResults
          Case "fsub_Project_Taxa"
            Set gvarRefTaxonCtl = Me.subQueryResults
          Case Else
            Set gvarRefForm = Me.Form
            Set gvarRefCtl = Me.subQueryResults
        End Select
        DoCmd.OpenForm varObject, , , , , , varArgs
    Else
    ' Filter by form and filter
        Select Case varObject
          Case "frm_Contacts"
            Set gvarRefContactCtl = Me.subQueryResults
          Case "fsub_Project_Taxa"
            Set gvarRefTaxonCtl = Me.subQueryResults
          Case Else
            Set gvarRefForm = Me.Form
            Set gvarRefCtl = Me.subQueryResults
        End Select
        DoCmd.OpenForm varObject, , , varFilter, , , varArgs
    End If

Exit_Procedure:
    On Error Resume Next
    Set ctlArgs = Nothing
    Set ctlFilter = Nothing
    Set ctlObject = Nothing
    Set qdfs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 2427   ' No records in the subform
        ' Do nothing ...
      Case 2465   ' Needed field is not present in the record set
        MsgBox "No form is specified for fixing these results", , "Missing query field"
      Case 2467   ' No subform recordset
        MsgBox "No query result set"
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdOpenBrowser_Click()
    On Error GoTo Err_Handler

    Set gvarRefForm = Me.Form
    Set gvarRefCtl = Me.subQueryResults
    ' Open to a blank record - to distinguish from opening to the selected record in the subform
    DoCmd.OpenForm "frm_Data_Browser", , , , acFormAdd, , "off"

Exit_Procedure:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox Err.Number & ": " & Err.Description
    End Select
    Resume Exit_Procedure

End Sub

Private Sub cmdExport_Click()
    On Error GoTo Err_Handler

    Dim strQName As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.selObject) Then GoTo Exit_Procedure
    ' Requery the selected record in the recordset, and update the subform
    Me.subQueryResults.Requery
    strQName = Me.selObject
    strSaveFile = CurrentProject.Path & "\" & strQName & "_" & _
        CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".xls"
    DoCmd.OutputTo acOutputQuery, strQName, acFormatXLS, strSaveFile, True
    MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdRequery_Click()
    On Error GoTo Err_Handler

    'Dim varReturn As Variant

    ' Bail out if no query is currently selected
    If IsNull(Me.selObject) Then GoTo Exit_Procedure
    ' Requery the selected record in the recordset, and update the subform
    Me.subQueryResults.Requery
    ' Run the function to update the current QA query record - commented out because this
    '   is done upon changing page tabs
    'varReturn = fxnUpdateQAResults(False, Me.selObject)

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub txtUser_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Prompt user to confirm before allowing edits in the QA user control
    If MsgBox("Are you sure you want to change the user name?", _
        vbYesNo, "Please confirm ...") = vbNo Then
        DoCmd.CancelEvent
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub txtQueryDesc_Dirty(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Prompt user to confirm before allowing edits in query definition control
    If MsgBox("Are you sure you want to change the query definition?", _
        vbYesNo, "Please confirm ...") = vbNo Then
        DoCmd.CancelEvent
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' PAGE NAME:    Data Tables Page (pgDataTables)
' Description:  allows the user to select and view the contents of individual data
'               tables to make data revisions
' Unbound ctls: selTable - combo box for selecting the table object by name
' Subforms:     subDataTables - subform showing the contents of the selected table
' =================================

Private Sub selTable_Enter()
     On Error GoTo Err_Handler

    Dim strSysTable As String

    strSysTable = "tsys_Link_Tables"     ' System table listing linked tables

    ' If the system table does not exist, replace the row source with one that doesn't use it
    If fxnTableExists(strSysTable) = False Then
        Me.selTable.RowSource = "SELECT MSysObjects.Name " & _
            "FROM MSysObjects " & _
            "WHERE (((MSysObjects.Name) Like 'tbl_*' " & _
            "And (MSysObjects.Name)<>'tbl_QA_Results')) " & _
            "OR (((MSysObjects.Name)='tlu_Project_Crew')) " & _
            "OR (((MSysObjects.Name)='tlu_Project_Taxa')) " & _
            "OR (((MSysObjects.Name)='tlu_Park_Taxa'));"
        Me.selTable.ColumnCount = 1
        Me.selTable.ListWidth = Me.selTable.Width
        Me.selTable.Requery
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub selTable_AfterUpdate()
    On Error GoTo Err_Handler

    ' Once a table is selected, bind the subform to this table
    If IsNull(Me.selTable) Then
    ' If none selected ...
        Me.subDataTables.SourceObject = ""
    Else
    ' If a table is selected ...
        If fxnTableExists(Me.selTable) Then
            Me.subDataTables.SourceObject = "Table." & Me.selTable.Value
        Else
            MsgBox "Unable to find the selected table in the database ...", , _
                "Table not found"
        End If
    End If

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

' =================================
' FUNCTION:     fxnUpdateQAResults
' Description:  Updates the data validation results table
'
'   This function requires that the database contain tbl_QA_Results with the
'   following fields:  Query_name (txt 100), Time_frame (txt 30), Data_scope (tinyint),
'   Query_type (txt 20), Query_result (txt 50), Query_run_time (date/time),
'   Query_description (memo), Query_expression (memo);
'   optional fields:  Remedy_desc (memo), Remedy_date (date/time), QA_user (txt 50),
'   Is_done (yes/no)
'
'   Also required is the query "qsys_QA_query_expressions":  SELECT MSysObjects.Name,
'   MSysQueries.Attribute, MSysQueries.Expression FROM MSysObjects LEFT JOIN MSysQueries
'   ON MSysObjects.Id = MSysQueries.ObjectId WHERE (((MSysObjects.Name) Like "qa*") And
'   ((MSysQueries.Attribute) = 8 Or (MSysQueries.Attribute) = 10) And ((MSysQueries.Expression)
'   Is Not Null)) ORDER BY MSysObjects.Name;
'
'   Also required is the query "qsys_QA_query_errors":
'   SELECT tbl_QA_Results.Query_name, "No longer exists, but in result set" AS Issue,
'   tbl_QA_Results.Time_frame FROM MSysObjects RIGHT JOIN tbl_QA_Results ON
'   MSysObjects.Name = tbl_QA_Results.Query_name WHERE (((tbl_QA_Results.Time_frame)
'   = [Forms]![frm_Switchboard]![cTimeframe]) And ((MSysObjects.Name) Is Null))
'   UNION SELECT MSysObjects.Name AS Query_name, "Not in result set" AS Issue,
'   tbl_QA_Results.Time_frame FROM MSysObjects LEFT JOIN tbl_QA_Results ON
'   MSysObjects.Name = tbl_QA_Results.Query_name WHERE (((MSysObjects.Name) Like "qa_*")
'   And ((tbl_QA_Results.Time_frame) = [Forms]![frm_Switchboard]![cTimeframe])
'   And ((tbl_QA_Results.Query_name) Is Null))
'   UNION SELECT tbl_QA_Results.Query_name, "Not running properly" AS Issue,
'   tbl_QA_Results.Time_frame FROM tbl_QA_Results WHERE (((tbl_QA_Results.Time_frame)
'   =[Forms]![frm_Switchboard]![cTimeframe]) AND ((tbl_QA_Results.Query_run_time) Is Null))
'   OR (((tbl_QA_Results.Time_frame)=[Forms]![frm_Switchboard]![cTimeframe]) AND
'   ((tbl_QA_Results.Query_result) Is Null));
'
'   The following code assumes the following naming convention for all validation queries:
'   1) prefix of "qa_" for all queries that are intended to return results to the user
'       (subqueries may have a prefix of "qasub_")
'   2) 4th character may be used for sorting queries hierarchically (e.g., a-z)
'   3) 5th and 6th characters are the sort order within each level of the hierarchy
'   4) 7th character indicates the severity of the error if it returns records:  1=critical,
'       2=warning, 3=information
'   5) 8th character is an underbar "_" and from the 9th character on is a descriptive name
'       with words separated by an underbar "_" character (no spaces or special characters!)
'
' Parameters:   blnUpdateAll - boolean, false if only one of the QA queries is to be updated
'               strSingleQName - string, the name of the single query to be updated
'               blnCreateNew - boolean, true if a new query needs to be created (given new
'                   filter criteria)
' Returns:      none
' Throws:       none
' References:   fxnChangeDelimiter
' Source/date:  John R. Boetsch, 2006 February
' Revisions:    JRB, 3/9/2006 - added a line to handle nulls for query descriptions
'               JRB, 5/9/2006 - added function call to clean the query expression string
'                   by replacing double quotes with single quotes (thanks to Simon Kingston)
'               JRB, 10/3/2006 - added code to include timeframe in the insert into statement
'               JRB, 11/14/2007 - revised the naming convention description above and
'                   updated Mid() statement to reflect revised naming conventions
'               JRB, 12/17/2007 - updated code to make sure records specify time frame as
'                   well as query name; code also now allows a single query to be updated
'                   instead of the full set, through use of blnUpdateAll and strSingleQName;
'                   code also allows user to use the Is_done flag to sort records
'               JRB, 5/14/2008 - updated qsys_QA_query_errors statement to filter results
'                   records by the current selected timeframe
'               JRB, 9/17/2008 - updated by adding reference to frm_Progress_Meter to show query
'                   names while running queries (helps for optimizing slow QA queries)
'               JRB, 9/19/2008 - updated tbl_QA_Results by adding Data_scope field
'               JRB, 2/20/2009 - added blnCreateNew
'               JRB, 6/10/2009 - added qdfs.Refresh to capture query description changes
' =================================

Private Function fxnUpdateQAResults(Optional blnUpdateAll As Boolean = True, _
    Optional strSingleQName As String, Optional blnCreateNew As Boolean = False)

    On Error GoTo Err_Handler

    Dim qdf As DAO.QueryDef     ' Individual query objects
    Dim qdfs As DAO.QueryDefs   ' The database query set
    Dim strTimeframe As String  ' Data timeframe, from the switchboard
    Dim intScope As Integer     ' Indicates whether or not certified records are included
                                '   in query runs: 0=no, 1=yes, 2=both certified and uncertified
    Dim strSQL As String        ' The SQL statement
    Dim strQName As String      ' Name of the query
    Dim strQType As String      ' Type of query (embedded in name; 1=critical, 2=warning, 3=info)
    Dim strQDesc As String      ' Description of the query
    Dim strQResult As String    ' N records currently returned by the query
    Dim strTResult As String    ' N records previously returned by the query, from the QA table
    Dim strQExp As String       ' WHERE clause expression of the query
    Dim dtRunTime As Date       ' Query run time
    Dim intNErrors As Integer   ' Number of queries that have update problems
    Dim intNQueries As Integer  ' Number of queries total
    Dim varReturn As Variant    ' For manipulating the system meter
    Dim intI As Integer         ' Counter for updating the system meter
    Dim frm As Form             ' Reference to the progress popup form
    Dim strProgForm As String   ' Name of the progress popup form
    Dim strProgress As String   ' Progress bar string

    Set qdfs = DBEngine(0)(0).QueryDefs
    qdfs.Refresh

    DoCmd.Hourglass True

    dtRunTime = Now()   ' Set the run time variable to now

    ' Initialize the progress popup form
    strProgForm = "frm_Progress_Meter"
    DoCmd.OpenForm strProgForm
    Set frm = Forms!frm_Progress_Meter
    frm.Caption = " Running validation queries"
    frm!txtPercent = 0
    intNQueries = 0

    For Each qdf In qdfs
        If Left(qdf.Name, 4) = "qQA_" Then intNQueries = intNQueries + 1
    Next qdf

    On Error Resume Next
    ' Initialize the system meter to indicate progress
    varReturn = SysCmd(acSysCmdInitMeter, "Running validation queries", intNQueries)
    intI = 0

    strTimeframe = "unknown"
    If IsNull(Me.cmbTimeframe) = False Then
        strTimeframe = Me.cmbTimeframe
    Else
        If IsNull(Forms!frm_Switchboard!cTimeframe) = False Then _
            strTimeframe = Forms!frm_Switchboard!cTimeframe
    End If
    intScope = Me.optgScope

    For Each qdf In qdfs
        If Left(qdf.Name, 4) = "qQA_" Then
            intI = intI + 1
            ' Update the percent complete in the progress popup
            frm!txtPercent = Round(100 * intI / intNQueries)
            ' Update the progress bar in the progress popup with sequential "Û" characters
            '   which look like a bar because of the font of the control (20 characters=100%)
            strProgress = String(Round(19 * intI / intNQueries), "Û")
            frm!txtProgress = strProgress
            ' Update the progress meter in the status bar
            varReturn = SysCmd(acSysCmdUpdateMeter, intI)
            strQName = qdf.Name
            ' Update the query name in the progress popup
            frm!txtMsg = strQName
            frm.Repaint
            ' Create the record if all queries are being updated
            If blnUpdateAll Or (blnCreateNew And strQName = strSingleQName) Then
                strQType = mid(strQName, 5, 1)
                If strQType = "" Then strQType = "0"
                ' Create the statement to insert new records
                strSQL = "INSERT INTO tbl_QA_Results " & _
                    "(Query_name, Time_frame, Data_scope, Query_type, Is_done) " & _
                    "SELECT """ & strQName & """ AS Query_name, """ & _
                    strTimeframe & """ AS Time_frame, " & intScope & _
                    " AS Data_scope, """ & strQType & _
                    """ AS Query_type, 0 AS Is_done;"
                ' Run the SQL code
                CurrentDb.Execute strSQL
            End If

            ' Run the following query if all queries are being updated, or if the current
            '   query matches the selected query in the form
            If blnUpdateAll Or strQName = strSingleQName Then
                ' Look up the number of records returned on the last run, for comparison
                strTResult = DLookup("Query_result", "tbl_QA_Results", _
                    "[Query_name]=""" & strQName & """ AND [Time_frame]=""" _
                    & strTimeframe & """ AND [Data_scope]=" & intScope)
                ' Update existing records to refresh the results
                strQResult = DCount("*", qdf.Name)  ' the number of records currently returned
                ' Create the statement to add the query description and expression
                '   (expression not always present)
                strQDesc = " - none defined - "         ' Default in case of error
                strQDesc = qdf.Properties("Description")    ' Query description
                ' Clean up any double-quotes in the description and change to single quotes
                strQDesc = fxnChangeDelimiter(strQDesc)
                strQExp = " - none defined - "          ' Default in case of error
                strQExp = DLookup("Expression", "qsys_QA_query_expressions", "[Name]=""" & _
                    strQName & """")
                ' Clean up any double-quotes in the expression and change to single quotes
                strQExp = fxnChangeDelimiter(strQExp)

                If strQResult = "0" And strQType <> "I" Then
                    ' If the number of records is zero and the query type is not 'information'
                    '   then set the Is_done flag to True
                    strSQL = "UPDATE tbl_QA_Results SET tbl_QA_Results.Query_expression = """ _
                        & strQExp & """, tbl_QA_Results.Query_description = """ & strQDesc & _
                        """, tbl_QA_Results.Query_result = """ & strQResult _
                        & """, tbl_QA_Results.Query_run_time = #" & dtRunTime & _
                        "#, tbl_QA_Results.Is_done = TRUE " & _
                        "WHERE (((tbl_QA_Results.Query_name)=""" & strQName & _
                        """) AND ((tbl_QA_Results.Time_frame)=""" & strTimeframe & _
                        """) AND ((tbl_QA_Results.Data_scope)=" & intScope & "));"

                ElseIf strTResult <> strQResult Then
                    ' If the number of records has changed then set Is_done flag to False
                    strSQL = "UPDATE tbl_QA_Results SET tbl_QA_Results.Query_expression = """ _
                        & strQExp & """, tbl_QA_Results.Query_description = """ & strQDesc & _
                        """, tbl_QA_Results.Query_result = """ & strQResult _
                        & """, tbl_QA_Results.Query_run_time = #" & dtRunTime & _
                        "#, tbl_QA_Results.Is_done = FALSE " & _
                        "WHERE (((tbl_QA_Results.Query_name)=""" & strQName & _
                        """) AND ((tbl_QA_Results.Time_frame)=""" & strTimeframe & _
                        """) AND ((tbl_QA_Results.Data_scope)=" & intScope & "));"

                Else
                    ' Build the update query without changing the Is_done flag
                    strSQL = "UPDATE tbl_QA_Results SET tbl_QA_Results.Query_expression = """ _
                        & strQExp & """, tbl_QA_Results.Query_description = """ & strQDesc & _
                        """, tbl_QA_Results.Query_result = """ & strQResult _
                        & """, tbl_QA_Results.Query_run_time = #" & dtRunTime & _
                        "#  WHERE (((tbl_QA_Results.Query_name)=""" & strQName & _
                        """) AND ((tbl_QA_Results.Time_frame)=""" & strTimeframe & _
                        """) AND ((tbl_QA_Results.Data_scope)=" & intScope & "));"
                End If
                ' Run the SQL code
                CurrentDb.Execute strSQL
            End If
        End If
    Next qdf

    On Error GoTo Err_Handler
    ' Notify the user if queries are not updating properly
    intNErrors = DCount("*", "qsys_QA_query_errors")
    If intNErrors > 0 Then
        If intNErrors = 1 Then
            MsgBox "There is 1 query not updating properly.", vbCritical, _
                "Validation query error"
        Else
            MsgBox "There are " & intNErrors & " queries not updating properly.", vbCritical, _
                "Validation query error"
        End If
        DoCmd.OpenQuery "qsys_QA_query_errors", , acReadOnly
    End If

    If blnUpdateAll Then
        ' Pause for a second before proceeding
        Dim varPause, varStart
        varPause = 1
        varStart = Timer
        Do While Timer < varStart + varPause
            DoEvents    ' Yield to other processes
        Loop
    End If

Exit_Procedure:
    On Error Resume Next
    varReturn = SysCmd(acSysCmdRemoveMeter)
    DoCmd.Close acForm, strProgForm, acSaveNo
    Set frm = Nothing
    DoCmd.Hourglass False
    Set qdfs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnUpdateQAResults)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnFilterRecords
' Description:  Filter the records by the indicated field
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, May 5, 2006
' Revisions:    JRB, May 2008 - made code more robust and error-proof
'               JRB, 7/1/2008 - updated by filtering on the subform rather than the form
'               JRB, 5/22/2009 - updated filter AND clauses
' =================================

Private Function fxnFilterRecords()
    On Error GoTo Err_Handler

    Dim strFilter As String
    Dim bFilterOn As Boolean

    bFilterOn = False
    strFilter = ""

    ' Save the record (to trigger validation)
    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord

    If Me.togFilterByType Then
        bFilterOn = True
        strFilter = strFilter & "[Query_type] = """ & Me.cmbTypeFilter & """"
    End If
    If Me.togFilterByDone Then
        If bFilterOn Then strFilter = strFilter & " AND "
        bFilterOn = True
        strFilter = strFilter & "[Is_done] = " & Me.cmbDoneFilter & ""
    End If

    ' Apply the filter
    'Me.Filter = strFilter
    'Me.FilterOn = bFilterOn
    Me.subResults.Form.Filter = strFilter
    Me.subResults.Form.FilterOn = bFilterOn

    ' Make the labels bold or not depending on filter settings
    Me.labTypeFilter.fontBold = Me.togFilterByType
    Me.labDoneFilter.fontBold = Me.togFilterByDone

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2001   ' Run time canceled event (validation error) - do nothing
        Me.togFilterByType = False
        Me.togFilterByDone = False
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnFilterRecords)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnSetQueryFlag
' Description:  Updates the flag to indicate whether or not the query results are editable
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 10/7/2008
' Revisions:    JRB, 2/8/2010 - updated flag from "X" to "_X" in of x as last letter in name
' =================================

Private Function fxnSetQueryFlag()
    On Error GoTo Err_Handler

    ' Update the visual flag to indicate whether or not the query results are editable
    '   Note: suffix of "_X" means that the query results may be edited
    If Right(Me.selObject.Value, 2) = "_X" Then
        Me.txtEditQuery = "OK"
        Me.txtEditQuery.ForeColor = 16777215   'white
        Me.txtEditQuery.BackColor = 4227072    'green
        ' Unlock the subform if in edit mode
        If Me.optgMode = 1 Then Me.subQueryResults.Locked = False _
            Else Me.subQueryResults.Locked = True
    Else
        Me.txtEditQuery = "No"
        Me.txtEditQuery.ForeColor = 16777215   'white
        Me.txtEditQuery.BackColor = 255        'red
        ' Lock the subform
        Me.subQueryResults.Locked = True
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnSetQueryFlag)"
    End Select
    Resume Exit_Procedure

End Function

' =================================
' FUNCTION:     fxnEnableAutoFix
' Description:  Enables or disables the control for running an action query to fix records
' Parameters:   none
' Returns:      none
' Throws:       none
' References:   none
' Source/date:  John R. Boetsch, 11/3/2009
' Revisions:    <name, date, desc>
' =================================

Private Function fxnEnableAutoFix()
    On Error GoTo Err_Handler

    Dim ctlAutoFix As Control

    Me.cmdAutoFix.Enabled = False

    ' The following looks for 'varAutoFix' field in the query results ...
    '   If it isn't there, it will throw a trapped error and the ctl will remain disabled
    Set ctlAutoFix = Forms!frm_Data_QA.subQueryResults!varAutoFix

    ' If no error, the field is there ... enable the ctl if user has sufficient rights
    Select Case Forms!frm_Switchboard!cAppMode
      Case "admin", "power user"
        Me.cmdAutoFix.Enabled = True
    End Select

Exit_Procedure:
    On Error Resume Next
    Set ctlAutoFix = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case 2465, 2467
        ' Do nothing ...
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fxnEnableAutoFix)"
    End Select
    Resume Exit_Procedure

End Function
