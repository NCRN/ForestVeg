Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14460
    DatasheetFontHeight =10
    ItemSuffix =111
    Left =6720
    Top =1980
    Right =21465
    Bottom =10080
    DatasheetGridlinesColor =12632256
    Filter ="[TSN]=28610"
    RecSrcDt = Begin
        0xc3438e61353de340
    End
    RecordSource ="tlu_Plants"
    Caption ="Plants"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =8085
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4080
                    Top =660
                    Width =1860
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cmbo_Family"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.Family FROM tlu_Plants GROUP BY tlu_Plants.Family ORDER BY tlu"
                        "_Plants.Family UNION SELECT \"*\" as Family From tlu_Plants;"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"*\""
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =4080
                    LayoutCachedTop =660
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =900
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3480
                            Top =660
                            Width =540
                            Height =255
                            ForeColor =1643706
                            Name ="Family_Label"
                            Caption ="Family"
                            LayoutCachedLeft =3480
                            LayoutCachedTop =660
                            LayoutCachedWidth =4020
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6660
                    Top =660
                    Width =2160
                    TabIndex =1
                    BoundColumn =1
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cmbo_Genus"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.Family, tlu_Plants.Genus FROM tlu_Plants GROUP BY tlu_Plants.F"
                        "amily, tlu_Plants.Genus HAVING (((tlu_Plants.Family) Like Forms!frm_Plants!cmbo_"
                        "Family)) ORDER BY tlu_Plants.Genus; "
                    ColumnWidths ="0;1"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"*\""
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =6660
                    LayoutCachedTop =660
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =900
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6000
                            Top =660
                            Width =600
                            Height =255
                            ForeColor =1643706
                            Name ="Genus_Label"
                            Caption ="Genus"
                            LayoutCachedLeft =6000
                            LayoutCachedTop =660
                            LayoutCachedWidth =6600
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =4
                    Left =9600
                    Top =660
                    Width =1860
                    TabIndex =2
                    BoundColumn =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cmbo_Species"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.Family, tlu_Plants.Genus, tlu_Plants.Species, Min(tlu_Plants.T"
                        "SN) AS MinOfTSN FROM tlu_Plants GROUP BY tlu_Plants.Family, tlu_Plants.Genus, tl"
                        "u_Plants.Species HAVING (((tlu_Plants.Family) Like Forms!frm_Plants!cmbo_Family)"
                        " And ((tlu_Plants.Genus) Like Forms!frm_Plants!cmbo_Genus)); "
                    ColumnWidths ="0;0;1440;0"
                    DefaultValue ="\"*\""
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =9600
                    LayoutCachedTop =660
                    LayoutCachedWidth =11460
                    LayoutCachedHeight =900
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8880
                            Top =660
                            Width =660
                            Height =255
                            ForeColor =1643706
                            Name ="Species_Label"
                            Caption ="Species"
                            LayoutCachedLeft =8880
                            LayoutCachedTop =660
                            LayoutCachedWidth =9540
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =1500
                    Width =10860
                    Height =480
                    FontSize =18
                    FontWeight =700
                    TabIndex =3
                    Name ="txt_Latin"
                    ControlSource ="Latin_Name"
                    FontName ="Tahoma"

                    LayoutCachedLeft =60
                    LayoutCachedTop =1500
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =1980
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =6
                    ListRows =30
                    ListWidth =11520
                    Left =1380
                    Top =600
                    Width =300
                    Height =360
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"7\";\"8\""
                    Name ="cmbo_PickAPlant"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Plants.ID, tlu_Plants.TSN, tlu_Plants.Rank_Name, tlu_Plants.Family, t"
                        "lu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Genus, tlu_Plants.Species, t"
                        "lu_Plants.Favorite, tlu_Plants.Woody, tlu_Plants.Shrub, tlu_Plants.Vine, tlu_Pla"
                        "nts.Herbaceous, tlu_Plants.Targeted_Herb, tlu_Plants.Exotic, tlu_Plants.Sensitiv"
                        "e FROM tlu_Plants WHERE (((tlu_Plants.Family) Like Forms!frm_Plants!cmbo_Family)"
                        " And ((tlu_Plants.Genus) Like Forms!frm_Plants!cmbo_Genus) And ((tlu_Plants.Spec"
                        "ies) Like Forms!frm_Plants!cmbo_Species) And ((tlu_Plants.Favorite) Like Forms!f"
                        "rm_Plants!chk_FIlter_Favorite Or (tlu_Plants.Favorite)=True) And ((tlu_Plants.Wo"
                        "ody) Like Forms!frm_Plants!chk_FIlter_Woody Or (tlu_Plants.Woody)=True) And ((tl"
                        "u_Plants.Shrub) Like Forms!frm_Plants!chk_FIlter_Shrub Or (tlu_Plants.Shrub)=Tru"
                        "e) And ((tlu_Plants.Vine) Like Forms!frm_Plants!chk_FIlter_Vine Or (tlu_Plants.V"
                        "ine)=True) And ((tlu_Plants.Herbaceous) Like Forms!frm_Plants!chk_FIlter_Herb Or"
                        " (tlu_Plants.Herbaceous)=True) And ((tlu_Plants.Targeted_Herb) Like Forms!frm_Pl"
                        "ants!chk_FIlter_Targeted_Herb Or (tlu_Plants.Targeted_Herb)=True) And ((tlu_Plan"
                        "ts.Exotic) Like Forms!frm_Plants!chk_FIlter_Exotic Or (tlu_Plants.Exotic)=True) "
                        "And ((tlu_Plants.Sensitive) Like Forms!frm_Plants!chk_FIlter_Sensitive Or (tlu_P"
                        "lants.Sensitive)=True)) ORDER BY tlu_Plants.Latin_Name; "
                    ColumnWidths ="0;792;936;1296;2736;4464"
                    AfterUpdate ="[Event Procedure]"
                    OnGotFocus ="[Event Procedure]"

                    LayoutCachedLeft =1380
                    LayoutCachedTop =600
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =660
                            Width =1275
                            Height =210
                            FontWeight =700
                            ForeColor =1643706
                            Name ="Label40"
                            Caption ="Find a Plant->"
                            LayoutCachedLeft =60
                            LayoutCachedTop =660
                            LayoutCachedWidth =1335
                            LayoutCachedHeight =870
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    Left =2160
                    Top =660
                    Width =1260
                    Height =210
                    FontWeight =700
                    ForeColor =1643706
                    Name ="Label41"
                    Caption ="Filtering By ->"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =660
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =870
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1800
                    Top =3000
                    Width =1020
                    Height =255
                    ColumnWidth =1020
                    TabIndex =5
                    Name ="txt_TSN_Accepted"
                    ControlSource ="TSN_Accepted"
                    StatusBarText ="ITIS TSN Accepted"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =3000
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =3255
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1800
                    Top =2640
                    Width =1020
                    Height =255
                    ColumnWidth =990
                    TabIndex =6
                    Name ="txt_TSN"
                    ControlSource ="TSN"
                    StatusBarText ="ITIS TSN"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =2640
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =2895
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =660
                            Top =2640
                            Width =1008
                            Height =240
                            Name ="Label43"
                            Caption ="TSN:"
                            LayoutCachedLeft =660
                            LayoutCachedTop =2640
                            LayoutCachedWidth =1668
                            LayoutCachedHeight =2880
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1380
                    Top =4020
                    Width =2760
                    Height =255
                    TabIndex =7
                    Name ="txt_Order"
                    ControlSource ="Order"
                    StatusBarText ="Order"

                    LayoutCachedLeft =1380
                    LayoutCachedTop =4020
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =4275
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =240
                            Top =4020
                            Width =1008
                            Height =240
                            Name ="Label44"
                            Caption ="Order:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4020
                            LayoutCachedWidth =1248
                            LayoutCachedHeight =4260
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1380
                    Top =4320
                    Width =2760
                    Height =255
                    ColumnWidth =1635
                    TabIndex =8
                    Name ="txt_Family"
                    ControlSource ="Family"
                    StatusBarText ="Family"

                    LayoutCachedLeft =1380
                    LayoutCachedTop =4320
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =4575
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =240
                            Top =4320
                            Width =1008
                            Height =240
                            Name ="Label45"
                            Caption ="Family:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4320
                            LayoutCachedWidth =1248
                            LayoutCachedHeight =4560
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1380
                    Top =4620
                    Width =2760
                    Height =255
                    TabIndex =9
                    Name ="txt_Genus"
                    ControlSource ="Genus"
                    StatusBarText ="Genus"

                    LayoutCachedLeft =1380
                    LayoutCachedTop =4620
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =4875
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =240
                            Top =4620
                            Width =1008
                            Height =240
                            Name ="Label46"
                            Caption ="Genus:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4620
                            LayoutCachedWidth =1248
                            LayoutCachedHeight =4860
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1380
                    Top =4920
                    Width =2760
                    Height =255
                    TabIndex =10
                    Name ="txt_Species"
                    ControlSource ="Species"
                    StatusBarText ="Species"

                    LayoutCachedLeft =1380
                    LayoutCachedTop =4920
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =5175
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =240
                            Top =4920
                            Width =1008
                            Height =240
                            Name ="Label47"
                            Caption ="Species:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =4920
                            LayoutCachedWidth =1248
                            LayoutCachedHeight =5160
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1380
                    Top =5220
                    Width =2760
                    Height =255
                    ColumnWidth =1080
                    TabIndex =11
                    Name ="txt_Subspecies"
                    ControlSource ="Subspecies"
                    StatusBarText ="Subspecies"

                    LayoutCachedLeft =1380
                    LayoutCachedTop =5220
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =5475
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =240
                            Top =5220
                            Width =1008
                            Height =240
                            Name ="Label48"
                            Caption ="Subspecies:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =5220
                            LayoutCachedWidth =1248
                            LayoutCachedHeight =5460
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =12420
                    Top =1620
                    Height =255
                    ColumnWidth =1230
                    FontWeight =700
                    TabIndex =12
                    Name ="Rank_Name"
                    ControlSource ="Rank_Name"
                    StatusBarText ="Taxonomic Rank"
                    FontName ="Tahoma"

                    LayoutCachedLeft =12420
                    LayoutCachedTop =1620
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1875
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10980
                            Top =1620
                            Width =1368
                            Height =240
                            FontWeight =700
                            Name ="Label49"
                            Caption ="Identified to:"
                            FontName ="Tahoma"
                            LayoutCachedLeft =10980
                            LayoutCachedTop =1620
                            LayoutCachedWidth =12348
                            LayoutCachedHeight =1860
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1800
                    Top =3360
                    Width =1020
                    Height =255
                    TabIndex =13
                    Name ="txt_PLANTS_Code"
                    ControlSource ="PLANTS_Code"
                    StatusBarText ="Code for taxonomic unit assigned by USDA PLANTS database"
                    AfterUpdate ="[Event Procedure]"
                    OnChange ="[Event Procedure]"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =3360
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =3615
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =600
                            Top =3360
                            Width =1128
                            Height =240
                            Name ="Label50"
                            Caption ="PLANTS Code:"
                            LayoutCachedLeft =600
                            LayoutCachedTop =3360
                            LayoutCachedWidth =1728
                            LayoutCachedHeight =3600
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =240
                    Top =3030
                    ColumnWidth =1560
                    TabIndex =14
                    Name ="chk_Accepted_Found"
                    ControlSource ="Accepted_Found"
                    StatusBarText ="Accepted TSN is in table"

                    LayoutCachedLeft =240
                    LayoutCachedTop =3030
                    LayoutCachedWidth =500
                    LayoutCachedHeight =3270
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =470
                            Top =3000
                            Width =1305
                            Height =240
                            Name ="Label51"
                            Caption ="Accepted Found"
                            LayoutCachedLeft =470
                            LayoutCachedTop =3000
                            LayoutCachedWidth =1775
                            LayoutCachedHeight =3240
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =720
                    Top =6150
                    ColumnWidth =825
                    TabIndex =15
                    Name ="chk_Favorite"
                    ControlSource ="Favorite"
                    StatusBarText ="Indicates that this plant should show up in the \"Quick List\" dropdowns"

                    LayoutCachedLeft =720
                    LayoutCachedTop =6150
                    LayoutCachedWidth =980
                    LayoutCachedHeight =6390
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =980
                            Top =6120
                            Width =645
                            Height =240
                            Name ="Label52"
                            Caption ="Favorite"
                            LayoutCachedLeft =980
                            LayoutCachedTop =6120
                            LayoutCachedWidth =1625
                            LayoutCachedHeight =6360
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =720
                    Top =6450
                    ColumnWidth =735
                    TabIndex =16
                    Name ="chk_Woody"
                    ControlSource ="Woody"
                    StatusBarText ="Include in Woody Vegetation Queries"

                    LayoutCachedLeft =720
                    LayoutCachedTop =6450
                    LayoutCachedWidth =980
                    LayoutCachedHeight =6690
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =980
                            Top =6420
                            Width =585
                            Height =240
                            Name ="Label53"
                            Caption ="Woody"
                            LayoutCachedLeft =980
                            LayoutCachedTop =6420
                            LayoutCachedWidth =1565
                            LayoutCachedHeight =6660
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =720
                    Top =6750
                    ColumnWidth =1200
                    TabIndex =17
                    Name ="chk_Herbaceous"
                    ControlSource ="Herbaceous"
                    StatusBarText ="Include in Herbaceous Vegetation Queries"

                    LayoutCachedLeft =720
                    LayoutCachedTop =6750
                    LayoutCachedWidth =980
                    LayoutCachedHeight =6990
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =980
                            Top =6720
                            Width =945
                            Height =240
                            Name ="Label54"
                            Caption ="Herbaceous"
                            LayoutCachedLeft =980
                            LayoutCachedTop =6720
                            LayoutCachedWidth =1925
                            LayoutCachedHeight =6960
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =720
                    Top =7050
                    TabIndex =18
                    Name ="chk_Targeted_Herb"
                    ControlSource ="Targeted_Herb"
                    StatusBarText ="Include in Targeted Herbaceous Vegetation Queries"

                    LayoutCachedLeft =720
                    LayoutCachedTop =7050
                    LayoutCachedWidth =980
                    LayoutCachedHeight =7290
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =980
                            Top =7020
                            Width =1155
                            Height =240
                            Name ="Label55"
                            Caption ="Targeted_Herb"
                            LayoutCachedLeft =980
                            LayoutCachedTop =7020
                            LayoutCachedWidth =2135
                            LayoutCachedHeight =7260
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =2700
                    Top =6150
                    ColumnWidth =540
                    TabIndex =19
                    Name ="chk_Vine"
                    ControlSource ="Vine"
                    StatusBarText ="Include in Vine Vegetation Queries"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =6150
                    LayoutCachedWidth =2960
                    LayoutCachedHeight =6390
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =2930
                            Top =6120
                            Width =390
                            Height =240
                            Name ="Label56"
                            Caption ="Vine"
                            LayoutCachedLeft =2930
                            LayoutCachedTop =6120
                            LayoutCachedWidth =3320
                            LayoutCachedHeight =6360
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =2700
                    Top =6450
                    ColumnWidth =675
                    TabIndex =20
                    Name ="chk_Shrub"
                    ControlSource ="Shrub"
                    StatusBarText ="Include in Shrub Vegetation Queries"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =6450
                    LayoutCachedWidth =2960
                    LayoutCachedHeight =6690
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =2930
                            Top =6420
                            Width =495
                            Height =240
                            Name ="Label57"
                            Caption ="Shrub"
                            LayoutCachedLeft =2930
                            LayoutCachedTop =6420
                            LayoutCachedWidth =3425
                            LayoutCachedHeight =6660
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =2700
                    Top =6750
                    ColumnWidth =660
                    TabIndex =21
                    Name ="chk_Exotic"
                    ControlSource ="Exotic"
                    StatusBarText ="Include in Exotic Vegetation Queries"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =6750
                    LayoutCachedWidth =2960
                    LayoutCachedHeight =6990
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =2930
                            Top =6720
                            Width =510
                            Height =240
                            Name ="Label58"
                            Caption ="Exotic"
                            LayoutCachedLeft =2930
                            LayoutCachedTop =6720
                            LayoutCachedWidth =3440
                            LayoutCachedHeight =6960
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =2700
                    Top =7050
                    ColumnWidth =930
                    TabIndex =22
                    Name ="chk_Sensitive"
                    ControlSource ="Sensitive"
                    StatusBarText ="Include in Sensitive Vegetation Queries"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =7050
                    LayoutCachedWidth =2960
                    LayoutCachedHeight =7290
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =2930
                            Top =7020
                            Width =720
                            Height =240
                            Name ="Label59"
                            Caption ="Sensitive"
                            LayoutCachedLeft =2930
                            LayoutCachedTop =7020
                            LayoutCachedWidth =3650
                            LayoutCachedHeight =7260
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    OverlapFlags =85
                    Left =60
                    Top =1380
                    Width =13800
                    BorderColor =1643706
                    Name ="Line60"
                    GridlineColor =1643706
                    LayoutCachedLeft =60
                    LayoutCachedTop =1380
                    LayoutCachedWidth =13860
                    LayoutCachedHeight =1380
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =14280
                    Height =540
                    FontSize =18
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="Label36"
                    Caption ="Vegetation Plant Summary"
                    FontName ="Tahoma"
                    LayoutCachedWidth =14280
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =13200
                    Top =120
                    Width =960
                    Height =300
                    TabIndex =23
                    Name ="cmd_Close_Form"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =13200
                    LayoutCachedTop =120
                    LayoutCachedWidth =14160
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Tab
                    OverlapFlags =85
                    Left =4605
                    Top =2460
                    Width =9765
                    Height =5565
                    TabIndex =24
                    Name ="TabCtl61"

                    LayoutCachedLeft =4605
                    LayoutCachedTop =2460
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =8025
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =4740
                            Top =2865
                            Width =9495
                            Height =5025
                            Name ="Occurrences"
                            ImageData = Begin
                                0x00000000
                            End
                            LayoutCachedLeft =4740
                            LayoutCachedTop =2865
                            LayoutCachedWidth =14235
                            LayoutCachedHeight =7890
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =4740
                                    Top =2865
                                    Width =9495
                                    Height =4980
                                    Name ="child_Occurrences"
                                    SourceObject ="Form.fsub_Occurrences"
                                    LinkChildFields ="TSN"
                                    LinkMasterFields ="TSN"

                                    LayoutCachedLeft =4740
                                    LayoutCachedTop =2865
                                    LayoutCachedWidth =14235
                                    LayoutCachedHeight =7845
                                End
                            End
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =3900
                    Width =4200
                    Height =1740
                    Name ="Box66"
                    LayoutCachedLeft =180
                    LayoutCachedTop =3900
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =5640
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =180
                    Top =5880
                    Width =4200
                    Height =1560
                    Name ="Box67"
                    LayoutCachedLeft =180
                    LayoutCachedTop =5880
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =7440
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    Left =2880
                    Top =3360
                    Width =450
                    Height =255
                    ForeColor =1279872587
                    Name ="lbl_Link_to_PLANTS"
                    Caption ="Web"
                    HyperlinkAddress ="http://plants.usda.gov/java/profile?symbol=FRAM2"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =3360
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =3615
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =93
                    Left =2880
                    Top =2640
                    Width =450
                    Height =255
                    ForeColor =1279872587
                    Name ="lbl_Link_to_ITIS"
                    Caption ="Web"
                    HyperlinkAddress ="http://www.itis.gov/servlet/SingleRpt/SingleRpt?search_topic=TSN&search_value=32"
                        "931"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =2640
                    LayoutCachedWidth =3330
                    LayoutCachedHeight =2895
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3840
                    Top =990
                    TabIndex =25
                    Name ="chk_Filter_Favorite"
                    DefaultValue ="True"

                    LayoutCachedLeft =3840
                    LayoutCachedTop =990
                    LayoutCachedWidth =4100
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4100
                            Top =960
                            Width =645
                            Height =240
                            Name ="Label89"
                            Caption ="Favorite"
                            LayoutCachedLeft =4100
                            LayoutCachedTop =960
                            LayoutCachedWidth =4745
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4980
                    Top =990
                    TabIndex =26
                    Name ="chk_Filter_Woody"
                    DefaultValue ="False"

                    LayoutCachedLeft =4980
                    LayoutCachedTop =990
                    LayoutCachedWidth =5240
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5240
                            Top =960
                            Width =645
                            Height =240
                            Name ="Label92"
                            Caption ="Woody"
                            LayoutCachedLeft =5240
                            LayoutCachedTop =960
                            LayoutCachedWidth =5885
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6300
                    Top =990
                    TabIndex =27
                    Name ="chk_Filter_Shrub"
                    DefaultValue ="False"

                    LayoutCachedLeft =6300
                    LayoutCachedTop =990
                    LayoutCachedWidth =6560
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6560
                            Top =960
                            Width =645
                            Height =240
                            Name ="Label94"
                            Caption ="Shrub"
                            LayoutCachedLeft =6560
                            LayoutCachedTop =960
                            LayoutCachedWidth =7205
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7560
                    Top =990
                    TabIndex =28
                    Name ="chk_Filter_Vine"
                    DefaultValue ="False"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =990
                    LayoutCachedWidth =7820
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7820
                            Top =960
                            Width =645
                            Height =240
                            Name ="Label96"
                            Caption ="Vine"
                            LayoutCachedLeft =7820
                            LayoutCachedTop =960
                            LayoutCachedWidth =8465
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8760
                    Top =990
                    TabIndex =29
                    Name ="chk_Filter_Herb"
                    DefaultValue ="False"

                    LayoutCachedLeft =8760
                    LayoutCachedTop =990
                    LayoutCachedWidth =9020
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9020
                            Top =960
                            Width =645
                            Height =240
                            Name ="Label98"
                            Caption ="Herb"
                            LayoutCachedLeft =9020
                            LayoutCachedTop =960
                            LayoutCachedWidth =9665
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9900
                    Top =990
                    TabIndex =30
                    Name ="chk_Filter_Targeted_Herb"
                    DefaultValue ="False"

                    LayoutCachedLeft =9900
                    LayoutCachedTop =990
                    LayoutCachedWidth =10160
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10155
                            Top =960
                            Width =1110
                            Height =240
                            Name ="Label100"
                            Caption ="Targeted Herb"
                            LayoutCachedLeft =10155
                            LayoutCachedTop =960
                            LayoutCachedWidth =11265
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =11580
                    Top =990
                    TabIndex =31
                    Name ="chk_Filter_Exotic"
                    DefaultValue ="False"

                    LayoutCachedLeft =11580
                    LayoutCachedTop =990
                    LayoutCachedWidth =11840
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11840
                            Top =960
                            Width =645
                            Height =240
                            Name ="Label102"
                            Caption ="Exotic"
                            LayoutCachedLeft =11840
                            LayoutCachedTop =960
                            LayoutCachedWidth =12485
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =12840
                    Top =990
                    TabIndex =32
                    Name ="chk_Filter_Sensitive"
                    DefaultValue ="False"

                    LayoutCachedLeft =12840
                    LayoutCachedTop =990
                    LayoutCachedWidth =13100
                    LayoutCachedHeight =1230
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13095
                            Top =960
                            Width =720
                            Height =240
                            Name ="Label104"
                            Caption ="Sensitive"
                            LayoutCachedLeft =13095
                            LayoutCachedTop =960
                            LayoutCachedWidth =13815
                            LayoutCachedHeight =1200
                        End
                    End
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =255
                    Left =60
                    Top =2460
                    Width =4440
                    Height =5580
                    Name ="Box_Attributes"
                    GridlineWidthLeft =2
                    GridlineWidthTop =2
                    GridlineWidthRight =2
                    GridlineWidthBottom =2
                    LayoutCachedLeft =60
                    LayoutCachedTop =2460
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =8040
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =180
                    Top =7560
                    Width =1560
                    Height =300
                    TabIndex =33
                    Name ="cmd_Unlock_Attributes"
                    Caption ="Unlock Attributes"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =7560
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =7860
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1860
                    Top =2040
                    Width =3840
                    Height =255
                    FontWeight =700
                    TabIndex =34
                    Name ="txt_Common_Preferred"
                    ControlSource ="Common"

                    LayoutCachedLeft =1860
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =2295
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2040
                            Width =1740
                            Height =240
                            Name ="Label108"
                            Caption ="NCRN Preferred Name:"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2040
                            LayoutCachedWidth =1800
                            LayoutCachedHeight =2280
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6780
                    Top =2040
                    Width =7590
                    Height =255
                    FontWeight =700
                    TabIndex =35
                    Name ="txt_Common_ITIS"
                    ControlSource ="NPSpecies_Common"

                    LayoutCachedLeft =6780
                    LayoutCachedTop =2040
                    LayoutCachedWidth =14370
                    LayoutCachedHeight =2295
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5760
                            Top =2040
                            Width =990
                            Height =240
                            Name ="Label110"
                            Caption ="NPS Names:"
                            LayoutCachedLeft =5760
                            LayoutCachedTop =2040
                            LayoutCachedWidth =6750
                            LayoutCachedHeight =2280
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmbo_Family_Change()
    Me!cmbo_Genus.Value = "*"
    Me!cmbo_Species.Value = "*"
End Sub

Private Sub cmbo_Family_AfterUpdate()
    Me!cmbo_Genus.Requery
    Me!cmbo_Species.Requery
 End Sub

Private Sub cmbo_Genus_AfterUpdate()
If Me!cmbo_Family = "" Or Me!cmbo_Family = "*" Or IsNull(Me!cmbo_Family) Then
    Me!cmbo_Family.Value = Me!cmbo_Genus.Column(0)
End If
    Me!cmbo_Species.Locked = False
    Me!cmbo_Species.Requery
End Sub

Private Sub cmbo_Genus_Change()
    Me!cmbo_Species.Value = "*"
End Sub

'Private Sub cmbo_Genus_GotFocus()
'    Me.cmbo_Genus.Requery
'If IsNull(Me!cmbo_Family) Then
'    With Me!cmbo_Genus
'    .RowSource = "SELECT tlu_Plants.Family, tlu_Plants.Genus FROM tlu_Plants GROUP BY tlu_Plants.Family, tlu_Plants.Genus, tlu_Plants.Woody ORDER BY tlu_Plants.Genus;"
'    .RowSourceType = "Table/Query"
'    .BoundColumn = 2
'    .ColumnCount = 2
'    .ColumnWidths = "0;1"
'    .Requery
'    End With
'Else
'    With Me!cmbo_Genus
'    .RowSource = "SELECT tlu_Plants.Family, tlu_Plants.Genus FROM tlu_Plants GROUP BY tlu_Plants.Family, tlu_Plants.Genus, tlu_Plants.Woody HAVING ((tlu_Plants.Family)=[Forms]![frm_Plants]![cmbo_Family]);"
'    .RowSourceType = "Table/Query"
'    .BoundColumn = 2
'    .ColumnCount = 2
'    .ColumnWidths = "0;1"
'    .Requery
'    End With
'End If
'End Sub

Private Sub cmbo_PickAPlant_AfterUpdate()
 ' Find the record that matches the control.
    Dim rs As Object
    Set rs = Me.Recordset.Clone
    rs.FindFirst "[ID] = " & Str(Nz(Me![cmbo_PickAPlant], 0))
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    
    If Me!txt_TSN = "" Or IsNull(Me!txt_TSN) Then
        lbl_Link_to_ITIS.HyperlinkAddress = "http://www.itis.gov"
    Else
        lbl_Link_to_ITIS.HyperlinkAddress = "http://www.itis.gov/servlet/SingleRpt/SingleRpt?search_topic=TSN&search_value=" & Me!txt_TSN
    End If
    
    If Me!txt_PLANTS_Code = "" Or IsNull(Me!txt_PLANTS_Code) Then
        lbl_Link_to_PLANTS.HyperlinkAddress = "http://plants.usda.gov"
    Else
        lbl_Link_to_PLANTS.HyperlinkAddress = "http://plants.usda.gov/java/profile?symbol=" & Me!txt_PLANTS_Code
    End If
End Sub

Private Sub cmbo_PickAPlant_GotFocus()
    Me!cmbo_PickAPlant.Requery
End Sub

Private Sub cmbo_Species_GotFocus()
If Me!cmbo_Genus = "" Or IsNull(Me!cmbo_Genus) Then
    MsgBox "You must enter a Genus prior to selecting a species.", , "Vegetation DB"
    Me!cmbo_Genus.SetFocus
End If
    Me!cmbo_Species.Requery
End Sub

Private Sub cmd_Close_Form_Click()
On Error GoTo Err_cmd_close_form_Click

    DoCmd.Close

Exit_cmd_close_form_Click:
    Exit Sub

Err_cmd_close_form_Click:
    MsgBox Err.Description
    Resume Exit_cmd_close_form_Click
End Sub

Private Sub cmd_Unlock_Attributes_Click()
    Dim LockStatus As Boolean
    If cmd_Unlock_Attributes.Caption = "Unlock Attributes" Then
        LockStatus = False
        cmd_Unlock_Attributes.Caption = "Lock Attributes"
        Box_Attributes.BorderColor = vbRed
     Else
        LockStatus = True
        cmd_Unlock_Attributes.Caption = "Unlock Attributes"
        Box_Attributes.BorderColor = vbBlack
    End If
    Me!txt_Common_Preferred.Locked = LockStatus
    Me!txt_Order.Locked = LockStatus
    Me!txt_Family.Locked = LockStatus
    Me!txt_Genus.Locked = LockStatus
    Me!txt_Species.Locked = LockStatus
    Me!txt_Subspecies.Locked = LockStatus
    Me!txt_TSN.Locked = LockStatus
    Me!txt_TSN_Accepted.Locked = LockStatus
    Me!txt_PLANTS_Code.Locked = LockStatus
    Me!chk_Favorite.Locked = LockStatus
    Me!chk_Woody.Locked = LockStatus
    Me!chk_Herbaceous.Locked = LockStatus
    Me!chk_Targeted_Herb.Locked = LockStatus
    Me!chk_Vine.Locked = LockStatus
    Me!chk_Shrub.Locked = LockStatus
    Me!chk_Exotic.Locked = LockStatus
    Me!chk_Sensitive.Locked = LockStatus
    Me!chk_Accepted_Found.Locked = LockStatus
End Sub

Private Sub txt_PLANTS_Code_AfterUpdate()
'If Me!txt_PLANTS_Code = "" Or IsNull(Me!txt_PLANTS_Code) Then
'    lbl_Link_to_PLANTS.HyperlinkAddress = "http://plants.usda.gov"
'Else
'    lbl_Link_to_PLANTS.HyperlinkAddress = "http://plants.usda.gov/java/nameSearch?keywordquery=" & Me!txt_PLANTS_Code & "&mode=symbol"
'End If
End Sub

Private Sub txt_PLANTS_Code_Change()
'If Me!txt_PLANTS_Code = "" Or IsNull(Me!txt_PLANTS_Code) Then
'    lbl_Link_to_PLANTS.HyperlinkAddress = "http://plants.usda.gov"
'Else
'    lbl_Link_to_PLANTS.HyperlinkAddress = "http://plants.usda.gov/java/nameSearch?keywordquery=" & Me!txt_PLANTS_Code & "&mode=symbol"
'End If
End Sub
