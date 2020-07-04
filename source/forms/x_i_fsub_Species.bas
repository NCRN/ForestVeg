Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7860
    DatasheetFontHeight =9
    ItemSuffix =74
    Left =3240
    Top =4110
    Right =11295
    Bottom =8745
    DatasheetGridlinesColor =12632256
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x3d9c36b74cece440
    End
    RecordSource ="usys_temp_speciescover"
    Caption ="fsub_Species"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontWeight =700
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
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
            Height =540
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =3120
                    Width =2700
                    Height =240
                    Name ="Nested_Quad_Label"
                    Caption ="% Cover in Classes"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3120
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5865
                    Top =60
                    Width =840
                    Height =420
                    Name ="Percent_Cover_Label"
                    Caption ="AverageCover"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5865
                    LayoutCachedTop =60
                    LayoutCachedWidth =6705
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =165
                    Top =240
                    Width =960
                    Height =240
                    Name ="lblSpecies"
                    Caption ="Species"
                    LayoutCachedLeft =165
                    LayoutCachedTop =240
                    LayoutCachedWidth =1125
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =95
                    Left =3105
                    Top =240
                    Width =900
                    Height =240
                    Name ="lblQ1"
                    Caption ="Q1@0m"
                    LayoutCachedLeft =3105
                    LayoutCachedTop =240
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =3885
                    Top =240
                    Width =915
                    Height =240
                    Name ="lblQ2"
                    Caption ="Q2@4.5m"
                    LayoutCachedLeft =3885
                    LayoutCachedTop =240
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =4890
                    Top =240
                    Width =915
                    Height =240
                    Name ="lblQ3"
                    Caption ="Q3@9.5m"
                    LayoutCachedLeft =4890
                    LayoutCachedTop =240
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =480
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2100
                    Top =60
                    Width =840
                    Height =420
                    Name ="lblIsDead"
                    Caption ="Dead or Alive?"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =480
                End
            End
        End
        Begin Section
            Height =780
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =7740
                    Height =645
                    TabIndex =13
                    BackColor =-2147483633
                    Name ="tbxRecord"
                    ConditionalFormat = Begin
                        0x010000007a000000010000000100000000000000000000000c00000001000000 ,
                        0x00000000ffcccc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2200740062007800440075007000650022003e00310000000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =705
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ffcccc000b0000002200 ,
                        0x740062007800440075007000650022003e003100000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =165
                    Top =60
                    Width =420
                    Height =255
                    ColumnWidth =2310
                    TabIndex =6
                    Name ="Species_ID"
                    ControlSource ="Species_ID"
                    StatusBarText ="Unique record identifier - primary key"

                    LayoutCachedLeft =165
                    LayoutCachedTop =60
                    LayoutCachedWidth =585
                    LayoutCachedHeight =315
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =705
                    Top =60
                    Width =360
                    Height =255
                    ColumnWidth =2310
                    TabIndex =7
                    Name ="Transect_ID"
                    ControlSource ="Transect_ID"
                    StatusBarText ="Foreign key to tbl_Quadrat_Transect"

                    LayoutCachedLeft =705
                    LayoutCachedTop =60
                    LayoutCachedWidth =1065
                    LayoutCachedHeight =315
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =7005
                    Top =60
                    Width =705
                    Height =300
                    TabIndex =5
                    ForeColor =255
                    Name ="btnDelete"
                    Caption ="Delete"

                    LayoutCachedLeft =7005
                    LayoutCachedTop =60
                    LayoutCachedWidth =7710
                    LayoutCachedHeight =360
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =21
                    Left =3105
                    Top =60
                    Width =900
                    TabIndex =2
                    BackColor =62207
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    ConditionalFormat = Begin
                        0x010000009c000000030000000100000000000000000000000c00000000000000 ,
                        0x00000000ffffff0001000000000000000d000000190000000000000000000000 ,
                        0xffffff0000000000040000001a0000001d0000000100000000000000ffffff00 ,
                        0x5b0074006200780049005300510031005d003d003000000000005b0074006200 ,
                        0x78004e004500510031005d003d003100000000002700270000000000
                    End
                    Name ="Q1_hm"
                    ControlSource ="Q1_0m"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Percent cover Q1 @ 0.5m"
                    BeforeUpdate ="[Event Procedure]"

                    LayoutCachedLeft =3105
                    LayoutCachedTop =60
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000400000001000000000000000000000000000000ffffff000b0000005b00 ,
                        0x74006200780049005300510031005d003d003000000000000000000000000000 ,
                        0x00000000000000000001000000000000000000000000000000ffffff000b0000 ,
                        0x005b007400620078004e004500510031005d003d003100000000000000000000 ,
                        0x00000000000000000000000000000000040000000100000000000000ffffff00 ,
                        0x0200000027002700000000000000000000000000000000000000000000010000 ,
                        0x00000000000100000000000000fff200000e0000004c0065006e0028005b0051 ,
                        0x0031005f0030006d005d0029003d003000000000000000000000000000000000 ,
                        0x000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =21
                    Left =4005
                    Top =60
                    Width =900
                    TabIndex =3
                    BackColor =62207
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    ConditionalFormat = Begin
                        0x010000009c000000030000000100000000000000000000000c00000000000000 ,
                        0x00000000ffffff0001000000000000000d000000190000000000000000000000 ,
                        0xffffff0000000000040000001a0000001d0000000100000000000000ffffff00 ,
                        0x5b0074006200780049005300510032005d003d003000000000005b0074006200 ,
                        0x78004e004500510032005d003d003100000000002700270000000000
                    End
                    Name ="Q2_5m"
                    ControlSource ="Q2_4_5m"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Percent cover Q2 @ 4.5m"
                    BeforeUpdate ="[Event Procedure]"

                    LayoutCachedLeft =4005
                    LayoutCachedTop =60
                    LayoutCachedWidth =4905
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000400000001000000000000000000000000000000ffffff000b0000005b00 ,
                        0x74006200780049005300510032005d003d003000000000000000000000000000 ,
                        0x00000000000000000001000000000000000000000000000000ffffff000b0000 ,
                        0x005b007400620078004e004500510032005d003d003100000000000000000000 ,
                        0x00000000000000000000000000000000040000000100000000000000ffffff00 ,
                        0x0200000027002700000000000000000000000000000000000000000000010000 ,
                        0x00000000000100000000000000fff200000e0000004c0065006e0028005b0051 ,
                        0x0032005f0035006d005d0029003d003000000000000000000000000000000000 ,
                        0x000000000000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =21
                    Left =4905
                    Top =60
                    Width =900
                    TabIndex =4
                    BackColor =62207
                    ColumnInfo ="\"\";\"\";\"6\";\"4\""
                    ConditionalFormat = Begin
                        0x010000009c000000030000000100000000000000000000000c00000000000000 ,
                        0x00000000ffffff0001000000000000000d000000190000000000000000000000 ,
                        0xffffff0000000000040000001a0000001d0000000100000000000000ffffff00 ,
                        0x5b0074006200780049005300510033005d003d003000000000005b0074006200 ,
                        0x78004e004500510033005d003d003100000000002700270000000000
                    End
                    Name ="Q3_10m"
                    ControlSource ="Q3_9_5m"
                    RowSourceType ="Table/Query"
                    RowSource ="tlu_Cover_Code"
                    StatusBarText ="Percent cover Q3 @ 9.5m"
                    BeforeUpdate ="[Event Procedure]"

                    LayoutCachedLeft =4905
                    LayoutCachedTop =60
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000400000001000000000000000000000000000000ffffff000b0000005b00 ,
                        0x74006200780049005300510033005d003d003000000000000000000000000000 ,
                        0x00000000000000000001000000000000000000000000000000ffffff000b0000 ,
                        0x005b007400620078004e004500510033005d003d003100000000000000000000 ,
                        0x00000000000000000000000000000000040000000100000000000000ffffff00 ,
                        0x0200000027002700000000000000000000000000000000000000000000010000 ,
                        0x00000000000100000000000000fff200000f0000004c0065006e0028005b0051 ,
                        0x0033005f00310030006d005d0029003d00300000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin ComboBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =4320
                    Left =165
                    Top =60
                    Width =1860
                    BackColor =62207
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"50\""
                    ConditionalFormat = Begin
                        0x01000000a0000000030000000100000000000000000000000d00000001000000 ,
                        0x00000000ffffff0001000000000000000e0000001b0000000100000000000000 ,
                        0xffffff0000000000040000001c0000001f0000000100000000000000ffffff00 ,
                        0x5b00740062007800530075006d00490053005d003e003000000000005b007400 ,
                        0x62007800530075006d004e0045005d003c003300000000002700270000000000
                    End
                    Name ="Plant_Code"
                    ControlSource ="PlantCode"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qry_sel_Species_Lookup.Master_PLANT_Code, qry_sel_Species_Lookup.LU_Code,"
                        " qry_sel_Species_Lookup.Utah_Species FROM qry_sel_Species_Lookup;"
                    ColumnWidths ="0;1728;2592"
                    BeforeUpdate ="[Event Procedure]"

                    LayoutCachedLeft =165
                    LayoutCachedTop =60
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000600000001000000000000000100000000000000ffffff000c0000005b00 ,
                        0x740062007800530075006d00490053005d003e00300000000000000000000000 ,
                        0x000000000000000000000001000000000000000100000000000000ffffff000c ,
                        0x0000005b00740062007800530075006d004e0045005d003c0033000000000000 ,
                        0x0000000000000000000000000000000000000000040000000100000000000000 ,
                        0xffffff0002000000270027000000000000000000000000000000000000000000 ,
                        0x0001000000000000000100000000000000fff20000130000004c0065006e0028 ,
                        0x005b0050006c0061006e0074005f0043006f00640065005d0029003d00300000 ,
                        0x0000000000000000000000000000000000000000010000000000000000000000 ,
                        0x00000000ffffff001500000049004900660028005b0074006200780053007500 ,
                        0x6d00490053005d003d0030002c0031002c003000290000000000000000000000 ,
                        0x000000000000000000000001000000000000000000000000000000ffffff0015 ,
                        0x00000049004900660028005b00740062007800530075006d004e0045005d003c ,
                        0x0033002c0031002c003000290000000000000000000000000000000000000000 ,
                        0x0000
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =21
                    Left =2100
                    Top =60
                    Width =900
                    TabIndex =1
                    BackColor =62207
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"510\""
                    ConditionalFormat = Begin
                        0x01000000a0000000030000000100000000000000000000000d00000001000000 ,
                        0x00000000ffffff0001000000000000000e0000001b0000000100000000000000 ,
                        0xffffff0000000000040000001c0000001f0000000100000000000000ffffff00 ,
                        0x5b00740062007800530075006d00490053005d003e003000000000005b007400 ,
                        0x62007800530075006d004e0045005d003c003300000000002200220000000000
                    End
                    Name ="cbxIsDead"
                    ControlSource ="IsDead"
                    RowSourceType ="Table/Query"
                    RowSource ="IsDead_Plus_Flags"
                    ColumnWidths ="0;1440"
                    BeforeUpdate ="[Event Procedure]"
                    ControlTipText ="Indicate if species is alive or dead (or the appropriate missing data flag)"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =300
                    ConditionalFormat14 = Begin
                        0x01000400000001000000000000000100000000000000ffffff000c0000005b00 ,
                        0x740062007800530075006d00490053005d003e00300000000000000000000000 ,
                        0x000000000000000000000001000000000000000100000000000000ffffff000c ,
                        0x0000005b00740062007800530075006d004e0045005d003c0033000000000000 ,
                        0x0000000000000000000000000000000000000000040000000100000000000000 ,
                        0xffffff0002000000220022000000000000000000000000000000000000000000 ,
                        0x0001000000000000000100000000000000fff20000120000004c0065006e0028 ,
                        0x005b006300620078004900730044006500610064005d0029003d003000000000 ,
                        0x000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5985
                    Top =60
                    Width =540
                    Height =255
                    TabIndex =8
                    Name ="tbxAvgCover"
                    ControlSource ="=IIf([tbxSumSampled]>0,[tbxSumCover]/[tbxSumSampled],0)"
                    StatusBarText ="Percent cover in 10 m2 quadrat"
                    ConditionalFormat = Begin
                        0x01000000b0000000020000000100000000000000000000001200000000000000 ,
                        0x00000000ffffff00010000000000000013000000270000000000000000000000 ,
                        0xffffff0000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00740062007800530075006d00530061006d0070006c00650064005d003d00 ,
                        0x3000000000005b007400620078004e006f00450078006f007400690063007300 ,
                        0x530075006d005d003d00330000000000
                    End

                    LayoutCachedLeft =5985
                    LayoutCachedTop =60
                    LayoutCachedWidth =6525
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x01000200000001000000000000000000000000000000ffffff00110000005b00 ,
                        0x740062007800530075006d00530061006d0070006c00650064005d003d003000 ,
                        0x0000000000000000000000000000000000000000000100000000000000000000 ,
                        0x0000000000ffffff00130000005b007400620078004e006f00450078006f0074 ,
                        0x00690063007300530075006d005d003d00330000000000000000000000000000 ,
                        0x0000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3120
                    Top =420
                    Width =660
                    TabIndex =9
                    ForeColor =8355711
                    Name ="tbxSpeciesCoverID_Q1"
                    ControlSource ="SpeciesCoverID_Q1"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =3120
                    LayoutCachedTop =420
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4020
                    Top =420
                    Width =660
                    TabIndex =10
                    ForeColor =8355711
                    Name ="tbxSpeciesCoverID_Q2"
                    ControlSource ="SpeciesCoverID_Q2"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =4020
                    LayoutCachedTop =420
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4980
                    Top =420
                    Width =660
                    TabIndex =11
                    ForeColor =8355711
                    Name ="tbxSpeciesCoverID_Q3"
                    ControlSource ="SpeciesCoverID_Q3"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =4980
                    LayoutCachedTop =420
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =660
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =420
                    Top =390
                    Width =660
                    TabIndex =12
                    ForeColor =8355711
                    Name ="tbxDupe"
                    ControlSource ="=IIf((Len([Plant_Code])>0) And (Len([cbxIsDead])>0),DCount(\"*\",\"TransectSpeci"
                        "esCover\",\"Event_ID = '\" & [Forms]![frm_Data_Entry]![tbxEventID] & \"' AND  Tr"
                        "ansect_ID = '\" & [Transect_ID] & \"' AND PlantCode =  '\" & [Plant_Code] & \"' "
                        "AND IsDead =  \" & [cbxIsDead]),0)"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =420
                    LayoutCachedTop =390
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =630
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =1020
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3000
                    Top =60
                    Width =900
                    Height =255
                    ForeColor =8355711
                    Name ="tbxQ1_Sampled"
                    ControlSource ="=Count(IIf(Len([Q1_0m])>0,1,Null))"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =3000
                    LayoutCachedTop =60
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3945
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =1
                    ForeColor =8355711
                    Name ="tbxQ2_Sampled"
                    ControlSource ="=Count(IIf(Len([Q2_4_5m])>0,1,Null))"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =3945
                    LayoutCachedTop =60
                    LayoutCachedWidth =4845
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ffffff00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4905
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =2
                    ForeColor =8355711
                    Name ="tbxQ3_Sampled"
                    ControlSource ="=Count(IIf(Len([Q3_9_5m])>0,1,Null))"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =4905
                    LayoutCachedTop =60
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ffffff00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5925
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =3
                    ForeColor =8355711
                    Name ="tbxSumSampled"
                    ControlSource ="=[tbxISQ1]+[tbxISQ2]+[tbxISQ3]"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =5925
                    LayoutCachedTop =60
                    LayoutCachedWidth =6825
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ffffff00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =6885
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =4
                    ForeColor =8355711
                    Name ="tbxSumCover"
                    ControlSource ="=IIf([Q1_hm]>0,[Q1_hm],0)+IIf([Q2_4_5m]>0,[Q2_4_5m],0)+IIf([Q3_9_5m]>0,[Q3_9_5m]"
                        ",0)"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =6885
                    LayoutCachedTop =60
                    LayoutCachedWidth =7785
                    LayoutCachedHeight =315
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ffffff00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =285
                    Top =60
                    Width =900
                    Height =255
                    TabIndex =5
                    ForeColor =12566463
                    Name ="tbxDevMode"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0xececec00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End

                    LayoutCachedLeft =285
                    LayoutCachedTop =60
                    LayoutCachedWidth =1185
                    LayoutCachedHeight =315
                    ConditionalFormat14 = Begin
                        0x010001000000000000000200000001000000ececec00ffffff00050000004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3000
                    Top =375
                    Width =900
                    Height =255
                    TabIndex =6
                    ForeColor =8355711
                    Name ="tbxISQ1"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =3000
                    LayoutCachedTop =375
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =630
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3000
                    Top =675
                    Width =900
                    Height =255
                    TabIndex =7
                    ForeColor =8355711
                    Name ="tbxNEQ1"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =3000
                    LayoutCachedTop =675
                    LayoutCachedWidth =3900
                    LayoutCachedHeight =930
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =360
                    Width =900
                    Height =255
                    TabIndex =8
                    ForeColor =8355711
                    Name ="tbxISQ2"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =3960
                    LayoutCachedTop =360
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =615
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =660
                    Width =900
                    Height =255
                    TabIndex =9
                    ForeColor =8355711
                    Name ="tbxNEQ2"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =3960
                    LayoutCachedTop =660
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4920
                    Top =360
                    Width =900
                    Height =255
                    TabIndex =10
                    ForeColor =8355711
                    Name ="tbxISQ3"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =4920
                    LayoutCachedTop =360
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =615
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4920
                    Top =660
                    Width =900
                    Height =255
                    TabIndex =11
                    ForeColor =8355711
                    Name ="tbxNEQ3"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =4920
                    LayoutCachedTop =660
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =915
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =375
                    Width =900
                    Height =255
                    TabIndex =12
                    ForeColor =8355711
                    Name ="tbxISSum"
                    ControlSource ="=[tbxISQ1]+[tbxISQ2]+[tbxISQ3]"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =1980
                    LayoutCachedTop =375
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =630
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1980
                    Top =675
                    Width =900
                    Height =255
                    TabIndex =13
                    ForeColor =8355711
                    Name ="tbxNESum"
                    ControlSource ="=[tbxNEQ1]+[tbxNEQ2]+[tbxNEQ3]"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0xececec00ececec00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004400650076004d006f00640065005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =1980
                    LayoutCachedTop =675
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =930
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ececec00ececec00120000005b00 ,
                        0x7400620078004400650076004d006f00640065005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
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
' Form:         fsub_Species
' Level:        Application form
' Version:      1.11
' Basis:        -
'
' Description:  Species subform object related properties, functions & procedures for UI display
'
' Source/date:  Russ DenBleyker, Unknown - for NCPN tools
' References:   -
' Revisions:    RDB - Unknown   - 1.00 - initial version
'               BLC - 3/8/2017  - 1.01 - added documentation, error handling
'               BLC - 4/21/2017 - 1.02 - added HasRecords, ParentForm properties
'               BLC - 7/5/2017  - 1.03 - removed warnings for deleting record
'               BLC - 7/12/2017 - 1.04 - replaced CalcAvgCover w/ refresh of tbxAvgCover
'                                        which calculates average cover on-the-fly based on
'                                        IsSampled_Q1-3 and Q1_hm, Q2_5m, Q3_10m values
'               BLC - 7/17/2017 - 1.05 - set controls enabled by default
'               BLC - 7/18/2017 - 1.06 - revised for species cover updates/deletes
'               BLC - 7/19/2017 - 1.07 - removed CalcAverageCover(), ParentForm_Current() & other cleanup
'               BLC - 7/24/2017 - 1.08 - revised btnDelete_Click to properly delete from db & usys_temp_speciescover
'               BLC - 7/27/2017 - 1.09 - update average cover after setting species cover, revised IsDuplicateSpeciesCover
'               BLC - 7/28/2017 - 1.10 - code cleanup
'               BLC - 7/31/2017 - 1.11 - code cleanup & fix tab navigation after requery
'                                        when user updates species cover
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
'Private WithEvents m_ParentForm As Form 'Form_frm_Quadrat_Transect
Private m_HasRecords As Boolean
Private m_HasRecordsQ1 As Boolean
Private m_HasRecordsQ2 As Boolean
Private m_HasRecordsQ3 As Boolean

'---------------------
' Event Declarations
'---------------------
Public Event InvalidHasRecords(Value As Boolean)
Public Event InvalidHasRecordsQ1(Value As Boolean)
Public Event InvalidHasRecordsQ2(Value As Boolean)
Public Event InvalidHasRecordsQ3(Value As Boolean)

'---------------------
' Properties
'---------------------
Public Property Let HasRecords(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecords = Value
    Else
        RaiseEvent InvalidHasRecords(Value)
    End If
End Property

Public Property Get HasRecords() As Boolean
    HasRecords = m_HasRecords
End Property

Public Property Let HasRecordsQ1(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecordsQ1 = Value
    Else
        RaiseEvent InvalidHasRecordsQ1(Value)
    End If
End Property

Public Property Get HasRecordsQ1() As Boolean
    HasRecordsQ1 = m_HasRecordsQ1
End Property

Public Property Let HasRecordsQ2(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecordsQ2 = Value
    Else
        RaiseEvent InvalidHasRecordsQ2(Value)
    End If
End Property

Public Property Get HasRecordsQ2() As Boolean
    HasRecordsQ2 = m_HasRecordsQ2
End Property

Public Property Let HasRecordsQ3(Value As Boolean)
    If varType(Value) = vbBoolean Then
        m_HasRecordsQ3 = Value
    Else
        RaiseEvent InvalidHasRecordsQ3(Value)
    End If
End Property

Public Property Get HasRecordsQ3() As Boolean
    HasRecordsQ3 = m_HasRecordsQ3
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 4/21/2017 - added setting HasRecordsQ1-3 properties
'   BLC - 7/17/2017 - set controls enabled by default
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'set fields enabled by default
    Me.Plant_Code.Enabled = True
    Me.Q1_hm.Enabled = True
    Me.Q2_5m.Enabled = True
    Me.Q3_10m.Enabled = True
    Me.cbxIsDead.Enabled = True
    
    'defaults
    Me.HasRecords = False
    Me.HasRecordsQ1 = False
    Me.HasRecordsQ2 = False
    Me.HasRecordsQ3 = False

    'determine if Q1-3 have records
    If Me.Form.Recordset.RecordCount > 0 And Not IsNull(Me.Plant_Code) Then Me.HasRecords = True
    
    'hide dev mode so it doesn't flash w/ @ transect
    If Not DEV_MODE Then Me.tbxDevMode.visible = False
    
    'set dev mode
    Me.tbxDevMode = DEV_MODE

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 21, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/21/2017 - initial version
' ---------------------------------
Private Sub Form_Current()
    On Error GoTo Err_Handler

    'defaults
    HasRecords = False
    HasRecordsQ1 = False
    HasRecordsQ2 = False
    HasRecordsQ3 = False
    
    If Me.Form.Recordset.RecordCount > 0 And Not IsNull(Me.Plant_Code) Then _
        Me.HasRecords = True
    
    'determine if any Q1-3 has values
    ' NOTE: must use Me.Controls("XX") to handle controls w/ underscore
    'Debug.Print "Q1_hm: " & Me.Controls("Q1_hm")
    If Not IsNull(Me.Controls("Q1_hm")) Then
        HasRecordsQ1 = True
    End If

    'Debug.Print "Q2_5m: " & Me.Controls("Q2_5m")
    If Not IsNull(Me.Controls("Q2_5m")) Then
        HasRecordsQ2 = True
    End If

    'Debug.Print "Q3_10m: " & Me.Controls("Q3_10m")
    If Not IsNull(Me.Controls("Q3_10m")) Then
        HasRecordsQ3 = True
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Current[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_BeforeInsert
' Description:  form before insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 7/18/2017 - revised for new tables
'   BLC - 7/19/2017 - removed GUID creation, no longer necessary for Species_ID
'                     since species cover uses a numeric (long) ID instead
' ---------------------------------
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo Err_Handler

    If IsNull(Me.Parent!Observer) Then
      MsgBox "You must enter Observer first."
      DoCmd.CancelEvent
      SendKeys "{ESC}"
      GoTo Exit_Handler
    End If
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeInsert[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_AfterInsert
' Description:  Form after insert actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  BLC, July 18, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/18/2017 - initial version
' ---------------------------------
Private Sub Form_AfterInsert()
    On Error GoTo Err_Handler

    Debug.Print "form_afterinsert"

    'check if duplicate species cover (skip the warning)
    'IsDuplicateSpeciesCover


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_AfterInsert[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Form_AfterUpdate
' Description:  Form after Update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  BLC, July 25, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/25/2017 - initial version
' ---------------------------------
Private Sub Form_AfterUpdate()
    On Error GoTo Err_Handler

    Debug.Print "form_afterUpdate"

    'check if duplicate species cover (skip the warning)
    'IsDuplicateSpeciesCover

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_AfterUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' FIX - MISSING CLASS
'' ---------------------------------
'' Sub:          btnDelete_Click
'' Description:  Delete button actions
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  NCPN, Unknown - for NCPN tools
'' Adapted:      -
'' Revisions:
''   NCPN - Unknown - initial version
''   BLC - 3/8/2017 - added documentation, error handling
''   BLC - 7/5/2017 - toggled warnings off before removing record to avoid 2nd "do you want to
''                    do this?" dialog
''   BLC - 7/18/2017 - revised for deletions from SpeciesCover
''   BLC - 7/24/2017 - revised to delete from usys_temp_speciescover before db
''                     to avoid passing NULL as SpeciesCoverID to
''                     InvasiveCoverSpecies.DeleteSpeciesCover
'' ---------------------------------
'Private Sub btnDelete_Click()
'On Error GoTo Err_Handler
'
'  Dim Reply As Integer
'  Reply = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Species Delete")
'
'  'do the deletion for temp & underlying tables
'  If Reply = 6 Then
'
'    'handle the db delete FIRST since doing it second results in
'    'NULL for all controls if usys_temp_speciescover record is deleted first
'
'    'do for @ Quadrat that has a SpeciesCover record
'    Dim i As Integer
'    Dim strControl As String
'
'    For i = 1 To QUADRATS_PER_TRANSECT
'
'        strControl = "tbxSpeciesCoverID_Q" & i
'
'        'only delete existing records (these should have a SpeciesCover ID value)
'        If Me.Controls(strControl) > 0 Then
'
'            Dim sp As New InvasiveCoverSpecies
'
'            With sp
'                .SpeciesCoverID = Me.Controls(strControl)
'
'                .DeleteSpeciesCover
'
'            End With
'
'        End If
'
'        Next
'
'        'RefreshTempTable "usys_temp_speciescover" << ERROR #3211
'        'cannot refresh now because usys_temp_speciescover is locked by form
'
'    'delete the record in usys_temp_speciescover
'    DoCmd.SetWarnings False
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 8, , acMenuVer70
'    DoCmd.DoMenuItem acFormBar, acEditMenu, 6, , acMenuVer70
'    DoCmd.SetWarnings True
'
'  End If
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - btnDelete_Click[fsub_Species form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Sub:          Plant_Code_BeforeUdpate
' Description:  Plant_Code combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 7/17/2017 - revised for normalized tables & new form fields
'   BLC - 7/27/2017 - revised duplicate check
'   BLC - 7/31/2017 - code cleanup
' ---------------------------------
Private Sub Plant_Code_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
  
    'check for duplicate species when IsDead is set
    If Not IsNull(cbxIsDead) Then
    
        IsDuplicateSpeciesCover
    
    End If
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Plant_Code_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cbxIsDead_BeforeUdpate
' Description:  IsDead combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  BLC, 7/18/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/18/2017 - initial version
'   BLC - 7/27/2017 - revised duplicate check
'   BLC - 7/31/2017 - code cleanup
' ---------------------------------
Private Sub cbxIsDead_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler
    
    'ensure there isn't a dupe of Species + IsDead
    If Len(Me.PlantCode) > 0 Then

        IsDuplicateSpeciesCover

    End If
  
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxIsDead_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q1_hm_BeforeUpdate
' Description:  Q1_hm combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 7/18/2017 - revise to check for both species & is dead flag
'   BLC - 7/31/2017 - code cleanup
' ---------------------------------
Private Sub Q1_hm_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q1_hm_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q2_5m_BeforeUpdate
' Description:  Q2_5m combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 7/18/2017 - revise to check for both species & is dead flag
'   BLC - 7/31/2017 - code cleanup
' ---------------------------------
Private Sub Q2_5m_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q2_5m_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Q3_10m_BeforeUpdate
' Description:  Q3_10m combobox actions before update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  NCPN, Unknown - for NCPN tools
' Adapted:      -
' Revisions:
'   NCPN - Unknown - initial version
'   BLC - 3/8/2017 - added documentation, error handling
'   BLC - 7/18/2017 - revise to check for both species & is dead flag
'   BLC - 7/31/2017 - code cleanup
' ---------------------------------
Private Sub Q3_10m_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Q3_10m_BeforeUpdate[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Sub
' FIX - MISSING CLASS
'' ---------------------------------
'' Sub:          Plant_Code_AfterUdpate
'' Description:  Plant_Code combobox actions after update
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  BLC, 7/18/2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 7/18/2017 - initial version
''   BLC - 7/24/2017 - revised to check for dupe species + IsDead
''   BLC - 7/31/2017 - code cleanup
'' ---------------------------------
'Private Sub Plant_Code_AfterUpdate()
'On Error GoTo Err_Handler
'
'    'ensure there isn't a dupe of Species + IsDead
'    If Len(Me.cbxIsDead) > 0 Then
'
'        SetSpeciesCover
'
'    End If
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - Plant_Code_AfterUpdate[fsub_Species form])"
'    End Select
'    Resume Exit_Handler
'End Sub
' FIX - MISSING CLASS
'' ---------------------------------
'' Sub:          cbxIsDead_AfterUdpate
'' Description:  IsDead combobox actions after update
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  BLC, 7/18/2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 7/18/2017 - initial version
'' ---------------------------------
'Private Sub cbxIsDead_AfterUpdate()
'On Error GoTo Err_Handler
'
'    SetSpeciesCover
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - cbxIsDead_AfterUpdate[fsub_Species form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' FIX - MISSING CLASS
'' ---------------------------------
'' Sub:          Q1_hm_AfterUpdate
'' Description:  Q1_hm combobox actions after update
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  NCPN, Unknown - for NCPN tools
'' Adapted:      -
'' Revisions:
''   NCPN - Unknown - initial version
''   BLC - 3/8/2017 - added documentation, error handling
''   BLC - 7/12/2017 - replaced CalcAvgCover w/ refresh of tbxAvgCover
'' ---------------------------------
'Private Sub Q1_hm_AfterUpdate()
'On Error GoTo Err_Handler
'
'    SetSpeciesCover
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - Q1_hm_AfterUpdate[fsub_Species form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' FIX - MISSING CLASS
'' ---------------------------------
'' Sub:          Q2_5m_AfterUpdate
'' Description:  Q2_5m combobox actions after update
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  NCPN, Unknown - for NCPN tools
'' Adapted:      -
'' Revisions:
''   NCPN - Unknown - initial version
''   BLC - 3/8/2017 - added documentation, error handling
''   BLC - 7/12/2017 - replaced CalcAvgCover w/ refresh of tbxAvgCover
'' ---------------------------------
'Private Sub Q2_5m_AfterUpdate()
'On Error GoTo Err_Handler
'
'    SetSpeciesCover
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - Q2_5m_AfterUpdate[fsub_Species form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' FIX - MISSING CLASS
'' ---------------------------------
'' Sub:          Q3_10m_AfterUpdate
'' Description:  Q3_10m combobox actions after update
'' Assumptions:  -
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  NCPN, Unknown - for NCPN tools
'' Adapted:      -
'' Revisions:
''   NCPN - Unknown - initial version
''   BLC - 3/8/2017 - added documentation, error handling
''   BLC - 7/12/2017 - replaced CalcAvgCover w/ refresh of tbxAvgCover
'' ---------------------------------
'Private Sub Q3_10m_AfterUpdate()
'On Error GoTo Err_Handler
'
'    SetSpeciesCover
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - Q3_10m_AfterUpdate[fsub_Species form])"
'    End Select
'    Resume Exit_Handler
'End Sub

' ---------------------------------
' Sub:          AddNewTransectQuadrats
' Description:  Add quadrat records for new transect
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  BLC, 7/5/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/5/2017 - initial version
'   BLC - 7/18/2017 - replace 3 w/ QUADRATS_PER_TRANSECT
' ---------------------------------
Public Function AddNewTransectQuadrats() As Single
    On Error GoTo Err_Handler
    
    Dim AvgCover As Single
    Dim TotalCover As Single
    Dim Count As Integer, i As Integer
    Dim strControl As String, strPosition As String
   
    Count = 0
    AvgCover = 0
    TotalCover = 0
    
    For i = 1 To QUADRATS_PER_TRANSECT
        'determine quadrat control
        Select Case i
            Case 1
                strPosition = "h"
            Case 2
                strPosition = "5"
            Case 3
                strPosition = "10"
        End Select
    
        strControl = "Q" & i & "_" & strPosition & "m"
    
        If Me.Controls(strControl).Enabled Then
            If Not IsNull(Me.Controls(strControl)) Then
                TotalCover = TotalCover + Me.Controls(strControl)
                Count = Count + 1
            End If
        End If
    Next
    
    If Count > 0 Then
        'calculate the average
        AvgCover = TotalCover / Count
    End If

    AddNewTransectQuadrats = AvgCover

Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddNewTransectQuadrats[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     IsDuplicateSpeciesCover
' Description:  Checks if species cover record exists for same species & is dead flag
' Assumptions:  -
' Parameters:   SkipWarning - whether to skip plant species/is dead warning (boolean, optional,
'                             default = False)
' Returns:      True or False depending on whether a record is found in SpeciesCover
'               for the same species and is dead flag setting
' Throws:       none
' References:   -
' Source/date:  BLC, 7/18/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/18/2017 - initial version
'   BLC - 7/27/2017 - revised to use query vs. DLookup for duplicates
'                     dupe = same species + is dead for same sampling event on transect
'                     so event & transect must be included
'   BLC - 7/28/2017 - code cleanup
' ---------------------------------
Private Function IsDuplicateSpeciesCover(Optional SkipWarning As Boolean = False) As Boolean
On Error GoTo Err_Handler
      
    'ensure plant code & is dead flag are set
    If IsNull(Me!Plant_Code) Or IsNull(Me!cbxIsDead) Then
        
        If Not (SkipWarning) Then
            MsgBox "Enter the species and set the IsDead flag first."
            
            DoCmd.CancelEvent
            SendKeys "{ESC}"
        End If
    
    End If
    
    Dim isDupe As Boolean
      
    'default
    isDupe = False
      
    'check when plant code & is dead are set only
    If Not IsNull(Me!Plant_Code) And Not IsNull(Me.cbxIsDead) Then
        
        'check if *any* of the Quadrats have this species in SpeciesCover
        Dim i, IsDead As Integer
        Dim strCriteria As String
            IsDead = IIf(Me.cbxIsDead = "Dead", 1, 0)

            'determine if duplicates exist
            Dim NumRecords As Integer
            Dim Template As String
            
            Template = "s_speciescover_dupes"
            
            Dim Params(0 To 4) As Variant
        
            With Me
                Params(0) = "TransectSpeciesCover"
                Params(1) = Me.Parent.Parent.tbxEventID
                Params(2) = Me.Parent.tbxTransectID
                Params(3) = Me.Plant_Code
                Params(4) = Me.cbxIsDead
                
                'retrieve the first record returned
                NumRecords = GetRecords(Template, Params).Fields(0)
                
Debug.Print "Plant - IsDead - NumRecords: " & Me.Plant_Code & " - " & Me.cbxIsDead & " - " & NumRecords
            
            End With

            If NumRecords > 1 Then
             
                isDupe = True
                 
                 If Not SkipWarning Then
                    MsgBox "Duplicate species dead/alive for this transect."
                    
                    DoCmd.CancelEvent
                    SendKeys "{ESC}"
                 End If
            End If
    
    End If
  
    IsDuplicateSpeciesCover = isDupe
  
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsDuplicateSpeciesCover[fsub_Species form])"
    End Select
    Resume Exit_Handler
End Function

' FIX - MISSING CLASS
'' ---------------------------------
'' Sub:          SetSpeciesCover
'' Description:  Add or update species cover record
'' Assumptions:  Quadrat species cover IDs will be:
''                NULL - when no species cover record exists
''                > 0  - when the species cover record exists
''               for the selected species, is dead flag, quadrat combination
''               enabled Q1-3 percent cover values indicate if record should
''               be added/updated
'' Parameters:   -
'' Returns:      -
'' Throws:       none
'' References:   -
'' Source/date:  BLC, 7/18/2017 - for NCPN tools
'' Adapted:      -
'' Revisions:
''   BLC - 7/18/2017 - initial version
''   BLC - 7/27/2017 - update average cover after setting species cover,
''                     do not re-check IsDuplicateSpeciesCover since that is
''                     checked already
''   BLC - 7/28/2017 - code cleanup
''   BLC - 7/31/2017 - fix so returns to user's record after requery (refresh instead)
'' ---------------------------------
'Private Sub SetSpeciesCover()
'On Error GoTo Err_Handler
'
'    'fetch the current record
'    Dim lngRecordNum As Long
'
'    lngRecordNum = Me.CurrentRecord
'
'    'check if plant code is set
'    If IsNull(Me.Plant_Code) = True Then GoTo Exit_Handler
'
'    Dim i As Integer
'    Dim strControl As String
'    Dim sc As New InvasiveCoverSpecies
'    Dim CoverID As Integer
'    Dim Pct As Double
'    Dim QID As Long
'    Dim SkipQuadrat As Boolean
'
'    'iterate through quadrats
'    For i = 1 To QUADRATS_PER_TRANSECT
'
'        SkipQuadrat = False
'
'        strControl = "tbxSpeciesCoverID_Q" & i
'
'        ' do the update/add
'        With sc
'
'            Select Case i
'                Case 1
'                    If Me.tbxISQ1 = 0 Or Me.tbxNEQ1 = 1 Then
'                        SkipQuadrat = True
'                    Else
'                        Pct = Nz(Me.Q1_hm, 0)
'                        CoverID = Nz(Me.tbxSpeciesCoverID_Q1, 0)
'                        QID = Parent.Form.Controls("tbxQ1")
'                    End If
'
'                Case 2
'                    If Me.tbxISQ2 = 0 Or Me.tbxNEQ2 = 1 Then
'                        SkipQuadrat = True
'                    Else
'                        Pct = Nz(Me.Q2_5m, 0)
'                        CoverID = Nz(Me.tbxSpeciesCoverID_Q2, 0)
'                        QID = Parent.Form.Controls("tbxQ2")
'                    End If
'
'                Case 3
'                    If Me.tbxISQ3 = 0 Or Me.tbxNEQ3 = 1 Then
'                        SkipQuadrat = True
'                    Else
'                        Pct = Nz(Me.Q3_10m, 0)
'                        CoverID = Nz(Me.tbxSpeciesCoverID_Q3, 0)
'                        QID = Parent.Form.Controls("tbxQ3")
'                    End If
'
'            End Select
'
'            .QuadratID = QID
'            .LUCode = Me.Plant_Code
'            .IsDead = Nz(Me.cbxIsDead, 0) 'IIf(Me.cbxIsDead = "Dead", 1, 0)
'            .PctCover = Pct
'
'            'take action if quadrat shouldn't be skipped
'            If Not SkipQuadrat Then
'
'                'check if update or new
'                Select Case Me.Controls(strControl)
'
'                    Case Is > 0 'Update
'
'                        .SpeciesCoverID = CoverID
'                        .UpdateSpeciesCover
'
'                    Case Else   'New
'
'                        .AddSpeciesCover
'
'                        'update the values for species cover ID
'                        Me.Controls(strControl) = .SpeciesCoverID
'
'                End Select
'
'                'update the underlying data << not now, it's in use!
'                'RefreshTempTable "usys_temp_speciescover"
'
'            End If
'
'        End With
'
'    Next
'
'    'update average cover
'    'Me.Requery
'    Me.Refresh
'
'    'return to user's selected record (otherwise returns to top of tab order)
'    'DoCmd.GoToRecord acActiveDataObject, Me.Name, acGoTo, lngRecordNum << error 2489 Me.Name not open
'    'Me.SetFocus
'    'DoCmd.GoToRecord , , lngRecordNum << goes to wrong record
'    'DoCmd.GoToRecord acActiveDataObject, , acGoTo, lngRecordNum << error 2449 invalid method in an expression
'
'
'Exit_Handler:
'    Exit Sub
'Err_Handler:
'    Select Case Err.Number
'      Case Else
'        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
'            "Error encountered (#" & Err.Number & " - SetSpeciesCover[fsub_Species form])"
'    End Select
'    Resume Exit_Handler
'End Sub
