Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10741
    DatasheetFontHeight =10
    ItemSuffix =57
    Left =405
    Top =1515
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd62d8ba872b2e340
    End
    RecordSource ="SELECT tlu_Enumerations.Sort_Order, tbl_Quadrat_Data.* FROM tbl_Quadrat_Data INN"
        "ER JOIN tlu_Enumerations ON tbl_Quadrat_Data.Quadrat_Number=tlu_Enumerations.Enu"
        "m_Code WHERE (((tlu_Enumerations.Enum_Group)=\"Quadrat Number\")) ORDER BY tlu_E"
        "numerations.Sort_Order; "
    Caption ="sfrm_Quadrats subreport"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e0010000680100006801000000000000f42900005004000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            KeepTogether =2
            ControlSource ="Sort_Order"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =15
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Name ="Line12"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            BackColor =15132390
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =900
                    Width =1680
                    Height =300
                    FontSize =11
                    FontWeight =700
                    Name ="txtQuadrat_Number"
                    ControlSource ="Quadrat_Number"
                    FontName ="Calibri"

                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =60
                    Width =840
                    Height =300
                    ForeColor =0
                    Name ="lblQuadrat"
                    Caption ="Quadrat"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1104
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5130
                    Width =360
                    FontSize =9
                    Name ="Percent_Trees"
                    ControlSource ="Percent_Trees"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00500065007200630065006e0074005f00 ,
                        0x540072006500650073005d00290000000000
                    End

                    LayoutCachedLeft =5130
                    LayoutCachedWidth =5490
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400170000004900 ,
                        0x73004e0075006c006c0028005b00500065007200630065006e0074005f005400 ,
                        0x72006500650073005d0029000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =10185
                    Top =300
                    Width =360
                    FontSize =9
                    TabIndex =1
                    Name ="txtPercent_Bryophytes"
                    ControlSource ="Percent_Bryophytes"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000a2000000010000000100000000000000000000002000000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004200720079006f007000680079007400650073005d0029000000 ,
                        0x0000
                    End

                    LayoutCachedLeft =10185
                    LayoutCachedTop =300
                    LayoutCachedWidth =10545
                    LayoutCachedHeight =540
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24001f0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004200720079006f007000680079007400650073005d00290000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =6600
                    Width =360
                    FontSize =9
                    TabIndex =2
                    Name ="txtPercent_Rock"
                    ControlSource ="Percent_Rock"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000096000000010000000100000000000000000000001a00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f0052006f0063006b005d00290000000000
                    End

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400190000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f0052006f0063006b005d00290000000000000000000000000000000000 ,
                        0x0000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8040
                    Width =360
                    FontSize =9
                    TabIndex =3
                    Name ="txtPercent_Woody_Debris"
                    ControlSource ="Percent_Woody_Debris"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000a6000000010000000100000000000000000000002200000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f0057006f006f00640079005f004400650062007200690073005d00 ,
                        0x290000000000
                    End

                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400210000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f0057006f006f00640079005f004400650062007200690073005d002900 ,
                        0x000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4380
                    Top =300
                    Width =360
                    FontSize =9
                    TabIndex =4
                    Name ="txtPercent_Grasses"
                    ControlSource ="Percent_Grasses"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000009c000000010000000100000000000000000000001d00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f0047007200610073007300650073005d00290000000000
                    End

                    LayoutCachedLeft =4380
                    LayoutCachedTop =300
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =540
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24001c0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f0047007200610073007300650073005d00290000000000000000000000 ,
                        0x0000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5820
                    Top =300
                    Width =360
                    FontSize =9
                    TabIndex =5
                    Name ="txtPercent_Sedges"
                    ControlSource ="Percent_Sedges"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000009a000000010000000100000000000000000000001c00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f005300650064006700650073005d00290000000000
                    End

                    LayoutCachedLeft =5820
                    LayoutCachedTop =300
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =540
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24001b0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f005300650064006700650073005d002900000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7260
                    Top =300
                    Width =360
                    FontSize =9
                    TabIndex =6
                    Name ="txtPercent_Herbs"
                    ControlSource ="Percent_Herbs"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f00480065007200620073005d00290000000000
                    End

                    LayoutCachedLeft =7260
                    LayoutCachedTop =300
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =540
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24001a0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f00480065007200620073005d0029000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8700
                    Top =300
                    Width =360
                    FontSize =9
                    TabIndex =7
                    Name ="txtPercent_Ferns"
                    ControlSource ="Percent_Ferns"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004600650072006e0073005d00290000000000
                    End

                    LayoutCachedLeft =8700
                    LayoutCachedTop =300
                    LayoutCachedWidth =9060
                    LayoutCachedHeight =540
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24001a0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004600650072006e0073005d0029000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =2040
                    Top =600
                    Width =8580
                    Height =60
                    FontSize =9
                    TabIndex =8
                    Name ="txtQuadrat_Notes"
                    ControlSource ="Quadrat_Notes"
                    FontName ="Calibri"

                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =4050
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label31"
                    Caption ="Trees"
                    FontName ="Calibri"
                    LayoutCachedLeft =4050
                    LayoutCachedWidth =5130
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =9105
                    Top =300
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label32"
                    Caption ="Bryophytes"
                    FontName ="Calibri"
                    LayoutCachedLeft =9105
                    LayoutCachedTop =300
                    LayoutCachedWidth =10185
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =5520
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label33"
                    Caption ="Rock"
                    FontName ="Calibri"
                    LayoutCachedLeft =5520
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =6960
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label34"
                    Caption ="CWD"
                    FontName ="Calibri"
                    LayoutCachedLeft =6960
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =3645
                    Top =300
                    Width =735
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label35"
                    Caption ="Grasses"
                    FontName ="Calibri"
                    LayoutCachedLeft =3645
                    LayoutCachedTop =300
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =4740
                    Top =300
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label36"
                    Caption ="Sedges"
                    FontName ="Calibri"
                    LayoutCachedLeft =4740
                    LayoutCachedTop =300
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =6180
                    Top =300
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label37"
                    Caption ="Herbs"
                    FontName ="Calibri"
                    LayoutCachedLeft =6180
                    LayoutCachedTop =300
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =7620
                    Top =300
                    Width =1080
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label38"
                    Caption ="Ferns"
                    FontName ="Calibri"
                    LayoutCachedLeft =7620
                    LayoutCachedTop =300
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =2040
                    Width =1440
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label39"
                    Caption ="Floor Condition %"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =2040
                    Top =300
                    Width =1440
                    Height =240
                    FontSize =9
                    ForeColor =0
                    Name ="Label40"
                    Caption ="Veg Cover %"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                End
                Begin Subform
                    Left =5400
                    Top =960
                    Width =5341
                    Height =144
                    TabIndex =9
                    Name ="rSub_Event_rSub_Quads_Herbaceous"
                    SourceObject ="Report.rSub_Event_rSub_Quads_Herbaceous"
                    LinkChildFields ="Quadrat_Data_ID"
                    LinkMasterFields ="Quadrat_Data_ID"

                    LayoutCachedLeft =5400
                    LayoutCachedTop =960
                    LayoutCachedWidth =10741
                    LayoutCachedHeight =1104
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =238
                            TextAlign =2
                            TextFontFamily =34
                            Left =5760
                            Top =720
                            Width =3000
                            Height =240
                            FontSize =9
                            FontWeight =400
                            ForeColor =0
                            Name ="lblHerbaceous_Subreport"
                            Caption ="T a r g e t e d   H e r b a c e o u s"
                            FontName ="Calibri"
                            LayoutCachedLeft =5760
                            LayoutCachedTop =720
                            LayoutCachedWidth =8760
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin Subform
                    Top =960
                    Width =5221
                    Height =144
                    TabIndex =10
                    Name ="rSub_Event_rSub_Quads_Seedlings"
                    SourceObject ="Report.rSub_Event_rSub_Quads_Seedlings"
                    LinkChildFields ="Quadrat_Data_ID"
                    LinkMasterFields ="Quadrat_Data_ID"

                    LayoutCachedTop =960
                    LayoutCachedWidth =5221
                    LayoutCachedHeight =1104
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =238
                            TextAlign =2
                            TextFontFamily =34
                            Left =60
                            Top =720
                            Width =4260
                            Height =240
                            FontSize =9
                            FontWeight =400
                            ForeColor =0
                            Name ="lblSeedling_Subreport"
                            Caption ="S e e d l i n g s"
                            FontName ="Calibri"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =10155
                    Width =360
                    FontSize =9
                    TabIndex =11
                    Name ="txtPercent_Other"
                    ControlSource ="Percent_Other"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740050006500720063006500 ,
                        0x6e0074005f004f0074006800650072005d00290000000000
                    End

                    LayoutCachedLeft =10155
                    LayoutCachedWidth =10515
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24001a0000004900 ,
                        0x73004e0075006c006c0028005b00740078007400500065007200630065006e00 ,
                        0x74005f004f0074006800650072005d0029000000000000000000000000000000 ,
                        0x00000000000000
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =9540
                    Width =600
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="lblOther"
                    Caption ="Other"
                    FontName ="Calibri"
                    LayoutCachedLeft =9540
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9000
                    Width =360
                    FontSize =9
                    TabIndex =12
                    Name ="txtPercent_Fine_Woody_Debris"
                    ControlSource ="Percent_Fine_Woody_Debris"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000aa000000010000000100000000000000000000002400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00500065007200630065006e0074005f00 ,
                        0x460069006e0065005f0057006f006f00640079005f0044006500620072006900 ,
                        0x73005d00290000000000
                    End

                    LayoutCachedLeft =9000
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400230000004900 ,
                        0x73004e0075006c006c0028005b00500065007200630065006e0074005f004600 ,
                        0x69006e0065005f0057006f006f00640079005f00440065006200720069007300 ,
                        0x5d002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            TextFontFamily =34
                            Left =8580
                            Width =420
                            Height =240
                            FontSize =9
                            FontWeight =400
                            ForeColor =0
                            Name ="Label55"
                            Caption ="FWD"
                            FontName ="Calibri"
                            LayoutCachedLeft =8580
                            LayoutCachedWidth =9000
                            LayoutCachedHeight =240
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =9000
                    Top =720
                    Width =1095
                    Height =240
                    FontSize =9
                    FontWeight =400
                    ForeColor =0
                    Name ="Label56"
                    Caption ="B r o w s e"
                    FontName ="Calibri"
                    LayoutCachedLeft =9000
                    LayoutCachedTop =720
                    LayoutCachedWidth =10095
                    LayoutCachedHeight =960
                End
            End
        End
        Begin PageFooter
            Height =15
            Name ="PageFooterSection"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =0
                    BorderColor =12632256
                    Name ="Line13"
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
