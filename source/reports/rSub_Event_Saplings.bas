﻿Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =127
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =40
    Top =600
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1a8ba6632e17e540
    End
    RecordSource ="SELECT tbl_Sapling_Data.Event_ID, tbl_Sapling_Data.Tag_ID, tbl_Tags.Location_ID,"
        " tbl_Tags.Tag, tlu_Plants.Latin_Name, [SaplingVigor] & \" \" & [TreeVigorClass] "
        "AS Vig, tbl_Sapling_Data.Sapling_Status, tbl_Sapling_Data.Sapling_Notes, tlu_Pla"
        "nts.Shrub, tbl_Sapling_Data.DRC, tbl_Tags.Microplot_Number, qCalc_Basal_Area_per"
        "_Sapling.Stems, tbl_Sapling_Data.Sapling_Data_ID, tbl_Sapling_Data.Habit, tbl_Sa"
        "pling_Data.Browsed, tbl_Sapling_Data.Browsable, qCalc_Basal_Area_per_Sapling.Equ"
        "iv_DBH_cm, Len(qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm) AS len_equiv_dbh, Type"
        "Name(qCalc_Basal_Area_per_Sapling.Equiv_DBH_cm) AS typename_eqiv_dbh, Left(qCalc"
        "_Basal_Area_per_Sapling.Equiv_DBH_cm,4) AS left4_equiv_dbh, CStr(qCalc_Basal_Are"
        "a_per_Sapling.Equiv_DBH_cm) AS cstr_equiv_dbh FROM (((tbl_Sapling_Data LEFT JOIN"
        " tbl_Tags ON tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID) LEFT JOIN qCalc_Basal_Ar"
        "ea_per_Sapling ON tbl_Sapling_Data.Sapling_Data_ID = qCalc_Basal_Area_per_Saplin"
        "g.Sapling_Data_ID) LEFT JOIN tlu_Plants ON tbl_Tags.TSN = tlu_Plants.TSN) LEFT J"
        "OIN tluTreeVigor ON tbl_Sapling_Data.SaplingVigor = tluTreeVigor.TreeVigorCode O"
        "RDER BY tbl_Tags.Tag;"
    Caption ="srpt_Microplots"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e0010000680100006801000000000000f42900001c02000001000000 ,
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
            ControlSource ="Microplot_Number"
        End
        Begin BreakLevel
            ControlSource ="Tag"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =660
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =180
                    Top =240
                    Width =735
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblTag"
                    Caption ="Tag"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =180
                    LayoutCachedTop =240
                    LayoutCachedWidth =915
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =975
                    Top =240
                    Width =1845
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblLatin_Name"
                    Caption ="Latin Name"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =975
                    LayoutCachedTop =240
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =5100
                    Top =240
                    Width =720
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStems"
                    Caption ="Stems"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5100
                    LayoutCachedTop =240
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =6960
                    Top =60
                    Width =930
                    Height =480
                    FontSize =10
                    ForeColor =0
                    Name ="lblSum_basal_area"
                    Caption ="Equivalent DBH (cm)"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6960
                    LayoutCachedTop =60
                    LayoutCachedWidth =7890
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =2760
                    Top =120
                    Width =975
                    Height =465
                    FontSize =10
                    ForeColor =0
                    Name ="lblBrowse"
                    Caption ="Browsable/Browsed"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2760
                    LayoutCachedTop =120
                    LayoutCachedWidth =3735
                    LayoutCachedHeight =585
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =9959
                    Top =299
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStatus"
                    Caption ="Status"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9959
                    LayoutCachedTop =299
                    LayoutCachedWidth =10739
                    LayoutCachedHeight =599
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =3720
                    Top =240
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblHabit"
                    Caption ="Habit"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3720
                    LayoutCachedTop =240
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =8205
                    Top =299
                    Width =645
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblVigor"
                    Caption ="Vigor"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8205
                    LayoutCachedTop =299
                    LayoutCachedWidth =8850
                    LayoutCachedHeight =599
                End
            End
        End
        Begin PageHeader
            Height =15
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Name ="Line14"
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
                    TextAlign =1
                    IMESentenceMode =3
                    Width =1500
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="txtMicroplot_Number"
                    ControlSource ="Microplot_Number"
                    StatusBarText ="Distance (m) from plot center to near EDGE of tree"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000056010000010000000100000000000000000000007a00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b004d006900630072006f0070006c006f00 ,
                        0x74005f004e0075006d006200650072005d00290020004f007200200049007300 ,
                        0x4e006f007400680069006e00670028005b004d006900630072006f0070006c00 ,
                        0x6f0074005f004e0075006d006200650072005d00290020004f00720020004900 ,
                        0x730045006d0070007400790028005b004d006900630072006f0070006c006f00 ,
                        0x74005f004e0075006d006200650072005d00290020004f007200200049007300 ,
                        0x4500720072006f00720028005b004d006900630072006f0070006c006f007400 ,
                        0x5f004e0075006d006200650072005d00290000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400790000004900 ,
                        0x73004e0075006c006c0028005b004d006900630072006f0070006c006f007400 ,
                        0x5f004e0075006d006200650072005d00290020004f0072002000490073004e00 ,
                        0x6f007400680069006e00670028005b004d006900630072006f0070006c006f00 ,
                        0x74005f004e0075006d006200650072005d00290020004f007200200049007300 ,
                        0x45006d0070007400790028005b004d006900630072006f0070006c006f007400 ,
                        0x5f004e0075006d006200650072005d00290020004f0072002000490073004500 ,
                        0x720072006f00720028005b004d006900630072006f0070006c006f0074005f00 ,
                        0x4e0075006d006200650072005d00290000000000000000000000000000000000 ,
                        0x0000000000
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =600
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =150
                    Width =795
                    FontSize =9
                    FontWeight =700
                    Name ="txtTag"
                    ControlSource ="Tag"
                    StatusBarText ="Number of physical tag attached to tree"
                    FontName ="Calibri"

                    LayoutCachedLeft =150
                    LayoutCachedWidth =945
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =1125
                    Width =1695
                    FontSize =9
                    TabIndex =1
                    Name ="txtLatin_Name"
                    ControlSource ="Latin_Name"
                    StatusBarText ="Genus of specimen"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b007400780074004c006100740069006e00 ,
                        0x5f004e0061006d0065005d00290000000000
                    End

                    LayoutCachedLeft =1125
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400170000004900 ,
                        0x73004e0075006c006c0028005b007400780074004c006100740069006e005f00 ,
                        0x4e0061006d0065005d0029000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4620
                    Width =2325
                    FontSize =9
                    TabIndex =2
                    Name ="txtStems"
                    ControlSource ="=MakeSaplingStemList([Event_ID],[Sapling_Data_ID])"
                    FontName ="Calibri"

                    LayoutCachedLeft =4620
                    LayoutCachedWidth =6945
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7080
                    Width =720
                    FontSize =9
                    TabIndex =3
                    Name ="txtSum_BasalArea"
                    Format ="Standard"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000f2000000020000000000000004000000000000000300000001000000 ,
                        0x00000000ed1c2400010000000000000004000000480000000100000000000000 ,
                        0xed1c240000000000000000000000000000000000000000000000000000000000 ,
                        0x31003000000000004e007a00280049006e00740028005b007400780074005300 ,
                        0x75006d005f0042006100730061006c0041007200650061005d0029002c003000 ,
                        0x29003e003100300020004f00720020004e007a00280049006e00740028005b00 ,
                        0x740078007400530075006d005f0042006100730061006c004100720065006100 ,
                        0x5d0029002c00300029003c00310000000000
                    End

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000200000000000000040000000100000000000000ed1c2400020000003100 ,
                        0x3000000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ed1c2400430000004e007a00280049006e00740028005b0074 ,
                        0x0078007400530075006d005f0042006100730061006c0041007200650061005d ,
                        0x0029002c00300029003e003100300020004f00720020004e007a00280049006e ,
                        0x00740028005b00740078007400530075006d005f0042006100730061006c0041 ,
                        0x007200650061005d0029002c00300029003c0031000000000000000000000000 ,
                        0x00000000000000000000
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =2400
                    Top =300
                    Width =7380
                    Height =0
                    TabIndex =4
                    Name ="txtSapling_Notes"
                    ControlSource ="Sapling_Notes"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9360
                    Width =1319
                    FontSize =9
                    TabIndex =5
                    Name ="txtSapling_Status"
                    ControlSource ="Sapling_Status"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000009a000000010000000100000000000000000000001c00000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b007400780074005300610070006c006900 ,
                        0x6e0067005f005300740061007400750073005d00290000000000
                    End

                    LayoutCachedLeft =9360
                    LayoutCachedWidth =10679
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24001b0000004900 ,
                        0x73004e0075006c006c0028005b007400780074005300610070006c0069006e00 ,
                        0x67005f005300740061007400750073005d002900000000000000000000000000 ,
                        0x000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3780
                    Width =615
                    FontSize =9
                    TabIndex =6
                    Name ="txtHabit"
                    ControlSource ="Habit"
                    Format ="Standard"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000002a010000030000000100000000000000000000002900000001000000 ,
                        0x00000000faf3e80001000000000000002a000000500000000100000000000000 ,
                        0xfaf3e800010000000000000051000000640000000100000000000000ed1c2400 ,
                        0x49004900660028004c0065006600740028005b005300610070006c0069006e00 ,
                        0x67005f005300740061007400750073005d002c00340029003d00220044004500 ,
                        0x4100440022002c0031002c0030002900000000005b005300610070006c006900 ,
                        0x6e0067005f005300740061007400750073005d003d002200520065006d006f00 ,
                        0x7600650064002000660072006f006d0020007300740075006400790022000000 ,
                        0x0000490073004e0075006c006c0028005b007400780074004800610062006900 ,
                        0x74005d00290000000000
                    End

                    LayoutCachedLeft =3780
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000faf3e800280000004900 ,
                        0x4900660028004c0065006600740028005b005300610070006c0069006e006700 ,
                        0x5f005300740061007400750073005d002c00340029003d002200440045004100 ,
                        0x440022002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x00000001000000000000000100000000000000faf3e800250000005b00530061 ,
                        0x0070006c0069006e0067005f005300740061007400750073005d003d00220052 ,
                        0x0065006d006f007600650064002000660072006f006d00200073007400750064 ,
                        0x0079002200000000000000000000000000000000000000000000010000000000 ,
                        0x00000100000000000000ed1c240012000000490073004e0075006c006c002800 ,
                        0x5b00740078007400480061006200690074005d00290000000000000000000000 ,
                        0x0000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2820
                    Width =825
                    FontSize =9
                    TabIndex =7
                    Name ="txtBrowse_Status"
                    ControlSource ="=[Browsable] & \"/\" & [Browsed]"
                    Format ="Standard"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000056010000030000000100000000000000000000002900000001000000 ,
                        0x00000000faf3e80001000000000000002a000000500000000100000000000000 ,
                        0xfaf3e8000100000000000000510000007a0000000100000000000000ed1c2400 ,
                        0x49004900660028004c0065006600740028005b005300610070006c0069006e00 ,
                        0x67005f005300740061007400750073005d002c00340029003d00220044004500 ,
                        0x4100440022002c0031002c0030002900000000005b005300610070006c006900 ,
                        0x6e0067005f005300740061007400750073005d003d002200520065006d006f00 ,
                        0x7600650064002000660072006f006d0020007300740075006400790022000000 ,
                        0x0000490073004e0075006c006c0028005b00420072006f007700730061006200 ,
                        0x6c0065005d00290020004f0072002000490073004e0075006c006c0028005b00 ,
                        0x420072006f0077007300650064005d00290000000000
                    End

                    LayoutCachedLeft =2820
                    LayoutCachedWidth =3645
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000faf3e800280000004900 ,
                        0x4900660028004c0065006600740028005b005300610070006c0069006e006700 ,
                        0x5f005300740061007400750073005d002c00340029003d002200440045004100 ,
                        0x440022002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x00000001000000000000000100000000000000faf3e800250000005b00530061 ,
                        0x0070006c0069006e0067005f005300740061007400750073005d003d00220052 ,
                        0x0065006d006f007600650064002000660072006f006d00200073007400750064 ,
                        0x0079002200000000000000000000000000000000000000000000010000000000 ,
                        0x00000100000000000000ed1c240028000000490073004e0075006c006c002800 ,
                        0x5b00420072006f0077007300610062006c0065005d00290020004f0072002000 ,
                        0x490073004e0075006c006c0028005b00420072006f0077007300650064005d00 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                End
                Begin Subform
                    Left =5760
                    Top =360
                    Width =2956
                    Height =120
                    TabIndex =8
                    Name ="rSub_Event_rSub_Sapling_Conditions"
                    SourceObject ="Report.rSub_Event_rSub_Sapling_Conditions"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =360
                    LayoutCachedWidth =8716
                    LayoutCachedHeight =480
                End
                Begin Subform
                    Left =1320
                    Top =360
                    Width =2956
                    Height =180
                    TabIndex =9
                    Name ="rSub_Event_rSub_Sapling_Vines"
                    SourceObject ="Report.rSub_Event_rSub_Sapling_Vines"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =360
                    LayoutCachedWidth =4276
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8100
                    Width =1140
                    FontSize =9
                    TabIndex =10
                    Name ="txtVigor"
                    ControlSource ="Vig"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000000c010000030000000100000000000000000000002900000001000000 ,
                        0x00000000ffffff0001000000000000002a000000500000000100000000000000 ,
                        0xfaf3e800000000000200000051000000550000000100000000000000ed1c2400 ,
                        0x49004900660028004c0065006600740028005b005300610070006c0069006e00 ,
                        0x67005f005300740061007400750073005d002c00340029003d00220044004500 ,
                        0x4100440022002c0031002c0030002900000000005b005300610070006c006900 ,
                        0x6e0067005f005300740061007400750073005d003d002200520065006d006f00 ,
                        0x7600650064002000660072006f006d0020007300740075006400790022000000 ,
                        0x000022002000220000000000
                    End

                    LayoutCachedLeft =8100
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffffff00280000004900 ,
                        0x4900660028004c0065006600740028005b005300610070006c0069006e006700 ,
                        0x5f005300740061007400750073005d002c00340029003d002200440045004100 ,
                        0x440022002c0031002c0030002900000000000000000000000000000000000000 ,
                        0x00000001000000000000000100000000000000faf3e800250000005b00530061 ,
                        0x0070006c0069006e0067005f005300740061007400750073005d003d00220052 ,
                        0x0065006d006f007600650064002000660072006f006d00200073007400750064 ,
                        0x0079002200000000000000000000000000000000000000000000000000000200 ,
                        0x00000100000000000000ed1c2400030000002200200022000000000000000000 ,
                        0x00000000000000000000000000
                    End
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
                    Name ="Line15"
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
