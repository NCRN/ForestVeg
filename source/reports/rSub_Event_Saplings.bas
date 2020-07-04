Version =21
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
    Width =11220
    DatasheetFontHeight =10
    ItemSuffix =54
    Left =150
    Top =2160
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x37263bef737de540
    End
    RecordSource ="SELECT sd.Event_ID, sd.Tag_ID, t.Location_ID, t.Tag, p.Latin_Name, [SaplingVigor"
        "] & \" \" & [TreeVigorClass] AS Vig, sd.Sapling_Status, sd.Sapling_Notes, p.Shru"
        "b, sd.DRC, t.Microplot_Number, ba.Stems, sd.Sapling_Data_ID, sd.Habit, sd.Browse"
        "d, sd.Browsable, ba.Equiv_DBH_cm, Len(ba.Equiv_DBH_cm) AS len_equiv_dbh, TypeNam"
        "e(ba.Equiv_DBH_cm) AS typename_eqiv_dbh, Left(ba.Equiv_DBH_cm,4) AS left4_equiv_"
        "dbh, CStr(ba.Equiv_DBH_cm) AS cstr_equiv_dbh, sd.Vines_Checked, sd.Conditions_Ch"
        "ecked, sd.Foliage_Conditions_Checked, sd.DBH_Check FROM (((tbl_Sapling_Data AS s"
        "d LEFT JOIN tbl_Tags AS t ON sd.Tag_ID = t.Tag_ID) LEFT JOIN qCalc_Basal_Area_pe"
        "r_Sapling AS ba ON sd.Sapling_Data_ID = ba.Sapling_Data_ID) LEFT JOIN tlu_Plants"
        " AS p ON t.TSN = p.TSN) LEFT JOIN tluTreeVigor AS v ON sd.SaplingVigor = v.TreeV"
        "igorCode ORDER BY t.Tag;"
    Caption ="srpt_Microplots"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e0010000680100006801000000000000d42b00001c02000001000000 ,
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
            Height =600
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
                    Left =960
                    Top =240
                    Width =1845
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblLatin_Name"
                    Caption ="Latin Name"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =960
                    LayoutCachedTop =240
                    LayoutCachedWidth =2805
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =5085
                    Top =240
                    Width =720
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStems"
                    Caption ="Stems"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5085
                    LayoutCachedTop =240
                    LayoutCachedWidth =5805
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =6945
                    Top =60
                    Width =930
                    Height =480
                    FontSize =10
                    ForeColor =0
                    Name ="lblSum_basal_area"
                    Caption ="Equivalent DBH (cm)"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6945
                    LayoutCachedTop =60
                    LayoutCachedWidth =7875
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =2745
                    Top =120
                    Width =975
                    Height =465
                    FontSize =10
                    ForeColor =0
                    Name ="lblBrowse"
                    Caption ="Browsable/Browsed"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2745
                    LayoutCachedTop =120
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =585
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =10140
                    Top =240
                    Width =600
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStatus"
                    Caption ="Status"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =10140
                    LayoutCachedTop =240
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =3705
                    Top =240
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblHabit"
                    Caption ="Habit"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3705
                    LayoutCachedTop =240
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =8625
                    Top =240
                    Width =645
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblVigor"
                    Caption ="Vigor"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8625
                    LayoutCachedTop =240
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =7920
                    Top =240
                    Width =615
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblVCF"
                    Caption ="V-C-F"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7920
                    LayoutCachedTop =240
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =540
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
            Height =720
            OnFormat ="[Event Procedure]"
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
                    Name ="tbxLatinName"
                    ControlSource ="Latin_Name"
                    StatusBarText ="Genus of specimen"
                    FontName ="Calibri"

                    LayoutCachedLeft =1125
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4620
                    Width =2325
                    FontSize =9
                    TabIndex =2
                    Name ="txtStems"
                    ControlSource ="=MakeSaplingStemList([Event_ID],[Sapling_Data_ID])"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000aa000000010000000100000000000000000000002400000001000000 ,
                        0x00000000fff20000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4c0065006e0028005400720069006d0028005b00740078007400530074006500 ,
                        0x6d0073005d00290029003d004c0065006e00280022004c003a00200020004400 ,
                        0x3a002200290000000000
                    End

                    LayoutCachedLeft =4620
                    LayoutCachedWidth =6945
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000fff20000230000004c00 ,
                        0x65006e0028005400720069006d0028005b007400780074005300740065006d00 ,
                        0x73005d00290029003d004c0065006e00280022004c003a002000200044003a00 ,
                        0x22002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7080
                    Top =120
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
                    LayoutCachedTop =120
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =360
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
                    Left =1440
                    Top =480
                    Width =8820
                    Height =0
                    TabIndex =4
                    Name ="txtSapling_Notes"
                    ControlSource ="Sapling_Notes"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =480
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9600
                    Width =1199
                    Height =420
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

                    LayoutCachedLeft =9600
                    LayoutCachedWidth =10799
                    LayoutCachedHeight =420
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
                    Left =3720
                    Top =540
                    Width =2956
                    Height =120
                    TabIndex =8
                    Name ="rSub_Event_rSub_Sapling_Conditions"
                    SourceObject ="Report.rSub_Event_rSub_Sapling_Conditions"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =540
                    LayoutCachedWidth =6676
                    LayoutCachedHeight =660
                End
                Begin Subform
                    Left =480
                    Top =540
                    Width =2956
                    Height =180
                    TabIndex =9
                    Name ="rSub_Event_rSub_Sapling_Vines"
                    SourceObject ="Report.rSub_Event_rSub_Sapling_Vines"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =480
                    LayoutCachedTop =540
                    LayoutCachedWidth =3436
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8700
                    Width =840
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

                    LayoutCachedLeft =8700
                    LayoutCachedWidth =9540
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
                Begin Subform
                    Left =6960
                    Top =540
                    Width =2956
                    Height =60
                    TabIndex =11
                    Name ="rSub_rSub_Sapling_Foliage"
                    SourceObject ="Report.rSub_Event_rSub_Sapling_Foliage"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =6960
                    LayoutCachedTop =540
                    LayoutCachedWidth =9916
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =1020
                    Width =1680
                    Height =270
                    FontSize =8
                    BackColor =721136
                    ForeColor =16777215
                    Name ="lblMissingID"
                    Caption ="M I S S I N G  I D"
                    FontName ="Arial"
                    LayoutCachedLeft =1020
                    LayoutCachedWidth =2700
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7920
                    Width =720
                    TabIndex =12
                    Name ="tbxCheckBackground"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x0100000028010000010000000100000000000000000000006300000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00630068006b00560069006e00650073005f0043006800650063006b006500 ,
                        0x64005d003d00460061006c007300650020004f00720020005b00630068006b00 ,
                        0x43006f006e0064006900740069006f006e0073005f0043006800650063006b00 ,
                        0x650064005d003d00460061006c007300650020004f00720020005b0063006800 ,
                        0x6b0046006f006c0069006100670065005f0043006f006e006400690074006900 ,
                        0x6f006e0073005f0043006800650063006b00650064005d003d00460061006c00 ,
                        0x7300650000000000
                    End

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400620000005b00 ,
                        0x630068006b00560069006e00650073005f0043006800650063006b0065006400 ,
                        0x5d003d00460061006c007300650020004f00720020005b00630068006b004300 ,
                        0x6f006e0064006900740069006f006e0073005f0043006800650063006b006500 ,
                        0x64005d003d00460061006c007300650020004f00720020005b00630068006b00 ,
                        0x46006f006c0069006100670065005f0043006f006e0064006900740069006f00 ,
                        0x6e0073005f0043006800650063006b00650064005d003d00460061006c007300 ,
                        0x6500000000000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    Left =8040
                    Top =15
                    TabIndex =13
                    Name ="chkVines_Checked"
                    ControlSource ="Vines_Checked"

                    LayoutCachedLeft =8040
                    LayoutCachedTop =15
                    LayoutCachedWidth =8300
                    LayoutCachedHeight =255
                End
                Begin CheckBox
                    Left =8220
                    Top =15
                    TabIndex =14
                    Name ="chkConditions_Checked"
                    ControlSource ="Conditions_Checked"

                    LayoutCachedLeft =8220
                    LayoutCachedTop =15
                    LayoutCachedWidth =8480
                    LayoutCachedHeight =255
                End
                Begin CheckBox
                    Left =8400
                    Top =15
                    TabIndex =15
                    Name ="chkFoliage_Conditions_Checked"
                    ControlSource ="Foliage_Conditions_Checked"

                    LayoutCachedLeft =8400
                    LayoutCachedTop =15
                    LayoutCachedWidth =8660
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7080
                    Width =720
                    FontSize =9
                    TabIndex =16
                    Name ="tbxEquivDBH"
                    ControlSource ="=GetEquivDBH([Sapling_Data_ID])"
                    Format ="Standard"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000de000000020000000000000004000000000000000300000001000000 ,
                        0x00000000ed1c24000100000000000000040000003e0000000100000000000000 ,
                        0xed1c240000000000000000000000000000000000000000000000000000000000 ,
                        0x31003000000000004e007a00280049006e00740028005b007400620078004500 ,
                        0x71007500690076004400420048005d0029002c00300029003e00310030002000 ,
                        0x4f00720020004e007a00280049006e00740028005b0074006200780045007100 ,
                        0x7500690076004400420048005d0029002c00300029003c00310000000000
                    End

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000200000000000000040000000100000000000000ed1c2400020000003100 ,
                        0x3000000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ed1c2400390000004e007a00280049006e00740028005b0074 ,
                        0x0062007800450071007500690076004400420048005d0029002c00300029003e ,
                        0x003100300020004f00720020004e007a00280049006e00740028005b00740062 ,
                        0x007800450071007500690076004400420048005d0029002c00300029003c0031 ,
                        0x00000000000000000000000000000000000000000000
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
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' REPORT:       rSub_Event_Saplings
' Level:        Application report
' Version:      1.01
'
' Description:  Report related functions & procedures for application
'
' Source/date:  Bonnie Campbell, October 24, 2018
' Revisions:    BLC - 10/24/2018 - 1.00 - initial version
' =================================

' ---------------------------------
' SUB:          Detail_Format
' Description:  report format actions
' Assumptions:  -
' Parameters:   Cancel - whether format action should be cancelled (boolean)
'               FormatCount - number of times a section (in this case the detail section)
'                             is formatted (integer)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 24, 2018
' Adapted:      -
' Revisions:
'   BLC - 10/24/2018 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    'turn on label if missing sapling ID (tbxLatinName)
    'visible IF there is no data (if no latin name = False, returns True & displays)
    lblMissingID.visible = IIf(Len(tbxLatinName) > 0, False, True)
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[rpt_rSub_Event_Saplings])"
    End Select
    Resume Exit_Handler
End Sub
