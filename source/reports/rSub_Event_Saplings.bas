Version =20
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
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =38
    Top =600
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xceb7046bf4efe440
    End
    RecordSource ="SELECT tbl_Sapling_Data.Event_ID, tbl_Sapling_Data.Tag_ID, tbl_Tags.Location_ID,"
        " tbl_Tags.Tag, tlu_Plants.Latin_Name, [SaplingVigor] & \" \" & [TreeVigorClass] "
        "AS Vig, tbl_Sapling_Data.Sapling_Status, tbl_Sapling_Data.Sapling_Notes, tlu_Pla"
        "nts.Shrub, tbl_Sapling_Data.DRC, tbl_Tags.Microplot_Number, qCalc_Basal_Area_per"
        "_Sapling.Stems, tbl_Sapling_Data.Sapling_Data_ID, tbl_Sapling_Data.Habit, tbl_Sa"
        "pling_Data.Browsed, tbl_Sapling_Data.Browsable, qCalc_Basal_Area_per_Sapling.Equ"
        "iv_DBH_cm FROM (((tbl_Sapling_Data LEFT JOIN tbl_Tags ON tbl_Sapling_Data.Tag_ID"
        " = tbl_Tags.Tag_ID) LEFT JOIN qCalc_Basal_Area_per_Sapling ON tbl_Sapling_Data.S"
        "apling_Data_ID = qCalc_Basal_Area_per_Sapling.Sapling_Data_ID) LEFT JOIN tlu_Pla"
        "nts ON tbl_Tags.TSN = tlu_Plants.TSN) LEFT JOIN tluTreeVigor ON tbl_Sapling_Data"
        ".SaplingVigor = tluTreeVigor.TreeVigorCode ORDER BY tbl_Tags.Tag;"
    Caption ="srpt_Microplots"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x55010000f0000000550100000301000000000000f42900001c02000001000000 ,
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
                    Left =5580
                    Top =240
                    Width =720
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStems"
                    Caption ="Stems"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5580
                    LayoutCachedTop =240
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =7380
                    Top =60
                    Width =930
                    Height =480
                    FontSize =10
                    ForeColor =0
                    Name ="lblSum_basal_area"
                    Caption ="Equivalent DBH (cm)"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7380
                    LayoutCachedTop =60
                    LayoutCachedWidth =8310
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    Left =2880
                    Top =120
                    Width =975
                    Height =465
                    FontSize =10
                    ForeColor =0
                    Name ="lblBrowse"
                    Caption ="Browsable/Browsed"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =120
                    LayoutCachedWidth =3855
                    LayoutCachedHeight =585
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =4620
                    Top =240
                    Width =465
                    Height =300
                    FontSize =10
                    ForeColor =7633277
                    Name ="lblDRC"
                    Caption ="DRC"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4620
                    LayoutCachedTop =240
                    LayoutCachedWidth =5085
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =9780
                    Top =240
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStatus"
                    Caption ="Status"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9780
                    LayoutCachedTop =240
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =3840
                    Top =240
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label29"
                    Caption ="Habit"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3840
                    LayoutCachedTop =240
                    LayoutCachedWidth =4620
                    LayoutCachedHeight =540
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =8640
                    Top =240
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label36"
                    Caption ="Vigor"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8640
                    LayoutCachedTop =240
                    LayoutCachedWidth =9420
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
            Height =540
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
                    Left =5100
                    Width =2325
                    FontSize =9
                    TabIndex =2
                    Name ="txtStems"
                    ControlSource ="=MakeSaplingStemList([Event_ID],[Sapling_Data_ID])"
                    FontName ="Calibri"

                    LayoutCachedLeft =5100
                    LayoutCachedWidth =7425
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7500
                    Width =660
                    FontSize =9
                    TabIndex =3
                    Name ="txtSum_BasalArea"
                    ControlSource ="=\"\""
                    Format ="Standard"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000ea000000010000000100000000000000000000004400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4e007a00280049006e00740028005b00740078007400530075006d005f004200 ,
                        0x6100730061006c0041007200650061005d0029002c00300029003e0031003000 ,
                        0x20004f00720020004e007a00280049006e00740028005b007400780074005300 ,
                        0x75006d005f0042006100730061006c0041007200650061005d0029002c003000 ,
                        0x29003c00310000000000
                    End

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =8160
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400430000004e00 ,
                        0x7a00280049006e00740028005b00740078007400530075006d005f0042006100 ,
                        0x730061006c0041007200650061005d0029002c00300029003e00310030002000 ,
                        0x4f00720020004e007a00280049006e00740028005b0074007800740053007500 ,
                        0x6d005f0042006100730061006c0041007200650061005d0029002c0030002900 ,
                        0x3c003100000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4560
                    Width =525
                    FontSize =9
                    TabIndex =4
                    ForeColor =7633277
                    Name ="txtDRC"
                    ControlSource ="DRC"
                    FontName ="Calibri"

                    LayoutCachedLeft =4560
                    LayoutCachedWidth =5085
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    CanGrow = NotDefault
                    IMESentenceMode =3
                    Left =2400
                    Top =300
                    Width =7380
                    Height =0
                    TabIndex =5
                    Name ="txtSapling_Notes"
                    ControlSource ="Sapling_Notes"

                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9360
                    Width =1320
                    FontSize =9
                    TabIndex =6
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
                    LayoutCachedWidth =10680
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
                    Left =3900
                    Width =615
                    FontSize =9
                    TabIndex =7
                    Name ="txtHabit"
                    ControlSource ="Habit"
                    Format ="Standard"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740048006100620069007400 ,
                        0x5d00290000000000
                    End

                    LayoutCachedLeft =3900
                    LayoutCachedWidth =4515
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400120000004900 ,
                        0x73004e0075006c006c0028005b00740078007400480061006200690074005d00 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2940
                    Width =825
                    FontSize =9
                    TabIndex =8
                    Name ="txtBrowse_Status"
                    ControlSource ="=[Browsable] & \"/\" & [Browsed]"
                    Format ="Standard"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000b4000000010000000100000000000000000000002900000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00420072006f0077007300610062006c00 ,
                        0x65005d00290020004f0072002000490073004e0075006c006c0028005b004200 ,
                        0x72006f0077007300650064005d00290000000000
                    End

                    LayoutCachedLeft =2940
                    LayoutCachedWidth =3765
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400280000004900 ,
                        0x73004e0075006c006c0028005b00420072006f0077007300610062006c006500 ,
                        0x5d00290020004f0072002000490073004e0075006c006c0028005b0042007200 ,
                        0x6f0077007300650064005d002900000000000000000000000000000000000000 ,
                        0x000000
                    End
                End
                Begin Subform
                    Left =5760
                    Top =360
                    Width =2956
                    Height =120
                    TabIndex =9
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
                    TabIndex =10
                    Name ="rSub_Event_rSub_Sapliing_Vines"
                    SourceObject ="Report.rSub_Event_rSub_Sapliing_Vines"
                    LinkChildFields ="Sapling_Data_ID"
                    LinkMasterFields ="Sapling_Data_ID"

                    LayoutCachedLeft =1320
                    LayoutCachedTop =360
                    LayoutCachedWidth =4276
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8220
                    Width =1140
                    FontSize =9
                    TabIndex =11
                    Name ="txtVigor"
                    ControlSource ="Vig"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000002000000000000000400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x22002000220000000000
                    End

                    LayoutCachedLeft =8220
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ed1c2400030000002200 ,
                        0x20002200000000000000000000000000000000000000000000
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
