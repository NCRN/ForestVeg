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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =43
    Left =30
    Top =-7441
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x159b418f4c1ce540
    End
    RecordSource ="SELECT t.Tag, p.Latin_Name,  q.Stems, q.Equiv_DBH_cm, [Crown_Class] & \" \" & [C"
        "rownClass] AS CC,  [TreeVigor] & \" \" & [TreeVigorClass] AS Vig,  td.Vines_Chec"
        "ked, td.Conditions_Checked, td.Foliage_Conditions_Checked, td.Tree_Status, t.Azi"
        "muth, t.Distance, td.Tree_Notes, td.Tree_Data_ID, td.Event_ID, MakeStemList('Tre"
        "e',[tbl_tree_data]![Event_ID],[tbl_tree_data]![Tree_Data_Id]) AS StemList, MakeL"
        "iveFlag('Tree',[tbl_tree_data]![Event_ID],[tbl_tree_data]![Tree_Data_Id]) AS Liv"
        "eFlag FROM (((tbl_Tree_Data td LEFT JOIN qCalc_Basal_Area_per_Tree q ON td.Tree_"
        "Data_ID = q.Tree_Data_ID)  LEFT JOIN tbl_Tags t ON td.Tag_ID = t.Tag_ID)  LEFT J"
        "OIN tlu_Plants p ON t.TSN = p.TSN)  LEFT JOIN tluTreeVigor tv ON td.TreeVigor = "
        "tv.TreeVigorCode ORDER BY t.Tag;"
    Caption ="srpt_Trees"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x0000000000000000000000000000000000000000302a00001c02000001000000 ,
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =660
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Top =300
                    Width =600
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblTag"
                    Caption ="Tag"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =300
                    LayoutCachedWidth =600
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =960
                    Top =300
                    Width =1020
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblLocation"
                    Caption ="Location"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =960
                    LayoutCachedTop =300
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =2100
                    Top =300
                    Width =1500
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblLatin_Name"
                    Caption ="Latin Name"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2100
                    LayoutCachedTop =300
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =3960
                    Top =300
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStems"
                    Caption ="Stems"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3960
                    LayoutCachedTop =300
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =5400
                    Top =60
                    Width =1035
                    Height =540
                    FontSize =10
                    ForeColor =0
                    Name ="lblBasal_Area"
                    Caption ="Equivalent DBH (cm)"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =60
                    LayoutCachedWidth =6435
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =6480
                    Top =60
                    Width =660
                    Height =540
                    FontSize =10
                    ForeColor =0
                    Name ="lblCrown_Class"
                    Caption ="Crown Class"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6480
                    LayoutCachedTop =60
                    LayoutCachedWidth =7140
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =9660
                    Top =300
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStatus"
                    Caption ="Status"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9660
                    LayoutCachedTop =300
                    LayoutCachedWidth =10440
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =7920
                    Top =300
                    Width =615
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblVCF"
                    Caption ="V-C-F"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7920
                    LayoutCachedTop =300
                    LayoutCachedWidth =8535
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =8685
                    Top =300
                    Width =645
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblVigor"
                    Caption ="Vigor"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8685
                    LayoutCachedTop =300
                    LayoutCachedWidth =9330
                    LayoutCachedHeight =600
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
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =540
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7800
                    Width =720
                    Height =239
                    TabIndex =14
                    Name ="tbxCheckBackground"
                    ControlSource ="=\"\""
                    ConditionalFormat = Begin
                        0x01000000de010000030000000100000000000000000000002600000001000000 ,
                        0x00000000faf3e8000100000000000000270000005a0000000100000000000000 ,
                        0xffffff0001000000000000005b000000be0000000100000000000000ed1c2400 ,
                        0x49004900660028004c0065006600740028005b0054007200650065005f005300 ,
                        0x740061007400750073005d002c00340029003d00220044004500410044002200 ,
                        0x2c0031002c00300029000000000049004900660028005b005400720065006500 ,
                        0x5f005300740061007400750073005d003d00220044006f0077006e0067007200 ,
                        0x6100640065006400200074006f0020004e006f006e002d00530061006d007000 ,
                        0x6c006500640022002c0031002c0030002900000000005b00630068006b005600 ,
                        0x69006e00650073005f0043006800650063006b00650064005d003d0046006100 ,
                        0x6c007300650020004f00720020005b00630068006b0043006f006e0064006900 ,
                        0x740069006f006e0073005f0043006800650063006b00650064005d003d004600 ,
                        0x61006c007300650020004f00720020005b00630068006b0046006f006c006900 ,
                        0x6100670065005f0043006f006e0064006900740069006f006e0073005f004300 ,
                        0x6800650063006b00650064005d003d00460061006c007300650000000000
                    End

                    LayoutCachedLeft =7800
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =239
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000faf3e800250000004900 ,
                        0x4900660028004c0065006600740028005b0054007200650065005f0053007400 ,
                        0x61007400750073005d002c00340029003d002200440045004100440022002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000010000 ,
                        0x00000000000100000000000000ffffff003200000049004900660028005b0054 ,
                        0x007200650065005f005300740061007400750073005d003d00220044006f0077 ,
                        0x006e00670072006100640065006400200074006f0020004e006f006e002d0053 ,
                        0x0061006d0070006c006500640022002c0031002c003000290000000000000000 ,
                        0x000000000000000000000000000001000000000000000100000000000000ed1c ,
                        0x2400620000005b00630068006b00560069006e00650073005f00430068006500 ,
                        0x63006b00650064005d003d00460061006c007300650020004f00720020005b00 ,
                        0x630068006b0043006f006e0064006900740069006f006e0073005f0043006800 ,
                        0x650063006b00650064005d003d00460061006c007300650020004f0072002000 ,
                        0x5b00630068006b0046006f006c0069006100670065005f0043006f006e006400 ,
                        0x6900740069006f006e0073005f0043006800650063006b00650064005d003d00 ,
                        0x460061006c0073006500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Width =720
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="tbxTag"
                    ControlSource ="Tag"
                    StatusBarText ="Number of physical tag attached to tree"
                    FontName ="Calibri"

                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =780
                    Width =1320
                    Height =270
                    FontSize =9
                    TabIndex =1
                    Name ="tbxLocation"
                    ControlSource ="=[Azimuth] & \"º  \" & Format([Distance],\"Fixed\") & \"m\""
                    StatusBarText ="Distance (m) from plot center to near EDGE of tree"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000b2000000010000000100000000000000000000002800000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00440069007300740061006e0063006500 ,
                        0x5d00290020004f0072002000490073004e0075006c006c0028005b0041007a00 ,
                        0x69006d007500740068005d00290000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400270000004900 ,
                        0x73004e0075006c006c0028005b00440069007300740061006e00630065005d00 ,
                        0x290020004f0072002000490073004e0075006c006c0028005b0041007a006900 ,
                        0x6d007500740068005d0029000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2160
                    Height =270
                    FontSize =9
                    TabIndex =2
                    Name ="tbxLatinName"
                    ControlSource ="Latin_Name"
                    StatusBarText ="Genus of specimen"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b007400620078004c006100740069006e00 ,
                        0x4e0061006d0065005d00290000000000
                    End

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =270
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400160000004900 ,
                        0x73004e0075006c006c0028005b007400620078004c006100740069006e004e00 ,
                        0x61006d0065005d002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3540
                    Width =1980
                    FontSize =9
                    TabIndex =3
                    Name ="tbxStems"
                    ControlSource ="=MakeStemList('Tree',[Event_ID],[Tree_Data_ID])"
                    FontName ="Calibri"

                    LayoutCachedLeft =3540
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5520
                    Width =720
                    FontSize =9
                    TabIndex =4
                    Name ="tbxSumBasalArea"
                    ControlSource ="Equiv_DBH_cm"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000d8010000030000000100000000000000000000002600000001000000 ,
                        0x00000000faf3e8000100000000000000270000005a0000000100000000000000 ,
                        0xffffff0001000000000000005b000000bb0000000100000000000000ed1c2400 ,
                        0x49004900660028004c0065006600740028005b0054007200650065005f005300 ,
                        0x740061007400750073005d002c00340029003d00220044004500410044002200 ,
                        0x2c0031002c00300029000000000049004900660028005b005400720065006500 ,
                        0x5f005300740061007400750073005d003d00220044006f0077006e0067007200 ,
                        0x6100640065006400200074006f0020004e006f006e002d00530061006d007000 ,
                        0x6c006500640022002c0031002c00300029000000000049004900660028005b00 ,
                        0x54007200650065005f005300740061007400750073005d003d00220044006500 ,
                        0x6100640020007300740061006e00640069006e006700220020004f0072002000 ,
                        0x5b0054007200650065005f005300740061007400750073005d003d0022004400 ,
                        0x65006100640020006c00650061006e0069006e0067002200200041006e006400 ,
                        0x20005b00740062007800530075006d0042006100730061006c00410072006500 ,
                        0x61005d003d00220022002c0031002c003000290000000000
                    End

                    LayoutCachedLeft =5520
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000500000001000000000000000100000000000000faf3e800250000004900 ,
                        0x4900660028004c0065006600740028005b0054007200650065005f0053007400 ,
                        0x61007400750073005d002c00340029003d002200440045004100440022002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000010000 ,
                        0x00000000000100000000000000ffffff003200000049004900660028005b0054 ,
                        0x007200650065005f005300740061007400750073005d003d00220044006f0077 ,
                        0x006e00670072006100640065006400200074006f0020004e006f006e002d0053 ,
                        0x0061006d0070006c006500640022002c0031002c003000290000000000000000 ,
                        0x000000000000000000000000000001000000000000000100000000000000ed1c ,
                        0x24005f00000049004900660028005b0054007200650065005f00530074006100 ,
                        0x7400750073005d003d002200440065006100640020007300740061006e006400 ,
                        0x69006e006700220020004f00720020005b0054007200650065005f0053007400 ,
                        0x61007400750073005d003d002200440065006100640020006c00650061006e00 ,
                        0x69006e0067002200200041006e00640020005b00740062007800530075006d00 ,
                        0x42006100730061006c0041007200650061005d003d00220022002c0031002c00 ,
                        0x3000290000000000000000000000000000000000000000000001000000000000 ,
                        0x000100000000000000ed1c24001f0000004e007a00280049006e00740028005b ,
                        0x00740062007800530075006d0042006100730061006c0041007200650061005d ,
                        0x0029002c00300029003c00310030000000000000000000000000000000000000 ,
                        0x0000000001000000000000000101000000000000ed1c24007a00000049004900 ,
                        0x66002800280028004c00650066007400240028005b0074006200780054007200 ,
                        0x650065005300740061007400750073005d002c00340029003d00270044006500 ,
                        0x610064002700200041006e006400200028005b004c0069007600650046006c00 ,
                        0x610067005d003e0030002900290020004f0072002000280028004c0065006600 ,
                        0x7400240028005b00740062007800540072006500650053007400610074007500 ,
                        0x73005d002c00350029003d00270041006c006900760065002700290020004100 ,
                        0x6e006400200028005b004c0069007600650046006c00610067005d003d003000 ,
                        0x2900290029002c0031002c003000290000000000000000000000000000000000 ,
                        0x0000000000
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =6360
                    Width =1260
                    FontSize =9
                    TabIndex =5
                    Name ="tbxCrownClass"
                    ControlSource ="CC"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000020010000030000000100000000000000000000002600000001000000 ,
                        0x00000000faf3e8000100000000000000270000005a0000000100000000000000 ,
                        0xffffff0000000000020000005b0000005f0000000100000000000000ed1c2400 ,
                        0x49004900660028004c0065006600740028005b0054007200650065005f005300 ,
                        0x740061007400750073005d002c00340029003d00220044004500410044002200 ,
                        0x2c0031002c00300029000000000049004900660028005b005400720065006500 ,
                        0x5f005300740061007400750073005d003d00220044006f0077006e0067007200 ,
                        0x6100640065006400200074006f0020004e006f006e002d00530061006d007000 ,
                        0x6c006500640022002c0031002c00300029000000000022002000220000000000
                    End

                    LayoutCachedLeft =6360
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000faf3e800250000004900 ,
                        0x4900660028004c0065006600740028005b0054007200650065005f0053007400 ,
                        0x61007400750073005d002c00340029003d002200440045004100440022002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000010000 ,
                        0x00000000000100000000000000ffffff003200000049004900660028005b0054 ,
                        0x007200650065005f005300740061007400750073005d003d00220044006f0077 ,
                        0x006e00670072006100640065006400200074006f0020004e006f006e002d0053 ,
                        0x0061006d0070006c006500640022002c0031002c003000290000000000000000 ,
                        0x000000000000000000000000000000000000020000000100000000000000ed1c ,
                        0x2400030000002200200022000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9480
                    Width =1319
                    FontSize =9
                    TabIndex =6
                    Name ="tbxTreeStatus"
                    ControlSource ="Tree_Status"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000ee010000030000000100000000000000000000003300000001000000 ,
                        0x00000000ffffcc000100000000000000340000004a0000000100000000000000 ,
                        0xed1c240001000000000000004b000000c60000000100000000000000ed1c2400 ,
                        0x49004900660028005b0054007200650065005f00530074006100740075007300 ,
                        0x5d003d00220044006f0077006e00670072006100640065006400200074006f00 ,
                        0x20004e006f006e002d00530061006d0070006c006500640022002c0031002c00 ,
                        0x3000290000000000490073004e0075006c006c0028005b005400720065006500 ,
                        0x5f005300740061007400750073005d0029000000000049004900660028002800 ,
                        0x28004c00650066007400240028005b0074006200780054007200650065005300 ,
                        0x740061007400750073005d002c00340029003d00270044006500610064002700 ,
                        0x200041006e006400200028005b004c0069007600650046006c00610067005d00 ,
                        0x3e0030002900290020004f0072002000280028004c0065006600740024002800 ,
                        0x5b0074006200780054007200650065005300740061007400750073005d002c00 ,
                        0x350029003d00270041006c0069007600650027002900200041006e0064002000 ,
                        0x28005b004c0069007600650046006c00610067005d003d003000290029002900 ,
                        0x2c0031002c003000290000000000
                    End

                    LayoutCachedLeft =9480
                    LayoutCachedWidth =10799
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000ffffcc00320000004900 ,
                        0x4900660028005b0054007200650065005f005300740061007400750073005d00 ,
                        0x3d00220044006f0077006e00670072006100640065006400200074006f002000 ,
                        0x4e006f006e002d00530061006d0070006c006500640022002c0031002c003000 ,
                        0x2900000000000000000000000000000000000000000000010000000000000001 ,
                        0x00000000000000ed1c240015000000490073004e0075006c006c0028005b0054 ,
                        0x007200650065005f005300740061007400750073005d00290000000000000000 ,
                        0x000000000000000000000000000001000000000000000100000000000000ed1c ,
                        0x24007a0000004900490066002800280028004c00650066007400240028005b00 ,
                        0x74006200780054007200650065005300740061007400750073005d002c003400 ,
                        0x29003d00270044006500610064002700200041006e006400200028005b004c00 ,
                        0x69007600650046006c00610067005d003e0030002900290020004f0072002000 ,
                        0x280028004c00650066007400240028005b007400620078005400720065006500 ,
                        0x5300740061007400750073005d002c00350029003d00270041006c0069007600 ,
                        0x650027002900200041006e006400200028005b004c0069007600650046006c00 ,
                        0x610067005d003d0030002900290029002c0031002c0030002900000000000000 ,
                        0x000000000000000000000000000000
                    End
                End
                Begin Subform
                    Left =1440
                    Top =300
                    Width =2956
                    Height =60
                    TabIndex =7
                    Name ="srpt_srpt_Tree_Vines"
                    SourceObject ="Report.rSub_Event_rSub_Tree_Vines"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =300
                    LayoutCachedWidth =4396
                    LayoutCachedHeight =360
                End
                Begin Subform
                    Left =4440
                    Top =300
                    Width =2956
                    Height =60
                    TabIndex =8
                    Name ="rSub_rSub_Tree_Conditions"
                    SourceObject ="Report.rSub_Event_rSub_Tree_Conditions"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                    LayoutCachedLeft =4440
                    LayoutCachedTop =300
                    LayoutCachedWidth =7396
                    LayoutCachedHeight =360
                End
                Begin Subform
                    Left =7680
                    Top =300
                    Width =2956
                    Height =60
                    TabIndex =9
                    Name ="rSub_rSub_Tree_Foliage"
                    SourceObject ="Report.rSub_Event_rSub_Tree_Foliage"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMESentenceMode =3
                    Top =300
                    Width =1560
                    Height =60
                    FontSize =6
                    TabIndex =10
                    ForeColor =8421504
                    Name ="tbxTreeNotes"
                    ControlSource ="Tree_Notes"
                    StatusBarText ="Number of physical tag attached to tree"

                End
                Begin CheckBox
                    Left =7920
                    TabIndex =11
                    Name ="chkVines_Checked"
                    ControlSource ="Vines_Checked"

                    LayoutCachedLeft =7920
                    LayoutCachedWidth =8180
                    LayoutCachedHeight =240
                End
                Begin CheckBox
                    Left =8100
                    TabIndex =12
                    Name ="chkConditions_Checked"
                    ControlSource ="Conditions_Checked"

                    LayoutCachedLeft =8100
                    LayoutCachedWidth =8360
                    LayoutCachedHeight =240
                End
                Begin CheckBox
                    Left =8280
                    TabIndex =13
                    Name ="chkFoliage_Conditions_Checked"
                    ControlSource ="Foliage_Conditions_Checked"

                    LayoutCachedLeft =8280
                    LayoutCachedWidth =8540
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8580
                    Width =1020
                    FontSize =9
                    TabIndex =15
                    Name ="tbxVigor"
                    ControlSource ="Vig"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000020010000030000000100000000000000000000002600000001000000 ,
                        0x00000000faf3e8000100000000000000270000005a0000000100000000000000 ,
                        0xffffff0000000000020000005b0000005f0000000100000000000000ed1c2400 ,
                        0x49004900660028004c0065006600740028005b0054007200650065005f005300 ,
                        0x740061007400750073005d002c00340029003d00220044004500410044002200 ,
                        0x2c0031002c00300029000000000049004900660028005b005400720065006500 ,
                        0x5f005300740061007400750073005d003d00220044006f0077006e0067007200 ,
                        0x6100640065006400200074006f0020004e006f006e002d00530061006d007000 ,
                        0x6c006500640022002c0031002c00300029000000000022002000220000000000
                    End

                    LayoutCachedLeft =8580
                    LayoutCachedWidth =9600
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000300000001000000000000000100000000000000faf3e800250000004900 ,
                        0x4900660028004c0065006600740028005b0054007200650065005f0053007400 ,
                        0x61007400750073005d002c00340029003d002200440045004100440022002c00 ,
                        0x31002c0030002900000000000000000000000000000000000000000000010000 ,
                        0x00000000000100000000000000ffffff003200000049004900660028005b0054 ,
                        0x007200650065005f005300740061007400750073005d003d00220044006f0077 ,
                        0x006e00670072006100640065006400200074006f0020004e006f006e002d0053 ,
                        0x0061006d0070006c006500640022002c0031002c003000290000000000000000 ,
                        0x000000000000000000000000000000000000020000000100000000000000ed1c ,
                        0x2400030000002200200022000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =34
                    Left =4500
                    Top =300
                    Width =2865
                    Height =225
                    FontSize =8
                    BackColor =2366701
                    ForeColor =16777215
                    Name ="lblNoTreeConditions"
                    Caption ="N o  T r e e   C o n d i t i o n s"
                    FontName ="Arial"
                    LayoutCachedLeft =4500
                    LayoutCachedTop =300
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =525
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
' REPORT:       rSub_Event_Trees
' Level:        Application report
' Version:      1.01
'
' Description:  Report related functions & procedures for application
'
' Source/date:  Bonnie Campbell, April 3, 2018
' Revisions:    BLC - 4/3/2018 - 1.00 - initial version
'               BLC - 5/21/2018 1.01 - revise to only show in red if trees are not dead/non-sampled (w/o EAB-infection)
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
' References:
'   Marshall Barton, January 14, 2013
'   https://answers.microsoft.com/en-us/msoffice/forum/msoffice_access-mso_other/show-subreport-when-there-are-no-records/0631e68d-45fc-4fcc-b49c-a8944bc47906
'   Duane (dhookom), February 7, 2013
'   http://www.tek-tips.com/viewthread.cfm?qid=1703869
' Source/date:  Bonnie Campbell, April 3, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/3/2018 - initial version
'   BLC - 5/21/2018 - revise to only show in red if trees are not dead/non-sampled (w/o EAB-infection)
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    'turn on label if *not* dead or non-sampled tree & no tree conditions
    If (Left(tbxTreeStatus, 4) <> "Dead") And _
       (Right(tbxTreeStatus, 11) <> "Non-Sampled") And _
        rSub_rSub_Tree_Conditions.Report.HasData = False Then
                
            'visible IF there is no data (if HasData = False, returns True & displays)
            lblNoTreeConditions.Visible = Not rSub_rSub_Tree_Conditions.Report.HasData
        
    Else
        'hide it
        lblNoTreeConditions.Visible = False
    End If

    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[rpt_rSub_Event_Trees])"
    End Select
    Resume Exit_Handler
End Sub
