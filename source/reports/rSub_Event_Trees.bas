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
    ItemSuffix =42
    Left =645
    Top =1365
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc47eed94f4efe440
    End
    RecordSource ="SELECT tbl_Tags.Tag, tlu_Plants.Latin_Name, qCalc_Basal_Area_per_Tree.Stems, qCa"
        "lc_Basal_Area_per_Tree.Equiv_DBH_cm, [Crown_Class] & \" \" & [CrownClass] AS CC,"
        " [TreeVigor] & \" \" & [TreeVigorClass] AS Vig, tbl_Tree_Data.Vines_Checked, tbl"
        "_Tree_Data.Conditions_Checked, tbl_Tree_Data.Foliage_Conditions_Checked, tbl_Tre"
        "e_Data.Tree_Status, tbl_Tags.Azimuth, tbl_Tags.Distance, tbl_Tree_Data.Tree_Note"
        "s, tbl_Tree_Data.Tree_Data_ID, tbl_Tree_Data.Event_ID, maketreestemlist([tbl_tre"
        "e_data]![Event_ID],[tbl_tree_data]![Tree_Data_Id]) AS StemList FROM (((tbl_Tree_"
        "Data LEFT JOIN qCalc_Basal_Area_per_Tree ON tbl_Tree_Data.Tree_Data_ID = qCalc_B"
        "asal_Area_per_Tree.Tree_Data_ID) LEFT JOIN tbl_Tags ON tbl_Tree_Data.Tag_ID = tb"
        "l_Tags.Tag_ID) LEFT JOIN tlu_Plants ON tbl_Tags.TSN = tlu_Plants.TSN) LEFT JOIN "
        "tluTreeVigor ON tbl_Tree_Data.TreeVigor = tluTreeVigor.TreeVigorCode ORDER BY tb"
        "l_Tags.Tag;"
    Caption ="srpt_Trees"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xf0000000f0000000190100000301000000000000302a0000a401000001000000 ,
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
                    Left =4080
                    Top =300
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStems"
                    Caption ="Stems"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4080
                    LayoutCachedTop =300
                    LayoutCachedWidth =4860
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =5520
                    Top =60
                    Width =1035
                    Height =540
                    FontSize =10
                    ForeColor =0
                    Name ="lblBasal_Area"
                    Caption ="Equivalent DBH (cm)"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5520
                    LayoutCachedTop =60
                    LayoutCachedWidth =6555
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =6660
                    Top =60
                    Width =660
                    Height =540
                    FontSize =10
                    ForeColor =0
                    Name ="lblCrown_Class"
                    Caption ="Crown Class"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6660
                    LayoutCachedTop =60
                    LayoutCachedWidth =7320
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =9960
                    Top =300
                    Width =780
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblStatus"
                    Caption ="Status"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9960
                    LayoutCachedTop =300
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextAlign =2
                    TextFontFamily =34
                    Left =8040
                    Top =300
                    Width =615
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="lblVCF"
                    Caption ="V-C-F"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8040
                    LayoutCachedTop =300
                    LayoutCachedWidth =8655
                    LayoutCachedHeight =600
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =8805
                    Top =300
                    Width =645
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label40"
                    Caption ="Vigor"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8805
                    LayoutCachedTop =300
                    LayoutCachedWidth =9450
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
            Height =420
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    IMESentenceMode =3
                    Left =7920
                    Width =720
                    Height =239
                    TabIndex =14
                    Name ="txtCheck_Background"
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
                    LayoutCachedHeight =239
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
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =1
                    IMESentenceMode =3
                    Width =720
                    Height =270
                    FontSize =9
                    FontWeight =700
                    Name ="txtTag"
                    ControlSource ="Tag"
                    StatusBarText ="Number of physical tag attached to tree"
                    FontName ="Calibri"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =780
                    Width =1320
                    Height =270
                    FontSize =9
                    TabIndex =1
                    Name ="txtLocation"
                    ControlSource ="=[Azimuth] & \"º  \" & Format([Distance],\"Fixed\") & \"m\""
                    StatusBarText ="Distance (m) from plot center to near EDGE of tree"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000b2000000010000000100000000000000000000002800000001000000 ,
                        0x00000000cf7b7900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00440069007300740061006e0063006500 ,
                        0x5d00290020004f0072002000490073004e0075006c006c0028005b0041007a00 ,
                        0x69006d007500740068005d00290000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000cf7b7900270000004900 ,
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

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =270
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
                    Left =3660
                    Width =1980
                    FontSize =9
                    TabIndex =3
                    Name ="txtStems"
                    ControlSource ="=MakeTreeStemList([Event_ID],[Tree_Data_ID])"
                    FontName ="Calibri"

                    LayoutCachedLeft =3660
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =5640
                    Width =720
                    FontSize =9
                    TabIndex =4
                    Name ="txtSum_BasalArea"
                    ControlSource ="Equiv_DBH_cm"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x01000000a4000000010000000100000000000000000000002100000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x4e007a00280049006e00740028005b00740078007400530075006d005f004200 ,
                        0x6100730061006c0041007200650061005d0029002c00300029003c0031003000 ,
                        0x00000000
                    End

                    LayoutCachedLeft =5640
                    LayoutCachedWidth =6360
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400200000004e00 ,
                        0x7a00280049006e00740028005b00740078007400530075006d005f0042006100 ,
                        0x730061006c0041007200650061005d0029002c00300029003c00310030000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =6540
                    Width =1260
                    FontSize =9
                    TabIndex =5
                    Name ="txtCrown_Class"
                    ControlSource ="CC"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000002000000000000000400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x22002000220000000000
                    End

                    LayoutCachedLeft =6540
                    LayoutCachedWidth =7800
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ed1c2400030000002200 ,
                        0x20002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =9780
                    Width =960
                    FontSize =9
                    TabIndex =6
                    Name ="txtTree_Status"
                    ControlSource ="Tree_Status"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0054007200650065005f00530074006100 ,
                        0x7400750073005d00290000000000
                    End

                    LayoutCachedLeft =9780
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400150000004900 ,
                        0x73004e0075006c006c0028005b0054007200650065005f005300740061007400 ,
                        0x750073005d002900000000000000000000000000000000000000000000
                    End
                End
                Begin Subform
                    Left =1560
                    Top =300
                    Width =2956
                    Height =60
                    TabIndex =7
                    Name ="srpt_srpt_Tree_Vines"
                    SourceObject ="Report.rSub_Event_rSub_Tree_Vines"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

                End
                Begin Subform
                    Left =4560
                    Top =300
                    Width =2956
                    Height =60
                    TabIndex =8
                    Name ="rSub_rSub_Tree_Conditions"
                    SourceObject ="Report.rSub_Event_rSub_Tree_Conditions"
                    LinkChildFields ="Tree_Data_ID"
                    LinkMasterFields ="Tree_Data_ID"

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
                    Name ="txtTree_Notes"
                    ControlSource ="Tree_Notes"
                    StatusBarText ="Number of physical tag attached to tree"

                End
                Begin CheckBox
                    Left =8040
                    TabIndex =11
                    Name ="chkVines_Checked"
                    ControlSource ="Vines_Checked"

                    LayoutCachedLeft =8040
                    LayoutCachedWidth =8300
                    LayoutCachedHeight =240
                End
                Begin CheckBox
                    Left =8220
                    TabIndex =12
                    Name ="chkConditions_Checked"
                    ControlSource ="Conditions_Checked"

                    LayoutCachedLeft =8220
                    LayoutCachedWidth =8480
                    LayoutCachedHeight =240
                End
                Begin CheckBox
                    Left =8400
                    TabIndex =13
                    Name ="chkFoliage_Conditions_Checked"
                    ControlSource ="Foliage_Conditions_Checked"

                    LayoutCachedLeft =8400
                    LayoutCachedWidth =8660
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    BackStyle =1
                    IMESentenceMode =3
                    Left =8700
                    Width =1020
                    FontSize =9
                    TabIndex =15
                    Name ="txtVigor"
                    ControlSource ="Vig"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000006a000000010000000000000002000000000000000400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x22002000220000000000
                    End

                    LayoutCachedLeft =8700
                    LayoutCachedWidth =9720
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
