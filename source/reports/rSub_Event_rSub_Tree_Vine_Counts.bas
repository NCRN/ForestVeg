Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =124
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1920
    DatasheetFontHeight =9
    ItemSuffix =7
    Left =2050
    Top =770
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xffc06a4957a2e540
    End
    RecordSource ="SELECT DISTINCT t.Tag, YEAR(e.Event_Date) AS SampleYear,  COUNT(tv.Tree_Data_ID)"
        " AS VineCount FROM (((tbl_Tree_Vines tv LEFT JOIN tbl_Tree_Data td ON td.Tree_Da"
        "ta_ID = tv.Tree_Data_ID) LEFT JOIN tbl_Tags t ON t.Tag_ID = td.Tag_ID) LEFT JOIN"
        " tbl_Events e ON e.Event_ID = td.Event_ID) WHERE  t.Tag =[tbxTag] GROUP BY tv.Tr"
        "ee_Data_ID, t.Tag, YEAR(e.Event_Date) ORDER BY YEAR(e.Event_Date) DESC;"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x0000000000000000000000000000000000000000800700000100000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin OptionGroup
            BackStyle =1
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =600
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Width =1920
                    Height =230
                    FontWeight =700
                    ForeColor =9868950
                    Name ="Label0"
                    Caption ="V i n e  Counts "
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =230
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =5
                    Top =240
                    Width =1010
                    Height =230
                    FontWeight =700
                    ForeColor =9868950
                    Name ="lblCurrentYear"
                    Caption ="Sample Year"
                    LayoutCachedLeft =5
                    LayoutCachedTop =240
                    LayoutCachedWidth =1015
                    LayoutCachedHeight =470
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =1240
                    Top =240
                    Width =520
                    Height =230
                    FontWeight =700
                    ForeColor =9868950
                    Name ="Label4"
                    Caption ="Count"
                    LayoutCachedLeft =1240
                    LayoutCachedTop =240
                    LayoutCachedWidth =1760
                    LayoutCachedHeight =470
                End
                Begin TextBox
                    CanGrow = NotDefault
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =660
                    Top =540
                    Width =720
                    Height =0
                    ForeColor =9868950
                    Name ="tbxTag"

                    LayoutCachedLeft =660
                    LayoutCachedTop =540
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =540
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Width =720
                    Height =0
                    ForeColor =9868950
                    Name ="tbxSampleYear"
                    ControlSource ="SampleYear"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =780
                End
                Begin TextBox
                    CanGrow = NotDefault
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1140
                    Width =720
                    Height =0
                    TabIndex =1
                    ForeColor =9868950
                    Name ="tbxVineCount"
                    ControlSource ="VineCount"

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =1860
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
