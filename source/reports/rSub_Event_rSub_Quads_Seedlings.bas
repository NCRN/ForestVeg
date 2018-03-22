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
    Width =5220
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =2040
    Top =1545
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf5157d435593e440
    End
    RecordSource ="SELECT tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID, StringFromGUID([Quadrat_Data_"
        "ID]) AS Quadrat_Data_txt, tbl_Quadrat_Seedlings_Data.TSN, tlu_Plants.Family, tlu"
        "_Plants.Genus, tlu_Plants.Species, tbl_Quadrat_Seedlings_Data.Height, [genus] & "
        "\" \" & [species] AS SciName, [Height] & \" cm\" AS Height_txt, tlu_Plants.Latin"
        "_Name, [Browsable] & \"/\" & [Browsed] AS Browse FROM tbl_Quadrat_Seedlings_Data"
        " LEFT JOIN tlu_Plants ON tbl_Quadrat_Seedlings_Data.TSN = tlu_Plants.TSN;"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x55010000f000000055010000f00000000000000064140000f000000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
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
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =120
                    Width =2520
                    Name ="Genus"
                    ControlSource ="Latin_Name"

                    LayoutCachedLeft =120
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3420
                    Width =720
                    TabIndex =1
                    Name ="txt_Height"
                    ControlSource ="Height_txt"

                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4140
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2640
                    Width =780
                    TabIndex =2
                    Name ="txtBrowse"
                    ControlSource ="Browse"

                    LayoutCachedLeft =2640
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =240
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
