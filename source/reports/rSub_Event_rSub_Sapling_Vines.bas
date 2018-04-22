Version =20
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
    Width =2940
    DatasheetFontHeight =9
    ItemSuffix =3
    Left =705
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbab7fa62d393e440
    End
    RecordSource ="SELECT tbl_Sapling_Vines.Sapling_Data_ID, tlu_Plants.Latin_Name, StringFromGUID("
        "[Sapling_Data_ID]) AS Sapling_Data_txt FROM tlu_Plants INNER JOIN tbl_Sapling_Vi"
        "nes ON tlu_Plants.TSN = tbl_Sapling_Vines.TSN;"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e00100006801000068010000000000007c0b0000b400000001000000 ,
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
            Height =225
            Name ="ReportHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =-15
                    Width =2865
                    Height =225
                    FontWeight =700
                    ForeColor =6250335
                    Name ="lblHdr"
                    Caption ="V i n e s"
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
            Height =180
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Width =2760
                    Height =120
                    ForeColor =6250335
                    Name ="tbxLatinName"
                    ControlSource ="Latin_Name"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =120
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
