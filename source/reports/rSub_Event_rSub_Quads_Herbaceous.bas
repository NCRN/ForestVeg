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
    Width =5460
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =2445
    Top =2325
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xae3757a21093e440
    End
    RecordSource ="SELECT tbl_Quadrat_Herbaceous_Data.Quadrat_Data_ID, StringFromGUID([Quadrat_Data"
        "_ID]) AS Quadrat_Data_txt, tbl_Quadrat_Herbaceous_Data.TSN, tbl_Quadrat_Herbaceo"
        "us_Data.Percent_Cover, [Percent_Cover] & \" %\" AS Perc_Cover_txt, tlu_Plants.La"
        "tin_Name, tbl_Quadrat_Herbaceous_Data.Browse FROM tbl_Quadrat_Herbaceous_Data LE"
        "FT JOIN tlu_Plants ON tbl_Quadrat_Herbaceous_Data.TSN = tlu_Plants.TSN;"
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
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =120
                    Width =2820
                    Name ="txtLatin_Name"
                    ControlSource ="Latin_Name"
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b007400780074004c006100740069006e00 ,
                        0x5f004e0061006d0065005d00290000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400170000004900 ,
                        0x73004e0075006c006c0028005b007400780074004c006100740069006e005f00 ,
                        0x4e0061006d0065005d0029000000000000000000000000000000000000000000 ,
                        0x00
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3120
                    Width =780
                    TabIndex =1
                    Name ="Percent_Cover"
                    ControlSource ="Perc_Cover_txt"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0050006500720063005f0043006f007600650072005f007400780074005d00 ,
                        0x3d002200300020002500220000000000
                    End

                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400160000005b00 ,
                        0x50006500720063005f0043006f007600650072005f007400780074005d003d00 ,
                        0x22003000200025002200000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4140
                    Width =660
                    TabIndex =2
                    Name ="txtBrowse"
                    ControlSource ="Browse"
                    ConditionalFormat = Begin
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00740078007400420072006f0077007300 ,
                        0x65005d00290000000000
                    End

                    LayoutCachedLeft =4140
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400130000004900 ,
                        0x73004e0075006c006c0028005b00740078007400420072006f00770073006500 ,
                        0x5d002900000000000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
