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
    ItemSuffix =5
    Left =735
    Top =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe218a1201317e540
    End
    RecordSource ="SELECT tbl_Tree_Conditions.Tree_Data_ID, tbl_Tree_Conditions.Condition, stringfr"
        "omGUID([Tree_Data_ID]) AS Tree_Data_txt FROM tbl_Tree_Conditions;"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x55010000f000000055010000f0000000000000007c0b00003c00000001000000 ,
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
                Begin TextBox
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =2940
                    Height =225
                    Name ="tbxTreeConditionsHdrBgd"
                    ConditionalFormat = Begin
                        0x010000008c000000010000000100000000000000000000001500000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x43006f0075006e00740028005b0043006f006e0064006900740069006f006e00 ,
                        0x5d0029003d00300000000000
                    End

                    LayoutCachedWidth =2940
                    LayoutCachedHeight =225
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400140000004300 ,
                        0x6f0075006e00740028005b0043006f006e0064006900740069006f006e005d00 ,
                        0x29003d003000000000000000000000000000000000000000000000
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =-15
                    Width =2865
                    Height =225
                    FontWeight =700
                    ForeColor =9868950
                    Name ="lblTreeConditions"
                    Caption ="T r e e   C o n d i t i o n s"
                    LayoutCachedLeft =-15
                    LayoutCachedWidth =2850
                    LayoutCachedHeight =225
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
            Height =60
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Width =2760
                    Height =0
                    ForeColor =9868950
                    Name ="tbxCondition"
                    ControlSource ="Condition"
                    ConditionalFormat = Begin
                        0x0100000086000000010000000100000000000000000000001200000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b0074006200780043006f006e0064006900740069006f006e005d003d002200 ,
                        0x220000000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2820
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400110000005b00 ,
                        0x74006200780043006f006e0064006900740069006f006e005d003d0022002200 ,
                        0x000000000000000000000000000000000000000000
                    End
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
