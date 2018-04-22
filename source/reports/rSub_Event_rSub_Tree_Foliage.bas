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
    Width =2865
    DatasheetFontHeight =9
    ItemSuffix =4
    Left =735
    Top =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xcea3444fd408e340
    End
    RecordSource ="SELECT tbl_Tree_Foliage_Conditions.Tree_Data_ID, stringfromGUID([Tree_Data_ID]) "
        "AS Tree_Data_txt, tbl_Tree_Foliage_Conditions.Condition, tbl_Tree_Foliage_Condit"
        "ions.Percent_Afflicted, [Condition] & \"  \" & [Percent_Afflicted] AS Cond_Perce"
        "nt FROM tbl_Tree_Foliage_Conditions; "
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xe0010000e0010000680100006801000000000000310b00000e00000001000000 ,
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
                    OverlapFlags =4
                    TextAlign =2
                    TextFontFamily =34
                    Left =-30
                    Width =2895
                    Height =225
                    FontWeight =700
                    ForeColor =6250335
                    Name ="lblHdr"
                    Caption ="F o l i a g e   C o n d i t i o n s"
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
            Height =14
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    OverlapFlags =4
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Width =1920
                    Height =0
                    ForeColor =6250335
                    Name ="tbxPercent"
                    ControlSource ="Cond_Percent"

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
