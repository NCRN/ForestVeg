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
    Width =10380
    DatasheetFontHeight =10
    ItemSuffix =240
    Left =810
    Top =2640
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x72153ccbad16e440
    End
    RecordSource ="SELECT qFiltered_Locations.Plot_Name, tbl_Tags.Tag, tbl_Tags.Microplot_Number, t"
        "bl_Tags.Tag_Status FROM qFiltered_Locations INNER JOIN tbl_Tags ON qFiltered_Loc"
        "ations.Location_ID = tbl_Tags.Location_ID WHERE (((tbl_Tags.Tag_Status)=\"Saplin"
        "g\"));"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x60030000800400006003000040020000000000008c2800000b04000001000000 ,
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
        Begin Rectangle
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            BorderLineStyle =0
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin PageHeader
            Height =1095
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    Left =120
                    Top =1080
                    Width =10200
                    Name ="Line11"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Left =60
                    Top =60
                    Width =2820
                    Height =600
                    FontSize =24
                    FontWeight =700
                    Name ="Label19"
                    Caption ="Saplings"
                    FontName ="Calibri"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =660
                End
                Begin Label
                    TextFontFamily =34
                    Left =6000
                    Top =660
                    Width =4080
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label127"
                    Caption ="Date:______________________"
                    FontName ="Calibri"
                    LayoutCachedLeft =6000
                    LayoutCachedTop =660
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =1020
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =660
                    Width =5835
                    Height =360
                    ColumnOrder =0
                    FontSize =14
                    FontWeight =700
                    Name ="Label26"
                    ControlSource ="Plot_Name"
                    FontName ="Calibri"

                    LayoutCachedLeft =60
                    LayoutCachedTop =660
                    LayoutCachedWidth =5895
                    LayoutCachedHeight =1020
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1035
            Name ="Detail"
            Begin
                Begin Line
                    Left =120
                    Top =1020
                    Width =10140
                    Name ="Line10"
                End
                Begin Label
                    Left =5760
                    Top =120
                    Width =4290
                    Height =225
                    Name ="Label14"
                    Caption ="DBH (of each stem)_______________________________"
                    LayoutCachedLeft =5760
                    LayoutCachedTop =120
                    LayoutCachedWidth =10050
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =5760
                    Top =420
                    Width =960
                    Height =225
                    Name ="Label15"
                    Caption ="Browsable?"
                    LayoutCachedLeft =5760
                    LayoutCachedTop =420
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =645
                End
                Begin Rectangle
                    Left =6840
                    Top =480
                    Width =120
                    Height =120
                    Name ="Box38"
                    LayoutCachedLeft =6840
                    LayoutCachedTop =480
                    LayoutCachedWidth =6960
                    LayoutCachedHeight =600
                End
                Begin Label
                    Left =2160
                    Top =420
                    Width =3420
                    Height =225
                    Name ="Label22"
                    Caption ="Habit: TREE or SHRUB"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =420
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =645
                End
                Begin Label
                    Left =2160
                    Top =120
                    Width =3420
                    Height =225
                    Name ="Label24"
                    Caption ="Species:______________________________"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =120
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =345
                End
                Begin Label
                    Left =2160
                    Top =720
                    Width =7860
                    Height =225
                    Name ="Label126"
                    Caption ="Status: ALIVE Standing/Leaning/Broken/Fallen, DEAD,  MISSING, DOWNGRADED TO NON-"
                        "SAMPLED"
                    LayoutCachedLeft =2160
                    LayoutCachedTop =720
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =945
                End
                Begin Label
                    TextFontFamily =34
                    Left =7200
                    Top =420
                    Width =840
                    Height =225
                    Name ="Label128"
                    Caption ="Browsed?"
                    LayoutCachedLeft =7200
                    LayoutCachedTop =420
                    LayoutCachedWidth =8040
                    LayoutCachedHeight =645
                End
                Begin Rectangle
                    Left =8100
                    Top =480
                    Width =120
                    Height =120
                    Name ="Box129"
                    LayoutCachedLeft =8100
                    LayoutCachedTop =480
                    LayoutCachedWidth =8220
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =420
                    Width =1395
                    Height =300
                    FontSize =11
                    FontWeight =700
                    Name ="Label25"
                    ControlSource ="Tag"

                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =1455
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =1395
                    Height =225
                    TabIndex =1
                    Name ="Label23"
                    ControlSource ="=\"MP: \" & [Microplot_Number]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =1455
                    LayoutCachedHeight =345
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
