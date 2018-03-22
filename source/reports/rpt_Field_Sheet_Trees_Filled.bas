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
    Width =10620
    DatasheetFontHeight =10
    ItemSuffix =187
    Left =1980
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x248c6545ac16e440
    End
    RecordSource ="SELECT qFiltered_Locations.Plot_Name, tbl_Tags.Tag, tbl_Tags.Azimuth, tbl_Tags.D"
        "istance, tbl_Tags.Tag_Status FROM qFiltered_Locations INNER JOIN tbl_Tags ON qFi"
        "ltered_Locations.Location_ID = tbl_Tags.Location_ID WHERE (((tbl_Tags.Tag_Status"
        ")=\"Tree\"));"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xd002000080040000d002000040020000000000007c2900006f09000001000000 ,
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
                Begin Line
                    Left =120
                    Top =1080
                    Width =10200
                    Name ="Line143"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1080
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =1080
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
                    Caption ="Trees"
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
            Height =2415
            Name ="Detail"
            Begin
                Begin Line
                    Left =120
                    Top =2400
                    Width =10140
                    Name ="Line10"
                End
                Begin Label
                    Left =6720
                    Top =60
                    Width =3660
                    Height =225
                    Name ="Label14"
                    Caption ="DBH of each stem: ________________________"
                    LayoutCachedLeft =6720
                    LayoutCachedTop =60
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =285
                End
                Begin Label
                    Left =180
                    Top =540
                    Width =10140
                    Height =225
                    Name ="Label16"
                    Caption ="STATUS: ALIVE Standing/Leaning/Broken/Fallen, DEAD,  MISSING, DOWNGRADED TO NON-"
                        "SAMPLED"
                    LayoutCachedLeft =180
                    LayoutCachedTop =540
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =765
                End
                Begin Label
                    Left =180
                    Top =840
                    Width =4020
                    Height =435
                    Name ="Label17"
                    Caption ="CROWN:  Codominant / Dominant / Intermediate / Open-grown / Overtopped / Light G"
                        "ap Exploiter / Edge Tree"
                    LayoutCachedLeft =180
                    LayoutCachedTop =840
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =1275
                End
                Begin Label
                    Left =4380
                    Top =840
                    Width =4500
                    Height =225
                    Name ="Label18"
                    Caption ="FOLIAGE CONDITIONS:  List Condition and Percent Afflicted"
                    LayoutCachedLeft =4380
                    LayoutCachedTop =840
                    LayoutCachedWidth =8880
                    LayoutCachedHeight =1065
                End
                Begin Label
                    TextAlign =1
                    Left =180
                    Top =2100
                    Width =10140
                    Height =225
                    Name ="Label39"
                    Caption ="Species of vine in tree: _______________________________________________________"
                        "______________________________________"
                    LayoutCachedLeft =180
                    LayoutCachedTop =2100
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =2325
                End
                Begin Label
                    TextAlign =1
                    Left =180
                    Top =1380
                    Width =10155
                    Height =645
                    Name ="Label40"
                    Caption ="TREE PESTS: Beech Bark Disease/ Butternut Canker/ Dogwood Antracnose/ Gypsy Moth"
                        "/Hemlock Scale/ Hemlock Wooly Adelgid/Other Insect Damage___________________ CON"
                        "DITIONS: Advanced Decay/ Primary branch broken/ Large dead brances/ Lightning da"
                        "mage/ Wind damage/ Open wound/ Vines in crown/ Other visible damage_____________"
                        "_______"
                    LayoutCachedLeft =180
                    LayoutCachedTop =1380
                    LayoutCachedWidth =10335
                    LayoutCachedHeight =2025
                End
                Begin Label
                    Left =2100
                    Top =60
                    Width =4560
                    Height =225
                    Name ="Label41"
                    Caption ="SPECIES: _________________________________________"
                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =6660
                    LayoutCachedHeight =285
                End
                Begin Label
                    Left =4380
                    Top =1080
                    Width =5925
                    Height =240
                    Name ="Label42"
                    Caption ="Chlorosis___ / Holes___ / Necrosis___ / Small Leaves___ / Wilting___ / Other___"
                    LayoutCachedLeft =4380
                    LayoutCachedTop =1080
                    LayoutCachedWidth =10305
                    LayoutCachedHeight =1320
                End
                Begin Label
                    TextFontFamily =34
                    Left =180
                    Top =300
                    Width =8325
                    Height =225
                    Name ="Label144"
                    Caption ="Please note any changes to tag distance, azimuth: ______________________________"
                        "___________"
                    LayoutCachedLeft =180
                    LayoutCachedTop =300
                    LayoutCachedWidth =8505
                    LayoutCachedHeight =525
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =780
                    Height =285
                    FontSize =11
                    FontWeight =700
                    Name ="txtTag"
                    ControlSource ="Tag"
                    FontName ="Calibri"

                    LayoutCachedWidth =780
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Left =840
                    Width =1200
                    Height =285
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    Name ="Text186"
                    ControlSource ="=[Azimuth] & \"°  \" & [Distance] & \"m\""
                    FontName ="Calibri"

                    LayoutCachedLeft =840
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =285
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
