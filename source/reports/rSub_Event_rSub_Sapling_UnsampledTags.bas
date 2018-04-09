﻿Version =20
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
    Width =4320
    DatasheetFontHeight =9
    ItemSuffix =9
    Left =2325
    Top =3165
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9d7646507217e540
    End
    RecordSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag, tbl_Tags.Tag_Status AS Class, tbl_Tags.Mic"
        "roplot_Number AS MP, IIf(IsNull([azimuth]),\"\",[Azimuth] & \" / \" & [distance]"
        " & \"m\") AS Azi_Dist, tbl_Tree_Data.Tree_Data_ID, tbl_Sapling_Data.Sapling_Data"
        "_ID FROM (((tbl_Tags LEFT JOIN qry_Status_Sapling_Current_Event ON tbl_Tags.Tag_"
        "ID = qry_Status_Sapling_Current_Event.Tag_ID) LEFT JOIN qry_Status_Tree_Current_"
        "Event ON tbl_Tags.Tag_ID = qry_Status_Tree_Current_Event.Tag_ID) INNER JOIN tbl_"
        "Tree_Data ON tbl_Tags.Tag_ID = tbl_Tree_Data.Tag_ID) INNER JOIN tbl_Sapling_Data"
        " ON tbl_Tags.Tag_ID = tbl_Sapling_Data.Tag_ID WHERE (((qry_Status_Sapling_Curren"
        "t_Event.Event_ID) Is Null) AND ((tbl_Tags.Location_ID)=[Forms]![frm_Events]![Loc"
        "ation_ID]) AND ((qry_Status_Tree_Current_Event.Event_ID) Is Null)) ORDER BY tbl_"
        "Tags.Tag_Status, tbl_Tags.Tag;"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x55010000f000000055010000f000000000000000e01000003c00000001000000 ,
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
                    Width =840
                    Height =225
                    FontWeight =700
                    ForeColor =5855577
                    Name ="lblHdrTag"
                    Caption ="Tag"
                    LayoutCachedWidth =840
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =900
                    Width =1440
                    Height =225
                    FontWeight =700
                    ForeColor =5855577
                    Name ="lblHdrClass"
                    Caption ="Class"
                    LayoutCachedLeft =900
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =2580
                    Width =840
                    Height =225
                    FontWeight =700
                    ForeColor =5855577
                    Name ="lblHdrAziDist"
                    Caption ="Azi/Dist"
                    LayoutCachedLeft =2580
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =34
                    Left =3420
                    Width =840
                    Height =225
                    FontWeight =700
                    ForeColor =5855577
                    Name ="lblHdrMP"
                    Caption ="MP"
                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =225
                    ForeThemeColorIndex =0
                    ForeTint =65.0
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
                    Width =780
                    Height =0
                    ForeColor =4210752
                    Name ="tbxTag"
                    ControlSource ="Tag"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =840
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =900
                    Height =0
                    TabIndex =1
                    ForeColor =4210752
                    Name ="tbxClass"
                    ControlSource ="Class"

                    LayoutCachedLeft =900
                    LayoutCachedWidth =2340
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2520
                    Width =780
                    Height =0
                    TabIndex =2
                    ForeColor =4210752
                    Name ="tbxAziDist"
                    ControlSource ="Azi_Dist"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =3300
                    ForeThemeColorIndex =0
                    ForeTint =75.0
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3420
                    Width =780
                    Height =0
                    TabIndex =3
                    ForeColor =4210752
                    Name ="tbxMP"
                    ControlSource ="MP"

                    LayoutCachedLeft =3420
                    LayoutCachedWidth =4200
                    ForeThemeColorIndex =0
                    ForeTint =75.0
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
