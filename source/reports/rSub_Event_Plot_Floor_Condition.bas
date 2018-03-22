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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =24
    Left =2130
    Top =4350
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x92fab1cccd93e440
    End
    RecordSource ="SELECT tbl_Plot_Floor_Condition_Data.Event_ID, tbl_Plot_Floor_Condition_Data.Roc"
        "k_Cover, tbl_Plot_Floor_Condition_Data.Bare_Soil_Cover, tbl_Plot_Floor_Condition"
        "_Data.Trampled, tbl_Plot_Floor_Condition_Data.Humus, tbl_Plot_Floor_Condition_Da"
        "ta.Earthworms, tbl_Events.Early_Detect, tbl_Events.Rare_Spp, tbl_Events.Plot_Mai"
        "nt FROM tbl_Events INNER JOIN tbl_Plot_Floor_Condition_Data ON tbl_Events.Event_"
        "ID = tbl_Plot_Floor_Condition_Data.Event_ID;"
    Caption ="sfrm_Plot_Floor_Condition_Data subreport"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x55010000f000000055010000f000000000000000302a00006801000001000000 ,
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
            Height =240
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Width =960
                    Height =240
                    FontSize =10
                    ForeColor =0
                    Name ="Rock_Cover_Label"
                    Caption ="Rock"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =960
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =1080
                    Width =1020
                    Height =240
                    FontSize =10
                    ForeColor =0
                    Name ="Bare_Soil_Cover_Label"
                    Caption ="Bare Soil"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1080
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =2220
                    Width =960
                    Height =240
                    FontSize =10
                    ForeColor =0
                    Name ="Trampled_Label"
                    Caption ="Trampled"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2220
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =6300
                    Width =1560
                    Height =240
                    FontSize =10
                    ForeColor =0
                    Name ="lblPlotMaint"
                    Caption ="Plot Maintenance"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =6300
                    LayoutCachedWidth =7860
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =4920
                    Width =1140
                    Height =240
                    FontSize =10
                    ForeColor =0
                    Name ="lblRareSpp"
                    Caption ="Rare Species"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4920
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =240
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =3360
                    Width =1320
                    Height =240
                    FontSize =10
                    ForeColor =0
                    Name ="lblEarlyDetcect"
                    Caption ="Early Detection"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =240
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
                    Name ="Line12"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =360
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Top =60
                    Width =900
                    Height =270
                    FontSize =10
                    Name ="Rock_Cover"
                    ControlSource ="Rock_Cover"
                    StatusBarText ="Percent of the plot covered by rocks"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000008c000000010000000100000000000000000000001500000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0052006f0063006b005f0043006f007600 ,
                        0x650072005d00290000000000
                    End

                    LayoutCachedTop =60
                    LayoutCachedWidth =900
                    LayoutCachedHeight =330
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400140000004900 ,
                        0x73004e0075006c006c0028005b0052006f0063006b005f0043006f0076006500 ,
                        0x72005d002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1080
                    Top =60
                    Width =1020
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="Bare_Soil_Cover"
                    ControlSource ="Bare_Soil_Cover"
                    StatusBarText ="Percent of the plot covered by bare soil"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000096000000010000000100000000000000000000001a00000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0042006100720065005f0053006f006900 ,
                        0x6c005f0043006f007600650072005d00290000000000
                    End

                    LayoutCachedLeft =1080
                    LayoutCachedTop =60
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =330
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ffffff00190000004900 ,
                        0x73004e0075006c006c0028005b0042006100720065005f0053006f0069006c00 ,
                        0x5f0043006f007600650072005d00290000000000000000000000000000000000 ,
                        0x0000000000
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2280
                    Top =60
                    Width =900
                    Height =270
                    FontSize =10
                    TabIndex =2
                    Name ="Trampled"
                    ControlSource ="Trampled"
                    StatusBarText ="Percent of the plot trampled"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b005400720061006d0070006c0065006400 ,
                        0x5d00290000000000
                    End

                    LayoutCachedLeft =2280
                    LayoutCachedTop =60
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =330
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400120000004900 ,
                        0x73004e0075006c006c0028005b005400720061006d0070006c00650064005d00 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    Left =3900
                    Top =60
                    Width =300
                    Height =300
                    TabIndex =3
                    Name ="chkEarlyDetect"
                    ControlSource ="Early_Detect"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =60
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    Left =6960
                    Top =60
                    Width =300
                    Height =300
                    TabIndex =4
                    Name ="chkPlotMaint"
                    ControlSource ="Plot_Maint"

                    LayoutCachedLeft =6960
                    LayoutCachedTop =60
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =360
                End
                Begin CheckBox
                    Left =5280
                    Top =60
                    Width =300
                    Height =300
                    TabIndex =5
                    Name ="chkRareSpp"
                    ControlSource ="Rare_Spp"

                    LayoutCachedLeft =5280
                    LayoutCachedTop =60
                    LayoutCachedWidth =5580
                    LayoutCachedHeight =360
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
                    Name ="Line13"
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
