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
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =705
    Top =3015
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x05acb872f5b1e340
    End
    RecordSource ="qry_Srpt_Transects"
    Caption ="sfrm_Plot_Floor_Condition_Data subreport"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x55010000f000000055010000f000000000000000f4290000f000000001000000 ,
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
        Begin BreakLevel
            ControlSource ="Transect_Azimuth"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =300
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Width =960
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Rock_Cover_Label"
                    Caption ="Transect"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedWidth =960
                    LayoutCachedHeight =300
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =1080
                    Width =1215
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Bare_Soil_Cover_Label"
                    Caption ="Decay Class"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1080
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =300
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =2340
                    Width =960
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Trampled_Label"
                    Caption ="Diameter"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =2340
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =300
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =3360
                    Width =750
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Humus_Label"
                    Caption ="Hollow"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4110
                    LayoutCachedHeight =300
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =5580
                    Width =1485
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Earthworms_Label"
                    Caption ="Latin Name"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5580
                    LayoutCachedWidth =7065
                    LayoutCachedHeight =300
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    Left =4440
                    Width =750
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label22"
                    Caption ="Tag"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4440
                    LayoutCachedWidth =5190
                    LayoutCachedHeight =300
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontFamily =34
                    Left =7500
                    Width =2370
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label24"
                    Caption ="Coarse Woody Debris Notes"
                    FontName ="Calibri"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =7500
                    LayoutCachedWidth =9870
                    LayoutCachedHeight =300
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
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Width =900
                    ColumnWidth =1755
                    FontSize =9
                    Name ="Transect_Azimuth"
                    ControlSource ="Transect_Azimuth"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000002010000010000000100000000000000000000005000000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b005400720061006e0073006500630074005f0041007a0069006d0075007400 ,
                        0x68005d003c003e00310032003000200041006e00640020005b00540072006100 ,
                        0x6e0073006500630074005f0041007a0069006d007500740068005d003c003e00 ,
                        0x320034003000200041006e00640020005b005400720061006e00730065006300 ,
                        0x74005f0041007a0069006d007500740068005d003c003e003300360030000000 ,
                        0x0000
                    End

                    LayoutCachedWidth =900
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c24004f0000005b00 ,
                        0x5400720061006e0073006500630074005f0041007a0069006d00750074006800 ,
                        0x5d003c003e00310032003000200041006e00640020005b005400720061006e00 ,
                        0x73006500630074005f0041007a0069006d007500740068005d003c003e003200 ,
                        0x34003000200041006e00640020005b005400720061006e007300650063007400 ,
                        0x5f0041007a0069006d007500740068005d003c003e0033003600300000000000 ,
                        0x0000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Width =1260
                    FontSize =9
                    TabIndex =1
                    Name ="Decay_Class"
                    ControlSource ="Decay_Class"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0x00000000cf7b7900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00440065006300610079005f0043006c00 ,
                        0x6100730073005d00290000000000
                    End

                    LayoutCachedLeft =1020
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000cf7b7900150000004900 ,
                        0x73004e0075006c006c0028005b00440065006300610079005f0043006c006100 ,
                        0x730073005d002900000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2340
                    Width =1020
                    FontSize =9
                    TabIndex =2
                    Name ="Diameter"
                    ControlSource ="Diameter"
                    StatusBarText ="The diameter of the debris at the intersection of the transect."
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b004400690061006d006500740065007200 ,
                        0x5d00290000000000
                    End

                    LayoutCachedLeft =2340
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =240
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400120000004900 ,
                        0x73004e0075006c006c0028005b004400690061006d0065007400650072005d00 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    Left =3600
                    TabIndex =3
                    Name ="Hollow"
                    ControlSource ="Hollow"
                    StatusBarText ="Considered hollow if cavity extends 0.5m along the central longitudinal axis of "
                        "the piece and the cavity entrance is at least 1/4 the diameter of the piece."

                    LayoutCachedLeft =3600
                    LayoutCachedWidth =3860
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextFontCharSet =238
                    IMESentenceMode =3
                    Left =5520
                    Width =2040
                    FontSize =9
                    TabIndex =4
                    Name ="LatinName"
                    ControlSource ="Latin_Name"
                    FontName ="Calibri"

                    LayoutCachedLeft =5520
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4260
                    Width =1020
                    FontSize =9
                    TabIndex =5
                    Name ="Text21"
                    ControlSource ="Tag"
                    StatusBarText ="The diameter of the debris at the intersection of the transect."
                    FontName ="Calibri"

                    LayoutCachedLeft =4260
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7500
                    Width =3180
                    FontSize =9
                    TabIndex =6
                    Name ="Text25"
                    ControlSource ="CWD_Notes"
                    FontName ="Calibri"

                    LayoutCachedLeft =7500
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =240
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
