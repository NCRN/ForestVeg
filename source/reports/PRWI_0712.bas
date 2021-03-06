﻿Version =21
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
    Width =11160
    DatasheetFontHeight =10
    ItemSuffix =62
    Left =3495
    Top =765
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1d220c6fb4dfe340
    End
    RecordSource ="SELECT qRpt_Event_Summary_Unfiltered.*\015\012FROM qRpt_Event_Summary_Unfiltered"
        "\015\012WHERE ([Event_ID]='{4B8624FC-39D6-4EE4-AA7A-E008C0326886}');"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xf801000038040000f80100003804000000000000982b00007413000001000000 ,
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin Subform
            BorderLineStyle =0
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =4980
            Name ="Detail"
            Begin
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =2475
                    Width =10800
                    Height =0
                    Name ="rSub_Plot_Floor_Condition"
                    SourceObject ="Report.rSub_Event_Plot_Floor_Condition"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =2475
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =2475
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =3840
                    Width =10800
                    Height =0
                    TabIndex =1
                    Name ="rSub_Event_Trees"
                    SourceObject ="Report.rSub_Event_Trees"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =3840
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =3840
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =2925
                    Width =10800
                    Height =0
                    TabIndex =2
                    Name ="srpt_Transects"
                    SourceObject ="Report.rSub_Event_CWD"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =2925
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =2925
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =4320
                    Width =10800
                    Height =0
                    TabIndex =3
                    Name ="srpt_Microplots"
                    SourceObject ="Report.rSub_Event_Saplings"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =4320
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =4320
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =4980
                    Width =10800
                    Height =0
                    TabIndex =4
                    Name ="rSub_Event_Quadrats"
                    SourceObject ="Report.rSub_Event_Quadrats"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =4980
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =4980
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =1380
                    Width =10800
                    Height =719
                    TabIndex =5
                    Name ="rSub_Event_Notes"
                    SourceObject ="Report.rSub_Event_Notes"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                    LayoutCachedTop =1380
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =2099
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Top =3090
                    Width =2400
                    Height =225
                    FontWeight =700
                    Name ="Label22"
                    Caption ="TRANSECTS CHECKED --->"
                    LayoutCachedTop =3090
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =3315
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Left =5490
                    Top =60
                    Width =5310
                    Height =1257
                    TabIndex =6
                    Name ="rSub_Event_Participants"
                    SourceObject ="Report.rSub_Event_Participants"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedLeft =5490
                    LayoutCachedTop =60
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =1317
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2820
                    Width =2175
                    FontSize =9
                    TabIndex =7
                    ForeColor =7633277
                    Name ="txtAdmin_Unit_Code"
                    ControlSource ="=\"Administered by \" & [Admin_Unit_Code]"
                    StatusBarText ="Unit Code of the park that manages this location"
                    FontName ="Calibri"

                    LayoutCachedLeft =2820
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2820
                    Top =480
                    Width =2175
                    FontSize =9
                    TabIndex =8
                    ForeColor =7633277
                    Name ="txtUTM_Coordinates"
                    ControlSource ="=\"UTM: \" & [X_Coord] & \", \" & [Y_Coord]"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    FontName ="Calibri"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =480
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =720
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2820
                    Top =240
                    Width =2175
                    FontSize =9
                    TabIndex =9
                    ForeColor =7633277
                    Name ="txtPanelAndFrame"
                    ControlSource ="=\"Panel: \" & [Panel] & \", Frame: \" & [Frame]"
                    StatusBarText ="Sampling Panel Number"
                    FontName ="Calibri"

                    LayoutCachedLeft =2820
                    LayoutCachedTop =240
                    LayoutCachedWidth =4995
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextFontCharSet =238
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Width =2685
                    Height =480
                    FontSize =22
                    FontWeight =700
                    TabIndex =10
                    Name ="Text5"
                    ControlSource ="=[Plot_Name]"
                    FontName ="Calibri"

                    LayoutCachedWidth =2685
                    LayoutCachedHeight =480
                End
                Begin TextBox
                    TextFontCharSet =204
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =0
                    IMESentenceMode =3
                    Top =525
                    Width =2685
                    Height =480
                    FontSize =18
                    FontWeight =700
                    TabIndex =11
                    Name ="Text8"
                    ControlSource ="Event_Date"
                    Format ="mm/dd/yyyy"
                    FontName ="Calibri"

                    LayoutCachedTop =525
                    LayoutCachedWidth =2685
                    LayoutCachedHeight =1005
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =238
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3480
                    Top =1080
                    Width =1065
                    Height =225
                    ColumnWidth =1245
                    TabIndex =12
                    Name ="txtEvent_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"
                    FontName ="Calibri"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =1080
                    LayoutCachedWidth =4545
                    LayoutCachedHeight =1305
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4605
                    Top =1080
                    Width =900
                    Height =225
                    TabIndex =13
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Location identifier (Location_ID)"
                    FontName ="Calibri"

                    LayoutCachedLeft =4605
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5505
                    LayoutCachedHeight =1305
                End
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextFontCharSet =238
                    TextAlign =1
                    TextFontFamily =34
                    Top =2580
                    Width =5940
                    Height =345
                    FontSize =14
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Coarse Woody Debris"
                    FontName ="Calibri"
                    LayoutCachedTop =2580
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =2925
                End
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextFontCharSet =238
                    TextAlign =1
                    TextFontFamily =34
                    Top =2115
                    Width =5475
                    Height =375
                    FontSize =14
                    FontWeight =700
                    Name ="Label41"
                    Caption ="Forest Floor Conditions and Plot Observations"
                    FontName ="Calibri"
                    LayoutCachedTop =2115
                    LayoutCachedWidth =5475
                    LayoutCachedHeight =2490
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =1020
                    Width =240
                    Height =259
                    FontSize =12
                    FontWeight =700
                    TabIndex =14
                    Name ="txtPictures_Taken"
                    ControlSource ="=IIf([Pictures_Taken],\"X\",\"\")"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x0100000090000000010000000100000000000000000000001700000001010000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00500069006300740075007200650073005f00540061006b0065006e005d00 ,
                        0x3d00460061006c007300650000000000
                    End

                    LayoutCachedLeft =60
                    LayoutCachedTop =1020
                    LayoutCachedWidth =300
                    LayoutCachedHeight =1279
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400160000005b00 ,
                        0x500069006300740075007200650073005f00540061006b0065006e005d003d00 ,
                        0x460061006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =360
                            Top =1020
                            Width =1260
                            Height =225
                            FontSize =10
                            Name ="lblPictures_Taken"
                            Caption ="Pictures Taken"
                            FontName ="Calibri"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1020
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =1245
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2580
                    Top =3075
                    Width =240
                    Height =259
                    FontSize =12
                    FontWeight =700
                    TabIndex =15
                    Name ="txtCWD_360"
                    ControlSource ="=IIf([CWD_Check_360],\"X\",\"\")"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001010000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004300570044005f0043006800650063006b005f003300360030005d003d00 ,
                        0x460061006c007300650000000000
                    End

                    LayoutCachedLeft =2580
                    LayoutCachedTop =3075
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =3334
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400150000005b00 ,
                        0x4300570044005f0043006800650063006b005f003300360030005d003d004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =2880
                            Top =3089
                            Width =360
                            Height =225
                            FontSize =10
                            Name ="lblCWD_360"
                            Caption ="360"
                            FontName ="Calibri"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =3089
                            LayoutCachedWidth =3240
                            LayoutCachedHeight =3314
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3420
                    Top =3075
                    Width =240
                    Height =259
                    FontSize =12
                    FontWeight =700
                    TabIndex =16
                    Name ="txtCWD_240"
                    ControlSource ="=IIf([CWD_Check_120],\"X\",\"\")"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001010000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004300570044005f0043006800650063006b005f003100320030005d003d00 ,
                        0x460061006c007300650000000000
                    End

                    LayoutCachedLeft =3420
                    LayoutCachedTop =3075
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =3334
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400150000005b00 ,
                        0x4300570044005f0043006800650063006b005f003100320030005d003d004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =3720
                            Top =3089
                            Width =360
                            Height =225
                            FontSize =10
                            Name ="lblCWD_240"
                            Caption ="120"
                            FontName ="Calibri"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =3089
                            LayoutCachedWidth =4080
                            LayoutCachedHeight =3314
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =3075
                    Width =240
                    Height =259
                    FontSize =12
                    FontWeight =700
                    TabIndex =17
                    Name ="txtCWD_120"
                    ControlSource ="=IIf([CWD_Check_240],\"X\",\"\")"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001010000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b004300570044005f0043006800650063006b005f003200340030005d003d00 ,
                        0x460061006c007300650000000000
                    End

                    LayoutCachedLeft =4260
                    LayoutCachedTop =3075
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =3334
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400150000005b00 ,
                        0x4300570044005f0043006800650063006b005f003200340030005d003d004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =4560
                            Top =3089
                            Width =360
                            Height =225
                            FontSize =10
                            Name ="lblCWD_120"
                            Caption ="240"
                            FontName ="Calibri"
                            LayoutCachedLeft =4560
                            LayoutCachedTop =3089
                            LayoutCachedWidth =4920
                            LayoutCachedHeight =3314
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Top =3420
                    Width =2295
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="lblHeading"
                    Caption ="Trees"
                    FontName ="Calibri"
                    LayoutCachedTop =3420
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =3780
                End
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Top =3960
                    Width =3195
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label52"
                    Caption ="Saplings"
                    FontName ="Calibri"
                    LayoutCachedTop =3960
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =4350
                End
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    Top =4500
                    Width =4200
                    Height =390
                    FontSize =14
                    FontWeight =700
                    Name ="Label53"
                    Caption ="Seedlings and Herbaceous"
                    FontName ="Calibri"
                    LayoutCachedTop =4500
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =4890
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2940
                    Top =1020
                    Width =300
                    Height =285
                    FontSize =11
                    FontWeight =700
                    TabIndex =18
                    Name ="Deer_Impact"
                    ControlSource ="Deer_Impact"
                    StatusBarText ="Deer impact classification (1-5)"
                    FontName ="Calibri"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000100000000000000000000001600000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0044006500650072005f0049006d007000 ,
                        0x6100630074005d00290000000000
                    End

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3240
                    LayoutCachedHeight =1305
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400150000004900 ,
                        0x73004e0075006c006c0028005b0044006500650072005f0049006d0070006100 ,
                        0x630074005d002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =1575
                            Top =1020
                            Width =1305
                            Height =285
                            FontSize =11
                            Name ="Label57"
                            Caption ="Deer Impact: "
                            FontName ="Calibri"
                            LayoutCachedLeft =1575
                            LayoutCachedTop =1020
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =1305
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3360
                    Top =720
                    Width =345
                    Height =225
                    TabIndex =19
                    ForeColor =8355711
                    Name ="txtSlope"
                    ControlSource ="Slope"
                    ConditionalFormat = Begin
                        0x0100000088000000010000000100000000000000000000001300000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740053006c006f0070006500 ,
                        0x5d00290000000000
                    End

                    LayoutCachedLeft =3360
                    LayoutCachedTop =720
                    LayoutCachedWidth =3705
                    LayoutCachedHeight =945
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400120000004900 ,
                        0x73004e0075006c006c0028005b0074007800740053006c006f00700065005d00 ,
                        0x2900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            TextFontFamily =34
                            Left =2820
                            Top =720
                            Width =480
                            Height =225
                            ForeColor =8355711
                            Name ="Label59"
                            Caption ="Slope:"
                            LayoutCachedLeft =2820
                            LayoutCachedTop =720
                            LayoutCachedWidth =3300
                            LayoutCachedHeight =945
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =1
                            ForeShade =50.0
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4440
                    Top =720
                    Width =345
                    Height =225
                    TabIndex =20
                    ForeColor =8355711
                    Name ="txtAspect"
                    ControlSource ="Aspect"
                    ConditionalFormat = Begin
                        0x010000008a000000010000000100000000000000000000001400000001000000 ,
                        0x00000000ed1c2400000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740041007300700065006300 ,
                        0x74005d00290000000000
                    End

                    LayoutCachedLeft =4440
                    LayoutCachedTop =720
                    LayoutCachedWidth =4785
                    LayoutCachedHeight =945
                    BackThemeColorIndex =1
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000ed1c2400130000004900 ,
                        0x73004e0075006c006c0028005b00740078007400410073007000650063007400 ,
                        0x5d002900000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            TextFontFamily =34
                            Left =3750
                            Top =720
                            Width =645
                            Height =225
                            ForeColor =8355711
                            Name ="Label61"
                            Caption ="Aspect:"
                            LayoutCachedLeft =3750
                            LayoutCachedTop =720
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =945
                            BackThemeColorIndex =1
                            ForeThemeColorIndex =1
                            ForeShade =50.0
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =420
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8880
                    Top =180
                    Width =2280
                    ForeColor =8421504
                    Name ="Text18"
                    ControlSource ="=Now()"

                    LayoutCachedLeft =8880
                    LayoutCachedTop =180
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =420
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            TextFontFamily =34
                            Left =7980
                            Top =180
                            Width =900
                            Height =225
                            ForeColor =8421504
                            Name ="Label19"
                            Caption ="Printed on:"
                            LayoutCachedLeft =7980
                            LayoutCachedTop =180
                            LayoutCachedWidth =8880
                            LayoutCachedHeight =405
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Top =180
                    Width =1380
                    TabIndex =1
                    ForeColor =8421504
                    Name ="Text54"
                    ControlSource ="Plot_Name"

                    LayoutCachedTop =180
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1440
                    Top =180
                    Width =2520
                    TabIndex =2
                    ForeColor =8421504
                    Name ="Text56"
                    ControlSource ="Event_Date"

                    LayoutCachedLeft =1440
                    LayoutCachedTop =180
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
