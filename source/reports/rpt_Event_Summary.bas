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
    ItemSuffix =56
    Left =2820
    Top =525
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x41ab439f33b2e340
    End
    RecordSource ="qRpt_Event_Summary"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xf801000038040000f80100003804000000000000302a0000cf21000001000000 ,
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
            Height =8655
            Name ="Detail"
            Begin
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =2475
                    Width =10800
                    Height =629
                    Name ="rSub_Plot_Floor_Condition"
                    SourceObject ="Report.rSub_Event_Plot_Floor_Condition"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =2475
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =3104
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =4440
                    Width =10800
                    Height =1095
                    TabIndex =1
                    Name ="rSub_Event_Trees"
                    SourceObject ="Report.rSub_Event_Trees"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =4440
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =5535
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =3450
                    Width =10800
                    Height =674
                    TabIndex =2
                    Name ="srpt_Transects"
                    SourceObject ="Report.rSub_Event_CWD"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =3450
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =4124
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =5550
                    Width =10680
                    Height =1335
                    TabIndex =3
                    Name ="srpt_Microplots"
                    SourceObject ="Report.rSub_Event_Saplings"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =5550
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =6885
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =6915
                    Width =10680
                    Height =1740
                    TabIndex =4
                    Name ="rSub_Event_Quadrats"
                    SourceObject ="Report.rSub_Event_Quadrats"
                    LinkChildFields ="Event_ID"
                    LinkMasterFields ="Event_ID"

                    LayoutCachedTop =6915
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =8655
                End
                Begin Subform
                    Locked = NotDefault
                    OldBorderStyle =0
                    Top =1440
                    Width =10800
                    Height =719
                    TabIndex =5
                    Name ="rSub_Event_Notes"
                    SourceObject ="Report.rSub_Event_Notes"
                    LinkChildFields ="Location_ID"
                    LinkMasterFields ="Location_ID"

                    LayoutCachedTop =1440
                    LayoutCachedWidth =10800
                    LayoutCachedHeight =2159
                End
                Begin Label
                    TextFontCharSet =186
                    TextAlign =1
                    TextFontFamily =34
                    Top =4155
                    Width =2400
                    Height =225
                    FontWeight =700
                    Name ="Label22"
                    Caption ="TRANSECTS CHECKED --->"
                    LayoutCachedTop =4155
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =4380
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
                    Left =2940
                    Top =1140
                    Width =1065
                    Height =225
                    ColumnWidth =1245
                    TabIndex =12
                    Name ="txtEvent_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"
                    FontName ="Calibri"

                    LayoutCachedLeft =2940
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4005
                    LayoutCachedHeight =1365
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =238
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4065
                    Top =1140
                    Width =900
                    Height =225
                    TabIndex =13
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Location identifier (Location_ID)"
                    FontName ="Calibri"

                    LayoutCachedLeft =4065
                    LayoutCachedTop =1140
                    LayoutCachedWidth =4965
                    LayoutCachedHeight =1365
                End
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextFontCharSet =238
                    TextAlign =1
                    TextFontFamily =34
                    Top =3105
                    Width =5940
                    Height =345
                    FontSize =14
                    FontWeight =700
                    Name ="Label14"
                    Caption ="Coarse Woody Debris"
                    FontName ="Calibri"
                    LayoutCachedTop =3105
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =3450
                End
                Begin Label
                    FontItalic = NotDefault
                    FontUnderline = NotDefault
                    TextFontCharSet =238
                    TextAlign =1
                    TextFontFamily =34
                    Top =2115
                    Width =3180
                    Height =360
                    FontSize =14
                    FontWeight =700
                    Name ="Label41"
                    Caption ="Forest Floor Conditions"
                    FontName ="Calibri"
                    LayoutCachedTop =2115
                    LayoutCachedWidth =3180
                    LayoutCachedHeight =2475
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =60
                    Top =1080
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
                    LayoutCachedTop =1080
                    LayoutCachedWidth =300
                    LayoutCachedHeight =1339
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400160000005b00 ,
                        0x500069006300740075007200650073005f00540061006b0065006e005d003d00 ,
                        0x460061006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =360
                            Top =1094
                            Width =1260
                            Height =225
                            FontSize =10
                            Name ="lblPictures_Taken"
                            Caption ="Pictures Taken"
                            FontName ="Calibri"
                            LayoutCachedLeft =360
                            LayoutCachedTop =1094
                            LayoutCachedWidth =1620
                            LayoutCachedHeight =1319
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2580
                    Top =4140
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
                    LayoutCachedTop =4140
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =4399
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400150000005b00 ,
                        0x4300570044005f0043006800650063006b005f003300360030005d003d004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =2880
                            Top =4154
                            Width =360
                            Height =225
                            FontSize =10
                            Name ="lblCWD_360"
                            Caption ="360"
                            FontName ="Calibri"
                            LayoutCachedLeft =2880
                            LayoutCachedTop =4154
                            LayoutCachedWidth =3240
                            LayoutCachedHeight =4379
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3420
                    Top =4140
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
                    LayoutCachedTop =4140
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =4399
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400150000005b00 ,
                        0x4300570044005f0043006800650063006b005f003100320030005d003d004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =3720
                            Top =4154
                            Width =360
                            Height =225
                            FontSize =10
                            Name ="lblCWD_240"
                            Caption ="120"
                            FontName ="Calibri"
                            LayoutCachedLeft =3720
                            LayoutCachedTop =4154
                            LayoutCachedWidth =4080
                            LayoutCachedHeight =4379
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    TextAlign =4
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4260
                    Top =4140
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
                    LayoutCachedTop =4140
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =4399
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000101000000000000ed1c2400150000005b00 ,
                        0x4300570044005f0043006800650063006b005f003200340030005d003d004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            TextFontFamily =34
                            Left =4560
                            Top =4154
                            Width =360
                            Height =225
                            FontSize =10
                            Name ="lblCWD_120"
                            Caption ="240"
                            FontName ="Calibri"
                            LayoutCachedLeft =4560
                            LayoutCachedTop =4154
                            LayoutCachedWidth =4920
                            LayoutCachedHeight =4379
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4500
                    Top =720
                    Width =480
                    Height =225
                    TabIndex =19
                    ForeColor =255
                    Name ="Text54"
                    ControlSource ="Aspect"

                    LayoutCachedLeft =4500
                    LayoutCachedTop =720
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =945
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            TextFontFamily =34
                            Left =3840
                            Top =720
                            Width =600
                            Height =225
                            ForeColor =255
                            Name ="Label55"
                            Caption ="Aspect:"
                            LayoutCachedLeft =3840
                            LayoutCachedTop =720
                            LayoutCachedWidth =4440
                            LayoutCachedHeight =945
                            BackThemeColorIndex =1
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3330
                    Top =735
                    Width =480
                    Height =225
                    TabIndex =18
                    ForeColor =255
                    Name ="Text52"
                    ControlSource ="Slope"

                    LayoutCachedLeft =3330
                    LayoutCachedTop =735
                    LayoutCachedWidth =3810
                    LayoutCachedHeight =960
                    BackThemeColorIndex =1
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            TextFontFamily =34
                            Left =2820
                            Top =735
                            Width =480
                            Height =225
                            ForeColor =255
                            Name ="Label53"
                            Caption ="Slope:"
                            LayoutCachedLeft =2820
                            LayoutCachedTop =735
                            LayoutCachedWidth =3300
                            LayoutCachedHeight =960
                            BackThemeColorIndex =1
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
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =900
                    Top =180
                    Width =2580
                    ForeColor =8421504
                    Name ="Text18"
                    ControlSource ="=Now()"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextAlign =3
                            TextFontFamily =34
                            Top =180
                            Width =840
                            Height =225
                            ForeColor =8421504
                            Name ="Label19"
                            Caption ="Printed on:"
                        End
                    End
                End
            End
        End
    End
End
