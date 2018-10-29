Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5325
    DatasheetFontHeight =10
    ItemSuffix =26
    Left =5850
    Top =1425
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2aea6abea030e540
    End
    RecordSource ="qRpt_sRpt_Participants"
    Caption ="srpt_Metadata"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x0000000000000000000000000000000000000000cd140000ff00000001000000 ,
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
            Height =255
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =204
                    TextAlign =3
                    TextFontFamily =34
                    Left =2610
                    Width =1215
                    Height =255
                    FontSize =10
                    ForeColor =0
                    Name ="Label24"
                    Caption ="Participants"
                    FontName ="Calibri"
                    LayoutCachedLeft =2610
                    LayoutCachedWidth =3825
                    LayoutCachedHeight =255
                End
                Begin Label
                    TextFontCharSet =204
                    TextAlign =2
                    TextFontFamily =34
                    Left =3870
                    Top =45
                    Width =1350
                    Height =210
                    FontSize =8
                    FontWeight =400
                    ForeColor =9870754
                    Name ="Label25"
                    Caption ="Initial after review"
                    FontName ="Calibri"
                    LayoutCachedLeft =3870
                    LayoutCachedTop =45
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =255
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
            CanGrow = NotDefault
            Height =255
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    TextFontCharSet =238
                    TextAlign =3
                    IMESentenceMode =3
                    Left =60
                    Width =3780
                    Height =243
                    ColumnWidth =1800
                    FontSize =10
                    Name ="Notes"
                    ControlSource ="=[First_Name] & \" \" & [Last_Name] & \" (\" & [Contact_Role] & \")\""
                    StatusBarText ="MA. General notes on the event (Ev_Notes)"
                    FontName ="Calibri"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =3840
                    LayoutCachedHeight =243
                End
                Begin Line
                    Left =3870
                    Top =240
                    Width =1350
                    BorderColor =5855577
                    Name ="Line23"
                    LayoutCachedLeft =3870
                    LayoutCachedTop =240
                    LayoutCachedWidth =5220
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
