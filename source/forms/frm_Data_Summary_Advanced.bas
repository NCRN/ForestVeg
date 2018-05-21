﻿Version =20
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =48
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =9
    ItemSuffix =66
    Left =3405
    Top =1455
    Right =17805
    Bottom =11340
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2680758ff389e340
    End
    Caption =" Data Summary Tool"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            CanGrow = NotDefault
            Height =9900
            BackColor =12574431
            Name ="Detail"
            Begin
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    IMESentenceMode =3
                    ListRows =24
                    Left =5040
                    Top =75
                    Width =7440
                    Height =300
                    FontSize =10
                    ColumnInfo ="\"\";\"\";\"10\";\"510\""
                    Name ="cbxSelectQuery"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MSysObjects.Name, MSysObjects.Type, * FROM MSysObjects WHERE (((MSysObjec"
                        "ts.Name) Like \"qSum_*\") AND ((MSysObjects.Type)=5)) ORDER BY MSysObjects.Name;"
                        " "
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    OnNotInList ="[Event Procedure]"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =75
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =375
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =186
                            TextAlign =3
                            Left =3420
                            Top =75
                            Width =1560
                            Height =240
                            FontSize =10
                            Name ="lblQuery"
                            Caption ="Select the query:"
                            FontName ="Calibri"
                            LayoutCachedLeft =3420
                            LayoutCachedTop =75
                            LayoutCachedWidth =4980
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =3465
                    Top =1440
                    Width =10860
                    Height =8460
                    TabIndex =8
                    Name ="subResults"

                    LayoutCachedLeft =3465
                    LayoutCachedTop =1440
                    LayoutCachedWidth =14325
                    LayoutCachedHeight =9900
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =13680
                    Top =660
                    Width =426
                    Height =426
                    FontSize =10
                    TabIndex =7
                    Name ="btnDesign"
                    Caption ="Design view"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000000d088888888888880a ,
                        0x080808080808080d000000000000000aa0eeeeeeee0dadadd0e0000ee0dadada ,
                        0xa0e0a0ee00adadadd0e00ee0d00adadaa0e0ee0da000adadd0eee0dad0b70ada ,
                        0xa0ee0dada0b80dadd0e0dadada0b70daa00dadadad0b00add0dadadadad0110a ,
                        0xadadadadada000ad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View the selected query in design view"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =12600
                    Top =120
                    Width =426
                    Height =426
                    FontSize =10
                    TabIndex =2
                    Name ="btnChart"
                    Caption ="Chart view"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada00000000000000000ad0c010d0c010da ,
                        0x0da0c010a0c010ad0ad0c010d0c010da0da0c010a0c010ad0ad0c010d0c010da ,
                        0x0da0c000a0c010ad0ad0c0dad0c010da0da0c0ada00010ad0ad0c0dadad010da ,
                        0x0da000adada010ad0adadadadad010da0dadadadada010ad0adadadadad000da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View the selected query in chart view"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =13140
                    Top =120
                    Width =426
                    Height =426
                    FontSize =10
                    TabIndex =3
                    Name ="btnPivotTable"
                    Caption ="Table view"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadadd00000000000000a ,
                        0xa0880fffffffff0dd0440f0f0f0f0f0aa0880fffffffff0dd0440f0f0f0f0f0a ,
                        0xa0880fffffffff0dd0440f0f0f0f0f0aa0880fffffffff0dd04400000000000a ,
                        0xa04448484848480dd04448484848480aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="View the selected query in pivot table view"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =13680
                    Top =120
                    Width =426
                    Height =426
                    FontSize =10
                    TabIndex =4
                    Name ="btnCloseup"
                    Caption ="Zoom"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada00adadadadadadad000adadadadadada ,
                        0xa000adadadadadadda000a700007dadaada0000888800daddada07ee888870da ,
                        0xada708e88888807ddad08e888888880aada088888888880ddad088888888e80a ,
                        0xada088888888e80ddad70888888ee07aadad07888eee70addadad00888800ada ,
                        0xadadad700007adad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Open the selected query in a new window"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =12600
                    Top =660
                    Width =426
                    Height =426
                    FontSize =10
                    TabIndex =5
                    Name ="btnExportExcel"
                    Caption ="Zoom"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                        0xad000dadadadadaddad0dadadadadadaadadadadad72727ddada2727272f272a ,
                        0xadad727272f272addada27272f2727daadada272f27272addadada2f2727dada ,
                        0xadada2f272727daddada2f27272727daadad72727d7272addada2727dad727da ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Export the selected query to Excel"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =13140
                    Top =660
                    Width =426
                    Height =426
                    FontSize =10
                    TabIndex =6
                    Name ="btnExportText"
                    Caption ="Zoom"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada0000000dadadadadd00000dadadadada ,
                        0xad000dad777777addad0dad00000077aadadad0ffffff07ddad000000888807a ,
                        0xad0e8e8e80fff07dda08e8e8e088807aad0e8e8e8e0ff07ddad0e0000808807a ,
                        0xada08e8e8e80f07ddada080000e0807aadad0e8e8e8e007ddadad0f0f0f000da ,
                        0xadadad0d0d0d0dad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Export the selected query to a text file"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    AllowAutoCorrect = NotDefault
                    SpecialEffect =3
                    OverlapFlags =85
                    TextFontCharSet =238
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5040
                    Top =435
                    Width =7440
                    Height =630
                    TabIndex =1
                    Name ="tbxDesc"
                    FontName ="Calibri"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =435
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =1065
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =3465
                    Top =1065
                    Width =1020
                    Height =317
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    ForeColor =0
                    Name ="btnRequery"
                    Caption ="Requery"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Requery the results set for the selected query"

                    LayoutCachedLeft =3465
                    LayoutCachedTop =1065
                    LayoutCachedWidth =4485
                    LayoutCachedHeight =1382
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1290
                    Top =510
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxParkFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Unit Code\")) ORDER BY tlu_Enumerations.Enum_Code; "
                    ColumnWidths ="1224"
                    StatusBarText ="Filter by park"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by park"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =510
                    LayoutCachedWidth =2514
                    LayoutCachedHeight =780
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =186
                            TextAlign =3
                            Left =480
                            Top =510
                            Width =750
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblParkFilter"
                            Caption ="Park:"
                            FontName ="Calibri"
                            LayoutCachedLeft =480
                            LayoutCachedTop =510
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =765
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =2610
                    Top =510
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =13
                    ForeColor =0
                    Name ="tglFilterByPark"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the park filter on or off"

                    LayoutCachedLeft =2610
                    LayoutCachedTop =510
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =810
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1290
                    Top =870
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =14
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxAdminParkFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Group FROM tlu_Enumerat"
                        "ions WHERE (((tlu_Enumerations.Enum_Group)=\"Unit Code\")) ORDER BY tlu_Enumerat"
                        "ions.Enum_Code; "
                    ColumnWidths ="1224"
                    StatusBarText ="Filter by admin park"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by admin park"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =870
                    LayoutCachedWidth =2514
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =186
                            TextAlign =3
                            Left =105
                            Top =870
                            Width =1125
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblAdminParkFilter"
                            Caption ="Admin Park:"
                            FontName ="Calibri"
                            LayoutCachedLeft =105
                            LayoutCachedTop =870
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1125
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =2610
                    Top =870
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =15
                    ForeColor =0
                    Name ="tglFilterByAdminPark"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the admin park filter on or off"

                    LayoutCachedLeft =2610
                    LayoutCachedTop =870
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =1170
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =20
                    Left =1296
                    Top =3255
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =20
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="cbxYearFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qfrm_Data_Gateway.Event_Year FROM qfrm_Data_Gateway WHERE (((qfrm_Data_Ga"
                        "teway.Event_Year) Is Not Null)) GROUP BY qfrm_Data_Gateway.Event_Year ORDER BY q"
                        "frm_Data_Gateway.Event_Year DESC; "
                    ColumnWidths ="1224"
                    StatusBarText ="Filter by event year"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by event year"

                    LayoutCachedLeft =1296
                    LayoutCachedTop =3255
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =3525
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =186
                            TextAlign =3
                            Left =636
                            Top =3255
                            Width =600
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblYearFilter"
                            Caption ="Year:"
                            FontName ="Calibri"
                            LayoutCachedLeft =636
                            LayoutCachedTop =3255
                            LayoutCachedWidth =1236
                            LayoutCachedHeight =3510
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =2616
                    Top =3255
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =21
                    ForeColor =0
                    Name ="tglFilterByYear"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the year filter on or off"

                    LayoutCachedLeft =2616
                    LayoutCachedTop =3255
                    LayoutCachedWidth =3096
                    LayoutCachedHeight =3555
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =135
                    Top =6165
                    Width =1486
                    Height =264
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    ForeColor =0
                    Name ="btnFiltersOff"
                    Caption ="Filters off"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Turn off all form filters"

                    LayoutCachedLeft =135
                    LayoutCachedTop =6165
                    LayoutCachedWidth =1621
                    LayoutCachedHeight =6429
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =2610
                    Top =1590
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =19
                    ForeColor =0
                    Name ="tglFilterByPanel"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the selected panel"

                    LayoutCachedLeft =2610
                    LayoutCachedTop =1590
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =1890
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1290
                    Top =1590
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =18
                    BackColor =-2147483643
                    ColumnInfo ="\"\";\"\";\"4\";\"4\""
                    Name ="cbxPanelFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Panel AS Expr1 FROM tbl_Locations GROUP BY tbl_Locations.Pa"
                        "nel ORDER BY tbl_Locations.Panel; "
                    ColumnWidths ="1224"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by panel"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1590
                    LayoutCachedWidth =2514
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =186
                            TextAlign =3
                            Left =450
                            Top =1590
                            Width =780
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblPanelFilter"
                            Caption ="Panel:"
                            FontName ="Calibri"
                            LayoutCachedLeft =450
                            LayoutCachedTop =1590
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1845
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1296
                    Top =3915
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =22
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="tbxStartDateFilter"
                    Format ="yyyy mmm dd"
                    StatusBarText ="Start date for filters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =1296
                    LayoutCachedTop =3915
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =4185
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =186
                            TextAlign =3
                            Left =270
                            Top =3915
                            Width =966
                            Height =252
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblStartDateFilter"
                            Caption ="From date:"
                            FontName ="Calibri"
                            LayoutCachedLeft =270
                            LayoutCachedTop =3915
                            LayoutCachedWidth =1236
                            LayoutCachedHeight =4167
                        End
                    End
                End
                Begin TextBox
                    AllowAutoCorrect = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1296
                    Top =4215
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =23
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="tbxEndDateFilter"
                    Format ="yyyy mmm dd"
                    StatusBarText ="End date for filters"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =1296
                    LayoutCachedTop =4215
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =4485
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =186
                            TextAlign =3
                            Left =273
                            Top =4215
                            Width =963
                            Height =252
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblEndDateFilter"
                            Caption ="To date:"
                            FontName ="Calibri"
                            LayoutCachedLeft =273
                            LayoutCachedTop =4215
                            LayoutCachedWidth =1236
                            LayoutCachedHeight =4467
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =60
                    Top =90
                    Width =3240
                    Height =300
                    FontSize =12
                    FontWeight =700
                    BackColor =15527148
                    Name ="lblLocFilters"
                    Caption ="F I L T E R S"
                    FontName ="Calibri"
                    LayoutCachedLeft =60
                    LayoutCachedTop =90
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =390
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =2616
                    Top =4035
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =24
                    ForeColor =0
                    Name ="tglFilterByRange"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the date range filter on or off"

                    LayoutCachedLeft =2616
                    LayoutCachedTop =4035
                    LayoutCachedWidth =3096
                    LayoutCachedHeight =4335
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    Left =1296
                    Top =3585
                    Width =1215
                    Height =255
                    FontSize =10
                    Name ="lblOr"
                    Caption ="Or"
                    FontName ="Calibri"
                    LayoutCachedLeft =1296
                    LayoutCachedTop =3585
                    LayoutCachedWidth =2511
                    LayoutCachedHeight =3840
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    ListRows =20
                    Left =1290
                    Top =1950
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =16
                    BackColor =-2147483643
                    ColumnInfo ="\"\";\"\";\"10\";\"32\""
                    Name ="cbxFrameFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Frame FROM tbl_Locations GROUP BY tbl_Locations.Frame ORDER"
                        " BY tbl_Locations.Frame; "
                    ColumnWidths ="1224"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by frame"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1950
                    LayoutCachedWidth =2514
                    LayoutCachedHeight =2220
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =186
                            TextAlign =3
                            Left =450
                            Top =1950
                            Width =780
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblFrameFilter"
                            Caption ="Frame:"
                            FontName ="Calibri"
                            LayoutCachedLeft =450
                            LayoutCachedTop =1950
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =2205
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =2610
                    Top =1950
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =17
                    ForeColor =0
                    Name ="tglFilterByFrame"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the selected frame"

                    LayoutCachedLeft =2610
                    LayoutCachedTop =1950
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =2250
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =450
                    Width =3240
                    Height =6060
                    Name ="boxFilter"
                    LayoutCachedLeft =60
                    LayoutCachedTop =450
                    LayoutCachedWidth =3300
                    LayoutCachedHeight =6510
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7590
                    Top =1110
                    Width =606
                    Height =255
                    FontSize =10
                    TabIndex =9
                    BackColor =8454143
                    Name ="tbxUnfilteredFlag"
                    FontName ="Calibri"
                    ControlTipText ="Indicates whether results for the selected query can be filtered"

                    LayoutCachedLeft =7590
                    LayoutCachedTop =1110
                    LayoutCachedWidth =8196
                    LayoutCachedHeight =1365
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =4875
                            Top =1110
                            Width =2655
                            Height =255
                            FontSize =10
                            Name ="lblUnfilteredFlag"
                            Caption ="Query returning filtered results?"
                            FontName ="Calibri"
                            ControlTipText ="Indicates whether results for the selected query can be filtered"
                            LayoutCachedLeft =4875
                            LayoutCachedTop =1110
                            LayoutCachedWidth =7530
                            LayoutCachedHeight =1365
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =7080
                    Width =1800
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =25
                    Name ="btnEventSummary"
                    Caption ="Event Summary"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =7080
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =7380
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    Left =30
                    Top =6600
                    Width =3285
                    Height =300
                    FontSize =12
                    FontWeight =700
                    BackColor =15527148
                    Name ="lblReportsExports"
                    Caption ="R E P O R T S   &&   E X P O R T S"
                    FontName ="Calibri"
                    LayoutCachedLeft =30
                    LayoutCachedTop =6600
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =6900
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =3
                    ListWidth =2880
                    Left =1980
                    Top =7140
                    FontSize =10
                    TabIndex =26
                    ColumnInfo ="\"Event ID\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cbxEventSelection"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qFiltered_Events.Event_ID, qFiltered_Locations.Plot_Name, Format([tbl_Eve"
                        "nts].[Event_Date],\"mm/dd/yyyy\") AS [Date] FROM qFiltered_Locations INNER JOIN "
                        "qFiltered_Events ON qFiltered_Locations.[Location_ID]=qFiltered_Events.Location_"
                        "ID ORDER BY qFiltered_Locations.Plot_Name, Format([tbl_Events].[Event_Date],\"mm"
                        "/dd/yyyy\") DESC; "
                    ColumnWidths ="0;1440;1440"
                    FontName ="Calibri"
                    LayoutCachedLeft =1980
                    LayoutCachedTop =7140
                    LayoutCachedWidth =3420
                    LayoutCachedHeight =7380
                End
                Begin ToggleButton
                    OverlapFlags =247
                    TextFontCharSet =186
                    Left =2595
                    Top =2310
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =27
                    ForeColor =0
                    Name ="tglFilterByStatus"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the plot status"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2595
                    LayoutCachedTop =2310
                    LayoutCachedWidth =3075
                    LayoutCachedHeight =2610
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =8640
                    Left =1275
                    Top =2310
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =28
                    BackColor =-2147483643
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbxStatusFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Location Status\")) ORDER BY "
                        "tlu_Enumerations.Sort_Order; "
                    ColumnWidths ="1080;7560"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by  Plot status"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =2310
                    LayoutCachedWidth =2499
                    LayoutCachedHeight =2580
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =186
                            TextAlign =3
                            Left =435
                            Top =2310
                            Width =780
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblStatusFilter"
                            Caption ="Status:"
                            FontName ="Calibri"
                            LayoutCachedLeft =435
                            LayoutCachedTop =2310
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =2565
                        End
                    End
                End
                Begin ComboBox
                    ColumnHeads = NotDefault
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =4
                    ListRows =20
                    ListWidth =3960
                    Left =1275
                    Top =2670
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =29
                    BackColor =-2147483643
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"\";\"10\";\"200\""
                    Name ="cbxLocationFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, tbl_Locations.Unit_Co"
                        "de, tbl_Locations.Admin_Unit_Code FROM tbl_Locations WHERE (((tbl_Locations.Unit"
                        "_Code) Like Nz([Forms]![frm_Data_Summary_Advanced]![cboParkFilter],\"*\")) AND ("
                        "(tbl_Locations.Admin_Unit_Code) Like Nz([Forms]![frm_Data_Summary_Advanced]![cbo"
                        "AdminParkFilter],\"*\"))) ORDER BY tbl_Locations.Plot_Name;"
                    ColumnWidths ="0;1440;1080;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by Plot Name"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =2670
                    LayoutCachedWidth =2499
                    LayoutCachedHeight =2940
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =186
                            TextAlign =3
                            Left =300
                            Top =2670
                            Width =915
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblLocationFilter"
                            Caption ="Location:"
                            FontName ="Calibri"
                            LayoutCachedLeft =300
                            LayoutCachedTop =2670
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =2925
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =247
                    TextFontCharSet =186
                    Left =2595
                    Top =2670
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =30
                    ForeColor =0
                    Name ="tglFilterByLocation"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Filter by the Plot Name"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2595
                    LayoutCachedTop =2670
                    LayoutCachedWidth =3075
                    LayoutCachedHeight =2970
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    TextFontCharSet =186
                    Left =1695
                    Top =6165
                    Width =1516
                    Height =264
                    FontSize =10
                    FontWeight =700
                    TabIndex =31
                    ForeColor =0
                    Name ="btnFiltersClear"
                    Caption ="Clear Filters"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Clear all form filters"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1695
                    LayoutCachedTop =6165
                    LayoutCachedWidth =3211
                    LayoutCachedHeight =6429
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AllowAutoCorrect = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =186
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =1290
                    Top =1230
                    Width =1224
                    Height =270
                    FontSize =10
                    TabIndex =32
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cbxSubunitFilter"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Subunit Code\")) ORDER BY tlu"
                        "_Enumerations.Enum_Code; "
                    ColumnWidths ="720;2880"
                    StatusBarText ="Filter by subunit"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Filter by unofficial subunit"

                    LayoutCachedLeft =1290
                    LayoutCachedTop =1230
                    LayoutCachedWidth =2514
                    LayoutCachedHeight =1500
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =186
                            TextAlign =3
                            Left =180
                            Top =1230
                            Width =1050
                            Height =255
                            FontSize =10
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblSubunitFilter"
                            Caption ="Subunit:"
                            FontName ="Calibri"
                            ControlTipText ="Unofficial subunit"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1230
                            LayoutCachedWidth =1230
                            LayoutCachedHeight =1485
                        End
                    End
                End
                Begin ToggleButton
                    OverlapFlags =247
                    TextFontCharSet =186
                    Left =2610
                    Top =1230
                    Width =480
                    Height =300
                    FontSize =10
                    TabIndex =33
                    ForeColor =0
                    Name ="tglFilterBySubunit"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    Caption ="Filter on"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadad0000adadaddadada0660dadadaadadad0660adadaddadada0f80dadada ,
                        0xadadad0f80adadaddadad088860adadaadad06888660adaddad068f888660ada ,
                        0xad068f88888660add068fff88886660aa00000000000000ddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Turn the subunit filter on or off"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2610
                    LayoutCachedTop =1230
                    LayoutCachedWidth =3090
                    LayoutCachedHeight =1530
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =120
                    Top =7440
                    Width =1800
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =34
                    Name ="btnRptTagHistory"
                    Caption ="Tag History"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =120
                    LayoutCachedTop =7440
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =7739
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin OptionGroup
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =247
                    Left =135
                    Top =4635
                    Width =3090
                    Height =1335
                    TabIndex =35
                    Name ="optgScope"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    ControlTipText ="Scope of the data included in the validation queries: uncertified events, certif"
                        "ied events, or both?"

                    LayoutCachedLeft =135
                    LayoutCachedTop =4635
                    LayoutCachedWidth =3225
                    LayoutCachedHeight =5970
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OverlapFlags =247
                            TextFontCharSet =238
                            Left =255
                            Top =4680
                            Width =2850
                            Height =255
                            FontSize =10
                            FontWeight =700
                            BackColor =13025979
                            Name ="lblIncludeCertified"
                            Caption ="D a t a    s c o p e"
                            FontName ="Calibri"
                            LayoutCachedLeft =255
                            LayoutCachedTop =4680
                            LayoutCachedWidth =3105
                            LayoutCachedHeight =4935
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =285
                            Top =5019
                            OptionValue =0
                            Name ="optUncertOnly"

                            LayoutCachedLeft =285
                            LayoutCachedTop =5019
                            LayoutCachedWidth =545
                            LayoutCachedHeight =5259
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =238
                                    Left =525
                                    Top =4995
                                    Width =2580
                                    Height =270
                                    FontSize =10
                                    Name ="lblUncertifiedOnly"
                                    Caption ="Use only uncertified data"
                                    FontName ="Calibri"
                                    ControlTipText ="Run queries only on uncertified events"
                                    LayoutCachedLeft =525
                                    LayoutCachedTop =4995
                                    LayoutCachedWidth =3105
                                    LayoutCachedHeight =5265
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =285
                            Top =5340
                            TabIndex =1
                            OptionValue =1
                            Name ="optBoth"

                            LayoutCachedLeft =285
                            LayoutCachedTop =5340
                            LayoutCachedWidth =545
                            LayoutCachedHeight =5580
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =238
                                    Left =525
                                    Top =5310
                                    Width =2580
                                    Height =270
                                    FontSize =10
                                    Name ="lblBoth"
                                    Caption ="Both uncertified and certified"
                                    FontName ="Calibri"
                                    ControlTipText ="Run queries only on certified and uncertified events"
                                    LayoutCachedLeft =525
                                    LayoutCachedTop =5310
                                    LayoutCachedWidth =3105
                                    LayoutCachedHeight =5580
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =247
                            Left =285
                            Top =5640
                            TabIndex =2
                            OptionValue =2
                            Name ="optCertOnly"

                            LayoutCachedLeft =285
                            LayoutCachedTop =5640
                            LayoutCachedWidth =545
                            LayoutCachedHeight =5880
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =238
                                    Left =525
                                    Top =5610
                                    Width =2580
                                    Height =270
                                    FontSize =10
                                    Name ="lblCertifiedOnly"
                                    Caption ="Use certified data only"
                                    FontName ="Calibri"
                                    ControlTipText ="Run queries only on certified events"
                                    LayoutCachedLeft =525
                                    LayoutCachedTop =5610
                                    LayoutCachedWidth =3105
                                    LayoutCachedHeight =5880
                                End
                            End
                        End
                    End
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =8265
                    Top =1110
                    Width =4245
                    Height =255
                    FontSize =10
                    Name ="lblNote"
                    Caption ="Note that Crosstab queries (_x) are never filtered."
                    FontName ="Calibri"
                    ControlTipText ="Indicates whether results for the selected query can be filtered"
                    LayoutCachedLeft =8265
                    LayoutCachedTop =1110
                    LayoutCachedWidth =12510
                    LayoutCachedHeight =1365
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =238
                    Left =120
                    Top =7800
                    Width =1800
                    Height =299
                    FontSize =10
                    FontWeight =700
                    TabIndex =36
                    Name ="btnExportProducts"
                    Caption ="Export 4 Yr Products"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =120
                    LayoutCachedTop =7800
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =8099
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =8160
                    Width =1800
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =37
                    Name ="btnExportAll"
                    Caption ="Export All Data"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =120
                    LayoutCachedTop =8160
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =8490
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =8520
                    Width =1800
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =38
                    Name ="btnOpenBasicSummaryForm"
                    Caption ="Basic Summaries"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =120
                    LayoutCachedTop =8520
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =8850
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' MODULE:       frm_Data_Summary_Advanced
' Level:        Application module
' Version:      1.02
'
' Description:  Standard form for summarizing/exploring project data
' Source/date:  John Boetsch, Jan 2010
'               Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, May 14, 2018
' Revisions:    JB/ML/GS - 1/2010+  - 1.00 - initial version
'               BLC   - 5/14/2018 - 1.01 - added documentation, error handling
'               BLC   - 5/15/2018 - 1.02 - revise to use ToggleFilters(), SetFilters()
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ----------------
'  Events
' ----------------

' ----------------
'  Form
' ----------------
' ---------------------------------
' SUB:          Form_Open
' Description:  form open actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John R. Boetsch, Jan 2010
'               Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   JB/ML/GS - 1/2010+ - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    ' Close the form if the switchboard is not open
    If fxnSwitchboardIsOpen = False Then
        MsgBox "The main database switchboard must be" & vbCrLf & _
            "open for this form to function properly.", , "Cannot open the form ..."
        DoCmd.CancelEvent
        GoTo Exit_Handler
    End If
    FilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSelectQuery_NotInList
' Description:  combobox not in list actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxSelectQuery_NotInList(NewData As String, Response As Integer)
On Error GoTo Err_Handler
    
    Me.ActiveControl.Undo
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectQuery_NotInList[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSelectQuery_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxSelectQuery_AfterUpdate()
On Error GoTo Err_Handler
    
    ' Exit if no query selected
    If IsNull(Me.cbxSelectQuery) Then
        Me.tbxUnfilteredFlag = ""
        Me.tbxUnfilteredFlag.ForeColor = 0          'black
        Me.tbxUnfilteredFlag.BackColor = 8454143    'yellow
        Me.subResults.SourceObject = ""
        GoTo Exit_Handler
    End If

    ' Update the description
    Me.tbxDesc = ""

    Dim qdf As DAO.QueryDef
    Dim qdfs As DAO.QueryDefs
    Set qdfs = DBEngine(0)(0).QueryDefs

    On Error Resume Next
    For Each qdf In qdfs
        If qdf.Name = Me.cbxSelectQuery.Value Then
            Me.tbxDesc = qdf.Properties("Description")
        End If
    Next qdf

    On Error GoTo Err_Handler
    ' Bind the subform to the newly-selected object
    Me.subResults.SourceObject = "Query." & Me.cbxSelectQuery.Value

    ' Update the visual flag to indicate whether or not the query returns filtered results
    '   Note: suffix of "_X" means that the query cannot accept parameters (e.g., crosstab)
    If Right(Me.cbxSelectQuery.Value, 2) = "_X" Then
        Me.tbxUnfilteredFlag = "No"
        Me.tbxUnfilteredFlag.ForeColor = 16777215   'white
        Me.tbxUnfilteredFlag.BackColor = 255        'red
    Else
        Me.tbxUnfilteredFlag = "Yes"
        Me.tbxUnfilteredFlag.ForeColor = 16777215   'white
        Me.tbxUnfilteredFlag.BackColor = 4227072    'green
    End If

    ' Set focus to the subform to allow scrolling, etc.
    Me.subResults.SetFocus
    
Exit_Handler:
    Set qdfs = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cbxSelectQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSelectQuery_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenBrowser_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnOpenBrowser_Click()
On Error GoTo Err_Handler
    
    Set gvarRefForm = Me.Form
    Set gvarRefCtl = Me.subResults
    ' Open to a blank record - to distinguish from opening to the selected record in the subform
    DoCmd.OpenForm "frm_Data_Browser", , , , acFormAdd, , "off"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "The table, query or form is no longer available in the application.", , _
            "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenBrowser_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnRequery_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnRequery_Click()
On Error GoTo Err_Handler
    
    ' Bail out if no query is currently selected
    If IsNull(Me.cbxSelectQuery) Then GoTo Exit_Handler

    ' Requery the selected record in the recordset, and update the subform
    Me.subResults.Requery
    Me.subResults.SetFocus
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRequery_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Filters
' ----------------
' ---------------------------------
' SUB:          btnFiltersOff_Click
' Description:  button click
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
'   BLC - 5/15/2018 - revise to use ToggleFilters()
' ---------------------------------
Private Sub btnFiltersOff_Click()
On Error GoTo Err_Handler

    ToggleFilters "off"
'    ' Turn off the filters
'    Me.btnRequery.SetFocus
'    ' Undo the filter toggles
'    Me.tglFilterByPark = False
'    Me.tglFilterByAdminPark = False
'    Me.tglFilterBySubunit = False
'    Me.tglFilterByPanel = False
'    Me.tglFilterByFrame = False
'    Me.tglFilterByStatus = False
'    Me.tglFilterByLocation = False
'    Me.tglFilterByYear = False
'    Me.tglFilterByRange = False
'
'    fxnFilterRecords
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnFiltersOff_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnFiltersClear_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnFiltersClear_Click()
On Error GoTo Err_Handler
    
    ToggleFilters "clear"
    
'    'Clear the filters
'    Me.btnRequery.SetFocus
'    Me.cbxParkFilter = Null
'    Me.cbxAdminParkFilter = Null
'    Me.cbxSubunitFilter = Null
'    Me.cbxPanelFilter = Null
'    Me.cbxFrameFilter = Null
'    Me.cbxStatusFilter = Null
'    Me.cbxLocationFilter = Null
'    Me.cbxYearFilter = Null
'    Me.tbxStartDateFilter = Null
'    Me.tbxEndDateFilter = Null
'
'    fxnFilterRecords
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnFiltersClear_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Filter Selection
' ----------------

' ---------------------------------
' SUB:          cbxParkFilter_AfterUpdate
' Description:  combobox after udpate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxParkFilter_AfterUpdate()
On Error GoTo Err_Handler
    
'    SelectFilter Me.cbxParkFilter
    
    Me.tglFilterByPark = Not IsNull(Me.cbxParkFilter)
    FilterRecords 'fxnFilterRecords
    Me.tglFilterByPark.SetFocus
    Me.cbxLocationFilter.Requery
    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxParkFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxAdminParkFilter_AfterUpdate
' Description:  combobox after update
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxAdminParkFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.cbxAdminParkFilter
    
'    Me.tglFilterByAdminPark = Not IsNull(Me.cbxAdminParkFilter)
'    fxnFilterRecords
'    Me.tglFilterByAdminPark.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxAdminParkFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxSubunitFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxSubunitFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.cbxSubunitFilter
        
'    Me.tglFilterBySubunit = Not IsNull(Me.cbxSubunitFilter)
'    fxnFilterRecords
'    Me.tglFilterBySubunit.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxSubunitFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxFrameFilter_AfterUpdate
' Description:  form open actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxFrameFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.cbxFrameFilter
    
'    Me.tglFilterByFrame = Not IsNull(Me.cbxFrameFilter)
'    fxnFilterRecords
'    Me.tglFilterByFrame.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxFrameFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxPanelFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxPanelFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.cbxPanelFilter
    
'    Me.tglFilterByPanel = Not IsNull(Me.cbxPanelFilter)
'    fxnFilterRecords
'    Me.tglFilterByPanel.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxPanelFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxStatusFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxStatusFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.cbxStatusFilter
    
'    Me.tglFilterByStatus = Not IsNull(Me.cbxStatusFilter)
'    fxnFilterRecords
'    Me.tglFilterByStatus.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxStatusFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxYearFilter_AfterUpdate
' Description:  combobox after udpate actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxYearFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.cbxYearFilter
    
'    Me.tglFilterByYear = Not IsNull(Me.cbxYearFilter)
'    If Me.tglFilterByYear = True Then Me.tglFilterByRange = False
'    fxnFilterRecords
'    Me.tglFilterByYear.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxYearFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxLocationFilter_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub cbxLocationFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.cbxLocationFilter
    
'    Me.tglFilterByLocation = Not IsNull(Me.cbxLocationFilter)
'    fxnFilterRecords
'    Me.tglFilterByLocation.SetFocus
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxLocationFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxStartDateFilter_AfterUpdate
' Description:  textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tbxStartDateFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.tbxStartDateFilter
    
'    Me.tglFilterByRange = (Not IsNull(Me.tbxStartDateFilter)) And (Not IsNull(Me.tbxEndDateFilter))
'    If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
'    fxnFilterRecords
'    Me.tglFilterByYear.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxStartDateFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxEndDateFilter_AfterUpdate
' Description:  textbox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tbxEndDateFilter_AfterUpdate()
On Error GoTo Err_Handler
    
    SelectFilter Me.tbxEndDateFilter
    
'    Me.tglFilterByRange = (Not IsNull(Me.tbxStartDateFilter)) And (Not IsNull(Me.tbxEndDateFilter))
'    If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
'    fxnFilterRecords
'    Me.tglFilterByYear.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxEndDateFilter_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Filter Toggles
' ----------------

' ---------------------------------
' SUB:          tglFilterByPark_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling

' ---------------------------------
Private Sub tglFilterByPark_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterByPark
    
'    If IsNull(Me.cbxParkFilter) = True Then Me.tglFilterByPark = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByPark_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByAdminPark_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterByAdminPark_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterByAdminPark
'    If IsNull(Me.cbxAdminParkFilter) = True Then Me.tglFilterByAdminPark = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByAdminPark_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterBySubunit_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterBySubunit_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterBySubunit
'    If IsNull(Me.cbxSubunitFilter) = True Then Me.tglFilterBySubunit = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterBySubunit_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByFrame_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterByFrame_AfterUpdate()
On Error GoTo Err_Handler

    SetFilter Me.tglFilterByFrame
'    If IsNull(Me.cbxFrameFilter) = True Then Me.tglFilterByFrame = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByFrame_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByPanel_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterByPanel_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterByPanel
'    If IsNull(Me.cbxPanelFilter) = True Then Me.tglFilterByPanel = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByPanel_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByStatus_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterByStatus_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterByStatus
'    If IsNull(Me.cbxStatusFilter) = True Then Me.tglFilterByStatus = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByStatus_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByLocation_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterByLocation_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterByLocation
    
'    If IsNull(Me.cbxLocationFilter) = True Then Me.tglFilterByLocation = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByLocation_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByYear_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterByYear_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterByYear
'    If IsNull(Me.cbxYearFilter) Then Me.tglFilterByYear = False
'    If Me.tglFilterByYear = True Then Me.tglFilterByRange = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByYear_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tglFilterByRange_AfterUpdate
' Description:  toggle after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub tglFilterByRange_AfterUpdate()
On Error GoTo Err_Handler
    
    SetFilter Me.tglFilterByRange
'    If IsNull(Me.tbxStartDateFilter) And IsNull(Me.tbxEndDateFilter) _
'        Then Me.tglFilterByRange = False
'    If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tglFilterByRange_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          optgScope_AfterUpdate
' Description:  option group after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub optgScope_AfterUpdate()
On Error GoTo Err_Handler
    
    'Revised MEL 8/23/2010 to better handle to three scope options
    Select Case Me.optgScope
      Case 0    'Selected Only Uncertified Data
         If MsgBox("Warning: The summary results may be based on data" & vbCrLf & _
            "that have not yet passed the quality review." & vbCrLf & vbCrLf & _
            "As such the results should be considered provisional" & vbCrLf & _
            "and should only be shared or reported on in a way" & vbCrLf & _
            "that clearly indicates this.", vbExclamation + vbOKCancel + vbDefaultButton2, _
            "Include uncertified data?") = vbCancel Then
            'Revert to Certified Data Only
            Me.optgScope = 2
            Me.lblUncertifiedOnly.FontBold = False
            Me.lblBoth.FontBold = False
            Me.lblCertifiedOnly.FontBold = True
        Else
            'Keep uncertified Data in selection
            Me.lblUncertifiedOnly.FontBold = True
            Me.lblBoth.FontBold = False
            Me.lblCertifiedOnly.FontBold = False
        End If
      Case 1    'Selected certified and uncertified data
        If MsgBox("Warning: The summary results may be based on data" & vbCrLf & _
            "that have not yet passed the quality review." & vbCrLf & vbCrLf & _
            "As such the results should be considered provisional" & vbCrLf & _
            "and should only be shared or reported on in a way" & vbCrLf & _
            "that clearly indicates this.", vbExclamation + vbOKCancel + vbDefaultButton2, _
            "Include uncertified data?") = vbCancel Then
            Me.optgScope = 2
            Me.lblUncertifiedOnly.FontBold = False
            Me.lblBoth.FontBold = False
            Me.lblCertifiedOnly.FontBold = True
        Else
            Me.lblUncertifiedOnly.FontBold = False
            Me.lblBoth.FontBold = True
            Me.lblCertifiedOnly.FontBold = False
        End If
      Case 2    'Selected certified data only
            Me.lblUncertifiedOnly.FontBold = False
            Me.lblBoth.FontBold = False
            Me.lblCertifiedOnly.FontBold = True
    End Select

    Me.cbxEventSelection.Requery
    Me.cbxLocationFilter.Requery

'    If Me.optgScope = 1 Then
'        If MsgBox("Warning: The summary results may be based on data" & vbCrLf & _
'            "that have not yet passed the quality review." & vbCrLf & vbCrLf & _
'            "As such the results should be considered provisional" & vbCrLf & _
'            "and should only be shared or reported on in a way" & vbCrLf & _
'            "that clearly indicates this.", vbExclamation + vbOKCancel + vbDefaultButton2, _
'            "Include uncertified data?") = vbCancel Then
'            Me.optgScope = 0
'            Me.labCertOnly.FontBold = True
'            Me.labBoth.FontBold = False
'            Me.labBoth.ForeColor = 0
'        Else
'            Me.labCertOnly.FontBold = False
'            Me.labBoth.FontBold = True
'            Me.labBoth.ForeColor = 255
'            Me.cboEvent_Selection.Requery
'        End If
'    Else
'        Me.labCertOnly.FontBold = True
'        Me.labBoth.FontBold = False
'        Me.labBoth.ForeColor = 0
'    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - optgScope_AfterUpdate[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Buttons
' ----------------
' ---------------------------------
' SUB:          btnViewExcluded_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnViewExcluded_Click()
On Error GoTo Err_Handler
    
    ' Open the query to view event records flagged for exclusion from summaries
    DoCmd.OpenQuery "qsub_Excluded_events", acViewNormal, acReadOnly
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cbxSelectQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewExcluded_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnChart_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnChart_Click()
On Error GoTo Err_Handler
    
    ' Open the selected query as a pivot chart after checking that a query is selected
    If IsNull(Me.cbxSelectQuery) = False Then
        DoCmd.OpenQuery Me.cbxSelectQuery.Value, acViewPivotChart, acReadOnly
        DoCmd.Maximize
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cbxSelectQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnChart_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPivotTable_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnPivotTable_Click()
On Error GoTo Err_Handler
    
    ' Open the selected query as a pivot table after checking that a query is selected
    If IsNull(Me.cbxSelectQuery) = False Then
        DoCmd.OpenQuery Me.cbxSelectQuery.Value, acViewPivotTable, acReadOnly
        DoCmd.Maximize
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cbxSelectQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnPivotTable_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCloseup_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnCloseup_Click()
On Error GoTo Err_Handler
    
    ' Open the selected query in a new window after checking that a query is selected
    If IsNull(Me.cbxSelectQuery) = False Then
        DoCmd.OpenQuery Me.cbxSelectQuery.Value, acViewNormal, acReadOnly
        DoCmd.Maximize
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCloseup_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnExportExcel_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnExportExcel_Click()
On Error GoTo Err_Handler
    
    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cbxSelectQuery) Then GoTo Exit_Handler
    
    strQryName = Me.cbxSelectQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".xls"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.xls)", "*.xls")
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatXLS, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnExportExcel_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnExportText_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnExportText_Click()
On Error GoTo Err_Handler
    
    Dim strQryName As String
    Dim strInitFile As String
    Dim strSaveFile As String

    ' Bail out if no query is currently selected
    If IsNull(Me.cbxSelectQuery) Then GoTo Exit_Handler

    strQryName = Me.cbxSelectQuery

    strInitFile = Application.CurrentProject.Path & "\" & _
        strQryName & "_" & CStr(Format(Now(), "yyyymmdd_hhnnss")) & ".txt"
    ' Open the save file dialog and update to the actual name given by the user
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.txt)", "*.txt")
    DoCmd.OutputTo acOutputQuery, strQryName, acFormatTXT, strSaveFile, True
    'MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 94, 2001
        ' User canceled dialog box - do nothing
      Case 2501
        ' Canceled open report action - do nothing
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnExportText_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDesign_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnDesign_Click()
On Error GoTo Err_Handler
    
    ' Open the selected query in design view after checking that a query is selected
    If IsNull(Me.cbxSelectQuery) = False Then _
        DoCmd.OpenQuery Me.cbxSelectQuery.Value, acViewDesign, acReadOnly
        
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3011, 7874   ' Object not found
        MsgBox "This query is not found in the application:" & _
            vbCrLf & """" & Me.cbxSelectQuery & """", , "Object not found"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDesign_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnEventSummary_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnEventSummary_Click()
On Error GoTo Err_Handler
    
    Dim sttDocName As String
    sttDocName = "rpt_Event_Summary"
    DoCmd.OpenReport sttDocName, acPreview
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnEventSummary_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnRptTagHistory_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnRptTagHistory_Click()
On Error GoTo Err_Handler
    
    Dim strDocName As String
    strDocName = "rpt_Tag_History"
    DoCmd.OpenReport strDocName, acPreview
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnRptTagHistory_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnExportProducts_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnExportProducts_Click()
On Error GoTo Err_Handler
    
    Dim strQryName(8, 2) As String
    Dim qNum As Integer
    Dim qDef As QueryDef
    Dim strParkName As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim strSaveFolder As String
    Dim db As DAO.Database
    Set db = CurrentDb
    
    ' Bail out if no start year is currently selected
    If IsNull(Me.cbxYearFilter) Then
        MsgBox ("A YEAR filter must be entered above for these products to be generated. Please enter the starting year of the 4 year period desired and try again")
        GoTo Exit_Handler
    End If
    'Set the name of the group of records to be exported to Region if all Parks, otherwise use the Park Code
    If IsNull(Me.cbxAdminParkFilter) Then
        strParkName = "REGION"
    Else
        strParkName = Me.cbxAdminParkFilter
    End If

    strQryName(0, 0) = "qSum_4YR_PRODUCT_Event_List_for_4_Year_Cycle"
    strQryName(0, 1) = "Events"
    strQryName(1, 0) = "qSum_4YR_PRODUCT_Trees"
    strQryName(1, 1) = "Trees"
    strQryName(2, 0) = "qSum_4YR_PRODUCT_Shrubs"
    strQryName(2, 1) = "Shrubs"
    strQryName(3, 0) = "qSum_4YR_PRODUCT_Herbaceous"
    strQryName(3, 1) = "Herbs"
    strQryName(4, 0) = "qSum_4YR_PRODUCT_Vines"
    strQryName(4, 1) = "Vines"
    strQryName(5, 0) = "qSum_4YR_PRODUCT_Pests_and_Conditions"
    strQryName(5, 1) = "Conditions"
    strQryName(6, 0) = "qSum_4YR_PRODUCT_All_Occurences"
    strQryName(6, 1) = "Species_by_Plot"

    'Generate the default output file name and allow user to edit it
'    strInitFile = Application.CurrentProject.Path & "\NCRN_ForestVeg_" & strParkName & "_" & Me.cboYearFilter & "-" & Me.cboYearFilter + 3 & "_" & CStr(Format(Now(), "yyyymmdd")) & ".xls"
'    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.xls)", "*.xls")
    
    'Cycle through queries and create an worksheet tab for each one
'   For qNum = 0 To 6
'       Set qDef = db.CreateQueryDef(strQryName(qNum, 1), CurrentDb.QueryDefs(strQryName(qNum, 0)).SQL)
'       DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, strQryName(qNum, 1), strSaveFile, True
'       DoCmd.DeleteObject acQuery, strQryName(qNum, 1)
'   Next
    
    'Generate the default output file name and allow user to edit it
    strInitFile = Application.CurrentProject.Path & "\Exports\NCRN_ForestVeg_All_" & strParkName & "_" & Me.cbxYearFilter & "-" & Me.cbxYearFilter + 3 & "_" & CStr(Format(Now(), "yyyymmdd")) & ".xlsx"
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.xls*)", "*.xls*")
    strSaveFolder = fPathParsing(strSaveFile, "D")
    'Cycle through queries and create an worksheet tab for each one
    For qNum = 0 To 6
        Set qDef = db.CreateQueryDef(strQryName(qNum, 1), CurrentDb.QueryDefs(strQryName(qNum, 0)).SQL)
        'Export each parameter to a seperate worksheet in an XLSX workbook (SpreadsheetType = '10' for .XLSX)
        DoCmd.TransferSpreadsheet acExport, 10, strQryName(qNum, 1), strSaveFile, True
        'Export each parameter to a seperate CSV file.
        DoCmd.TransferText acExportDelim, , strQryName(qNum, 1), strSaveFolder & "\" & strQryName(qNum, 1) & "_" & CStr(Format(Now(), "yyyymmdd")) & ".csv", True
        DoCmd.DeleteObject acQuery, strQryName(qNum, 1)
    Next
    
    MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnExportProducts_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnExportAll_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnExportAll_Click()
On Error GoTo Err_Handler
    
'This routines exports all data to a single XLSX file as well as individual CSV files and is typically triggered from a button on the Data Summary form.
    
    Dim strQryName(15, 2) As String
    Dim qNum As Integer
    Dim qDef As QueryDef
    Dim strParkName As String
    Dim strInitFile As String
    Dim strSaveFile As String
    Dim strSaveFolder As String
    Dim db As DAO.Database
    Set db = CurrentDb

    'Populate an array with the name of a query and the worksheet name to be used for the results of this query.
    strQryName(0, 0) = "qExport_All_Plots"
    strQryName(0, 1) = "Plots"
    strQryName(1, 0) = "qExport_All_Events"
    strQryName(1, 1) = "Events"
    strQryName(2, 0) = "qExport_All_Trees"
    strQryName(2, 1) = "Trees"
    strQryName(3, 0) = "qExport_All_Saplings"
    strQryName(3, 1) = "Saplings"
    strQryName(4, 0) = "qExport_All_Stems"
    strQryName(4, 1) = "Stems"
    strQryName(5, 0) = "qExport_All_Seedlings"
    strQryName(5, 1) = "Seedlings"
    strQryName(6, 0) = "qExport_All_Shrubs"
    strQryName(6, 1) = "Shrubs"
    strQryName(7, 0) = "qExport_All_Shrub_Seedlings"
    strQryName(7, 1) = "Shrub_Seedlings"
    strQryName(8, 0) = "qExport_Conditions"
    strQryName(8, 1) = "Tree_Sapling_Conditions"
    strQryName(9, 0) = "qExport_FoliageConditions"
    strQryName(9, 1) = "Foliage_Conditions"
    strQryName(10, 0) = "qExport_AllVines"
    strQryName(10, 1) = "Vines"
    strQryName(11, 0) = "qExport_All_Herbaceous"
    strQryName(11, 1) = "Herbs"
    strQryName(12, 0) = "qExport_All_Quadrat_Conditions"
    strQryName(12, 1) = "Quadrat_Conditions"
    strQryName(13, 0) = "qExport_All_Plot_Floor_Conditions"
    strQryName(13, 1) = "Plot_Floor"
    strQryName(14, 0) = "qExport_All_CWD"
    strQryName(14, 1) = "CWD"
    strQryName(15, 0) = "qExport_Tag_Status_by_Cycle_x"
    strQryName(15, 1) = "Tag_History"

    
    'Generate the default output file name and allow user to edit it
    strInitFile = Application.CurrentProject.Path & "\Exports\NCRN_ForestVeg_All_Data_" & CStr(Format(Now(), "yyyymmdd")) & ".xlsx"
    strSaveFile = fxnSaveFile(strInitFile, "Microsoft Excel (*.xls*)", "*.xls*")
    strSaveFolder = fPathParsing(strSaveFile, "D")
    'Cycle through queries and create an worksheet tab for each one
    For qNum = 0 To 15
        Set qDef = db.CreateQueryDef(strQryName(qNum, 1), CurrentDb.QueryDefs(strQryName(qNum, 0)).SQL)
        'Export each parameter to a seperate worksheet in an XLSX workbook (SpreadsheetType = '10' for .XLSX)
        DoCmd.TransferSpreadsheet acExport, 10, strQryName(qNum, 1), strSaveFile, True
        'Export each parameter to a seperate CSV file.
        DoCmd.TransferText acExportDelim, , strQryName(qNum, 1), strSaveFolder & "\" & strQryName(qNum, 1) & "_" & CStr(Format(Now(), "yyyymmdd")) & ".csv", True
        DoCmd.DeleteObject acQuery, strQryName(qNum, 1)
    Next
    
    MsgBox "File saved to:" & vbCrLf & vbCrLf & strSaveFile
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnExportAll_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnOpenBasicSummaryForm_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   ML/GS - unknown - initial version
'   BLC - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Sub btnOpenBasicSummaryForm_Click()
On Error GoTo Err_Handler
    
    'record what the current record is so we can go back to that record on return
    DoCmd.Close acForm, "frm_Data_Summary_Advanced"
    DoCmd.OpenForm "frm_Data_Summary_Basic"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnOpenBasicSummaryForm_Click[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Sub

' ----------------
'  Functions
' ----------------
' ---------------------------------
' FUNCTION:     FilterRecords
' Description:  filter records on the desired field
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  John Boetsch, May 5, 2006
'               Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      -
' Revisions:
'   JB    - 5/5/2010 - initial version adapted to summarization tool, mainly for formatting filters
'   ML/GS - unknown - initial version
'   BLC   - 5/14/2018 - documentation, error handling
' ---------------------------------
Private Function FilterRecords()
On Error GoTo Err_Handler
    
    Dim bFilterOn As Boolean

    bFilterOn = False

    ' If any toggles are on, the filter is on
    'If Me.togFilterByPark Or Me.togFilterByType Or Me.togFilterByStatus Or _
        Me.togFilterByLoc Or Me.togFilterByStratum Then bFilterOn = True
    ' And for loc filters that allow null values ...
    'If Me.togFilterByRegion Or Me.togFilterByPanelType Or _
    '     Me.togFilterByPanelName Then bFilterOn = True
    '  And for event filters
    ' If Me.togFilterByYear Or Me.togFilterByRange Then bFilterOn = True
    ' Non-standard fields
    'If Me.togFilterByWatershed Then bFilterOn = True

Reformat_controls:
    ' Enable/disable the command button accordingly
    'Me.cmdFiltersOff.Enabled = bFilterOn
 
    ' Make the labels bold or not depending on filter settings
    Me.lblParkFilter.FontBold = Me.tglFilterByPark
    Me.lblAdminParkFilter.FontBold = Me.tglFilterByAdminPark
    Me.lblSubunitFilter.FontBold = Me.tglFilterBySubunit
    Me.lblStatusFilter.FontBold = Me.tglFilterByStatus
    Me.lblLocationFilter.FontBold = Me.tglFilterByLocation
    Me.lblFrameFilter.FontBold = Me.tglFilterByFrame
    Me.lblPanelFilter.FontBold = Me.tglFilterByPanel
    Me.lblYearFilter.FontBold = Me.tglFilterByYear
    Me.lblStartDateFilter.FontBold = Me.tglFilterByRange
    Me.lblEndDateFilter.FontBold = Me.tglFilterByRange
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FilterRecords[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SetFilter
' Description:  prepare filter display based on filter being set
' Assumptions:  -
' Parameters:   ctl - control being toggled (control)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 15, 2018
' Adapted:      -
' Revisions:
'   BLC   - 5/15/2018 - initial version
' ---------------------------------
Private Function SetFilter(ctl As Control)
On Error GoTo Err_Handler
    
'    Select Case filter
'        Case "park"
'        Case "adminpark"
'        Case "subunit"
'        Case "panel"
'        Case "frame"
'        Case "status"
'        Case "location"
'        Case "eventyr"
'        Case "eventdate"
'    End Select
'
    Dim cbx As String
    Dim filterby As String
    
    With Me
        
        'determine filter
        filterby = Replace(ctl.Name, "tglFilterBy", "")
        cbx = "cbx" & Replace(ctl.Name, "tglFilterBy", "") & "Filter"
        
        Select Case filterby
            Case "year" 'tglFilterByYear
                If IsNull(.Controls(cbx)) = True Then ctl = False
                If ctl = True Then .tglFilterByRange = False
            Case "range" 'tglFilterByRange
                If IsNull(.tbxStartDateFilter) And IsNull(.tbxEndDateFilter) _
                    Then ctl = False
                If ctl = True Then .tglFilterByYear = False
            Case Else
                If IsNull(.Controls(cbx)) = True Then ctl = False
        End Select
        
'        'handle date ranges
'        If Left$(ctl.Name, 3) = "tbx" Then
'            tgl = "tglFilterByRange"
'            .Controls(tgl) = (Not IsNull(Me.tbxStartDateFilter)) And (Not IsNull(Me.tbxEndDateFilter))
'        Else
'            tgl = "tglFilterBy" & Remove(Remove(ctl.Name, "cbx", ""), "Filter", "")
'            .Controls(tgl) = Not IsNull(.Controls(ctl.Name))
'        End If
'
'        'handle year & range filters
'        'year -> If Me.tglFilterByYear = True Then Me.tglFilterByRange = False
'        'range ->If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
'        If .Controls(tgl) = True Then
'            'set tglOr = range if year, year if range
'            tglOr = IIf(filterby = "year", "range", "year")
'
'            .Controls(tglOr) = False
'        End If
    
        FilterRecords
        
    End With

'    If IsNull(Me.cbxParkFilter) = True Then Me.tglFilterByPark = False
'    fxnFilterRecords

'    If IsNull(Me.cbxAdminParkFilter) = True Then Me.tglFilterByAdminPark = False
'    fxnFilterRecords

'    If IsNull(Me.cbxSubunitFilter) = True Then Me.tglFilterBySubunit = False
'    fxnFilterRecords

'    If IsNull(Me.cbxFrameFilter) = True Then Me.tglFilterByFrame = False
'    fxnFilterRecords

'    If IsNull(Me.cbxPanelFilter) = True Then Me.tglFilterByPanel = False
'    fxnFilterRecords

'    If IsNull(Me.cbxLocationFilter) = True Then Me.tglFilterByLocation = False
'    fxnFilterRecords

'    If IsNull(Me.cbxStatusFilter) = True Then Me.tglFilterByStatus = False
'    fxnFilterRecords

'    If IsNull(Me.cbxYearFilter) Then Me.tglFilterByYear = False
'    If Me.tglFilterByYear = True Then Me.tglFilterByRange = False
'    fxnFilterRecords

'    If IsNull(Me.tbxStartDateFilter) And IsNull(Me.tbxEndDateFilter) _
'        Then Me.tglFilterByRange = False
'    If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
'    fxnFilterRecords
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetFilter[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     ToggleFilters
' Description:  prepare filter display based on filter being set
' Assumptions:  -
' Parameters:   state - whether the filters should be on, off, or clear
' Returns:      -
' Throws:       none
' References:
'   Microsoft,  unknown
'   https://msdn.microsoft.com/en-us/library/office/aa224135(v=office.11).aspx
' Source/date:  Bonnie Campbell, May 15, 2018
' Adapted:      -
' Revisions:
'   BLC   - 5/15/2018 - initial version
' ---------------------------------
Private Function ToggleFilters(state As String)
On Error GoTo Err_Handler
    
    Dim FilterState As Variant
    Dim ctl As Control
    
    Select Case state
        Case "on"
            FilterState = True
        Case "off"
            FilterState = False
        Case "clear"
            FilterState = Null
    End Select

    With Me
        
        .btnRequery.SetFocus
        
        'iterate through the controls to set to appropriate state
        For Each ctl In .Controls
        
            Select Case ctl.ControlType
                Case acComboBox, acTextBox, acToggleButton
                    If Left(ctl.Name, 11) = "tglFilterBy" And state = "off" Then
                        ctl = FilterState
                    ElseIf Right(ctl.Name, 6) = "Filter" And state = "clear" Then
                        ctl = FilterState
                    End If
                Case Else
            End Select
        Next
        
'   OFF
'        .tglFilterByPark
'        .tglFilterByAdminPark
'        .tglFilterBySubunit
'        .tglFilterByPanel
'        .tglFilterByFrame
'        .tglFilterByStatus
'        .tglFilterByLocation
'        .tglFilterByYear
'        .tglFilterByRange
        
'   CLEAR
'    'Clear the filters
'    Me.cbxParkFilter = Null
'    Me.cbxAdminParkFilter = Null
'    Me.cbxSubunitFilter = Null
'    Me.cbxPanelFilter = Null
'    Me.cbxFrameFilter = Null
'    Me.cbxStatusFilter = Null
'    Me.cbxLocationFilter = Null
'    Me.cbxYearFilter = Null
'    Me.tbxStartDateFilter = Null
'    Me.tbxEndDateFilter = Null
        
        
        FilterRecords
        
        cbxLocationFilter.Requery
        cbxEventSelection.Requery
        
    End With
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ToggleFilters[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SelectFilter
' Description:  prepare filter display based on filter being set
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 15, 2018
' Adapted:      -
' Revisions:
'   BLC   - 5/15/2018 - initial version
' ---------------------------------
Private Function SelectFilter(ctl As Control)
On Error GoTo Err_Handler
    
    Dim tgl As String
    Dim tglOr As String
    Dim filterby As String
    
    With Me
        
        'default
        tglOr = ""
        
        'determine filter
        filterby = Replace(Replace(ctl.Name, "cbx", ""), "Filter", "")

        'handle date ranges
        If Left$(ctl.Name, 3) = "tbx" Then
            tgl = "tglFilterByRange"
            .Controls(tgl) = (Not IsNull(Me.tbxStartDateFilter)) And (Not IsNull(Me.tbxEndDateFilter))
        Else
            tgl = "tglFilterBy" & Replace(Replace(ctl.Name, "cbx", ""), "Filter", "")
            .Controls(tgl) = Not IsNull(.Controls(ctl.Name))
        End If
    
        'handle year & range filters
        'year -> If Me.tglFilterByYear = True Then Me.tglFilterByRange = False
        'range ->If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
        If .Controls(tgl) = True And (filterby = "year" Or filterby = "range") Then
            'set tglOr = range if year, year if range
            tglOr = IIf(filterby = "year", "range", "year")
            
            tglOr = "tglFilterBy" & tglOr
            
            .Controls(tglOr) = False
        End If
    
    
        FilterRecords
        
        If Len(tglOr) = 0 Then
            .Controls(tgl).SetFocus
        Else
            .Controls(tglOr).SetFocus
        End If
        
        'handle location filters
        If Not ctl.Name = "cbxLocationFilter" Then
            .cbxLocationFilter.Requery
            .cbxEventSelection.Requery
        End If
        
    End With
    
'   .tglFilterByPark = Not IsNull(.cbxParkFilter)
'    FilterRecords
'    .tglFilterByPark.SetFocus
'    .cbxLocationFilter.Requery
'    .cbxEventSelection.Requery

'    Me.tglFilterByAdminPark = Not IsNull(Me.cbxAdminParkFilter)
'    fxnFilterRecords
'    Me.tglFilterByAdminPark.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

'    Me.tglFilterBySubunit = Not IsNull(Me.cbxSubunitFilter)
'    fxnFilterRecords
'    Me.tglFilterBySubunit.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

'    Me.tglFilterByFrame = Not IsNull(Me.cbxFrameFilter)
'    fxnFilterRecords
'    Me.tglFilterByFrame.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

'    Me.tglFilterByPanel = Not IsNull(Me.cbxPanelFilter)
'    fxnFilterRecords
'    Me.tglFilterByPanel.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

'    Me.tglFilterByStatus = Not IsNull(Me.cbxStatusFilter)
'    fxnFilterRecords
'    Me.tglFilterByStatus.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

'    Me.tglFilterByLocation = Not IsNull(Me.cbxLocationFilter)
'    fxnFilterRecords
'    Me.tglFilterByLocation.SetFocus

'    Me.tglFilterByYear = Not IsNull(Me.cbxYearFilter)
'    If Me.tglFilterByYear = True Then Me.tglFilterByRange = False
'    fxnFilterRecords
'    Me.tglFilterByYear.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

'    Me.tglFilterByRange = (Not IsNull(Me.tbxStartDateFilter)) And (Not IsNull(Me.tbxEndDateFilter))
'    If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
'    fxnFilterRecords
'    Me.tglFilterByYear.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

'    Me.tglFilterByRange = (Not IsNull(Me.tbxStartDateFilter)) And (Not IsNull(Me.tbxEndDateFilter))
'    If Me.tglFilterByRange = True Then Me.tglFilterByYear = False
'    fxnFilterRecords
'    Me.tglFilterByYear.SetFocus
'    Me.cbxLocationFilter.Requery
'    Me.cbxEventSelection.Requery

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SelectFilter[frm_Data_Summary_Advanced])"
    End Select
    Resume Exit_Handler
End Function
