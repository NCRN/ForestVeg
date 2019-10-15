Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    DataEntry = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10080
    DatasheetFontHeight =11
    ItemSuffix =31
    Left =7470
    Top =2280
    Right =17805
    Bottom =9045
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0xe7620d04d25ae540
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =255
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =2100
            BackColor =0
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    OverlapFlags =93
                    Left =240
                    Top =120
                    Width =1785
                    Height =495
                    FontSize =18
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="RIO Check"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =240
                    LayoutCachedTop =120
                    LayoutCachedWidth =2025
                    LayoutCachedHeight =615
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =300
                    Top =1080
                    Width =1245
                    Height =345
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblTag"
                    Caption ="# RIO Tags"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =300
                    LayoutCachedTop =1080
                    LayoutCachedWidth =1545
                    LayoutCachedHeight =1425
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =93
                    Left =3360
                    Top =1080
                    Width =1920
                    Height =345
                    FontWeight =600
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblInOffice"
                    Caption ="Actually In Office"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =3360
                    LayoutCachedTop =1080
                    LayoutCachedWidth =5280
                    LayoutCachedHeight =1425
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =223
                    Left =420
                    Top =600
                    Width =6450
                    Height =360
                    FontSize =12
                    FontWeight =500
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblDirections"
                    Caption ="Are all the Retired in Office (RIO) tags actually IN the office?"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =420
                    LayoutCachedTop =600
                    LayoutCachedWidth =6870
                    LayoutCachedHeight =960
                    ForeThemeColorIndex =-1
                    ForeTint =20.0
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8100
                    Width =420
                    Height =300
                    ColumnOrder =2
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="tbxDevMode"
                    DefaultValue ="0"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x010000006e000000010000000000000002000000000000000600000001000000 ,
                        0x00000000ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x460061006c007300650000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =8100
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =300
                    BorderThemeColorIndex =0
                    BorderTint =50.0
                    BorderShade =100.0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x01000100000000000000020000000100000000000000ffffff00050000004600 ,
                        0x61006c0073006500000000000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5280
                    Top =1080
                    Width =960
                    Height =345
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =16776960
                    Name ="tbxCount"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =5280
                    LayoutCachedTop =1080
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =1425
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =215
                            Left =5280
                            Top =720
                            Width =690
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblCount"
                            Caption ="# Tags"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =5280
                            LayoutCachedTop =720
                            LayoutCachedWidth =5970
                            LayoutCachedHeight =1035
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7200
                    Top =1020
                    Width =2580
                    FontSize =12
                    TabIndex =2
                    Name ="btnTagList"
                    Caption ="Save/Print RIO Tag List (PDF)"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open or print current list of tags with RIO status"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7200
                    LayoutCachedTop =1020
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =1380
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1680
                    Top =1080
                    Width =960
                    Height =345
                    ColumnOrder =0
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =62207
                    Name ="tbxRIOTagCount"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =1425
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =223
                            Left =1560
                            Top =720
                            Width =1425
                            Height =345
                            BorderColor =15921906
                            ForeColor =12566463
                            Name ="lblRIOCount"
                            Caption ="RIO Tag Count"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =1560
                            LayoutCachedTop =720
                            LayoutCachedWidth =2985
                            LayoutCachedHeight =1065
                            BorderThemeColorIndex =1
                            BorderTint =100.0
                            BorderShade =95.0
                            ForeThemeColorIndex =1
                            ForeTint =100.0
                            ForeShade =75.0
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =8880
                    Top =120
                    Width =1020
                    Height =450
                    FontSize =12
                    TabIndex =4
                    Name ="btnClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Close form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =8880
                    LayoutCachedTop =120
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =570
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
                Begin Label
                    OverlapFlags =85
                    Left =360
                    Top =1620
                    Width =9300
                    Height =480
                    FontSize =9
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblTotalRIOCount"
                    Caption ="* # of RIO tags is the maximum for the listbox below. Actual Total # of RIO tags"
                        " is found by printing full tag list. >> Actual # ="
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =360
                    LayoutCachedTop =1620
                    LayoutCachedWidth =9660
                    LayoutCachedHeight =2100
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
                Begin Label
                    OverlapFlags =215
                    Left =1500
                    Top =1020
                    Width =180
                    Height =285
                    BorderColor =8355711
                    ForeColor =65535
                    Name ="lblAsterisk"
                    Caption ="*"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =1500
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =1305
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =4260
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =300
                    Top =120
                    Width =1620
                    Height =4020
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxRIOtags"
                    RowSourceType ="Value List"
                    RowSource ="{1763A2E5-28C3-4C53-866D-6CB58C55AAF4};30;GWMP-0314;20101217085739-705547511.577"
                        "606;59;'';20101217085839-533424019.813538;66;'';20101217085857-579518616.199493;"
                        "72;'';{811199DD-A4FF-4FE3-AD48-9E109BA0EF77};82;NACE-0004;{1F0C862F-75A5-4D68-A9"
                        "2D-58B95607CA64};89;NACE-0004;{FCCB9FB2-1085-466D-A0BC-0614289C15E2};90;NACE-000"
                        "4;{BA1048C7-1B2D-4127-844E-CE08AFB51252};126;MANA-0027;{66FE7E59-053F-4DA3-93F8-"
                        "2F6A2630F47E};127;MANA-0027;{0A49842D-7402-4767-AFF6-5BB930350E32};128;MANA-0027"
                        ";{437A4DCA-D79E-4A42-869F-5A3DE942297B};130;MANA-0027;20101217085938-289562463.7"
                        "60376;140;'';20101217090009-301948010.921478;176;'';{8F8577EC-1D12-4CE8-8E56-6CF"
                        "58BAB3641};181;MANA-0002;{65404C05-F9F6-4F24-A99F-229DB5931E25};182;MANA-0002;{E"
                        "130916F-D65E-4680-9B6E-9261A18700B4};183;MANA-0002;{C1198A6F-E664-40F0-B0CC-7563"
                        "85E612B1};189;MANA-0027;20101217090036-774740099.906921;195;'';{2A7CB20B-92AA-48"
                        "77-96DB-3D4A1771B8A7};196;MANA-0002;{B96245B3-0649-417C-9B0F-9E713F4B419B};222;C"
                        "HOH-0577;{02E9DCCD-4920-453F-B417-5AE635951E90};229;CHOH-0577;20101217090127-140"
                        "17641.544342;231;'';{BB967F7C-2C87-4F64-8C7C-D5B8F1926563};321;PRWI-0238;{A5801B"
                        "70-D0E9-4DDF-AC33-5995973A9CB3};332;NACE-0341;{642B9AD7-B812-4378-B1C0-8F3EC6145"
                        "0CB};339;GWMP-0072;{3A5F82A5-DC88-4D90-AC14-743428515E1F};352;CHOH-0539;20101216"
                        "134655-705547511.577606;357;'';20101215130117-705547511.577606;363;'';{66E3C553-"
                        "FBFD-4698-A472-89A85D733D97};378;CATO-0330;20101217090144-760723590.85083;443;''"
                        ";{AD985F4F-0B3B-48F1-A1EA-8391D07D8D5C};451;PRWI-0722;{6619D036-654C-4DD1-AA1C-9"
                        "4987D2FC843};455;GWMP-0207;{DE0F9820-F08F-44D5-ABEF-284F1D41D695};462;GWMP-0207;"
                        "{972487EC-0D31-4551-BB91-7179C34B3409};464;GWMP-0207;{85469327-A3C0-402E-AB30-53"
                        "3EE9423010};465;GWMP-0207;20101216134745-533424019.813538;484;'';{CF9FEC17-24C4-"
                        "4344-A80E-C6BCD65F7338};567;MANA-0060;{DABE9900-B4ED-4CE2-A711-1C7A2FB54F48};578"
                        ";MANA-0025;{C5BE9BA3-5736-49ED-861A-85BABF042B1D};580;MANA-0054;{ADCE7C73-33F9-4"
                        "B90-A6D2-534D243C6593};582;MANA-0054;20101217090226-814490020.275116;585;'';{97A"
                        "EE9CB-0068-4772-9633-E00437373790};588;MANA-0060;{6C1912C4-1160-4062-BFE5-8B3972"
                        "6F28B5};589;MANA-0060;{6F53FBC3-CEA2-487C-B7A2-4BDF36E3F38C};590;MANA-0060;{C94F"
                        "80A9-F6D6-401F-A9AC-60112F37AF95};591;MANA-0060;{0F179D24-CD9F-4E0D-80D4-D850789"
                        "A63DA};594;MANA-0060;{EA089DCD-0D5A-44E0-858E-1504EC1D0385};595;MANA-0060;{D5789"
                        "609-CCD4-41FA-886D-16C243D8D85B};597;MANA-0060;{A86103A3-AC42-461A-8158-1EDA598E"
                        "A1F9};598;MANA-0060;{392CA244-8391-4627-9063-6C3FDE3E8CA9};599;MANA-0060;{E9433A"
                        "B5-1919-42B5-BA9E-9F823993E013};638;MANA-0253;{F32990B9-4978-4C0D-B4B4-2838C22F9"
                        "B36};642;MANA-0253;{C155BB62-B6DC-46DA-A8B6-C2E85CF63DE2};643;MANA-0253;{A721F03"
                        "0-EFB3-4DCB-8E05-2C7E66634979};645;MANA-0253;{32C8DA0D-8848-4FEB-80F9-E55A0603E6"
                        "94};647;MANA-0253;20101217090250-709037899.971008;650;'';20101217090311-45352756"
                        ".9770813;681;'';{E32CB599-7B10-42EE-8097-37A19D94ABF7};697;WOTR-0008;20101217090"
                        "405-414032697.677612;706;'';20101217090457-862619340.419769;716;'';2010121709051"
                        "7-790480017.662048;734;'';{FF8544B8-356D-42CF-A5C4-41BDF0E5F33B};738;CHOH-1201;2"
                        "0101217090552-373536169.528961;751;'';{84E3A402-00BB-4F36-AD2F-C25D0736D4A1};767"
                        ";CHOH-1191;{90B100A0-F761-42C0-99DF-FFF07CBA335D};768;CHOH-1191;20101217090806-9"
                        "61953163.146973;769;'';{1EBE99C3-2DF1-4FE9-9D72-591BF78351AC};815;MONO-0044;{5BE"
                        "6E479-77C0-49A4-A6DD-4F30CE62D32F};848;CHOH-1045;{2AE3A1C0-43E0-4971-92E7-FDDD30"
                        "900233};849;CHOH-1045;{4235F137-B0CE-4C61-A68E-CAAD04B48643};850;CHOH-1045;{F374"
                        "5B65-0CC4-46BA-901E-FECAFE715B0C};852;CHOH-1045;{CE6FFB83-4A41-4424-9EF9-22A2588"
                        "0EEE2};853;CHOH-1045;{982B7667-946F-4263-96C3-B11C5BEA848B};854;CHOH-1045;201012"
                        "17090824-871445834.636688;855;'';{C3B811F5-8866-4B82-A463-D6075CA75128};857;CHOH"
                        "-1045;{920587E0-CF90-4C02-81B0-417BE9E446E8};858;CHOH-1045;{770D648A-50C3-4D39-8"
                        "7BA-A47728A47B41};860;CHOH-1045;{F28DC5D4-4604-47E5-A731-F86EE0894532};861;CHOH-"
                        "1045;{C0684932-8389-4BBD-8DA3-E34E8F363EDB};862;CHOH-1045;{45A4DB0E-FD2A-425E-A5"
                        "B9-10BF92B99A05};863;CHOH-1045;{3B36D3E7-5F3E-4365-A62A-98D34D87CE1A};864;CHOH-1"
                        "045;{A11BE655-95F8-4CEE-8E3F-B7BC28218EEC};866;CHOH-1045;{B3D9178F-6BEF-43CA-89F"
                        "C-E3701F3274D6};867;CHOH-1045;{542675FC-DCD9-450D-BE84-3D69E08281E3};869;CHOH-10"
                        "45;{F873D34A-29A4-4D5F-9864-58DD42A146BC};870;CHOH-1045;{84ECD7AE-19EA-4F79-8556"
                        "-71EBA3873FEA};872;CHOH-1045;{7798BB31-8091-4D15-8BED-DB0A12AD7A5C};873;CHOH-104"
                        "5;{5F6A357D-6AA6-43B0-A8F3-58AE3B348E30};874;CHOH-1045;{7407F482-11C4-4977-B31B-"
                        "44382E793449};875;CHOH-1045;{49BF19B0-3359-4820-BD34-3BCB0B6DE25A};876;CHOH-1045"
                        ";{D113C1DF-167A-4BA0-8C9C-B75213358ABA};877;CHOH-1045;{863B39D4-AB74-41B1-B08B-C"
                        "68A60A3CD5A};878;CHOH-1045;{38E384F0-7FE9-4E1E-8ADE-9D8ED588C543};879;CHOH-1045;"
                        "{8BBFA9E3-3B91-4C5A-B11E-002894174787};880;CHOH-1045;{67DB243E-C800-4D78-B0D2-89"
                        "88B008BAB2};881;CHOH-1045;{3328593F-EA48-4AD0-839F-45457E92D63B};883;CHOH-1045;{"
                        "38EC74AC-35E1-4C4A-9ED2-ABDF0FFBDE7D};884;CHOH-1045;{91D22EA7-2251-46EF-99B7-D59"
                        "FF8F55E04};886;CHOH-1045;{65A24BC7-1948-4D3E-A1F4-3D1CF45B0A9B};887;CHOH-1045;{8"
                        "0705C5E-6666-4E07-814B-F59D3BEEE237};889;CHOH-1045;{9A2837EA-1217-451D-97F3-195E"
                        "7B813E6D};891;CHOH-1045;{1B0BF311-0B78-4CF7-BF2F-AEFDA3B261C1};894;CHOH-1045;{8E"
                        "66FDE2-1F78-4F7B-B594-385D9C5304C7};895;CHOH-1045;{9251277A-33EF-4568-B294-40D8A"
                        "25FD822};896;CHOH-1045;{E3F2EDED-B459-4AEE-A4C7-CA2E87E7ED9D};897;CHOH-1045;{4F6"
                        "8B14F-7A0C-421E-8BE9-9A2797D466D2};898;CHOH-1045;{BE50C783-E3E8-4B82-8D91-5D5099"
                        "F95948};899;CHOH-1045;{06F61E35-B195-48FE-AC2A-94DD28CC1029};900;CHOH-1045;{E979"
                        "5F51-B1C3-48F9-8A0D-919A506B89A1};901;CHOH-1045;{93F909DF-052A-4568-8AE3-67FC7D6"
                        "8375C};902;CHOH-1045;{E67142B4-EA08-4CA8-ADB5-B11805B9E6E6};903;CHOH-1045;{2DF82"
                        "506-DB46-45EB-B07E-F3B53273FC46};904;CHOH-1045;{8C5C62E1-FA9D-4271-8C46-A5F8A14E"
                        "E4A9};905;CHOH-1045;{57B0AB24-05FB-4466-A533-6027850ED62C};906;CHOH-1045;{AB0A3A"
                        "22-601B-4B4D-9503-D3E9FBD06E00};907;CHOH-1045;{8AD63D7C-B82F-423A-B5B8-A8A41D78B"
                        "05D};908;CHOH-1045;{81A38AC9-91B6-448F-BA00-66D06700A3BF};909;CHOH-1045;{B07BF84"
                        "1-E6C3-473C-A2CF-349223EECED6};910;CHOH-1045;{7C436B3C-917A-42E3-B1F0-C453EF15BF"
                        "78};911;CHOH-1045;{D7CB10DB-D827-4D08-AEA4-7B2C0150E0DB};913;CHOH-1045;{10B3B2D9"
                        "-6B3E-4DD7-ACB1-EB2DFE3CE2B6};914;CHOH-1045;{157BADCC-A7B9-4F2A-BC38-B3BBCFF2066"
                        "C};915;CHOH-1045;{00BFBD44-A135-493D-AFDB-2CB8DBB7363D};916;CHOH-1045;{D1E4D3A8-"
                        "11DD-48E5-BC38-231859A368EC};917;CHOH-1045;{2A92A705-D764-45EA-B464-068D0D669115"
                        "};918;CHOH-1045;{34ADF7AF-801A-4E69-8B14-6885067606BA};919;CHOH-1045;{E91E0803-9"
                        "088-4F6B-8A39-C7D7FC2A7CEC};920;CHOH-1045;{D3434992-AF03-4C28-B3E3-E10DB246DFC6}"
                        ";921;CHOH-1045;{132B0113-E31D-46B9-9D3C-A188127B17BD};922;CHOH-1045;{B1787BC9-BE"
                        "02-4666-B58D-F416E40FCC49};923;CHOH-1045;{D7EBC52F-100F-489C-A13E-F7C1F7FF97AA};"
                        "925;CHOH-1045;{4561A20D-BD65-4883-8865-D0FE39BE5785};926;CHOH-1045;{442FEFD9-255"
                        "D-44B5-8E88-2FCD3698FC2E};927;CHOH-1045;{98FEE5D3-CC0D-4952-BF7E-3B8E037C5701};9"
                        "28;CHOH-1045;{4F6D0A47-989F-4C80-99F1-39AE59436C16};929;CHOH-1045;{BF17F9B4-DE23"
                        "-4911-9EB6-DD5BC0F7C7A7};930;CHOH-1045;{8B500C3F-3F4E-47C3-A667-2EB0F134D355};93"
                        "1;CHOH-1045;{DC7E0831-59A4-4777-A037-FA3523728503};932;CHOH-1045;{8B9AD4EA-1031-"
                        "41BA-95A4-F1BB216FAA5E};933;CHOH-1045;{D78AEAAA-3290-48CA-9898-E595E647AC1E};934"
                        ";CHOH-1045;{20AB759B-94DB-4C8B-8332-5DF192E946EB};935;CHOH-1045;{F76BFD91-915A-4"
                        "895-A23C-8C2E5BA5D235};936;CHOH-1045;{16528833-7559-4E05-8A55-8AA9B5AC23CB};937;"
                        "CHOH-1045;{DDDBEC3D-AFC4-495D-96B8-B598E59BC4BB};938;CHOH-1045;{787E9AFB-E206-4F"
                        "A8-BF31-054D9A037481};939;CHOH-1045;{96E3F77E-A4CF-4BED-B886-F06E05FA6E73};940;C"
                        "HOH-1045;{ED0D8B9C-D97A-474F-926A-49ACA820CD28};941;CHOH-1045;{66071FD1-0425-414"
                        "4-B3AC-20818D885581};942;CHOH-1045;{8F78575C-8D22-4AEA-BC8F-CA0607AB7B8B};943;CH"
                        "OH-1045;{01F67C03-6C73-4071-B833-AB37DFA4EEBF};944;CHOH-1045;{5213107C-02F1-4312"
                        "-9133-10B91D54B410};945;CHOH-1045;{18BC6962-2C31-434B-9EA7-FFA8B19BDBFD};946;CHO"
                        "H-1045;{8A73D5EF-863B-4D16-AF55-048D42842172};947;CHOH-1045;{175EE491-3FB5-47CB-"
                        "A08D-A8DA82A9D7EE};948;CHOH-1045;{751FD698-7FB7-4C3D-980F-C850E572997C};949;CHOH"
                        "-1045;{9FA5537F-DF9F-4391-A4C4-9E8E1DE57C53};950;CHOH-1045;{DA8BC292-4FC5-4208-A"
                        "402-31603DF470D9};951;CHOH-1045;{61DB1722-3EB4-4073-AB18-7A56BB57F5A0};952;CHOH-"
                        "1045;{A47F6B71-66BD-449A-BA65-08299D826AEF};953;CHOH-1045;{A9A5754E-585F-460B-AE"
                        "6D-87F656CEC141};954;CHOH-1045;{8E3688EF-4782-4367-9F8A-DB559CAD2375};955;CHOH-1"
                        "045;{B3FFD8C6-CA5E-43B5-8C77-3758AF8AB17F};957;CHOH-1045;{FF4CD415-D305-4756-BBD"
                        "6-C4CD94164CC6};958;CHOH-1045;{C1B3CC8E-006F-47B2-B6A3-78F32D748FD4};959;CHOH-10"
                        "45;{DB42FD4D-F2EC-4AA1-88C1-6DDB60ABFC53};960;CHOH-1045;{CBABC76A-F3CC-49A7-A542"
                        "-AC20798C4F4A};961;CHOH-1045;{5971E5FB-C4CA-4942-A2A6-BB86F189B583};962;CHOH-104"
                        "5;{1843E3F5-0F30-457B-A97E-1AE208C6F1CD};963;CHOH-1045;{4612CB3E-86CD-45EB-AE08-"
                        "A858B2ECC72C};964;CHOH-1045;{D990DB0B-6017-4C9A-ACAA-C16431E21D4A};966;CHOH-1045"
                        ";{75090CF9-1885-4360-8541-215715531489};967;CHOH-1045;{7BE99780-2EED-4B33-A1D6-0"
                        "CF4A465B476};968;CHOH-1045;{387C45F6-0AC2-461B-8DFC-A8DE2978F48F};969;CHOH-1045;"
                        "{C346F405-36A9-49A5-9355-518EEFD5BCA7};970;CHOH-1045;{D1A3A7B6-D63F-4216-A4BB-9A"
                        "BD057DD573};971;CHOH-1045;{BE786D96-9A0F-48D0-9DFB-E47E397B5268};972;CHOH-1045;{"
                        "02E49129-3482-4F97-8C72-9EB11E7C6E5E};973;CHOH-1045;{6C8E800D-5AC0-43FA-9E31-AFE"
                        "48862D047};974;CHOH-1045;{1A53DAFD-6613-4B06-B095-B15FF9BB0597};975;CHOH-1045;{2"
                        "4AA7A02-DEEC-4251-BBEE-C198F6399138};976;CHOH-1045;{E66A67CC-CC85-41AE-8CEF-1702"
                        "2984D04F};977;CHOH-1045;{3E1A330B-4345-44FE-8138-84C90DA0F033};979;CHOH-1045;{53"
                        "51F318-00BA-464C-886E-3D4B3E0DBDE6};980;CHOH-1045;{4259097D-140D-4CA0-AC9D-74E58"
                        "8AC7DCD};981;CHOH-1045;{C606CD0A-1D33-4506-9874-813E4F2513AE};982;CHOH-1045;{232"
                        "CC5D7-4034-46D7-B520-EAD545588AF5};983;CHOH-1045;{49610D26-8D2A-43D2-ACF1-D756EF"
                        "905B6C};985;CHOH-1045;{AF0F2919-8D2A-4E33-9D16-4C3FE7A2E377};986;CHOH-1045;{F6B8"
                        "3FD8-025A-4320-93CB-C2F77BD24A78};987;CHOH-1045;{B612F3AE-1E0C-49EB-BD8C-ADC92BB"
                        "D6EF4};988;CHOH-1045;{67B85C1A-0F63-4FC4-B9F2-A628C1CBF449};989;CHOH-1045;{4144E"
                        "994-9010-4304-9460-279CC14F1876};990;CHOH-1045;{3E0AEFBF-6046-49B6-AE19-7112AE06"
                        "33AC};991;CHOH-1045;{7ACA9A64-A1BE-4AE3-A75F-DE1FA226D38C};992;CHOH-1045;{FB938C"
                        "84-F48D-47C9-8C70-DB580DCF43BC};993;CHOH-1045;{7AE59BEF-807F-4BB6-BDEF-38A11F76F"
                        "925};994;CHOH-1045;{193786C6-C42C-4804-A814-CEF839581343};995;CHOH-1045;{163C843"
                        "0-A513-4E0F-BD83-14AC6396E052};996;CHOH-1045;{7C39D8CF-B747-4888-A668-64BECA1640"
                        "13};997;CHOH-1045;{4DF8ADF3-68E7-4383-B2E1-F4E80202C230};998;CHOH-1045;{CD5014C8"
                        "-D693-497D-99CF-7C68B908D7A9};999;CHOH-1045;{F9CFD8C8-621D-470D-B298-E3B5ABD59D6"
                        "7};1000;CHOH-1045;{26E6D684-159E-4F0B-B083-BF47460970CF};1002;CHOH-1045;{0D91897"
                        "A-9EF2-4DE2-93F8-61ED96209218};1003;CHOH-1045;{9460877F-C20A-4F1F-A4DF-986A2741D"
                        "F63};1004;CHOH-1045;{25169879-B6B3-40C7-ADF9-259D78B9A56C};1005;CHOH-1045;{E3FB7"
                        "614-10CD-4793-91C6-A8B611B44D2B};1006;CHOH-1045;{A179FFC3-34D9-4AE3-99A0-8319C49"
                        "21214};1007;CHOH-1045;{C0AFEFCF-A5D2-49A1-BEC1-DEE1CF27FBC6};1008;CHOH-1045;{E92"
                        "CF4E8-86C0-45A1-A1FA-6038E99943A2};1009;CHOH-1045;{392BD440-6B12-4F5C-AC71-F436F"
                        "04EC943};1011;CHOH-1045;{DBE3F35F-89CD-4919-B6B4-B9BDA9F99ABA};1012;CHOH-1045;{0"
                        "9F32A5E-FE4D-4799-94F8-4B9B0D51E8B3};1013;CHOH-1045;{8204B161-ABB4-404A-86DC-DC0"
                        "AFB0A82B3};1014;CHOH-1045;{4CCF3ADD-6249-4776-AF08-777047C89325};1015;CHOH-1045;"
                        "20101217090932-56236863.1362915;1016;'';{BA154219-CDBC-4C56-86BB-24BCB9A4E195};1"
                        "017;CHOH-1045;{DEBDD660-D4C8-4195-9E6E-0C0696B028E4};1018;CHOH-1045;{A5352673-12"
                        "4B-4560-8784-FE7C96E2C7FB};1019;CHOH-1045;{17F8424D-3AEF-4044-9DE3-B4A91FEC826C}"
                        ";1020;CHOH-1045;{2089BE6E-73B6-4170-B74B-EBF2784465DF};1021;CHOH-1045;{5E6339D1-"
                        "1D5A-4788-9352-681E7F073FDA};1022;CHOH-1045;{F4DFF3C5-DA59-4570-8768-0E4C72C37F1"
                        "B};1024;CHOH-1045;{BED687F7-947F-4B89-BA55-3EC74A917469};1025;CHOH-1045;{0F0A177"
                        "3-72C0-4246-917B-DA067F987D57};1026;CHOH-1045;{A09AFCB4-F0E1-4C39-A245-C33FA258B"
                        "CE5};1027;CHOH-1045;{31F20D28-DD78-4DCD-B96B-7049F775F4EC};1028;CHOH-1045;{DF87F"
                        "DE0-42B6-434C-81C8-B5653A1EE941};1029;CHOH-1045;{E2F3A20D-2A34-45ED-B0F3-7B97394"
                        "D9EA3};1030;CHOH-1045;{8670E9EA-64C1-4C80-A100-0235E3C25107};1031;CHOH-1045;2010"
                        "1217091531-949556648.731232;1032;'';{726BB18F-5B59-48A7-985C-E1102468DDE4};1033;"
                        "CHOH-1045;{2BD0B62C-59D3-4C2C-90B6-5E7145D6B86A};1034;CHOH-1045;{56556814-C6DE-4"
                        "128-BD93-90C3D85FFD7F};1035;CHOH-1045;{E0463024-D5F9-430F-9D90-E516984CAA2B};103"
                        "6;CHOH-1045;{AC3C6149-AA7F-4932-9615-6145C0950926};1037;CHOH-1045;{C9661088-A438"
                        "-41A3-B083-B77C37918F45};1038;CHOH-1045;{EF4FC39D-7A3D-4429-9E95-75020403F8A7};1"
                        "039;CHOH-1045;{1B0B48DA-6E89-4FC5-A6D5-BFE473EDE4C9};1040;CHOH-1045;{64B72FDC-E0"
                        "AF-4F2F-B504-F4E6AA678C11};1041;CHOH-1045;{6AA00706-EB55-4AD5-A2DE-AB72663E1D58}"
                        ";1042;CHOH-1045;{EFC0C5C1-F0CF-4A57-9118-E1F7E0D35000};1043;CHOH-1045;{EE2FE5DB-"
                        "EEA3-4B5D-A19A-11081B3C8D4F};1044;CHOH-1045;{92DD4EBC-27C8-466B-B817-EDDD0AD5649"
                        "C};1045;CHOH-1045;{4623FBC6-28F4-40DD-BB0D-0153F0070716};1046;CHOH-1045;{172D407"
                        "8-61C9-445B-9B07-8BBEAEBF4BDF};1047;CHOH-1045;{176E1512-3ED5-4F7D-A370-6C65D08A6"
                        "1D6};1048;CHOH-1045;{5F0BF1F5-D836-4421-A8E1-219958F7A987};1049;CHOH-1045;{F260F"
                        "3D7-2382-4EE7-8280-7274848355C6};1050;CHOH-1045;{282AB080-CE14-46A7-9C23-E613480"
                        "17358};1051;CHOH-1045;{6B177557-44DC-409D-A695-950BE8BF5BE2};1052;CHOH-1045;{A2A"
                        "68078-E5DB-4590-89D2-F51B6CD98CC0};1053;CHOH-1045;20101216092933-705547511.57760"
                        "6;1054;'';{E986F8B5-D4F5-4405-847A-F644B5DF04DE};1055;CHOH-1045;{B6EA8057-5CA0-4"
                        "424-8D58-F774DF6D05C2};1056;CHOH-1045;{5175A744-BA30-42CB-98E6-B4E2E6F0178F};105"
                        "7;CHOH-1045;{30F2DF5D-8CEC-4581-ADE5-3EBD6EBB45BE};1058;CHOH-1045;{5DC4247A-8E39"
                        "-417C-8A1B-F5CBF672D86E};1060;CHOH-1045;{33952F05-F1A1-4D12-B216-7BF3558B0D43};1"
                        "061;CHOH-1045;{7A65BCDB-35FC-4431-AEC7-1AB07F24ACDD};1062;CHOH-1045;{E9BE539C-30"
                        "24-4B2C-9D76-EB13AA81A4B0};1063;CHOH-1045;{EBBE517B-4107-45FB-B908-948ECDEAF145}"
                        ";1064;CHOH-1045;{3E311D14-B682-4E53-A795-4AB0CE0D3964};1065;CHOH-1045;{7A72DCBD-"
                        "2659-4BE5-8921-8E0B6D8BE483};1066;CHOH-1045;{2CCA31D9-BC47-42E5-8A7A-A79A4B3F511"
                        "8};1067;CHOH-1045;{DC35A53E-80E5-4DC6-A60B-975BBB0CB98B};1068;CHOH-1045;{7E351B0"
                        "B-4C5B-4344-A4BF-5ABDC2A5DE08};1069;CHOH-1045;{1C6D314E-5A57-4372-B617-382FDED74"
                        "A13};1070;CHOH-1045;{F432CEF2-CA41-44E8-8CFA-E6FD6E13202B};1071;CHOH-1045;{DD2D7"
                        "A91-5C8E-4710-8A6F-EB852DC0694A};1072;CHOH-1045;{FE116B59-7D87-4712-BD0A-CEE974C"
                        "8953C};1073;CHOH-1045;{BA52CC4F-B9F7-4FBC-9C65-3CF3DC736AC4};1074;CHOH-1045;{279"
                        "8B4A9-57E3-4981-A5D3-45CC012B727D};1075;CHOH-1045;{5C9C315F-F269-4DF5-AB38-E6D6E"
                        "A18D269};1076;CHOH-1045;{9A341E3C-AF4E-4109-BCD8-FD36438500EF};1077;CHOH-1045;{F"
                        "17EB81F-A24F-4CC1-A6F5-A9E391341244};1078;CHOH-1045;{188B3134-EF12-4D4C-B1B7-9E0"
                        "CBA3FA66F};1079;CHOH-1045;{8E0F638A-F50A-4E24-BBCC-32AB361471FD};1080;CHOH-1045;"
                        "{37173424-C858-481F-B042-62A22519D369};1082;CHOH-1045;{53C710C5-16E3-46A8-944A-7"
                        "6B625C2C419};1084;CHOH-1045;{6D2B00F0-E435-462D-8337-A355465DF877};1085;CHOH-104"
                        "5;{3F2A8DB1-B8A8-434F-AF54-AA8A1B210295};1086;CHOH-1045;{C51DBD28-8114-41A6-A281"
                        "-BCC71F032D38};1087;CHOH-1045;{F490103F-6CF1-49A6-BFEB-C0C0A3ACDCDC};1088;CHOH-1"
                        "045;{D0BAA3C5-7C13-42FE-9ADC-F6732FA3E4B5};1089;CHOH-1045;{839D34EE-FECF-4794-94"
                        "24-A1270CDE598C};1091;CHOH-1045;{4153D66B-2FA6-49B2-912C-DA355E0AF04D};1092;CHOH"
                        "-1045;{C476A064-A6F4-4A4F-85C8-3D7E4B501D60};1093;CHOH-1045;{9928D4A8-BA02-423B-"
                        "94D5-F3861BC48405};1100;CHOH-1045;{D7A3E743-3619-45BF-A5AB-EABDDE046805};1124;NA"
                        "CE-0491;{961939E6-8B25-4EF2-A639-3C8CB58584A5};1128;CHOH-1055;{570AF8B1-DD72-431"
                        "D-B079-F8E67EBE5F24};1129;CHOH-1055;{A4880FAE-0001-4B97-B677-9AEB5B820CFA};1179;"
                        "CHOH-1055;{449A3D1E-5AC8-432A-828A-8F6E8FCA7700};1180;CHOH-1055;{029DC260-E227-4"
                        "FB0-9F4B-94F6326CE92F};1182;CHOH-1055;{521083E9-58D2-4403-AD4A-E14E4C32B6DE};118"
                        "3;CHOH-1055;{8F636C13-09DB-44CA-85C0-0A67F832A28A};1186;CHOH-1055;{8BA10A01-D732"
                        "-48D9-833E-FC266850BBBC};1187;CHOH-1055;{BE1A3673-925A-4916-A141-C87BC63695D5};1"
                        "188;CHOH-1055;{6929D1C0-69B9-4F67-8914-F800C80FE5AB};1189;CHOH-1055;{601992B3-3F"
                        "8F-4724-9025-6705D8FE15E6};1191;CHOH-1055;{F66F5819-5D10-4590-B29C-E8A2E97AC9D7}"
                        ";1193;CHOH-1055;{78EE21D5-8C82-48B1-BE7C-F93203EE0F29};1215;CHOH-1063;{015FC0B3-"
                        "55DF-42D8-A154-4C1D7F17F517};1216;CHOH-1063;{0EE4DFD0-07BF-4D3D-AA31-083DBDC59FD"
                        "2};1217;CHOH-1063;{C1F35917-9939-45F7-BE66-B0301618E56F};1218;CHOH-1063;{7610DC2"
                        "4-DABE-428C-8AA9-3DD7800FBD1F};1219;CHOH-1063;{A4A4021E-A16F-48C7-9748-57F99D25F"
                        "5F5};1221;CHOH-1063;{192DA6F1-6FA5-4B69-8FDC-0F49D76829A4};1222;CHOH-1063;{6C89F"
                        "1C3-FC45-4EF3-8F9D-AB1DEE3C6D55};1223;CHOH-1063;{EB91463D-E65D-45EB-ADE3-B1C1C9D"
                        "1510E};1253;GWMP-0208;{593B8369-A1AF-4EDC-8D0D-0551BA2EA94B};1254;GWMP-0208;{82E"
                        "1B176-C539-412A-9EAE-7D20AE01A14E};1255;GWMP-0208;{0BAF0BC6-4A73-4FAC-9589-B9724"
                        "B9BB919};1256;GWMP-0208;{AA1A72EA-6231-415E-9B88-F34C94BF94EA};1258;GWMP-0208;{1"
                        "8566D5A-D7A6-48F8-B11C-D2B348140E8D};1260;GWMP-0208;{963AB190-D935-4FCA-BDCD-394"
                        "15402CB40};1273;NACE-0245;20101217091550-364018678.665161;1288;'';{99649B9D-CAE8"
                        "-45AD-B399-4F40FE10B975};1291;CHOH-0776;{5BEB250D-D8DA-4BC3-8161-3842A5604C7F};1"
                        "293;CHOH-0776;{B70444AA-498C-478B-9DFC-B1ECAF25FAAD};1294;CHOH-0776;{C4A5DE6D-9A"
                        "E3-4F02-8B4B-652DD5180B79};1300;CHOH-0776;{D3C5DE05-ED48-495E-97B2-C0A3515511D6}"
                        ";1301;CHOH-0776;{214585F0-321F-4228-9F9E-9153B4AA5B90};1302;CHOH-0776;{648A6AA8-"
                        "5799-4AA5-ABE2-4D7C2453F8BF};1303;CHOH-0776;{534C24D6-98F4-459D-A6E8-1C12014FAB9"
                        "7};1304;CHOH-0776;{4F573DE8-36F0-4AA8-929B-ABE60B9DFDB1};1305;CHOH-0776;20101217"
                        "091844-767111659.049988;1307;'';{B45646CF-6B24-47F0-A6FA-D4B9444C324C};1311;CHOH"
                        "-0776;{D079092E-3B85-4DE1-B033-9FCFFDB5484E};1312;CHOH-0776;20101217091926-53504"
                        "526.6151428;1361;'';{606C6B87-E89B-4171-A579-839069D6847D};1364;CHOH-0788;{E593D"
                        "5E7-11DF-4485-B840-7D8D9DE49C08};1399;CHOH-1338;{22C7412B-8B7D-466F-8CAD-2B450C1"
                        "3AD8F};1400;CHOH-1338;{30911FF8-2CF3-4E87-B522-A695117ADD44};1401;CHOH-1338;{746"
                        "38E7D-DA0E-4EB5-8FD4-68472964E864};1402;CHOH-1338;{8C4BC8D9-D0BF-4349-B9CA-C3BC3"
                        "EA03685};1404;NACE-0087;20101217092005-592458248.138428;1416;'';{F2B602BC-CC86-4"
                        "E78-805B-C0D11408DB83};1427;NACE-0087;{16787EF5-7609-43B3-99BD-8CD6D8DB05AE};145"
                        "0;NACE-0087;{D0A656EF-8BF0-4E4F-9FBF-307698022EC9};1455;CATO-0365;{6DB088CE-5E2B"
                        "-46A2-9657-9FE08973063A};1462;CATO-0365;{56103498-C9DE-4170-BD0E-10F395859D2C};1"
                        "475;CATO-0365;{A3DDEFEC-4269-4FF7-B13F-3A78801D2916};1477;CATO-0365;{6A5E2513-FD"
                        "6E-4823-893D-2F48EBC9DC6C};1478;CATO-0365;{8B80D041-C0D0-4311-80DA-42DFD81E7DDC}"
                        ";1479;CATO-0365;{190B5B80-3382-479C-9449-0A2325840EC1};1480;CATO-0365;{0B016440-"
                        "14AB-4D61-81A5-85020DC81FBE};1481;CATO-0365;{3163702C-B43C-4070-94D4-B8005497DD9"
                        "9};1482;CATO-0365;{D6115F26-3988-4C99-AC43-65DFC339AFF3};1483;CATO-0365;{86D0689"
                        "2-BB06-4692-886A-41A22BB07F63};1484;CATO-0365;{5B27FEBB-36F2-4E69-AEA8-754B8B95C"
                        "F76};1485;CATO-0365;{895451D1-2CAE-4BEB-AD2E-05E3777266D1};1486;CATO-0365;{44923"
                        "969-CB49-418B-A727-43FDDFA3E6C5};1508;NACE-0337;{E2BA70CA-1E41-42E5-BBE8-E93E219"
                        "D1BD2};1509;NACE-0337;{BCFBB747-29DF-4584-ABFA-0995A58AB7A2};1514;NACE-0623;{DBB"
                        "91B05-1FB3-4B3A-A204-E2BE66BFB65C};1519;NACE-0623;{69CC35FF-CC6D-40D6-BFFA-36CE9"
                        "7A556C9};1522;GWMP-0008;{2D7EA8BD-DA39-4F77-A666-5AF5AC31C400};1523;GWMP-0008;{8"
                        "8067801-A1DC-4C60-85F3-4F8A3C706EAF};1524;GWMP-0008;{BCC7C6C6-07AA-4D40-B419-DB2"
                        "CF4ACCA84};1525;GWMP-0008;{86CD2EF1-A1F1-4712-946D-18EDD7D1CA41};1526;GWMP-0008;"
                        "{ADEEDD83-3B0A-423E-80DA-4BC8C11A6DE5};1527;GWMP-0008;{A8BC5993-8B53-49C0-AE1E-6"
                        "6ED5DEAAD99};1528;GWMP-0008;{FBD75FF3-7DD1-42AB-9DAA-B575C8937600};1529;GWMP-000"
                        "8;{6B4E354B-FC34-49F0-9D7B-B7E5581D6ADF};1530;GWMP-0008;{F19C3A37-9CD8-4F81-9D42"
                        "-DB191D3528E7};1531;GWMP-0008;{F2FDC458-8FDB-48F0-811F-FD6D22738BE3};1532;GWMP-0"
                        "008;{77FD22F8-4F08-4B44-BDB1-7E12DFD5A1D4};1533;GWMP-0008;{41F076A5-EFC6-4595-AD"
                        "BA-0B8F79C8B41D};1534;GWMP-0008;{66D24A93-C37F-4D9C-859C-4EC55C3C130D};1535;GWMP"
                        "-0008;{66D19566-0D99-43C6-B7D2-E2B7A4AF9814};1536;GWMP-0008;{800A3567-F3C1-42A1-"
                        "8937-2FA35812BFDD};1537;GWMP-0008;{F017B967-9FE8-44C7-813F-0F773E7BA1AE};1538;GW"
                        "MP-0008;{0DF249FF-A3BF-496D-850F-5220E140D9E1};1539;GWMP-0008;{12DC8542-53B2-49A"
                        "4-BEF0-EFBE2FEDC6A8};1540;GWMP-0008;{FB8B6B21-87DA-4CEA-ACB1-9E7BB9144EB3};1541;"
                        "GWMP-0008;{503D0124-9CCB-40F4-ABF2-888F00C2047F};1542;GWMP-0008;{2934F2E2-0F22-4"
                        "A48-9E6F-C271387DBDEF};1543;GWMP-0008;{4FDA2204-4DC4-42B2-B6A3-BD76F30EFC55};154"
                        "4;GWMP-0008;{26B541D2-AF59-4360-A2A6-42F8B365ABE3};1545;GWMP-0008;{91908879-D8C8"
                        "-4FAF-A904-29F6E978DDA7};1546;GWMP-0008;{EC723889-74DE-4ADA-990D-B3D284769922};1"
                        "547;GWMP-0008;{487EC2A6-E9AF-44C4-BDAE-0786BBBE752B};1548;GWMP-0008;{A2D7F91C-59"
                        "E6-4BCB-8385-EB729F32D25D};1549;GWMP-0008;{CD3E4312-626F-44C6-B4D5-E028108BF3CF}"
                        ";1550;GWMP-0008;{F948F902-FF0A-44D9-871A-3ECCFCA9AE4B};1551;GWMP-0008;{536294ED-"
                        "CED7-41E2-A117-423403E45EB4};1554;GWMP-0008;{450F4E61-4E64-48F6-BBC2-B66C82312D7"
                        "9};1555;GWMP-0008;{0E9C999B-1F7E-414F-A684-17A483C27DC5};1557;GWMP-0008;{15D3132"
                        "8-FAC4-4D0C-A1E1-2BCB5CE66E47};1560;GWMP-0008;{436B9C83-C896-4EEA-B754-F74F0C0EF"
                        "855};1562;NACE-0081;{A5AC973F-0128-47ED-88CF-F73BD7F948B5};1592;NACE-0081;{EAA7C"
                        "462-A158-4D62-A0EF-174374CF0A2B};1593;NACE-0081;{1FF154C1-999D-49A1-841C-B724258"
                        "CC1DD};1596;NACE-0081;{2F4F4617-9C89-4A53-ADF0-8709745CEC99};1598;NACE-0081;2010"
                        "1216134956-579518616.199493;1600;'';{F48BC7E7-BAA8-4DBD-8A3C-0A2396619673};1602;"
                        "HAFE-0215;{68921A17-0598-4093-8134-80DBFB15E8E2};1604;HAFE-0215;{32E78BCC-5A2D-4"
                        "EB9-B61D-9F122F1ECA38};1605;HAFE-0215;{E6DD4483-C544-4071-8FAD-8276A79B223A};160"
                        "8;HAFE-0215;{A866C3FE-4EA5-49DC-A62F-8E8391A8A1C7};1609;HAFE-0215;{EBE5AA73-66E0"
                        "-4F9B-93CA-F1B535D89528};1611;HAFE-0215;{04E180A3-D3C0-4D19-8FCB-07E29A168922};1"
                        "741;HAFE-0215;{77DB5B66-1B52-49E8-8E27-D9D8895755CA};1758;HAFE-0215;{F243C1F0-D8"
                        "3C-4BAE-AECC-6F4BCE8E4281};1762;HAFE-0215;{32A0D547-1E5D-4199-8FC7-2410095D5762}"
                        ";1763;HAFE-0215;{2E9D9339-30BF-4169-8528-BEB5A5A40F36};1764;HAFE-0215;{38B89FD5-"
                        "4956-4E53-BA83-69D4BE34930B};1767;HAFE-0215;{9D15EBD3-72D1-4304-9622-896EDE71672"
                        "4};1770;HAFE-0215;20101217092123-468700110.912323;1782;'';{74D2B909-B187-498B-84"
                        "8F-0D5BA7FB3A1E};1824;CATO-0294;{0112333B-D8BB-4B82-8A85-993F6EEA68E6};1836;NACE"
                        "-0399;{0F4DF3BF-8521-4A14-9A0D-641C57E315B8};1849;NACE-0399;{28AEB677-4014-4C29-"
                        "9C82-72CE2B95A68E};1861;HAFE-0074;{D018104F-EFDF-44F9-969A-6F267640E988};1896;CH"
                        "OH-0983;{909CC740-80A8-4C45-96CA-B4C34D113F19};1897;CHOH-0983;{1BA1F7ED-89BD-400"
                        "4-95FF-05142FD7700F};1899;CATO-0365;{26CFC989-0739-4010-9AD2-B0929473B633};1900;"
                        "CATO-0365;{96F1AC5E-C759-4455-92DE-FFD29D3F560F};1902;NACE-0399;{6CCF52FB-14D2-4"
                        "351-8715-27986C60A299};1904;NACE-0399;{326B6DD8-C5FA-4365-AFF9-57170882A130};190"
                        "5;NACE-0399;{DE60EC45-D555-4982-9DE3-971859449862};1907;NACE-0399;{C7137D07-2590"
                        "-4FDD-87D2-A2FD4B443608};1908;NACE-0399;{9100655E-03F6-458A-9751-A9EFCBB9E42F};1"
                        "910;NACE-0399;{6F9C717A-E087-4B08-B4FF-C59FE2EC13F5};1911;NACE-0399;201012170921"
                        "46-298165440.559387;1914;'';{7287F7BF-A539-4542-A927-195CC8D559DF};1920;CHOH-132"
                        "8;{9DBF8D58-AD95-4A5B-9308-224277E26910};1976;GWMP-0062;{B27F257A-86D3-41E0-9008"
                        "-E743C4783967};1978;GWMP-0062;{A967BE6F-F3FA-404E-A837-B132CE00D96C};1991;GWMP-0"
                        "094;{E0C7ABC8-5A1C-42BC-BECE-3986BD4C331E};1992;GWMP-0094;{73BC9EE4-D17F-41A4-8F"
                        "BD-7B2ECF9E5811};1994;GWMP-0094;{81507F4C-D83A-4894-B44B-9A6FADF1CAF5};1998;GWMP"
                        "-0094;{C93FD002-B6A5-4C7A-91F9-A06127919508};2000;GWMP-0094;{B2D8B76F-DDC6-4CAE-"
                        "9C6E-4220CD5FCF43};2002;GWMP-0094;{4A28A89D-B0E2-4260-9EB0-C579FA5E227D};2007;GW"
                        "MP-0094;{085BFB74-819F-4304-9C9B-F872978A6538};2009;GWMP-0094;{587EE971-25B7-4EB"
                        "6-9087-AD051785BFB1};2010;GWMP-0094;{6AC1D6D9-B636-4FC7-9C76-0799D44E4E05};2011;"
                        "GWMP-0094;{81DF55D8-B539-4EA3-AD8B-0F4E04217B8D};2012;GWMP-0094;{71486BFD-6207-4"
                        "E53-9ACA-31AD72CAEEFD};2013;GWMP-0094;{21DE83C3-F5B7-4440-9E58-B22CBBE621AD};201"
                        "4;GWMP-0094;{68836BE9-8AE9-4986-849A-C16ED7C8217A};2016;GWMP-0094;{A304B7F0-E2A3"
                        "-4416-B6E4-C983A81546DD};2031;CHOH-1328;{7738C7E4-F84E-440E-8DF6-86D198BB997E};2"
                        "076;PRWI-0277;{33182A9A-0C83-4E14-8E7E-38C6266B00D3};2077;PRWI-0277;{57043EF9-84"
                        "68-4CAE-AC7B-C8B2E43F2083};2079;PRWI-0277;{5A8346A1-BE96-4578-BF6D-A1B38838F4A3}"
                        ";2080;PRWI-0277;{7659F777-442E-4679-840D-295522D3B974};2081;PRWI-0277;{7D0E01C9-"
                        "4D62-4A0A-8156-323959FF6504};2083;PRWI-0277;{0820FAC5-1C0C-437D-95C0-4960D2CC6C0"
                        "A};2085;PRWI-0277;{115226DE-5987-4C8B-B454-5B422C86B1F1};2087;PRWI-0277;{1C8F310"
                        "3-CDBB-40C4-88E8-AA922B100536};2115;PRWI-0621;{42D01751-AC51-484B-A7D9-697EAA313"
                        "9E8};2126;PRWI-0175;{AB06481E-C9FC-42DA-806F-D39DC504E049};2128;PRWI-0175;201012"
                        "17092210-622696697.711945;2136;'';{F0A40EB1-1073-4A9A-ADDE-7CC6B6A99B2A};2145;PR"
                        "WI-0173;{C18C9306-3C74-4365-A4A6-07F01E93C59B};2146;PRWI-0173;20101217092232-647"
                        "821187.973022;2148;'';20101217092252-263792932.033539;2150;'';{40FB02DE-F027-4F5"
                        "F-AC24-87F3ED1BCC4B};2194;PRWI-0491;{4972E813-867B-45CC-97F1-4215FB0DC4E7};2197;"
                        "PRWI-0491;{D0D5CE80-9E26-4B68-B680-9EE03AEEFEAC};2199;PRWI-0491;{35CD7FEC-66BC-4"
                        "017-B9F3-943FDA48A685};2201;PRWI-0491;{B228D2BE-E523-4E59-B0F5-4D73B1090D3C};220"
                        "4;PRWI-0491;{15C672FA-9069-4152-9935-2800639A4404};2205;PRWI-0491;{2E045E34-C250"
                        "-47B8-BF26-0730000DF6D1};2208;PRWI-0491;{4D79F0E8-352E-4655-A48A-F00A3E11EBD9};2"
                        "209;PRWI-0491;20101217092314-279342055.32074;2220;'';20101217092336-829801619.05"
                        "2887;2224;'';20101217092418-824602127.075195;2232;'';{CBB83D3F-5A00-4564-8D34-7F"
                        "2CD497C258};2239;PRWI-0233;{57DD3833-428B-44AD-801E-30EDD1A23875};2280;PRWI-0796"
                        ";{2DA868C4-FE9E-4330-AA1F-11954400EA28};2282;PRWI-0796;{CCE665A4-F8F5-44E4-A6F8-"
                        "972A56D808EE};2340;PRWI-0282;{1B896D29-4263-413D-A5C8-E47C673A18B3};2346;PRWI-02"
                        "82;{5F3C1207-2C3B-47A9-A69C-849E4E74A286};2347;PRWI-0282;{C443AC99-2E3D-417E-97F"
                        "6-093B33945BA4};2368;PRWI-0728;{50E90F7D-6F54-4AD4-9D9D-22A036CC61A8};2374;PRWI-"
                        "0728;{0228BD9A-E23C-42B9-B554-0AEC16EBD7F1};2375;PRWI-0728;{5F0B5C45-AF51-4CAE-8"
                        "FB9-B4006ABCE06C};2411;PRWI-0199;{1D1AA1F2-BC1C-4187-A08C-BF9A9704BF44};2412;PRW"
                        "I-0199;{1945E116-E42F-431D-B689-09078104D2EC};2413;PRWI-0199;{C98785C4-781B-430D"
                        "-B148-073EE642D0B4};2415;PRWI-0199;{80CD8146-CC69-46EF-8E98-88929BFCE856};2416;P"
                        "RWI-0199;{005C5E09-0853-4768-9438-665D68B5E805};2417;PRWI-0199;{7F8D83BF-EFD1-4B"
                        "04-819E-AA78E3735BC9};2420;PRWI-0199;{EEF833E1-2088-4E7F-985E-F2DA3FBC938C};2446"
                        ";PRWI-0796;{5B8BE00C-2D09-444A-ACE1-D40AFB77EF0B};2494;PRWI-0181;{6D7AF34C-6B5F-"
                        "40B5-BA63-3FD795C395C6};2502;PRWI-0463;{F6F82047-E5D5-4CA1-B5CF-7C47680B431B};25"
                        "04;PRWI-0463;{22A5305C-D7B7-482A-9426-4A92150FC0B4};2505;PRWI-0463;{54AAB3E6-1DA"
                        "5-437A-BBDA-982E421FE88F};2506;PRWI-0463;{F9296A06-E6F4-456B-875A-A25B36D0E4BD};"
                        "2509;PRWI-0463;20101217092443-589163005.35202;2570;'';{611806EC-0BFD-4C08-9B90-E"
                        "96C2DF19825};2571;PRWI-0508;{D608A903-B444-4BF8-ADC0-A33590350D01};2601;PRWI-072"
                        "8;{C4B5AE52-0DC9-4728-90DD-D0F3E2606BFA};2604;PRWI-0728;20101217092506-986093163"
                        ".490295;2610;'';{13590D69-9AA0-4DE5-8C11-B1D65CA2ED2C};2639;PRWI-0751;2010121709"
                        "2526-910964310.16922;2670;'';20101217092544-226866006.851196;2683;'';20101217092"
                        "605-695115506.649017;2710;'';{F4E0EDB9-8CFA-42C5-8854-9B803EF2CCD2};2735;PRWI-02"
                        "23;{76531CBE-1634-449C-9409-3B2B8DDBFE6A};2739;PRWI-0223;{CE67810D-0ADF-41FE-83F"
                        "B-BD7439BCB603};2740;PRWI-0223;{3355245F-8FDA-457E-867A-99D396EDFD6D};2741;PRWI-"
                        "0223;{25C9BA45-59DE-4534-BC44-4AD07C231D7B};2742;PRWI-0223;{1614AEB7-C2EB-4420-B"
                        "9C1-52F59CB0F0D8};2743;PRWI-0223;{F824D09C-71CB-49F5-B376-5176FCC0C72E};2744;PRW"
                        "I-0223;{B69D463E-ADDC-4FC0-A3E4-4619BCF09843};2749;PRWI-0223;{B902FBB3-BEFC-498A"
                        "-A270-1E3363DDC84A};2752;PRWI-0223;{25BA8D2D-F3CD-4416-AB8F-484585A4EA36};2753;P"
                        "RWI-0223;{DD61111C-B5A6-4A2A-A31F-772A16AB7C4F};2754;PRWI-0223;{F24468EE-8DA1-47"
                        "BB-9DE5-B45861791C1C};2755;PRWI-0223;{39275F20-285B-4795-8A36-CB69C11C740C};2756"
                        ";PRWI-0223;{FEF472C0-7110-4F62-BD4F-84FE6F3D222D};2757;PRWI-0223;{639CDA57-AC36-"
                        "4C65-9B5A-CA9C1199FE54};2759;PRWI-0223;{999578F1-9D9F-48FB-957E-01554B370C62};27"
                        "60;PRWI-0223;{46846CC9-E9CF-4D44-B4A1-7EFC417518EB};2803;PRWI-0338;{A40E5000-2E6"
                        "4-4046-80C6-9281BFA02566};2806;PRWI-0338;{553462BD-C8D3-40F6-8A98-26EED31D8AEC};"
                        "2808;PRWI-0338;{648AF137-47DF-4960-9894-812E58581E61};2809;PRWI-0338;{C71E282B-3"
                        "947-4D4A-A65A-BBFFE4FDABFD};2810;PRWI-0338;{BE6E070C-3077-46A2-AAAD-208ABD0A9134"
                        "};2811;PRWI-0338;{3FD5937F-F2B0-4098-B8EC-23AFDBD2D3FA};2812;PRWI-0338;{591A44E5"
                        "-6014-4302-B996-651BBA804199};2813;PRWI-0338;{61778110-C49B-4215-B04E-6DC36E17E3"
                        "08};2835;PRWI-0145;{5C1A73EB-4DC4-42F7-8CB1-F61866E862A1};2842;PRWI-0145;{0C9A79"
                        "E2-FCFA-47C5-A23D-199C9969CA46};2843;PRWI-0145;{5158DEEA-033F-471B-8FCC-2E9B8EA7"
                        "0AD9};2875;PRWI-0085;{4B02E756-4F0D-4B02-A394-DFD61C8A4ED0};2877;PRWI-0085;{2C5B"
                        "FC45-3294-4821-AF01-23D4535AC3CC};2878;PRWI-0085;{029F7646-2D0A-4DEB-8E7B-80D8E2"
                        "DF9E29};2901;PRWI-0062;{80799BF3-5EAA-4FF4-9F3E-B607A2D2AF56};2903;PRWI-0062;{2B"
                        "E6CE4C-5BE2-4974-8231-93E19CCC8CA6};2904;PRWI-0062;{93DB3519-D20F-4996-86AE-A1C9"
                        "3067B20A};2905;PRWI-0062;20101217092623-980003237.724304;2906;'';{9AB058C1-9377-"
                        "4094-80B9-429D967C2220};2908;PRWI-0062;{185C6888-2F8C-45F1-978C-80D084A7B73B};29"
                        "09;PRWI-0062;{685D96FE-3F05-4987-A3AC-9B247B89BEDD};2910;PRWI-0062;{D36AE4D4-C50"
                        "6-4994-A993-D3B8B8E18FBC};2912;PRWI-0062;{876C1806-A54E-40B6-B850-9105CD4A7685};"
                        "2913;PRWI-0062;{F74469ED-6179-4B3E-ADEC-A305DE938A8E};2914;PRWI-0062;{CB2B8221-3"
                        "311-4ED2-A63E-36BF3C9CFF2F};2917;PRWI-0062;{1875481A-C3D5-425C-A8F0-CDA391913CA3"
                        "};2918;PRWI-0062;{92D68D87-34D4-44E2-929A-439142BD4E41};2955;PRWI-0075;{FECEB1B4"
                        "-877A-4399-8760-57CC9527E87A};2956;PRWI-0075;{DEDF9798-8B68-4E6C-B377-4FD61AAA92"
                        "32};2957;PRWI-0075;{04758A2B-D17F-43A5-9B27-FB2B3FC4F1DE};2959;PRWI-0075;{A984DB"
                        "94-9D1B-436E-9F5B-69B351933868};2994;PRWI-0093;{6034CF03-E1BF-453C-982D-A322AE4D"
                        "E076};2996;PRWI-0093;{48BF4E8B-D1BE-4A82-8975-E9D2545AFDA7};2999;PRWI-0093;{02CE"
                        "9859-3484-4B11-A098-042CD918B0C8};3093;PRWI-0238;{0DAE6B73-8186-436E-B83F-09C7AE"
                        "74D9DE};3108;PRWI-0238;{62C4413B-8D58-4276-AC02-20E889E146CF};3111;PRWI-0238;{5E"
                        "78C44E-675D-44FC-ABE5-8B8E324A4AC1};3114;PRWI-0238;{BC609F3F-3A20-4314-875F-F16B"
                        "3DC7BCA2};3131;PRWI-0051;{38C46633-76C7-49F8-ADB8-08B606B9D9B9};3133;PRWI-0051;{"
                        "B8739678-BD20-44FC-8CC1-41EBC04478D5};3134;PRWI-0051;{136235BC-BF57-47A6-92F1-EC"
                        "E312C57981};3142;PRWI-0051;{4B237531-15F5-413E-9FC3-43BDCCC17C6E};3176;PRWI-0080"
                        ";{499CC64B-0338-40F6-988A-BBB06B59F9E0};3180;PRWI-0080;{56B64C47-557D-411F-AF52-"
                        "3F2CFEE5F0FC};3200;PRWI-0398;{70BC2872-1A04-48D5-A283-65955BB4A977};3226;PRWI-03"
                        "98;{7435118E-BD19-42AC-895F-2360E3C98C4D};3256;PRWI-0398;{C957968A-DA6A-4785-903"
                        "1-F121B7722F92};3274;PRWI-0333;{0FBFC04B-923E-41FF-B3F1-16E6791CCD63};3280;PRWI-"
                        "0333;{613D2044-3439-4E43-A8C5-ECB07560097D};3281;PRWI-0333;{3EF0EDCD-CAC2-42D4-9"
                        "61A-735BC6CD8491};3283;PRWI-0333;{40D67A2B-3387-47C6-BB3C-6F166D08840D};3286;PRW"
                        "I-0333;{8130921D-C93F-4886-8107-A369C5716ADB};3289;PRWI-0333;{1644A632-DFDD-403E"
                        "-BFBB-B15CE9C64466};3326;NACE-0174;{C7681E7E-A6F1-4EBC-9019-30F0009D7BAF};3327;N"
                        "ACE-0174;{9FB29A1C-3AD8-4F48-9AB4-C1E3A27242D9};3330;NACE-0174;{B3F30786-E9C9-4D"
                        "B6-B592-57C453E72A94};3334;NACE-0174;{DAC7AD41-75B1-4A73-9D93-4BD0C1BE69C6};3338"
                        ";NACE-0174;{0AA4A4A1-5CC2-4C14-A747-691BD0DE2FE9};3340;NACE-0174;20101217092717-"
                        "243931353.092194;3342;'';{7B84841E-EE44-4B66-9514-EC033E4A5B78};3344;NACE-0174;2"
                        "0101217092736-533873081.207275;3345;'';{B1E7CD06-8D0E-4945-8F3A-42C70C643883};33"
                        "55;NACE-0174;{323EAE25-772D-4027-9324-C9F821F3FE18};3357;NACE-0233;{47DC4D81-B37"
                        "F-4B71-85D4-B2E0EAB48471};3363;NACE-0233;{6E964920-4183-4B9B-842E-C7A6A2A791AB};"
                        "3364;NACE-0233;{4FDF0F4D-E341-4EF2-AEA5-999396D00496};3369;NACE-0233;{3C5486B3-B"
                        "638-4B83-B7C5-7F9C97D660F7};3372;NACE-0233;{66CF0991-CB14-44DD-85D9-2407AAE652D5"
                        "};3374;NACE-0233;{99EE4A4B-DF37-43DA-B9DD-2998F904F392};3375;NACE-0233;{E2DD5920"
                        "-DC46-4A69-ADF9-3A372FAB0CA6};3403;NACE-0233;{6A33A218-EFB3-4F24-8C33-21F4256EAE"
                        "09};3404;NACE-0233;{FFF3A016-8286-4FBD-A35B-E08EEB1FB844};3405;NACE-0233;{D1B091"
                        "95-C025-4F50-9105-B296B8E0086A};3407;NACE-0233;{AE1C10E0-7D2D-4DB3-AD6B-A5AE81DC"
                        "13D2};3408;NACE-0233;{5AFE05D2-CD88-42F8-8F64-AC88241BBFEC};3410;NACE-0233;{0C36"
                        "C519-6D72-42CB-A80F-DEF000F65187};3411;NACE-0233;20101217092758-106369674.20578;"
                        "3418;'';{60B4D9F9-29B8-4DE2-BC3F-B2DB6BCD6E51};3432;PRWI-0436;{A3A74C8F-392C-494"
                        "8-9D2B-05940DF57A33};3434;PRWI-0436;{F1A99B2B-969E-4051-954B-0107B8F1A201};3435;"
                        "PRWI-0436;{69C02C22-B4D2-474A-8E91-9076EA0AB10F};3437;PRWI-0436;{6BA3D4FF-FF7A-4"
                        "D30-8311-7CC386B3ED6F};3438;PRWI-0436;{454A84C2-F777-448F-837D-1533115D7174};343"
                        "9;PRWI-0436;{0EAC4769-96E7-4BFA-AEE2-EFCCF5F841FA};3455;PRWI-0712;{04431296-F835"
                        "-4BB1-BCA7-65CFBE6A47A5};3486;PRWI-0712;{FE65D755-A72B-4F7F-A86F-5A4084CAF1F6};3"
                        "488;PRWI-0712;{4AC0022B-6ACE-439A-8CF5-8BD29320C008};3492;PRWI-0712;201012170928"
                        "18-999414563.179016;3495;'';{A61BE73F-D902-40B7-90AE-81E40595F418};3496;PRWI-071"
                        "2;{67026878-506F-438D-BA32-4B5AD18D07B8};3498;PRWI-0712;{EF43D706-68D4-4AE5-AEFF"
                        "-5395F4F96917};3510;CHOH-0440;{0A0AA766-3F0A-449A-8F43-4F2417EF2F07};3511;CHOH-0"
                        "440;{BA56A990-80EF-414E-BA91-638BAB5A2185};3512;CHOH-0440;{B66C146E-49CC-4B5D-92"
                        "64-68D353B16B2F};3513;CHOH-0440;{37BB8566-6BBE-4306-B10C-332DEBD056AE};3514;CHOH"
                        "-0440;{53FAAD53-782E-4477-930C-1704808F2DF7};3515;CHOH-0440;{AA3785E6-F1C4-472D-"
                        "8D3F-4B120D08D318};3516;CHOH-0440;{48184C7A-E8C5-4C74-9E61-6D40CEA36C5B};3517;CH"
                        "OH-0440;{C6795501-95FE-43AC-973A-9F7D170AA60D};3558;CHOH-0239;{00141FF6-9CA2-45C"
                        "4-BA47-E7B503271BAC};3559;CHOH-0239;{F3DF0448-2487-4D37-B41C-B5AD8E9B58D2};3591;"
                        "ROCR-0094"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    OnClick ="[Event Procedure]"
                    RowSource ="{1763A2E5-28C3-4C53-866D-6CB58C55AAF4};30;GWMP-0314;20101217085739-705547511.577"
                        "606;59;'';20101217085839-533424019.813538;66;'';20101217085857-579518616.199493;"
                        "72;'';{811199DD-A4FF-4FE3-AD48-9E109BA0EF77};82;NACE-0004;{1F0C862F-75A5-4D68-A9"
                        "2D-58B95607CA64};89;NACE-0004;{FCCB9FB2-1085-466D-A0BC-0614289C15E2};90;NACE-000"
                        "4;{BA1048C7-1B2D-4127-844E-CE08AFB51252};126;MANA-0027;{66FE7E59-053F-4DA3-93F8-"
                        "2F6A2630F47E};127;MANA-0027;{0A49842D-7402-4767-AFF6-5BB930350E32};128;MANA-0027"
                        ";{437A4DCA-D79E-4A42-869F-5A3DE942297B};130;MANA-0027;20101217085938-289562463.7"
                        "60376;140;'';20101217090009-301948010.921478;176;'';{8F8577EC-1D12-4CE8-8E56-6CF"
                        "58BAB3641};181;MANA-0002;{65404C05-F9F6-4F24-A99F-229DB5931E25};182;MANA-0002;{E"
                        "130916F-D65E-4680-9B6E-9261A18700B4};183;MANA-0002;{C1198A6F-E664-40F0-B0CC-7563"
                        "85E612B1};189;MANA-0027;20101217090036-774740099.906921;195;'';{2A7CB20B-92AA-48"
                        "77-96DB-3D4A1771B8A7};196;MANA-0002;{B96245B3-0649-417C-9B0F-9E713F4B419B};222;C"
                        "HOH-0577;{02E9DCCD-4920-453F-B417-5AE635951E90};229;CHOH-0577;20101217090127-140"
                        "17641.544342;231;'';{BB967F7C-2C87-4F64-8C7C-D5B8F1926563};321;PRWI-0238;{A5801B"
                        "70-D0E9-4DDF-AC33-5995973A9CB3};332;NACE-0341;{642B9AD7-B812-4378-B1C0-8F3EC6145"
                        "0CB};339;GWMP-0072;{3A5F82A5-DC88-4D90-AC14-743428515E1F};352;CHOH-0539;20101216"
                        "134655-705547511.577606;357;'';20101215130117-705547511.577606;363;'';{66E3C553-"
                        "FBFD-4698-A472-89A85D733D97};378;CATO-0330;20101217090144-760723590.85083;443;''"
                        ";{AD985F4F-0B3B-48F1-A1EA-8391D07D8D5C};451;PRWI-0722;{6619D036-654C-4DD1-AA1C-9"
                        "4987D2FC843};455;GWMP-0207;{DE0F9820-F08F-44D5-ABEF-284F1D41D695};462;GWMP-0207;"
                        "{972487EC-0D31-4551-BB91-7179C34B3409};464;GWMP-0207;{85469327-A3C0-402E-AB30-53"
                        "3EE9423010};465;GWMP-0207;20101216134745-533424019.813538;484;'';{CF9FEC17-24C4-"
                        "4344-A80E-C6BCD65F7338};567;MANA-0060;{DABE9900-B4ED-4CE2-A711-1C7A2FB54F48};578"
                        ";MANA-0025;{C5BE9BA3-5736-49ED-861A-85BABF042B1D};580;MANA-0054;{ADCE7C73-33F9-4"
                        "B90-A6D2-534D243C6593};582;MANA-0054;20101217090226-814490020.275116;585;'';{97A"
                        "EE9CB-0068-4772-9633-E00437373790};588;MANA-0060;{6C1912C4-1160-4062-BFE5-8B3972"
                        "6F28B5};589;MANA-0060;{6F53FBC3-CEA2-487C-B7A2-4BDF36E3F38C};590;MANA-0060;{C94F"
                        "80A9-F6D6-401F-A9AC-60112F37AF95};591;MANA-0060;{0F179D24-CD9F-4E0D-80D4-D850789"
                        "A63DA};594;MANA-0060;{EA089DCD-0D5A-44E0-858E-1504EC1D0385};595;MANA-0060;{D5789"
                        "609-CCD4-41FA-886D-16C243D8D85B};597;MANA-0060;{A86103A3-AC42-461A-8158-1EDA598E"
                        "A1F9};598;MANA-0060;{392CA244-8391-4627-9063-6C3FDE3E8CA9};599;MANA-0060;{E9433A"
                        "B5-1919-42B5-BA9E-9F823993E013};638;MANA-0253;{F32990B9-4978-4C0D-B4B4-2838C22F9"
                        "B36};642;MANA-0253;{C155BB62-B6DC-46DA-A8B6-C2E85CF63DE2};643;MANA-0253;{A721F03"
                        "0-EFB3-4DCB-8E05-2C7E66634979};645;MANA-0253;{32C8DA0D-8848-4FEB-80F9-E55A0603E6"
                        "94};647;MANA-0253;20101217090250-709037899.971008;650;'';20101217090311-45352756"
                        ".9770813;681;'';{E32CB599-7B10-42EE-8097-37A19D94ABF7};697;WOTR-0008;20101217090"
                        "405-414032697.677612;706;'';20101217090457-862619340.419769;716;'';2010121709051"
                        "7-790480017.662048;734;'';{FF8544B8-356D-42CF-A5C4-41BDF0E5F33B};738;CHOH-1201;2"
                        "0101217090552-373536169.528961;751;'';{84E3A402-00BB-4F36-AD2F-C25D0736D4A1};767"
                        ";CHOH-1191;{90B100A0-F761-42C0-99DF-FFF07CBA335D};768;CHOH-1191;20101217090806-9"
                        "61953163.146973;769;'';{1EBE99C3-2DF1-4FE9-9D72-591BF78351AC};815;MONO-0044;{5BE"
                        "6E479-77C0-49A4-A6DD-4F30CE62D32F};848;CHOH-1045;{2AE3A1C0-43E0-4971-92E7-FDDD30"
                        "900233};849;CHOH-1045;{4235F137-B0CE-4C61-A68E-CAAD04B48643};850;CHOH-1045;{F374"
                        "5B65-0CC4-46BA-901E-FECAFE715B0C};852;CHOH-1045;{CE6FFB83-4A41-4424-9EF9-22A2588"
                        "0EEE2};853;CHOH-1045;{982B7667-946F-4263-96C3-B11C5BEA848B};854;CHOH-1045;201012"
                        "17090824-871445834.636688;855;'';{C3B811F5-8866-4B82-A463-D6075CA75128};857;CHOH"
                        "-1045;{920587E0-CF90-4C02-81B0-417BE9E446E8};858;CHOH-1045;{770D648A-50C3-4D39-8"
                        "7BA-A47728A47B41};860;CHOH-1045;{F28DC5D4-4604-47E5-A731-F86EE0894532};861;CHOH-"
                        "1045;{C0684932-8389-4BBD-8DA3-E34E8F363EDB};862;CHOH-1045;{45A4DB0E-FD2A-425E-A5"
                        "B9-10BF92B99A05};863;CHOH-1045;{3B36D3E7-5F3E-4365-A62A-98D34D87CE1A};864;CHOH-1"
                        "045;{A11BE655-95F8-4CEE-8E3F-B7BC28218EEC};866;CHOH-1045;{B3D9178F-6BEF-43CA-89F"
                        "C-E3701F3274D6};867;CHOH-1045;{542675FC-DCD9-450D-BE84-3D69E08281E3};869;CHOH-10"
                        "45;{F873D34A-29A4-4D5F-9864-58DD42A146BC};870;CHOH-1045;{84ECD7AE-19EA-4F79-8556"
                        "-71EBA3873FEA};872;CHOH-1045;{7798BB31-8091-4D15-8BED-DB0A12AD7A5C};873;CHOH-104"
                        "5;{5F6A357D-6AA6-43B0-A8F3-58AE3B348E30};874;CHOH-1045;{7407F482-11C4-4977-B31B-"
                        "44382E793449};875;CHOH-1045;{49BF19B0-3359-4820-BD34-3BCB0B6DE25A};876;CHOH-1045"
                        ";{D113C1DF-167A-4BA0-8C9C-B75213358ABA};877;CHOH-1045;{863B39D4-AB74-41B1-B08B-C"
                        "68A60A3CD5A};878;CHOH-1045;{38E384F0-7FE9-4E1E-8ADE-9D8ED588C543};879;CHOH-1045;"
                        "{8BBFA9E3-3B91-4C5A-B11E-002894174787};880;CHOH-1045;{67DB243E-C800-4D78-B0D2-89"
                        "88B008BAB2};881;CHOH-1045;{3328593F-EA48-4AD0-839F-45457E92D63B};883;CHOH-1045;{"
                        "38EC74AC-35E1-4C4A-9ED2-ABDF0FFBDE7D};884;CHOH-1045;{91D22EA7-2251-46EF-99B7-D59"
                        "FF8F55E04};886;CHOH-1045;{65A24BC7-1948-4D3E-A1F4-3D1CF45B0A9B};887;CHOH-1045;{8"
                        "0705C5E-6666-4E07-814B-F59D3BEEE237};889;CHOH-1045;{9A2837EA-1217-451D-97F3-195E"
                        "7B813E6D};891;CHOH-1045;{1B0BF311-0B78-4CF7-BF2F-AEFDA3B261C1};894;CHOH-1045;{8E"
                        "66FDE2-1F78-4F7B-B594-385D9C5304C7};895;CHOH-1045;{9251277A-33EF-4568-B294-40D8A"
                        "25FD822};896;CHOH-1045;{E3F2EDED-B459-4AEE-A4C7-CA2E87E7ED9D};897;CHOH-1045;{4F6"
                        "8B14F-7A0C-421E-8BE9-9A2797D466D2};898;CHOH-1045;{BE50C783-E3E8-4B82-8D91-5D5099"
                        "F95948};899;CHOH-1045;{06F61E35-B195-48FE-AC2A-94DD28CC1029};900;CHOH-1045;{E979"
                        "5F51-B1C3-48F9-8A0D-919A506B89A1};901;CHOH-1045;{93F909DF-052A-4568-8AE3-67FC7D6"
                        "8375C};902;CHOH-1045;{E67142B4-EA08-4CA8-ADB5-B11805B9E6E6};903;CHOH-1045;{2DF82"
                        "506-DB46-45EB-B07E-F3B53273FC46};904;CHOH-1045;{8C5C62E1-FA9D-4271-8C46-A5F8A14E"
                        "E4A9};905;CHOH-1045;{57B0AB24-05FB-4466-A533-6027850ED62C};906;CHOH-1045;{AB0A3A"
                        "22-601B-4B4D-9503-D3E9FBD06E00};907;CHOH-1045;{8AD63D7C-B82F-423A-B5B8-A8A41D78B"
                        "05D};908;CHOH-1045;{81A38AC9-91B6-448F-BA00-66D06700A3BF};909;CHOH-1045;{B07BF84"
                        "1-E6C3-473C-A2CF-349223EECED6};910;CHOH-1045;{7C436B3C-917A-42E3-B1F0-C453EF15BF"
                        "78};911;CHOH-1045;{D7CB10DB-D827-4D08-AEA4-7B2C0150E0DB};913;CHOH-1045;{10B3B2D9"
                        "-6B3E-4DD7-ACB1-EB2DFE3CE2B6};914;CHOH-1045;{157BADCC-A7B9-4F2A-BC38-B3BBCFF2066"
                        "C};915;CHOH-1045;{00BFBD44-A135-493D-AFDB-2CB8DBB7363D};916;CHOH-1045;{D1E4D3A8-"
                        "11DD-48E5-BC38-231859A368EC};917;CHOH-1045;{2A92A705-D764-45EA-B464-068D0D669115"
                        "};918;CHOH-1045;{34ADF7AF-801A-4E69-8B14-6885067606BA};919;CHOH-1045;{E91E0803-9"
                        "088-4F6B-8A39-C7D7FC2A7CEC};920;CHOH-1045;{D3434992-AF03-4C28-B3E3-E10DB246DFC6}"
                        ";921;CHOH-1045;{132B0113-E31D-46B9-9D3C-A188127B17BD};922;CHOH-1045;{B1787BC9-BE"
                        "02-4666-B58D-F416E40FCC49};923;CHOH-1045;{D7EBC52F-100F-489C-A13E-F7C1F7FF97AA};"
                        "925;CHOH-1045;{4561A20D-BD65-4883-8865-D0FE39BE5785};926;CHOH-1045;{442FEFD9-255"
                        "D-44B5-8E88-2FCD3698FC2E};927;CHOH-1045;{98FEE5D3-CC0D-4952-BF7E-3B8E037C5701};9"
                        "28;CHOH-1045;{4F6D0A47-989F-4C80-99F1-39AE59436C16};929;CHOH-1045;{BF17F9B4-DE23"
                        "-4911-9EB6-DD5BC0F7C7A7};930;CHOH-1045;{8B500C3F-3F4E-47C3-A667-2EB0F134D355};93"
                        "1;CHOH-1045;{DC7E0831-59A4-4777-A037-FA3523728503};932;CHOH-1045;{8B9AD4EA-1031-"
                        "41BA-95A4-F1BB216FAA5E};933;CHOH-1045;{D78AEAAA-3290-48CA-9898-E595E647AC1E};934"
                        ";CHOH-1045;{20AB759B-94DB-4C8B-8332-5DF192E946EB};935;CHOH-1045;{F76BFD91-915A-4"
                        "895-A23C-8C2E5BA5D235};936;CHOH-1045;{16528833-7559-4E05-8A55-8AA9B5AC23CB};937;"
                        "CHOH-1045;{DDDBEC3D-AFC4-495D-96B8-B598E59BC4BB};938;CHOH-1045;{787E9AFB-E206-4F"
                        "A8-BF31-054D9A037481};939;CHOH-1045;{96E3F77E-A4CF-4BED-B886-F06E05FA6E73};940;C"
                        "HOH-1045;{ED0D8B9C-D97A-474F-926A-49ACA820CD28};941;CHOH-1045;{66071FD1-0425-414"
                        "4-B3AC-20818D885581};942;CHOH-1045;{8F78575C-8D22-4AEA-BC8F-CA0607AB7B8B};943;CH"
                        "OH-1045;{01F67C03-6C73-4071-B833-AB37DFA4EEBF};944;CHOH-1045;{5213107C-02F1-4312"
                        "-9133-10B91D54B410};945;CHOH-1045;{18BC6962-2C31-434B-9EA7-FFA8B19BDBFD};946;CHO"
                        "H-1045;{8A73D5EF-863B-4D16-AF55-048D42842172};947;CHOH-1045;{175EE491-3FB5-47CB-"
                        "A08D-A8DA82A9D7EE};948;CHOH-1045;{751FD698-7FB7-4C3D-980F-C850E572997C};949;CHOH"
                        "-1045;{9FA5537F-DF9F-4391-A4C4-9E8E1DE57C53};950;CHOH-1045;{DA8BC292-4FC5-4208-A"
                        "402-31603DF470D9};951;CHOH-1045;{61DB1722-3EB4-4073-AB18-7A56BB57F5A0};952;CHOH-"
                        "1045;{A47F6B71-66BD-449A-BA65-08299D826AEF};953;CHOH-1045;{A9A5754E-585F-460B-AE"
                        "6D-87F656CEC141};954;CHOH-1045;{8E3688EF-4782-4367-9F8A-DB559CAD2375};955;CHOH-1"
                        "045;{B3FFD8C6-CA5E-43B5-8C77-3758AF8AB17F};957;CHOH-1045;{FF4CD415-D305-4756-BBD"
                        "6-C4CD94164CC6};958;CHOH-1045;{C1B3CC8E-006F-47B2-B6A3-78F32D748FD4};959;CHOH-10"
                        "45;{DB42FD4D-F2EC-4AA1-88C1-6DDB60ABFC53};960;CHOH-1045;{CBABC76A-F3CC-49A7-A542"
                        "-AC20798C4F4A};961;CHOH-1045;{5971E5FB-C4CA-4942-A2A6-BB86F189B583};962;CHOH-104"
                        "5;{1843E3F5-0F30-457B-A97E-1AE208C6F1CD};963;CHOH-1045;{4612CB3E-86CD-45EB-AE08-"
                        "A858B2ECC72C};964;CHOH-1045;{D990DB0B-6017-4C9A-ACAA-C16431E21D4A};966;CHOH-1045"
                        ";{75090CF9-1885-4360-8541-215715531489};967;CHOH-1045;{7BE99780-2EED-4B33-A1D6-0"
                        "CF4A465B476};968;CHOH-1045;{387C45F6-0AC2-461B-8DFC-A8DE2978F48F};969;CHOH-1045;"
                        "{C346F405-36A9-49A5-9355-518EEFD5BCA7};970;CHOH-1045;{D1A3A7B6-D63F-4216-A4BB-9A"
                        "BD057DD573};971;CHOH-1045;{BE786D96-9A0F-48D0-9DFB-E47E397B5268};972;CHOH-1045;{"
                        "02E49129-3482-4F97-8C72-9EB11E7C6E5E};973;CHOH-1045;{6C8E800D-5AC0-43FA-9E31-AFE"
                        "48862D047};974;CHOH-1045;{1A53DAFD-6613-4B06-B095-B15FF9BB0597};975;CHOH-1045;{2"
                        "4AA7A02-DEEC-4251-BBEE-C198F6399138};976;CHOH-1045;{E66A67CC-CC85-41AE-8CEF-1702"
                        "2984D04F};977;CHOH-1045;{3E1A330B-4345-44FE-8138-84C90DA0F033};979;CHOH-1045;{53"
                        "51F318-00BA-464C-886E-3D4B3E0DBDE6};980;CHOH-1045;{4259097D-140D-4CA0-AC9D-74E58"
                        "8AC7DCD};981;CHOH-1045;{C606CD0A-1D33-4506-9874-813E4F2513AE};982;CHOH-1045;{232"
                        "CC5D7-4034-46D7-B520-EAD545588AF5};983;CHOH-1045;{49610D26-8D2A-43D2-ACF1-D756EF"
                        "905B6C};985;CHOH-1045;{AF0F2919-8D2A-4E33-9D16-4C3FE7A2E377};986;CHOH-1045;{F6B8"
                        "3FD8-025A-4320-93CB-C2F77BD24A78};987;CHOH-1045;{B612F3AE-1E0C-49EB-BD8C-ADC92BB"
                        "D6EF4};988;CHOH-1045;{67B85C1A-0F63-4FC4-B9F2-A628C1CBF449};989;CHOH-1045;{4144E"
                        "994-9010-4304-9460-279CC14F1876};990;CHOH-1045;{3E0AEFBF-6046-49B6-AE19-7112AE06"
                        "33AC};991;CHOH-1045;{7ACA9A64-A1BE-4AE3-A75F-DE1FA226D38C};992;CHOH-1045;{FB938C"
                        "84-F48D-47C9-8C70-DB580DCF43BC};993;CHOH-1045;{7AE59BEF-807F-4BB6-BDEF-38A11F76F"
                        "925};994;CHOH-1045;{193786C6-C42C-4804-A814-CEF839581343};995;CHOH-1045;{163C843"
                        "0-A513-4E0F-BD83-14AC6396E052};996;CHOH-1045;{7C39D8CF-B747-4888-A668-64BECA1640"
                        "13};997;CHOH-1045;{4DF8ADF3-68E7-4383-B2E1-F4E80202C230};998;CHOH-1045;{CD5014C8"
                        "-D693-497D-99CF-7C68B908D7A9};999;CHOH-1045;{F9CFD8C8-621D-470D-B298-E3B5ABD59D6"
                        "7};1000;CHOH-1045;{26E6D684-159E-4F0B-B083-BF47460970CF};1002;CHOH-1045;{0D91897"
                        "A-9EF2-4DE2-93F8-61ED96209218};1003;CHOH-1045;{9460877F-C20A-4F1F-A4DF-986A2741D"
                        "F63};1004;CHOH-1045;{25169879-B6B3-40C7-ADF9-259D78B9A56C};1005;CHOH-1045;{E3FB7"
                        "614-10CD-4793-91C6-A8B611B44D2B};1006;CHOH-1045;{A179FFC3-34D9-4AE3-99A0-8319C49"
                        "21214};1007;CHOH-1045;{C0AFEFCF-A5D2-49A1-BEC1-DEE1CF27FBC6};1008;CHOH-1045;{E92"
                        "CF4E8-86C0-45A1-A1FA-6038E99943A2};1009;CHOH-1045;{392BD440-6B12-4F5C-AC71-F436F"
                        "04EC943};1011;CHOH-1045;{DBE3F35F-89CD-4919-B6B4-B9BDA9F99ABA};1012;CHOH-1045;{0"
                        "9F32A5E-FE4D-4799-94F8-4B9B0D51E8B3};1013;CHOH-1045;{8204B161-ABB4-404A-86DC-DC0"
                        "AFB0A82B3};1014;CHOH-1045;{4CCF3ADD-6249-4776-AF08-777047C89325};1015;CHOH-1045;"
                        "20101217090932-56236863.1362915;1016;'';{BA154219-CDBC-4C56-86BB-24BCB9A4E195};1"
                        "017;CHOH-1045;{DEBDD660-D4C8-4195-9E6E-0C0696B028E4};1018;CHOH-1045;{A5352673-12"
                        "4B-4560-8784-FE7C96E2C7FB};1019;CHOH-1045;{17F8424D-3AEF-4044-9DE3-B4A91FEC826C}"
                        ";1020;CHOH-1045;{2089BE6E-73B6-4170-B74B-EBF2784465DF};1021;CHOH-1045;{5E6339D1-"
                        "1D5A-4788-9352-681E7F073FDA};1022;CHOH-1045;{F4DFF3C5-DA59-4570-8768-0E4C72C37F1"
                        "B};1024;CHOH-1045;{BED687F7-947F-4B89-BA55-3EC74A917469};1025;CHOH-1045;{0F0A177"
                        "3-72C0-4246-917B-DA067F987D57};1026;CHOH-1045;{A09AFCB4-F0E1-4C39-A245-C33FA258B"
                        "CE5};1027;CHOH-1045;{31F20D28-DD78-4DCD-B96B-7049F775F4EC};1028;CHOH-1045;{DF87F"
                        "DE0-42B6-434C-81C8-B5653A1EE941};1029;CHOH-1045;{E2F3A20D-2A34-45ED-B0F3-7B97394"
                        "D9EA3};1030;CHOH-1045;{8670E9EA-64C1-4C80-A100-0235E3C25107};1031;CHOH-1045;2010"
                        "1217091531-949556648.731232;1032;'';{726BB18F-5B59-48A7-985C-E1102468DDE4};1033;"
                        "CHOH-1045;{2BD0B62C-59D3-4C2C-90B6-5E7145D6B86A};1034;CHOH-1045;{56556814-C6DE-4"
                        "128-BD93-90C3D85FFD7F};1035;CHOH-1045;{E0463024-D5F9-430F-9D90-E516984CAA2B};103"
                        "6;CHOH-1045;{AC3C6149-AA7F-4932-9615-6145C0950926};1037;CHOH-1045;{C9661088-A438"
                        "-41A3-B083-B77C37918F45};1038;CHOH-1045;{EF4FC39D-7A3D-4429-9E95-75020403F8A7};1"
                        "039;CHOH-1045;{1B0B48DA-6E89-4FC5-A6D5-BFE473EDE4C9};1040;CHOH-1045;{64B72FDC-E0"
                        "AF-4F2F-B504-F4E6AA678C11};1041;CHOH-1045;{6AA00706-EB55-4AD5-A2DE-AB72663E1D58}"
                        ";1042;CHOH-1045;{EFC0C5C1-F0CF-4A57-9118-E1F7E0D35000};1043;CHOH-1045;{EE2FE5DB-"
                        "EEA3-4B5D-A19A-11081B3C8D4F};1044;CHOH-1045;{92DD4EBC-27C8-466B-B817-EDDD0AD5649"
                        "C};1045;CHOH-1045;{4623FBC6-28F4-40DD-BB0D-0153F0070716};1046;CHOH-1045;{172D407"
                        "8-61C9-445B-9B07-8BBEAEBF4BDF};1047;CHOH-1045;{176E1512-3ED5-4F7D-A370-6C65D08A6"
                        "1D6};1048;CHOH-1045;{5F0BF1F5-D836-4421-A8E1-219958F7A987};1049;CHOH-1045;{F260F"
                        "3D7-2382-4EE7-8280-7274848355C6};1050;CHOH-1045;{282AB080-CE14-46A7-9C23-E613480"
                        "17358};1051;CHOH-1045;{6B177557-44DC-409D-A695-950BE8BF5BE2};1052;CHOH-1045;{A2A"
                        "68078-E5DB-4590-89D2-F51B6CD98CC0};1053;CHOH-1045;20101216092933-705547511.57760"
                        "6;1054;'';{E986F8B5-D4F5-4405-847A-F644B5DF04DE};1055;CHOH-1045;{B6EA8057-5CA0-4"
                        "424-8D58-F774DF6D05C2};1056;CHOH-1045;{5175A744-BA30-42CB-98E6-B4E2E6F0178F};105"
                        "7;CHOH-1045;{30F2DF5D-8CEC-4581-ADE5-3EBD6EBB45BE};1058;CHOH-1045;{5DC4247A-8E39"
                        "-417C-8A1B-F5CBF672D86E};1060;CHOH-1045;{33952F05-F1A1-4D12-B216-7BF3558B0D43};1"
                        "061;CHOH-1045;{7A65BCDB-35FC-4431-AEC7-1AB07F24ACDD};1062;CHOH-1045;{E9BE539C-30"
                        "24-4B2C-9D76-EB13AA81A4B0};1063;CHOH-1045;{EBBE517B-4107-45FB-B908-948ECDEAF145}"
                        ";1064;CHOH-1045;{3E311D14-B682-4E53-A795-4AB0CE0D3964};1065;CHOH-1045;{7A72DCBD-"
                        "2659-4BE5-8921-8E0B6D8BE483};1066;CHOH-1045;{2CCA31D9-BC47-42E5-8A7A-A79A4B3F511"
                        "8};1067;CHOH-1045;{DC35A53E-80E5-4DC6-A60B-975BBB0CB98B};1068;CHOH-1045;{7E351B0"
                        "B-4C5B-4344-A4BF-5ABDC2A5DE08};1069;CHOH-1045;{1C6D314E-5A57-4372-B617-382FDED74"
                        "A13};1070;CHOH-1045;{F432CEF2-CA41-44E8-8CFA-E6FD6E13202B};1071;CHOH-1045;{DD2D7"
                        "A91-5C8E-4710-8A6F-EB852DC0694A};1072;CHOH-1045;{FE116B59-7D87-4712-BD0A-CEE974C"
                        "8953C};1073;CHOH-1045;{BA52CC4F-B9F7-4FBC-9C65-3CF3DC736AC4};1074;CHOH-1045;{279"
                        "8B4A9-57E3-4981-A5D3-45CC012B727D};1075;CHOH-1045;{5C9C315F-F269-4DF5-AB38-E6D6E"
                        "A18D269};1076;CHOH-1045;{9A341E3C-AF4E-4109-BCD8-FD36438500EF};1077;CHOH-1045;{F"
                        "17EB81F-A24F-4CC1-A6F5-A9E391341244};1078;CHOH-1045;{188B3134-EF12-4D4C-B1B7-9E0"
                        "CBA3FA66F};1079;CHOH-1045;{8E0F638A-F50A-4E24-BBCC-32AB361471FD};1080;CHOH-1045;"
                        "{37173424-C858-481F-B042-62A22519D369};1082;CHOH-1045;{53C710C5-16E3-46A8-944A-7"
                        "6B625C2C419};1084;CHOH-1045;{6D2B00F0-E435-462D-8337-A355465DF877};1085;CHOH-104"
                        "5;{3F2A8DB1-B8A8-434F-AF54-AA8A1B210295};1086;CHOH-1045;{C51DBD28-8114-41A6-A281"
                        "-BCC71F032D38};1087;CHOH-1045;{F490103F-6CF1-49A6-BFEB-C0C0A3ACDCDC};1088;CHOH-1"
                        "045;{D0BAA3C5-7C13-42FE-9ADC-F6732FA3E4B5};1089;CHOH-1045;{839D34EE-FECF-4794-94"
                        "24-A1270CDE598C};1091;CHOH-1045;{4153D66B-2FA6-49B2-912C-DA355E0AF04D};1092;CHOH"
                        "-1045;{C476A064-A6F4-4A4F-85C8-3D7E4B501D60};1093;CHOH-1045;{9928D4A8-BA02-423B-"
                        "94D5-F3861BC48405};1100;CHOH-1045;{D7A3E743-3619-45BF-A5AB-EABDDE046805};1124;NA"
                        "CE-0491;{961939E6-8B25-4EF2-A639-3C8CB58584A5};1128;CHOH-1055;{570AF8B1-DD72-431"
                        "D-B079-F8E67EBE5F24};1129;CHOH-1055;{A4880FAE-0001-4B97-B677-9AEB5B820CFA};1179;"
                        "CHOH-1055;{449A3D1E-5AC8-432A-828A-8F6E8FCA7700};1180;CHOH-1055;{029DC260-E227-4"
                        "FB0-9F4B-94F6326CE92F};1182;CHOH-1055;{521083E9-58D2-4403-AD4A-E14E4C32B6DE};118"
                        "3;CHOH-1055;{8F636C13-09DB-44CA-85C0-0A67F832A28A};1186;CHOH-1055;{8BA10A01-D732"
                        "-48D9-833E-FC266850BBBC};1187;CHOH-1055;{BE1A3673-925A-4916-A141-C87BC63695D5};1"
                        "188;CHOH-1055;{6929D1C0-69B9-4F67-8914-F800C80FE5AB};1189;CHOH-1055;{601992B3-3F"
                        "8F-4724-9025-6705D8FE15E6};1191;CHOH-1055;{F66F5819-5D10-4590-B29C-E8A2E97AC9D7}"
                        ";1193;CHOH-1055;{78EE21D5-8C82-48B1-BE7C-F93203EE0F29};1215;CHOH-1063;{015FC0B3-"
                        "55DF-42D8-A154-4C1D7F17F517};1216;CHOH-1063;{0EE4DFD0-07BF-4D3D-AA31-083DBDC59FD"
                        "2};1217;CHOH-1063;{C1F35917-9939-45F7-BE66-B0301618E56F};1218;CHOH-1063;{7610DC2"
                        "4-DABE-428C-8AA9-3DD7800FBD1F};1219;CHOH-1063;{A4A4021E-A16F-48C7-9748-57F99D25F"
                        "5F5};1221;CHOH-1063;{192DA6F1-6FA5-4B69-8FDC-0F49D76829A4};1222;CHOH-1063;{6C89F"
                        "1C3-FC45-4EF3-8F9D-AB1DEE3C6D55};1223;CHOH-1063;{EB91463D-E65D-45EB-ADE3-B1C1C9D"
                        "1510E};1253;GWMP-0208;{593B8369-A1AF-4EDC-8D0D-0551BA2EA94B};1254;GWMP-0208;{82E"
                        "1B176-C539-412A-9EAE-7D20AE01A14E};1255;GWMP-0208;{0BAF0BC6-4A73-4FAC-9589-B9724"
                        "B9BB919};1256;GWMP-0208;{AA1A72EA-6231-415E-9B88-F34C94BF94EA};1258;GWMP-0208;{1"
                        "8566D5A-D7A6-48F8-B11C-D2B348140E8D};1260;GWMP-0208;{963AB190-D935-4FCA-BDCD-394"
                        "15402CB40};1273;NACE-0245;20101217091550-364018678.665161;1288;'';{99649B9D-CAE8"
                        "-45AD-B399-4F40FE10B975};1291;CHOH-0776;{5BEB250D-D8DA-4BC3-8161-3842A5604C7F};1"
                        "293;CHOH-0776;{B70444AA-498C-478B-9DFC-B1ECAF25FAAD};1294;CHOH-0776;{C4A5DE6D-9A"
                        "E3-4F02-8B4B-652DD5180B79};1300;CHOH-0776;{D3C5DE05-ED48-495E-97B2-C0A3515511D6}"
                        ";1301;CHOH-0776;{214585F0-321F-4228-9F9E-9153B4AA5B90};1302;CHOH-0776;{648A6AA8-"
                        "5799-4AA5-ABE2-4D7C2453F8BF};1303;CHOH-0776;{534C24D6-98F4-459D-A6E8-1C12014FAB9"
                        "7};1304;CHOH-0776;{4F573DE8-36F0-4AA8-929B-ABE60B9DFDB1};1305;CHOH-0776;20101217"
                        "091844-767111659.049988;1307;'';{B45646CF-6B24-47F0-A6FA-D4B9444C324C};1311;CHOH"
                        "-0776;{D079092E-3B85-4DE1-B033-9FCFFDB5484E};1312;CHOH-0776;20101217091926-53504"
                        "526.6151428;1361;'';{606C6B87-E89B-4171-A579-839069D6847D};1364;CHOH-0788;{E593D"
                        "5E7-11DF-4485-B840-7D8D9DE49C08};1399;CHOH-1338;{22C7412B-8B7D-466F-8CAD-2B450C1"
                        "3AD8F};1400;CHOH-1338;{30911FF8-2CF3-4E87-B522-A695117ADD44};1401;CHOH-1338;{746"
                        "38E7D-DA0E-4EB5-8FD4-68472964E864};1402;CHOH-1338;{8C4BC8D9-D0BF-4349-B9CA-C3BC3"
                        "EA03685};1404;NACE-0087;20101217092005-592458248.138428;1416;'';{F2B602BC-CC86-4"
                        "E78-805B-C0D11408DB83};1427;NACE-0087;{16787EF5-7609-43B3-99BD-8CD6D8DB05AE};145"
                        "0;NACE-0087;{D0A656EF-8BF0-4E4F-9FBF-307698022EC9};1455;CATO-0365;{6DB088CE-5E2B"
                        "-46A2-9657-9FE08973063A};1462;CATO-0365;{56103498-C9DE-4170-BD0E-10F395859D2C};1"
                        "475;CATO-0365;{A3DDEFEC-4269-4FF7-B13F-3A78801D2916};1477;CATO-0365;{6A5E2513-FD"
                        "6E-4823-893D-2F48EBC9DC6C};1478;CATO-0365;{8B80D041-C0D0-4311-80DA-42DFD81E7DDC}"
                        ";1479;CATO-0365;{190B5B80-3382-479C-9449-0A2325840EC1};1480;CATO-0365;{0B016440-"
                        "14AB-4D61-81A5-85020DC81FBE};1481;CATO-0365;{3163702C-B43C-4070-94D4-B8005497DD9"
                        "9};1482;CATO-0365;{D6115F26-3988-4C99-AC43-65DFC339AFF3};1483;CATO-0365;{86D0689"
                        "2-BB06-4692-886A-41A22BB07F63};1484;CATO-0365;{5B27FEBB-36F2-4E69-AEA8-754B8B95C"
                        "F76};1485;CATO-0365;{895451D1-2CAE-4BEB-AD2E-05E3777266D1};1486;CATO-0365;{44923"
                        "969-CB49-418B-A727-43FDDFA3E6C5};1508;NACE-0337;{E2BA70CA-1E41-42E5-BBE8-E93E219"
                        "D1BD2};1509;NACE-0337;{BCFBB747-29DF-4584-ABFA-0995A58AB7A2};1514;NACE-0623;{DBB"
                        "91B05-1FB3-4B3A-A204-E2BE66BFB65C};1519;NACE-0623;{69CC35FF-CC6D-40D6-BFFA-36CE9"
                        "7A556C9};1522;GWMP-0008;{2D7EA8BD-DA39-4F77-A666-5AF5AC31C400};1523;GWMP-0008;{8"
                        "8067801-A1DC-4C60-85F3-4F8A3C706EAF};1524;GWMP-0008;{BCC7C6C6-07AA-4D40-B419-DB2"
                        "CF4ACCA84};1525;GWMP-0008;{86CD2EF1-A1F1-4712-946D-18EDD7D1CA41};1526;GWMP-0008;"
                        "{ADEEDD83-3B0A-423E-80DA-4BC8C11A6DE5};1527;GWMP-0008;{A8BC5993-8B53-49C0-AE1E-6"
                        "6ED5DEAAD99};1528;GWMP-0008;{FBD75FF3-7DD1-42AB-9DAA-B575C8937600};1529;GWMP-000"
                        "8;{6B4E354B-FC34-49F0-9D7B-B7E5581D6ADF};1530;GWMP-0008;{F19C3A37-9CD8-4F81-9D42"
                        "-DB191D3528E7};1531;GWMP-0008;{F2FDC458-8FDB-48F0-811F-FD6D22738BE3};1532;GWMP-0"
                        "008;{77FD22F8-4F08-4B44-BDB1-7E12DFD5A1D4};1533;GWMP-0008;{41F076A5-EFC6-4595-AD"
                        "BA-0B8F79C8B41D};1534;GWMP-0008;{66D24A93-C37F-4D9C-859C-4EC55C3C130D};1535;GWMP"
                        "-0008;{66D19566-0D99-43C6-B7D2-E2B7A4AF9814};1536;GWMP-0008;{800A3567-F3C1-42A1-"
                        "8937-2FA35812BFDD};1537;GWMP-0008;{F017B967-9FE8-44C7-813F-0F773E7BA1AE};1538;GW"
                        "MP-0008;{0DF249FF-A3BF-496D-850F-5220E140D9E1};1539;GWMP-0008;{12DC8542-53B2-49A"
                        "4-BEF0-EFBE2FEDC6A8};1540;GWMP-0008;{FB8B6B21-87DA-4CEA-ACB1-9E7BB9144EB3};1541;"
                        "GWMP-0008;{503D0124-9CCB-40F4-ABF2-888F00C2047F};1542;GWMP-0008;{2934F2E2-0F22-4"
                        "A48-9E6F-C271387DBDEF};1543;GWMP-0008;{4FDA2204-4DC4-42B2-B6A3-BD76F30EFC55};154"
                        "4;GWMP-0008;{26B541D2-AF59-4360-A2A6-42F8B365ABE3};1545;GWMP-0008;{91908879-D8C8"
                        "-4FAF-A904-29F6E978DDA7};1546;GWMP-0008;{EC723889-74DE-4ADA-990D-B3D284769922};1"
                        "547;GWMP-0008;{487EC2A6-E9AF-44C4-BDAE-0786BBBE752B};1548;GWMP-0008;{A2D7F91C-59"
                        "E6-4BCB-8385-EB729F32D25D};1549;GWMP-0008;{CD3E4312-626F-44C6-B4D5-E028108BF3CF}"
                        ";1550;GWMP-0008;{F948F902-FF0A-44D9-871A-3ECCFCA9AE4B};1551;GWMP-0008;{536294ED-"
                        "CED7-41E2-A117-423403E45EB4};1554;GWMP-0008;{450F4E61-4E64-48F6-BBC2-B66C82312D7"
                        "9};1555;GWMP-0008;{0E9C999B-1F7E-414F-A684-17A483C27DC5};1557;GWMP-0008;{15D3132"
                        "8-FAC4-4D0C-A1E1-2BCB5CE66E47};1560;GWMP-0008;{436B9C83-C896-4EEA-B754-F74F0C0EF"
                        "855};1562;NACE-0081;{A5AC973F-0128-47ED-88CF-F73BD7F948B5};1592;NACE-0081;{EAA7C"
                        "462-A158-4D62-A0EF-174374CF0A2B};1593;NACE-0081;{1FF154C1-999D-49A1-841C-B724258"
                        "CC1DD};1596;NACE-0081;{2F4F4617-9C89-4A53-ADF0-8709745CEC99};1598;NACE-0081;2010"
                        "1216134956-579518616.199493;1600;'';{F48BC7E7-BAA8-4DBD-8A3C-0A2396619673};1602;"
                        "HAFE-0215;{68921A17-0598-4093-8134-80DBFB15E8E2};1604;HAFE-0215;{32E78BCC-5A2D-4"
                        "EB9-B61D-9F122F1ECA38};1605;HAFE-0215;{E6DD4483-C544-4071-8FAD-8276A79B223A};160"
                        "8;HAFE-0215;{A866C3FE-4EA5-49DC-A62F-8E8391A8A1C7};1609;HAFE-0215;{EBE5AA73-66E0"
                        "-4F9B-93CA-F1B535D89528};1611;HAFE-0215;{04E180A3-D3C0-4D19-8FCB-07E29A168922};1"
                        "741;HAFE-0215;{77DB5B66-1B52-49E8-8E27-D9D8895755CA};1758;HAFE-0215;{F243C1F0-D8"
                        "3C-4BAE-AECC-6F4BCE8E4281};1762;HAFE-0215;{32A0D547-1E5D-4199-8FC7-2410095D5762}"
                        ";1763;HAFE-0215;{2E9D9339-30BF-4169-8528-BEB5A5A40F36};1764;HAFE-0215;{38B89FD5-"
                        "4956-4E53-BA83-69D4BE34930B};1767;HAFE-0215;{9D15EBD3-72D1-4304-9622-896EDE71672"
                        "4};1770;HAFE-0215;20101217092123-468700110.912323;1782;'';{74D2B909-B187-498B-84"
                        "8F-0D5BA7FB3A1E};1824;CATO-0294;{0112333B-D8BB-4B82-8A85-993F6EEA68E6};1836;NACE"
                        "-0399;{0F4DF3BF-8521-4A14-9A0D-641C57E315B8};1849;NACE-0399;{28AEB677-4014-4C29-"
                        "9C82-72CE2B95A68E};1861;HAFE-0074;{D018104F-EFDF-44F9-969A-6F267640E988};1896;CH"
                        "OH-0983;{909CC740-80A8-4C45-96CA-B4C34D113F19};1897;CHOH-0983;{1BA1F7ED-89BD-400"
                        "4-95FF-05142FD7700F};1899;CATO-0365;{26CFC989-0739-4010-9AD2-B0929473B633};1900;"
                        "CATO-0365;{96F1AC5E-C759-4455-92DE-FFD29D3F560F};1902;NACE-0399;{6CCF52FB-14D2-4"
                        "351-8715-27986C60A299};1904;NACE-0399;{326B6DD8-C5FA-4365-AFF9-57170882A130};190"
                        "5;NACE-0399;{DE60EC45-D555-4982-9DE3-971859449862};1907;NACE-0399;{C7137D07-2590"
                        "-4FDD-87D2-A2FD4B443608};1908;NACE-0399;{9100655E-03F6-458A-9751-A9EFCBB9E42F};1"
                        "910;NACE-0399;{6F9C717A-E087-4B08-B4FF-C59FE2EC13F5};1911;NACE-0399;201012170921"
                        "46-298165440.559387;1914;'';{7287F7BF-A539-4542-A927-195CC8D559DF};1920;CHOH-132"
                        "8;{9DBF8D58-AD95-4A5B-9308-224277E26910};1976;GWMP-0062;{B27F257A-86D3-41E0-9008"
                        "-E743C4783967};1978;GWMP-0062;{A967BE6F-F3FA-404E-A837-B132CE00D96C};1991;GWMP-0"
                        "094;{E0C7ABC8-5A1C-42BC-BECE-3986BD4C331E};1992;GWMP-0094;{73BC9EE4-D17F-41A4-8F"
                        "BD-7B2ECF9E5811};1994;GWMP-0094;{81507F4C-D83A-4894-B44B-9A6FADF1CAF5};1998;GWMP"
                        "-0094;{C93FD002-B6A5-4C7A-91F9-A06127919508};2000;GWMP-0094;{B2D8B76F-DDC6-4CAE-"
                        "9C6E-4220CD5FCF43};2002;GWMP-0094;{4A28A89D-B0E2-4260-9EB0-C579FA5E227D};2007;GW"
                        "MP-0094;{085BFB74-819F-4304-9C9B-F872978A6538};2009;GWMP-0094;{587EE971-25B7-4EB"
                        "6-9087-AD051785BFB1};2010;GWMP-0094;{6AC1D6D9-B636-4FC7-9C76-0799D44E4E05};2011;"
                        "GWMP-0094;{81DF55D8-B539-4EA3-AD8B-0F4E04217B8D};2012;GWMP-0094;{71486BFD-6207-4"
                        "E53-9ACA-31AD72CAEEFD};2013;GWMP-0094;{21DE83C3-F5B7-4440-9E58-B22CBBE621AD};201"
                        "4;GWMP-0094;{68836BE9-8AE9-4986-849A-C16ED7C8217A};2016;GWMP-0094;{A304B7F0-E2A3"
                        "-4416-B6E4-C983A81546DD};2031;CHOH-1328;{7738C7E4-F84E-440E-8DF6-86D198BB997E};2"
                        "076;PRWI-0277;{33182A9A-0C83-4E14-8E7E-38C6266B00D3};2077;PRWI-0277;{57043EF9-84"
                        "68-4CAE-AC7B-C8B2E43F2083};2079;PRWI-0277;{5A8346A1-BE96-4578-BF6D-A1B38838F4A3}"
                        ";2080;PRWI-0277;{7659F777-442E-4679-840D-295522D3B974};2081;PRWI-0277;{7D0E01C9-"
                        "4D62-4A0A-8156-323959FF6504};2083;PRWI-0277;{0820FAC5-1C0C-437D-95C0-4960D2CC6C0"
                        "A};2085;PRWI-0277;{115226DE-5987-4C8B-B454-5B422C86B1F1};2087;PRWI-0277;{1C8F310"
                        "3-CDBB-40C4-88E8-AA922B100536};2115;PRWI-0621;{42D01751-AC51-484B-A7D9-697EAA313"
                        "9E8};2126;PRWI-0175;{AB06481E-C9FC-42DA-806F-D39DC504E049};2128;PRWI-0175;201012"
                        "17092210-622696697.711945;2136;'';{F0A40EB1-1073-4A9A-ADDE-7CC6B6A99B2A};2145;PR"
                        "WI-0173;{C18C9306-3C74-4365-A4A6-07F01E93C59B};2146;PRWI-0173;20101217092232-647"
                        "821187.973022;2148;'';20101217092252-263792932.033539;2150;'';{40FB02DE-F027-4F5"
                        "F-AC24-87F3ED1BCC4B};2194;PRWI-0491;{4972E813-867B-45CC-97F1-4215FB0DC4E7};2197;"
                        "PRWI-0491;{D0D5CE80-9E26-4B68-B680-9EE03AEEFEAC};2199;PRWI-0491;{35CD7FEC-66BC-4"
                        "017-B9F3-943FDA48A685};2201;PRWI-0491;{B228D2BE-E523-4E59-B0F5-4D73B1090D3C};220"
                        "4;PRWI-0491;{15C672FA-9069-4152-9935-2800639A4404};2205;PRWI-0491;{2E045E34-C250"
                        "-47B8-BF26-0730000DF6D1};2208;PRWI-0491;{4D79F0E8-352E-4655-A48A-F00A3E11EBD9};2"
                        "209;PRWI-0491;20101217092314-279342055.32074;2220;'';20101217092336-829801619.05"
                        "2887;2224;'';20101217092418-824602127.075195;2232;'';{CBB83D3F-5A00-4564-8D34-7F"
                        "2CD497C258};2239;PRWI-0233;{57DD3833-428B-44AD-801E-30EDD1A23875};2280;PRWI-0796"
                        ";{2DA868C4-FE9E-4330-AA1F-11954400EA28};2282;PRWI-0796;{CCE665A4-F8F5-44E4-A6F8-"
                        "972A56D808EE};2340;PRWI-0282;{1B896D29-4263-413D-A5C8-E47C673A18B3};2346;PRWI-02"
                        "82;{5F3C1207-2C3B-47A9-A69C-849E4E74A286};2347;PRWI-0282;{C443AC99-2E3D-417E-97F"
                        "6-093B33945BA4};2368;PRWI-0728;{50E90F7D-6F54-4AD4-9D9D-22A036CC61A8};2374;PRWI-"
                        "0728;{0228BD9A-E23C-42B9-B554-0AEC16EBD7F1};2375;PRWI-0728;{5F0B5C45-AF51-4CAE-8"
                        "FB9-B4006ABCE06C};2411;PRWI-0199;{1D1AA1F2-BC1C-4187-A08C-BF9A9704BF44};2412;PRW"
                        "I-0199;{1945E116-E42F-431D-B689-09078104D2EC};2413;PRWI-0199;{C98785C4-781B-430D"
                        "-B148-073EE642D0B4};2415;PRWI-0199;{80CD8146-CC69-46EF-8E98-88929BFCE856};2416;P"
                        "RWI-0199;{005C5E09-0853-4768-9438-665D68B5E805};2417;PRWI-0199;{7F8D83BF-EFD1-4B"
                        "04-819E-AA78E3735BC9};2420;PRWI-0199;{EEF833E1-2088-4E7F-985E-F2DA3FBC938C};2446"
                        ";PRWI-0796;{5B8BE00C-2D09-444A-ACE1-D40AFB77EF0B};2494;PRWI-0181;{6D7AF34C-6B5F-"
                        "40B5-BA63-3FD795C395C6};2502;PRWI-0463;{F6F82047-E5D5-4CA1-B5CF-7C47680B431B};25"
                        "04;PRWI-0463;{22A5305C-D7B7-482A-9426-4A92150FC0B4};2505;PRWI-0463;{54AAB3E6-1DA"
                        "5-437A-BBDA-982E421FE88F};2506;PRWI-0463;{F9296A06-E6F4-456B-875A-A25B36D0E4BD};"
                        "2509;PRWI-0463;20101217092443-589163005.35202;2570;'';{611806EC-0BFD-4C08-9B90-E"
                        "96C2DF19825};2571;PRWI-0508;{D608A903-B444-4BF8-ADC0-A33590350D01};2601;PRWI-072"
                        "8;{C4B5AE52-0DC9-4728-90DD-D0F3E2606BFA};2604;PRWI-0728;20101217092506-986093163"
                        ".490295;2610;'';{13590D69-9AA0-4DE5-8C11-B1D65CA2ED2C};2639;PRWI-0751;2010121709"
                        "2526-910964310.16922;2670;'';20101217092544-226866006.851196;2683;'';20101217092"
                        "605-695115506.649017;2710;'';{F4E0EDB9-8CFA-42C5-8854-9B803EF2CCD2};2735;PRWI-02"
                        "23;{76531CBE-1634-449C-9409-3B2B8DDBFE6A};2739;PRWI-0223;{CE67810D-0ADF-41FE-83F"
                        "B-BD7439BCB603};2740;PRWI-0223;{3355245F-8FDA-457E-867A-99D396EDFD6D};2741;PRWI-"
                        "0223;{25C9BA45-59DE-4534-BC44-4AD07C231D7B};2742;PRWI-0223;{1614AEB7-C2EB-4420-B"
                        "9C1-52F59CB0F0D8};2743;PRWI-0223;{F824D09C-71CB-49F5-B376-5176FCC0C72E};2744;PRW"
                        "I-0223;{B69D463E-ADDC-4FC0-A3E4-4619BCF09843};2749;PRWI-0223;{B902FBB3-BEFC-498A"
                        "-A270-1E3363DDC84A};2752;PRWI-0223;{25BA8D2D-F3CD-4416-AB8F-484585A4EA36};2753;P"
                        "RWI-0223;{DD61111C-B5A6-4A2A-A31F-772A16AB7C4F};2754;PRWI-0223;{F24468EE-8DA1-47"
                        "BB-9DE5-B45861791C1C};2755;PRWI-0223;{39275F20-285B-4795-8A36-CB69C11C740C};2756"
                        ";PRWI-0223;{FEF472C0-7110-4F62-BD4F-84FE6F3D222D};2757;PRWI-0223;{639CDA57-AC36-"
                        "4C65-9B5A-CA9C1199FE54};2759;PRWI-0223;{999578F1-9D9F-48FB-957E-01554B370C62};27"
                        "60;PRWI-0223;{46846CC9-E9CF-4D44-B4A1-7EFC417518EB};2803;PRWI-0338;{A40E5000-2E6"
                        "4-4046-80C6-9281BFA02566};2806;PRWI-0338;{553462BD-C8D3-40F6-8A98-26EED31D8AEC};"
                        "2808;PRWI-0338;{648AF137-47DF-4960-9894-812E58581E61};2809;PRWI-0338;{C71E282B-3"
                        "947-4D4A-A65A-BBFFE4FDABFD};2810;PRWI-0338;{BE6E070C-3077-46A2-AAAD-208ABD0A9134"
                        "};2811;PRWI-0338;{3FD5937F-F2B0-4098-B8EC-23AFDBD2D3FA};2812;PRWI-0338;{591A44E5"
                        "-6014-4302-B996-651BBA804199};2813;PRWI-0338;{61778110-C49B-4215-B04E-6DC36E17E3"
                        "08};2835;PRWI-0145;{5C1A73EB-4DC4-42F7-8CB1-F61866E862A1};2842;PRWI-0145;{0C9A79"
                        "E2-FCFA-47C5-A23D-199C9969CA46};2843;PRWI-0145;{5158DEEA-033F-471B-8FCC-2E9B8EA7"
                        "0AD9};2875;PRWI-0085;{4B02E756-4F0D-4B02-A394-DFD61C8A4ED0};2877;PRWI-0085;{2C5B"
                        "FC45-3294-4821-AF01-23D4535AC3CC};2878;PRWI-0085;{029F7646-2D0A-4DEB-8E7B-80D8E2"
                        "DF9E29};2901;PRWI-0062;{80799BF3-5EAA-4FF4-9F3E-B607A2D2AF56};2903;PRWI-0062;{2B"
                        "E6CE4C-5BE2-4974-8231-93E19CCC8CA6};2904;PRWI-0062;{93DB3519-D20F-4996-86AE-A1C9"
                        "3067B20A};2905;PRWI-0062;20101217092623-980003237.724304;2906;'';{9AB058C1-9377-"
                        "4094-80B9-429D967C2220};2908;PRWI-0062;{185C6888-2F8C-45F1-978C-80D084A7B73B};29"
                        "09;PRWI-0062;{685D96FE-3F05-4987-A3AC-9B247B89BEDD};2910;PRWI-0062;{D36AE4D4-C50"
                        "6-4994-A993-D3B8B8E18FBC};2912;PRWI-0062;{876C1806-A54E-40B6-B850-9105CD4A7685};"
                        "2913;PRWI-0062;{F74469ED-6179-4B3E-ADEC-A305DE938A8E};2914;PRWI-0062;{CB2B8221-3"
                        "311-4ED2-A63E-36BF3C9CFF2F};2917;PRWI-0062;{1875481A-C3D5-425C-A8F0-CDA391913CA3"
                        "};2918;PRWI-0062;{92D68D87-34D4-44E2-929A-439142BD4E41};2955;PRWI-0075;{FECEB1B4"
                        "-877A-4399-8760-57CC9527E87A};2956;PRWI-0075;{DEDF9798-8B68-4E6C-B377-4FD61AAA92"
                        "32};2957;PRWI-0075;{04758A2B-D17F-43A5-9B27-FB2B3FC4F1DE};2959;PRWI-0075;{A984DB"
                        "94-9D1B-436E-9F5B-69B351933868};2994;PRWI-0093;{6034CF03-E1BF-453C-982D-A322AE4D"
                        "E076};2996;PRWI-0093;{48BF4E8B-D1BE-4A82-8975-E9D2545AFDA7};2999;PRWI-0093;{02CE"
                        "9859-3484-4B11-A098-042CD918B0C8};3093;PRWI-0238;{0DAE6B73-8186-436E-B83F-09C7AE"
                        "74D9DE};3108;PRWI-0238;{62C4413B-8D58-4276-AC02-20E889E146CF};3111;PRWI-0238;{5E"
                        "78C44E-675D-44FC-ABE5-8B8E324A4AC1};3114;PRWI-0238;{BC609F3F-3A20-4314-875F-F16B"
                        "3DC7BCA2};3131;PRWI-0051;{38C46633-76C7-49F8-ADB8-08B606B9D9B9};3133;PRWI-0051;{"
                        "B8739678-BD20-44FC-8CC1-41EBC04478D5};3134;PRWI-0051;{136235BC-BF57-47A6-92F1-EC"
                        "E312C57981};3142;PRWI-0051;{4B237531-15F5-413E-9FC3-43BDCCC17C6E};3176;PRWI-0080"
                        ";{499CC64B-0338-40F6-988A-BBB06B59F9E0};3180;PRWI-0080;{56B64C47-557D-411F-AF52-"
                        "3F2CFEE5F0FC};3200;PRWI-0398;{70BC2872-1A04-48D5-A283-65955BB4A977};3226;PRWI-03"
                        "98;{7435118E-BD19-42AC-895F-2360E3C98C4D};3256;PRWI-0398;{C957968A-DA6A-4785-903"
                        "1-F121B7722F92};3274;PRWI-0333;{0FBFC04B-923E-41FF-B3F1-16E6791CCD63};3280;PRWI-"
                        "0333;{613D2044-3439-4E43-A8C5-ECB07560097D};3281;PRWI-0333;{3EF0EDCD-CAC2-42D4-9"
                        "61A-735BC6CD8491};3283;PRWI-0333;{40D67A2B-3387-47C6-BB3C-6F166D08840D};3286;PRW"
                        "I-0333;{8130921D-C93F-4886-8107-A369C5716ADB};3289;PRWI-0333;{1644A632-DFDD-403E"
                        "-BFBB-B15CE9C64466};3326;NACE-0174;{C7681E7E-A6F1-4EBC-9019-30F0009D7BAF};3327;N"
                        "ACE-0174;{9FB29A1C-3AD8-4F48-9AB4-C1E3A27242D9};3330;NACE-0174;{B3F30786-E9C9-4D"
                        "B6-B592-57C453E72A94};3334;NACE-0174;{DAC7AD41-75B1-4A73-9D93-4BD0C1BE69C6};3338"
                        ";NACE-0174;{0AA4A4A1-5CC2-4C14-A747-691BD0DE2FE9};3340;NACE-0174;20101217092717-"
                        "243931353.092194;3342;'';{7B84841E-EE44-4B66-9514-EC033E4A5B78};3344;NACE-0174;2"
                        "0101217092736-533873081.207275;3345;'';{B1E7CD06-8D0E-4945-8F3A-42C70C643883};33"
                        "55;NACE-0174;{323EAE25-772D-4027-9324-C9F821F3FE18};3357;NACE-0233;{47DC4D81-B37"
                        "F-4B71-85D4-B2E0EAB48471};3363;NACE-0233;{6E964920-4183-4B9B-842E-C7A6A2A791AB};"
                        "3364;NACE-0233;{4FDF0F4D-E341-4EF2-AEA5-999396D00496};3369;NACE-0233;{3C5486B3-B"
                        "638-4B83-B7C5-7F9C97D660F7};3372;NACE-0233;{66CF0991-CB14-44DD-85D9-2407AAE652D5"
                        "};3374;NACE-0233;{99EE4A4B-DF37-43DA-B9DD-2998F904F392};3375;NACE-0233;{E2DD5920"
                        "-DC46-4A69-ADF9-3A372FAB0CA6};3403;NACE-0233;{6A33A218-EFB3-4F24-8C33-21F4256EAE"
                        "09};3404;NACE-0233;{FFF3A016-8286-4FBD-A35B-E08EEB1FB844};3405;NACE-0233;{D1B091"
                        "95-C025-4F50-9105-B296B8E0086A};3407;NACE-0233;{AE1C10E0-7D2D-4DB3-AD6B-A5AE81DC"
                        "13D2};3408;NACE-0233;{5AFE05D2-CD88-42F8-8F64-AC88241BBFEC};3410;NACE-0233;{0C36"
                        "C519-6D72-42CB-A80F-DEF000F65187};3411;NACE-0233;20101217092758-106369674.20578;"
                        "3418;'';{60B4D9F9-29B8-4DE2-BC3F-B2DB6BCD6E51};3432;PRWI-0436;{A3A74C8F-392C-494"
                        "8-9D2B-05940DF57A33};3434;PRWI-0436;{F1A99B2B-969E-4051-954B-0107B8F1A201};3435;"
                        "PRWI-0436;{69C02C22-B4D2-474A-8E91-9076EA0AB10F};3437;PRWI-0436;{6BA3D4FF-FF7A-4"
                        "D30-8311-7CC386B3ED6F};3438;PRWI-0436;{454A84C2-F777-448F-837D-1533115D7174};343"
                        "9;PRWI-0436;{0EAC4769-96E7-4BFA-AEE2-EFCCF5F841FA};3455;PRWI-0712;{04431296-F835"
                        "-4BB1-BCA7-65CFBE6A47A5};3486;PRWI-0712;{FE65D755-A72B-4F7F-A86F-5A4084CAF1F6};3"
                        "488;PRWI-0712;{4AC0022B-6ACE-439A-8CF5-8BD29320C008};3492;PRWI-0712;201012170928"
                        "18-999414563.179016;3495;'';{A61BE73F-D902-40B7-90AE-81E40595F418};3496;PRWI-071"
                        "2;{67026878-506F-438D-BA32-4B5AD18D07B8};3498;PRWI-0712;{EF43D706-68D4-4AE5-AEFF"
                        "-5395F4F96917};3510;CHOH-0440;{0A0AA766-3F0A-449A-8F43-4F2417EF2F07};3511;CHOH-0"
                        "440;{BA56A990-80EF-414E-BA91-638BAB5A2185};3512;CHOH-0440;{B66C146E-49CC-4B5D-92"
                        "64-68D353B16B2F};3513;CHOH-0440;{37BB8566-6BBE-4306-B10C-332DEBD056AE};3514;CHOH"
                        "-0440;{53FAAD53-782E-4477-930C-1704808F2DF7};3515;CHOH-0440;{AA3785E6-F1C4-472D-"
                        "8D3F-4B120D08D318};3516;CHOH-0440;{48184C7A-E8C5-4C74-9E61-6D40CEA36C5B};3517;CH"
                        "OH-0440;{C6795501-95FE-43AC-973A-9F7D170AA60D};3558;CHOH-0239;{00141FF6-9CA2-45C"
                        "4-BA47-E7B503271BAC};3559;CHOH-0239;{F3DF0448-2487-4D37-B41C-B5AD8E9B58D2};3591;"
                        "ROCR-0094"
                    GridlineColor =10921638

                    LayoutCachedLeft =300
                    LayoutCachedTop =120
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =4140
                End
                Begin ListBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3480
                    Top =120
                    Width =1620
                    Height =4020
                    TabIndex =1
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lbxActuallyInOffice"
                    RowSourceType ="Value List"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    OnClick ="[Event Procedure]"
                    Tag ="622"
                    GridlineColor =10921638

                    LayoutCachedLeft =3480
                    LayoutCachedTop =120
                    LayoutCachedWidth =5100
                    LayoutCachedHeight =4140
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =204
                    Left =6420
                    Top =3600
                    Width =1260
                    FontSize =14
                    TabIndex =2
                    Name ="btnSave"
                    Caption ="Save"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Save this year's RIO tags"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6420
                    LayoutCachedTop =3600
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =3960
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6420
                    Top =120
                    Width =2400
                    Height =450
                    FontSize =12
                    TabIndex =3
                    Name ="btnTagEdit"
                    Caption ="Edit Last Selected Tag"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Edit the tag selected in the RIO tag list"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6420
                    LayoutCachedTop =120
                    LayoutCachedWidth =8820
                    LayoutCachedHeight =570
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =60.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8400
                    Top =840
                    Width =1500
                    Height =345
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="tbxTagLastSelected"
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638

                    LayoutCachedLeft =8400
                    LayoutCachedTop =840
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =1185
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6480
                            Top =840
                            Width =1770
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="lblTagLastSelected"
                            Caption ="Last Selected Tag"
                            FontName ="Franklin Gothic Book"
                            GridlineColor =10921638
                            LayoutCachedLeft =6480
                            LayoutCachedTop =840
                            LayoutCachedWidth =8250
                            LayoutCachedHeight =1155
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =420
            BackColor =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    Visible = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =3060
                    Top =60
                    Width =6450
                    Height =360
                    FontSize =10
                    FontWeight =500
                    LeftMargin =36
                    TopMargin =36
                    BackColor =13754878
                    BorderColor =8355711
                    ForeColor =16711680
                    Name ="lblLoading"
                    Caption ="Please be patient, I'm still loading tags..."
                    FontName ="Franklin Gothic Book"
                    GridlineColor =10921638
                    LayoutCachedLeft =3060
                    LayoutCachedTop =60
                    LayoutCachedWidth =9510
                    LayoutCachedHeight =420
                    BackThemeColorIndex =5
                    BackTint =20.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
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
' MODULE:       RIOCheck
' Level:        Application module
' Version:      1.01
'
' Description:  RIO (retired in office) tag check related functions & procedures
'
' Source/date:  Bonnie Campbell, September 26, 2019
' Adapted:      -
' Revisions:    BLC - 9/26/2019 - 1.00 - initial version
'               BLC - 10/1/2019 - 1.01 - add list sort
' =================================

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String
Private m_Cols As Integer

Private m_SaveOK As Boolean 'ok to save record (prevents bound form from immediately updating)

'---------------------
' Event Declarations
'---------------------
Public Event InvalidTitle(Value As String)
Public Event InvalidDirections(Value As String)
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let Title(Value As String)
    If Len(Value) > 0 Then
        m_Title = Value

        'set the form title & caption
        Me.lblTitle.Caption = m_Title
        'Me.Caption = m_Title
    Else
        RaiseEvent InvalidTitle(Value)
    End If
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Directions(Value As String)
    If Len(Value) > 0 Then
        m_Directions = Value

        'set the form directions
        Me.lblDirections.Caption = m_Directions
    Else
        RaiseEvent InvalidDirections(Value)
    End If
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let CallingForm(Value As String)
        m_CallingForm = Value
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

Public Property Let Cols(Value As Integer)
        m_Cols = Value
End Property

Public Property Get Cols() As Integer
    Cols = m_Cols
End Property

' ----------------
'  Events
' ----------------

' ---------------------------------
' Sub:          Form_Open
' Description:  form opening actions
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Microsoft, Unknown
'   https://docs.microsoft.com/en-us/office/vba/api/access.listbox.rowsourcetype
'   https://docs.microsoft.com/en-us/office/vba/api/access.listbox.additem
' Source/date:  Bonnie Campbell, September 2019
' Adapted:      -
' Revisions:
'   BLC - 9/26/2019 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    'default
'    Me.CallingForm = "Main"
'
'    If Len(Me.OpenArgs) > 0 Then Me.CallingForm = Me.OpenArgs
'
'    'minimize calling form
'    ToggleForm Me.CallingForm, -1
    
    'set defaults
    Me.RecordSource = "tbl_Tags"
    Me.Filter = "Tag_Status = 'Retired (In Office)'"
    Me.FilterOn = True
    Me.FilterOnLoad = True
    Me.OrderBy = "Tag"
    Me.OrderByOn = True
    Me.OrderByOnLoad = True
    
    'copy rs for populating lbx
    Dim rsCopy As DAO.Recordset
    Set rsCopy = Me.RecordsetClone
    
    'clear tags
    lbxRIOtags.RowSourceType = "Value List"
    lbxRIOtags.RowSource = ""
    lbxActuallyInOffice.RowSourceType = "Value List"
    lbxActuallyInOffice.RowSource = ""
    
    'status bar message
    DoCmd.Hourglass True
    Application.SysCmd acSysCmdSetStatus, "Loading tags..."
    Me.lblLoading.Visible = True
    
    'dev mode
    tbxDevMode = DEV_MODE
                
    Title = "RIO Check"
    'lblTitle.Caption = "" 'clear header title
    Directions = "Are all the Retired in Office (RIO) tags actually IN the office?"
    
    'defaults
    lblDirections.ForeColor = lngWhite
'    rctPseudoEvent.BackColor = lngLtTan
    btnSave.HoverColor = lngGreen
'    btnReportRIOHistory.HoverColor = lngGreen

    'recordcount
    Dim rs As DAO.Recordset
    Dim rsTags As DAO.Recordset
    Dim iCount As Integer
    Dim sql As String
    Dim itm As String
    Dim tags As Variant
    
    'listbox prep
    Dim widths As String
    
    Cols = 2 '# of columns to use 2 => Tag_ID, Tag    3 => Tag_ID, Tag, Plot
    
    Select Case Cols
        Case 2
            widths = "0"";1"""
            sql = "SELECT Tag_ID, Tag FROM tbl_Tags t " _
                    & "WHERE t.Tag_Status IN ('Retired (In Office)') " _
                    & "ORDER BY Tag;"
        
        Case 3
            widths = "0"";1"";0"""
            sql = "SELECT Tag_ID, Tag, l.Plot_Name AS Plot FROM tbl_Tags t " _
                    & "LEFT JOIN tbl_Locations l ON l.Location_ID = t.Location_ID " _
                    & "WHERE t.Tag_Status IN ('Retired (In Office)') " _
                    & "ORDER BY Tag;"
    End Select
    
    lbxRIOtags.ColumnCount = Cols
    lbxActuallyInOffice.ColumnCount = Cols
    lbxRIOtags.ColumnWidths = widths
    lbxActuallyInOffice.ColumnWidths = widths
    
    'default
    iCount = 0

Debug.Print sql

'    Set rs = CurrentDb.OpenRecordset(sql)
    'use the dupe
    Set rs = rsCopy
    
'    With lbxRIOtags
        'populate RIO tags
        lbxRIOtags.RowSourceType = "Value List"
'        .RowSource = "SELECT t.Tag_ID, t.Tag, l.Plot_Name FROM tbl_Tags t " _
'                                & "LEFT JOIN tbl_Locations l ON l.Location_ID = t.Location_ID " _
'                                & "WHERE t.Tag_Status IN ('Retired (In Office)') " _
'                                & " ORDER BY Tag;"
'        .RowSource = sql
'    End With
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        rs.MoveFirst
        
        'store count
        SetTempVar "TotalRIOs", rs.RecordCount
        lblTotalRIOCount.Caption = "# of RIO tags is the maximum for the listbox below. " & _
                                    "Actual Total # of RIO tags is found by printing the full tag list. " & _
                                    ">> Actual # RIO Tags = "
        lblTotalRIOCount.Caption = lblTotalRIOCount.Caption & " " & TempVars("TotalRIOs")
        
        Do While Not rs.EOF
            Select Case Cols
                Case 3
                    itm = rs("Tag_ID") & ";" & rs("Tag") & ";" & rs("Plot")
                Case 2
                    itm = rs("Tag_ID") & ";" & rs("Tag")
            End Select
'Debug.Print itm
            lbxRIOtags.AddItem itm
            'tags = tags & "," & itm
            rs.MoveNext
        Loop
    End If
    
    Debug.Print "list count = " & lbxRIOtags.ListCount
    
'    Dim ary As Variant
'    Set ary = Split("Tag_ID,Tag,Plot_Name", ",")
    
    
    'Set rsTags = ArrayToRecordset(ary, Split(tags, ","))
    
    'populate for accurate record count
'    If Not rsTags.BOF And rsTags.EOF Then
'        rsTags.MoveLast
'        rsTags.MoveFirst
'    End If
    
'    Do While Not rsTags.EOF
'        lbxRIOtags.AddItem rsTags(0)
'        rsTags.MoveNext
'    Loop
    
    'lbxRIOtags.RowSource = tags
'    lbxRIOtags.RowSource = Split(tags, ",")
    
    tbxRIOTagCount = ">>" & lbxRIOtags.ListCount
    tbxCount = ">>" & lbxActuallyInOffice.ListCount
    
    'set the original RIOTagCount
    lbxActuallyInOffice.Tag = lbxRIOtags.ListCount
    
    'status bar message
    Application.SysCmd acSysCmdSetStatus, lbxRIOtags.ListCount & " tags loaded!"

    Me.lblLoading.Visible = False

Exit_Handler:
    Application.SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
    Set rs = Nothing
    Set rsTags = Nothing
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 2019
' Adapted:      -
' Revisions:
'   BLC - 9/26/2019 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Load[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lbxRIOtags_DblClick
' Description:  listbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Pat Jones, May 1, 2008
'   https://bytes.com/topic/access/insights/795313-double-clicking-move-item-between-list-boxes
' Source/date:  Bonnie Campbell, September 2019
' Adapted:      -
' Revisions:
'   BLC - 9/26/2019 - initial version
'   BLC - 10/1/2019 - added list sort
' ---------------------------------
Private Sub lbxRIOtags_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    MoveListItem lbxRIOtags, lbxActuallyInOffice
    
    SortList lbxActuallyInOffice
    tbxRIOTagCount = ">>" & lbxRIOtags.ListCount
    tbxCount = ">>" & lbxActuallyInOffice.ListCount
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxRIOtags_DblClick[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lbxActuallyInOffice_DblClick
' Description:  listbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Pat Jones, May 1, 2008
'   https://bytes.com/topic/access/insights/795313-double-clicking-move-item-between-list-boxes
' Source/date:  Bonnie Campbell, September 2019
' Adapted:      -
' Revisions:
'   BLC - 9/26/2019 - initial version
'   BLC - 10/1/2019 - added list sort
' ---------------------------------
Private Sub lbxActuallyInOffice_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    MoveListItem lbxActuallyInOffice, lbxRIOtags

'    'SortList lbxRIOtags
'    tbxRIOTagCount.Value = ">>" & lbxRIOtags.ListCount
'    tbxCount = ">>" & lbxActuallyInOffice.ListCount
'
'    If lbxActuallyInOffice.Tag = lbxActuallyInOffice.ListCount Then
'        lbxActuallyInOffice.BackColor = lngLtLime
'    Else
'        lbxActuallyInOffice.BackColor = lngWhite
'    End If

    SetListCounts

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxActuallyInOffice_DblClick[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lbxRIOtags_Click
' Description:  listbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Private Sub lbxRIOtags_Click()
On Error GoTo Err_Handler

    SetTempVar "SelectedTagID", lbxRIOtags.Value
'Debug.Print TempVars("SelectedTagID")

    'populate last selected tag
    tbxTagLastSelected = lbxRIOtags.Column(1)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxRIOtags_Click[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lbxActuallyInOffice_Click
' Description:  listbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Private Sub lbxActuallyInOffice_Click()
On Error GoTo Err_Handler

    SetTempVar "SelectedTagID", lbxActuallyInOffice.Value
'Debug.Print TempVars("SelectedTagID")

    'populate last selected tag
    tbxTagLastSelected = lbxActuallyInOffice.Column(1)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lbxActuallyInOffice_Click[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTagList_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 2019
' Adapted:      -
' Revisions:
'   BLC - 9/30/2019 - initial version
' ---------------------------------
Private Sub btnTagList_Click()
On Error GoTo Err_Handler

'    'prep data columns
'    SetTempVar "ActuallyInOffice", lbxActuallyInOffice.RowSource
'    SetTempVar "RIOs", lbxRIOtags.RowSource
'
'    'open RIO tag report
'    DoCmd.OpenReport "RIOCheck", acViewNormal, , , acDialog

    btnSaveToPDF_Click

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTagList_Click[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnSaveToPDF_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   sxschech, August 14, 2016
'   https://www.tek-tips.com/viewthread.cfm?qid=1768657
'   boblarson, November 8, 2010
'   https://access-programmers.co.uk/forums/showthread.php?t=201223
'   Albert D. Kallal, October 13, 2016
'   https://answers.microsoft.com/en-us/msoffice/forum/all/print-an-access-report-as-pdf-with-vba-how-to-get/75e156f3-59bd-4589-93fe-9175ba3b30af
' Source/date:  Bonnie Campbell, September 2019
' Adapted:      -
' Revisions:
'   BLC - 10/2/2019 - initial version
' ---------------------------------
Private Sub btnSaveToPDF_Click()
On Error GoTo Err_Handler

    Dim FNameExists As Boolean
    Dim result As Variant
    Dim rpt As String
    Dim rptName As String
    Dim strFileName As String
    Dim strExportPath As String
    Dim strExportFolder As String
    Dim strExportFileName As String
    
    rptName = "RIOCheck"
    rpt = Format(Now, "YYYYMMDD_hhmm") & "_ForestVeg_" & rptName
    
    FNameExists = False
    
FNAME:
    Do While FNameExists = False
        strFileName = InputBox("Enter the name for this report." & vbCrLf & vbCrLf & _
                        "On the next screen, choose the directory location " & _
                        "where you want to save the report file.", "Save Report", rpt)
        strFileName = Replace(strFileName, "-", "_")
        If strFileName = "" Then
            MsgBox "No filename was chosen, or the action was canceled by the user.", vbOKOnly, "Missing File Name"
            FNameExists = True
        Else
            strExportPath = SelectFolder()
            strExportFileName = strExportPath & "\" & strFileName & ".pdf"
SaveOrReplace:
            If Dir(strExportFileName) = "" Then
                'DoCmd.OutputTo acOutputReport, rpt, "PDFFormat(*.pdf)", strExportFileName, ShowPDF, "", 0, acExportQualityPrint
                'DoCmd.OutputTo acOutputReport, rpt, acFormatPDF, strExportFileName, -1, "", , acExportQualityPrint
'                DoCmd.OutputTo ObjectType:=acOutputReport, ObjectName:=rpt, OutputFormat:=acFormatPDF, OutputFile:=strExportFileName, _
'                        AutoStart:=True, TemplateFile:="", OutputQuality:=acExportQualityPrint
                DoCmd.OutputTo ObjectType:=acOutputReport, ObjectName:=rptName, OutputFormat:=acFormatPDF, _
                        OutputFile:=strExportFileName, _
                        OutputQuality:=acExportQualityPrint
                
                FNameExists = True
            Else
                result = MsgBox("File " & strExportFileName & " already exists. " & vbCrLf & vbCrLf & _
                            "Would you like to REPLACE this file?", vbYesNo + vbQuestion, _
                            "File Already Exists")
                If result = vbYes Then
                    DeleteFile (strExportFileName)
                    GoTo SaveOrReplace
                Else
                    GoTo FNAME
                End If
            End If
            
            'MsgBox "File " & strExportFileName & " has been created.", vbOKOnly, "Report Saved!"
'            MsgBox "File " & strExportFileName & " has been created." _
'            & vbCrLf & vbCrLf & " Would you like to open the file?", vbDefaultButton1 + vbYesNo + vbQuestion, "Report Saved!"
            
            Dim strPrompt As String
            Dim strTitle As String
            strTitle = "RIO List Saved!"
            strPrompt = "File " & strExportFileName & " has been created." _
                & vbCrLf & vbCrLf & " Select Print or Open to print or open the file?"
            
            'use custom messagebox
            Dim CC As clsMsgBox
            Dim iR As Integer

            Set CC = New clsMsgBox
            
            With CC
                .UseCancel = True
                .Title = strTitle
                .Prompt = strPrompt
                .Icon = Question + DefaultButton3
                .ButtonText1 = "Open PDF"
                .ButtonText2 = "Print PDF"
                .ButtonText3 = "Cancel"
                iR = .MessageBox()
                Select Case iR
                    Case Button1    'open PDF
                        FollowURL strExportFileName
                    Case Button2    'print PDF
                        DoCmd.OpenReport "RIOCheck", acViewNormal, , , acDialog
                    Case Button3    'cancel
                        'do nothing
                End Select
            End With
            
'            If iR = Button1 Then
'                Debug.Print "Open PDF clicked"
'            ElseIf iR = Button2 Then
'                Debug.Print "Print PDF Clicked"
'            ElseIf iR = Button3 Then
'                Debug.Print "Cancelled"
'                Debug.Print "Cancel Clicked"
'            End If
        
        End If
    Loop

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 2501 'ignore error if user cancels print
        GoTo Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnSaveToPDF_Click[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTagEdit_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Private Sub btnTagEdit_Click()
On Error GoTo Err_Handler
  
    'DoCmd.Minimize
    
    'open tag history
    DoCmd.OpenForm "frm_Tags", acNormal, , "Tag_ID='" & TempVars("SelectedTagID") & "'", acFormEdit, acDialog

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTagEdit_Click[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub


' ---------------------------------
' SUB:          btnClose_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/2/2019 - initial version
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          MoveListItem
' Description:  move item from one listbox to another
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Pat Jones, May 1, 2008
'   https://bytes.com/topic/access/insights/795313-double-clicking-move-item-between-list-boxes
' Source/date:  Bonnie Campbell, September 2019
' Adapted:      -
' Revisions:
'   BLC - 9/26/2019 - initial version
' ---------------------------------
Private Sub MoveListItem(lbxFrom As ListBox, lbxTo As ListBox)
On Error GoTo Err_Handler

    ' listbox has 3 columns - 0 = bound value, 1 = tag #, 2 =
    ' because listbox has multiple columns, lbxFrom.Value cannot be added alone - need strItem to populate all columns
    Dim strItem As String
    strItem = lbxFrom.Column(0) & ";" & lbxFrom.Column(1) & IIf(Cols = 3, ";" & lbxFrom.Column(2), "")
    
    lbxTo.AddItem (strItem) '(lbxFrom.Value)
    lbxFrom.RemoveItem (lbxFrom.Value)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - MoveListItem[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SortList
' Description:  sort listbox items
' Assumptions:  listbox has 3 columns: Tag_ID, Tag #, Plot
' Parameters:   lbx - listbox to sort (listbox)
' Returns:      -
' Throws:       none
' References:
'   Fionnuala, October 12, 2011
'   https://stackoverflow.com/questions/7738811/access-listbox-based-on-value-list-sorting-on-column
'   MajP, March 22, 2012
'   https://www.tek-tips.com/viewthread.cfm?qid=1677888
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 10/1/2019 - initial version
' ---------------------------------
Private Sub SortList(lbx As ListBox)
On Error GoTo Err_Handler

'  Dim strTemp As String
'  Dim strTemp2 As String
'  Dim i As Integer
'  Dim j As Integer
'
'    With lbx
'        For i = 0 To .ListCount - 1
'            For j = i + 1 To .ListCount - 1
'                If .ItemData(i) > .ItemData(j) Then
'                    strTemp = .Column(0, i) & ";" & .Column(1, i) & ";" & .Column(2, i) '.ItemData(i)
'                    strTemp2 = .Column(0, j - 1) & ";" & .Column(1, j - 1) & ";" & .Column(2, j - 1)
'                    .RemoveItem (i)
'                    .AddItem strTemp2, i '.ItemData(j - 1), i
'                    .RemoveItem (j)
'                    .AddItem strTemp, j - 1
'                End If
'            Next j
'        Next i
'    End With
    
    Dim rs As ADODB.Recordset
    Dim slist As String
    Dim r As Integer
    Dim c As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ary As Variant
    
'    Set rs = lbx.Recordset
    'default
    slist = ""
    
    With lbx
        For r = 0 To lbx.ListCount - 1
            For c = 0 To Cols - 1 'Cols = 2 or 3 column lbx
                slist = slist & lbx.Column(c, r) & ","
            Next
            'slist = slist & ","
        Next
        slist = Left(slist, Len(slist) - 1)
    End With
'Debug.Print "slist = " & slist
    
    Set rs = New ADODB.Recordset

    With rs
        .ActiveConnection = Nothing
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        With .Fields
            .Append "Tag_ID", adVarChar, 255
            .Append "Tag", adInteger
            If Cols = 3 Then .Append "Plot_Name", adVarChar, 255
        End With
        .Open
        
        ary = Split(slist, ",")
        
        For j = 0 To UBound(ary)
            .AddNew
            For i = 0 To Cols - 1 'Cols =  2 or 3 column lbx
'Debug.Print "i=" & i
'Debug.Print "j=" & j
                .Fields(i).Value = ary(j)
                If j < UBound(ary) Then j = j + 1 'next j
'Debug.Print "j=" & j
            
            Next
            If j < UBound(ary) Then j = j - 1 'next j
'Debug.Print "j=" & j
        
        Next
    
        .Sort = "Tag"
    
    End With
    
    slist = rs.GetString(, , ",", ",")
    slist = Left(slist, Len(slist) - 1)
    
'    Debug.Print slist
    
    'repopulate the listbox
    lbx.RowSource = slist
    
Exit_Handler:
    'cleanup
    Set rs = Nothing
    Set ary = Nothing
    slist = ""
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortList[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetListCounts
' Description:  set current counts for listboxes
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 2019
' Adapted:      -
' Revisions:
'   BLC - 9/26/2019 - initial version
' ---------------------------------
Private Sub SetListCounts()
On Error GoTo Err_Handler

    'set values
    tbxRIOTagCount.Value = ">>" & lbxRIOtags.ListCount
    tbxCount = ">>" & lbxActuallyInOffice.ListCount
    
    With lbxActuallyInOffice
        If .Tag = .ListCount Then
            .BackColor = lngLtLime
        Else
            .BackColor = lngWhite
        End If
    End With

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetListCounts[RIOCheck form])"
    End Select
    Resume Exit_Handler
End Sub
