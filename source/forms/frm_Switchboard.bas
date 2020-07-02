Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =3
    PictureSizeMode =3
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =10215
    DatasheetFontHeight =10
    ItemSuffix =104
    Left =720
    Top =1275
    Right =10935
    Bottom =6840
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3b82f36a92b1e340
    End
    RecordSource ="tsys_App_Defaults"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    PictureSizeMode =3
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin CheckBox
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
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin ListBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =11056034
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =5580
            BackColor =0
            Name ="Detail"
            BackThemeColorIndex =0
            Begin
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OldBorderStyle =0
                    OverlapFlags =93
                    Width =10140
                    Height =2085
                    BackColor =0
                    Name ="boxBanner"
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =2085
                End
                Begin Tab
                    OverlapFlags =223
                    TextFontCharSet =204
                    Top =1800
                    Width =10215
                    Height =3780
                    FontWeight =700
                    Name ="tabMenu"
                    FontName ="Arial"

                    LayoutCachedTop =1800
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =5580
                    UseTheme =255
                    BackColor =14277081
                    BackThemeColorIndex =1
                    BackShade =85.0
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =13355721
                    HoverThemeColorIndex =4
                    HoverTint =40.0
                    PressedColor =15921906
                    PressedThemeColorIndex =1
                    PressedShade =95.0
                    HoverForeColor =4342595
                    HoverForeThemeColorIndex =2
                    PressedForeColor =0
                    PressedForeThemeColorIndex =0
                    ForeThemeColorIndex =0
                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =75
                            Top =2220
                            Width =10065
                            Height =3285
                            Name ="pagMain"
                            Caption =" Main menu"
                            LayoutCachedLeft =75
                            LayoutCachedTop =2220
                            LayoutCachedWidth =10140
                            LayoutCachedHeight =5505
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =4260
                                    Top =2415
                                    Width =1530
                                    Height =1319
                                    FontSize =15
                                    FontWeight =700
                                    ForeColor =0
                                    Name ="btnGateway"
                                    Caption ="Browse"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    ControlTipText ="Browse existing plot and sampling event data"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =4260
                                    LayoutCachedTop =2415
                                    LayoutCachedWidth =5790
                                    LayoutCachedHeight =3734
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =8289145
                                    BackThemeColorIndex =4
                                    BorderColor =8289145
                                    BorderThemeColorIndex =4
                                    HoverColor =65280
                                    PressedColor =6644321
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =2040
                                    Top =2400
                                    Width =1530
                                    Height =1319
                                    FontSize =15
                                    FontWeight =700
                                    TabIndex =1
                                    ForeColor =0
                                    Name ="btnAddEvent"
                                    Caption ="Create"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    ControlTipText ="Create a new sampling event"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =2040
                                    LayoutCachedTop =2400
                                    LayoutCachedWidth =3570
                                    LayoutCachedHeight =3719
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =8289145
                                    BackThemeColorIndex =4
                                    BorderColor =8289145
                                    BorderThemeColorIndex =4
                                    HoverColor =65280
                                    PressedColor =6644321
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =6420
                                    Top =2415
                                    Width =1530
                                    Height =1319
                                    FontSize =15
                                    FontWeight =700
                                    TabIndex =2
                                    ForeColor =0
                                    Name ="btnDataSummary"
                                    Caption ="Summarize"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Calibri"
                                    ControlTipText ="Summarize data using standard queries"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =6420
                                    LayoutCachedTop =2415
                                    LayoutCachedWidth =7950
                                    LayoutCachedHeight =3734
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =8289145
                                    BackThemeColorIndex =4
                                    BorderColor =8289145
                                    BorderThemeColorIndex =4
                                    HoverColor =65280
                                    PressedColor =6644321
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =180
                                    Top =3840
                                    Width =786
                                    Height =786
                                    FontSize =12
                                    FontWeight =700
                                    TabIndex =3
                                    ForeColor =0
                                    Name ="btnUtilities"
                                    Caption ="Setup"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000020000000200000000100180000000000000c0000c40e0000c40e0000 ,
                                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffafafaefefefe9e9e9e9e9e9e9e9e9e9e9e9e9e9e9efef ,
                                        0xeffafafaffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffefefefcececebdbdbdbcbcbcbcbcbcbcbcbcbdbdbdcece ,
                                        0xceefefefffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffffffffffdfdfdf6f6f6eeeeee ,
                                        0xf2f2f2fbfbfbffffffe9e9e97472706a68666765637774726765636a68667472 ,
                                        0x70e9e9e9fffffffbfbfbf2f2f2eeeeeef6f6f6fdfdfdffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffffbfbfbeeeeeed8d8d8c7c7c7 ,
                                        0xd1d1d1ebebebfcfcfce9e9e96b6967e6e6e3e0dfdddedddde0dfdde6e6e36b69 ,
                                        0x67e9e9e9fcfcfcebebebd1d1d1c7c7c7d8d8d8eeeeeefbfbfbffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffbfbfbebebebcfcfcf9d9d9c706c6a ,
                                        0x868483cbcbcbe2e2e2d8d8d86b6866e0dfdededcdbdcdad9dedcdbe0dfde6b67 ,
                                        0x65d8d8d8e2e2e2cbcbcb868483706c6a9d9d9ccfcfcfebebebfbfbfbffffffff ,
                                        0xfffffffffffffffffffffffffffffdfdfdeeeeeecfcfcf85848392908ec9c8c6 ,
                                        0x82807e868584b8b8b8b7b7b7696765dfdedcdad8d7d8d6d5dad8d7dededb716f ,
                                        0x6db7b7b7b8b8b885848382807ec9c8c692908e858483cfcfcfeeeeeefdfdfdff ,
                                        0xfffffffffffffffffffffffffffff6f6f6d8d8d8888685989694e2e1dfdedcdb ,
                                        0xdadad88f8d8b6b69666a68669c9a98dddbdad7d5d4d6d4d3d7d5d4dddbda9c9a ,
                                        0x9869676572706e979592dad9d7dedcdbe2e1df989694888685d8d8d8f6f6f6ff ,
                                        0xffffffffffffffffffffffffffffeeeeeeaaa9a8949290e2e1e0d8d5d4d5d4d3 ,
                                        0xd8d5d4dededde2e1dfe1e1dedddbdad7d5d4d4d2d1d4d2d1d4d2d1d7d5d4dddb ,
                                        0xdae1e0dee1e1dededddcd8d5d4d5d4d3d8d5d4e2e1e0949290aaa9a8eeeeeeff ,
                                        0xfffffffffffffffffffffffffffff2f2f2757371cccbc9dad9d7d2d1d0d2d0cf ,
                                        0xd2d0cfd3d1d0d3d1d0d3d1d0d2d0cfd2d0cfd2d0cfd2d0cfd2d0cfd2d0cfd2d0 ,
                                        0xcfd3d1d0d3d1d0d3d1d0d2d0cfd2d0cfd3d1d0dcdad9cccbc9757371f2f2f2ff ,
                                        0xfffffffffffffffffffffffffffffbfbfb9d9b9a878583e4e2e1d4d2d1d0cecd ,
                                        0xd0cecdd0cecdd0cecdd0cecdd0cecdd0cecdd0cecdd0cecdd0cecdd0cecdd0ce ,
                                        0xcdd0cecdd0cecdd0cecdd0cecdd0cecdd4d2d1dcdad98785839c9a99fbfbfbff ,
                                        0xfffffffffffffffffffffffffffffffffffcfcfc9d9b9a9d9a98e0e0dfcecccb ,
                                        0xcecccbcecccbcecccbcecccbcecccbcecccbcecccbd3d2d0cecccbcecccbcecc ,
                                        0xcbcecccbcecccbcecccbcecccbcecccbe0dedd9c9a979b9a99f9f9f9ffffffff ,
                                        0xfffffffffffffffffafafaefefefe9e9e9e9e9e9d8d8d872706ee5e5e2cccac9 ,
                                        0xcccac9cccac9cccac9cccac9d2d0cfdedddbe6e5e3d6d5d3e6e5e3dedddbd2d0 ,
                                        0xcfcccac9cccac9cccac9cccac9cccac9e0e0dd908e8cc6c6c5e8e8e8e9e9e9ef ,
                                        0xefeffafafaffffffefefefcececebdbdbdbcbcbcb7b7b773706ee7e6e4cac8c7 ,
                                        0xcac8c7cac8c7c9c8c7d4d2d1e0dfdd9f9d9b7472707e7c7a7472709f9d9bdfdf ,
                                        0xddd4d2d1c9c7c7cac8c7cac8c7cac8c7e5e2e27a7876b7b7b7bcbcbcbdbdbdce ,
                                        0xceceefefefffffffe9e9e9827f7d7a77757875727c7a78a6a4a2dedddbc8c5c4 ,
                                        0xc9c6c5c9c5c5d0cecce1dfdf7b7977b7b5b4fcfcfcfffffffcfcfcb7b5b48b89 ,
                                        0x87e1dfded0ceccc9c5c5c9c6c5c9c5c4dedddba6a4a27572707875737a777582 ,
                                        0x7f7de9e9e9ffffffe9e9e97b7977eeedede9e9e8e8e7e7dedddbd1cfcdc6c3c1 ,
                                        0xc7c4c2c6c3c1dddcdaa19f9db1b0affcfcfcfffffffffffffffffffcfcfcb1b0 ,
                                        0xafa19f9ddddcdac6c3c1c7c4c2c7c4c1cfcdccdededbe8e8e7eae9e8eeeeed7b ,
                                        0x7977e9e9e9ffffffe9e9e97a7775ebeae8c3c0bec4c1bec4c1bfc4c1bfc5c2c0 ,
                                        0xc5c2c0c4c1beeae9e7787674e9e9e9ffffffffffffffffffffffffffffffe9e9 ,
                                        0xe9787674eae9e7c4c0bec5c2c0c5c2c0c5c1bfc4c1bfc4c1bec3c0beebeae87a ,
                                        0x7876e9e9e9ffffffe9e9e98a8886ebeae8c1bdbbc2bfbdc2bfbdc3c0bec3c0be ,
                                        0xc3c0bec6c4c2dbdbda807d7be4e4e4fcfcfcfffffffffffffffffffcfcfce4e4 ,
                                        0xe4807d7bdbdbdac9c7c5c3c0bec3c0bec3c0bec2bfbdc2bfbdc1bebcebebe97c ,
                                        0x7977e9e9e9ffffffe9e9e97d7b79ebebeabebbb9bfbbb9bfbcbabfbcbac1bebc ,
                                        0xc1bdbbc0bcbaecebea7c7a78d0d0d0efefeffcfcfcfffffffcfcfcefefefd0d0 ,
                                        0xd07c7a78ebebeabfbcbac0bdbbc1bebcc0bcbabfbcbabfbbb9bebbb9ecebea7d ,
                                        0x7b79e9e9e9ffffffefefef807e7cf1f0efedecebebebeaeae9e9cfcdcdbdbab8 ,
                                        0xbebbb9bdbab8dedcdba7a5a39d9c9bd0d0d0e4e4e4e9e9e9e4e4e4d0d0d09d9c ,
                                        0x9ba7a5a3e0dedebdbab8bebbb9bdbab8cdcbcadfdfddecebeaedecebf1f0ef80 ,
                                        0x7e7cefefeffffffffafafa8b898783807d807d7b7e7b799e9c9ae3e2e0bbb8b6 ,
                                        0xbcb9b7bcb9b7c7c4c2e7e6e693918f9b9a99babababebebebababa9b9a998583 ,
                                        0x80e7e7e6c7c4c2bcb9b7bcb9b7bbb8b6e0dfddadacaa7e7b79807d7b83807d8b ,
                                        0x8987fafafafffffffffffffffffffffffffbfbfbe2e2e2807e7ce7e7e6bfbcb9 ,
                                        0xbab7b5bab7b5b9b6b4cccac8e8e7e7a9a7a5807e7c83817f807e7ca9a7a5e9e8 ,
                                        0xe7cccac8b9b6b4bab7b5bab7b5bfbcb9e7e6e67f7d7be1e1e1fafafaffffffff ,
                                        0xfffffffffffffffffffffffffffffbfbfbeaeaeacacaca878483d3d1d0ceccca ,
                                        0xb8b5b3b9b6b4b8b5b3b7b4b2c4c2c0e1dfdef1f0eee2e0e0f1f0eee1dfdec4c2 ,
                                        0xc0b7b4b2b8b5b3b9b6b4b8b5b3cecdcad3d1d0868482c8c8c8e9e9e9fbfbfbff ,
                                        0xfffffffffffffffffffffffffffff2f2f2d0d0d0949190acaaa8e5e3e3b8b5b3 ,
                                        0xb6b3b1b7b4b2b7b4b2b6b3b1b5b2b0b5b2afb4b1afc3c0bfb4b1afb5b2afb5b2 ,
                                        0xb0b6b3b1b7b4b2b7b4b2b6b3b1b4b1afe5e3e2bab8b68f8d8bd0d0d0f2f2f2ff ,
                                        0xffffffffffffffffffffffffffffeeeeee9e9c9b9a9896efeeedb9b6b4b3b0ae ,
                                        0xb4b1afb4b1afb4b1afb4b1afb4b1afb4b1afb4b1afb4b1afb4b1afb4b1afb4b1 ,
                                        0xafb4b1afb4b1afb4b1afb4b1afb3b0aeb5b2b0eeeeeda8a6a49e9b9aeeeeeeff ,
                                        0xfffffffffffffffffffffffffffff6f6f68a8785dbdad9c8c6c4b0adabb2afad ,
                                        0xb2afadb1aeacb0adabb0adabb1aeacb2afadb3b0aeb4b1afb3b0aeb2afadb1ae ,
                                        0xacb0adabb0adabb1aeacb2afadb2afadb0adabc8c5c3dbd9d88a8785f6f6f6ff ,
                                        0xfffffffffffffffffffffffffffffdfdfdcac9c8aaa8a6edebeabbb9b7afacaa ,
                                        0xbbb9b7e3e1e0f0efedf4f3f2e3e1e0c9c8c6b2afadb3b0aeb2afadc9c7c5e2e1 ,
                                        0xe0efefede7e5e4ebe9e8bbbab7b0acaabbb9b7f0f0eeaaa8a6cccbcafdfdfdff ,
                                        0xfffffffffffffffffffffffffffffffffffafafa959290b2b0aeecebebc8c6c4 ,
                                        0xf0efeeb7b5b3908d8b888583a9a6a4e3e2e1b1aeacb3b0aeb1aeace3e1e0b5b3 ,
                                        0xb18f8b89a4a2a0b0aeaceae9e8cdcbc8f0f0efb2b0aeacaaa9fbfbfbffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffffafafa979492adaaa9ebe9e8 ,
                                        0xb3b1af979593fafafae9e9e98b8886e4e4e3b0adabb2afadb0adabe3e2e1a4a2 ,
                                        0xa0e8e8e8e7e6e6aeadac9f9d9bdddcdcaeaaa9aeabaafbfbfbffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffffffffffdfdfdd2d1d0918e8c ,
                                        0x989594fafafaffffffe9e9e98e8b89e5e5e2aeaba9afacaaaeaba8e3e3e1a7a4 ,
                                        0xa2e9e9e9fffffffbfbfbb1afad918e8bd2d1d0fdfdfdffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffefefef908d8bfaf9f9f7f6f6eeedecf7f6f5f9f8f7aaa7 ,
                                        0xa6efefefffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffafafa9a9796928f8d918d8b908d8b908d8b928f8d9896 ,
                                        0x94fafafaffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff
                                    End
                                    FontName ="Calibri"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Advanced utilities"
                                    Picture ="Farm-Fresh_cog.bmp"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =3840
                                    LayoutCachedWidth =966
                                    LayoutCachedHeight =4626
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =8289145
                                    BackThemeColorIndex =4
                                    BorderColor =8289145
                                    BorderThemeColorIndex =4
                                    HoverColor =65280
                                    PressedColor =6644321
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    Left =9180
                                    Top =3840
                                    Width =786
                                    Height =786
                                    FontSize =12
                                    FontWeight =700
                                    TabIndex =4
                                    ForeColor =0
                                    Name ="btnExit"
                                    Caption ="Exit"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000030000000300000000100180000000000001b0000c40e0000c40e0000 ,
                                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffa4a4a4a2a2a2a2a2a2a0a0 ,
                                        0xa0ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffff9f9f9fa2a2a2a2a2a29f9f9feeeeeeffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffff9e9e9e9a9a9a9a9a9a999999ffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff9898 ,
                                        0x989a9a9a9a9a9a989898edededffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffff9696968f8f8f8f8f8f8f8f ,
                                        0x8fffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffff8f8f8f8f8f8f8f8f8f8f8f8fedededffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffff9595958e8e8e8e8e8e8e8e8effffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8e8e ,
                                        0x8e8e8e8e8e8e8e8e8e8eedededffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffff9696968f8f8f8f8f8f8f8f ,
                                        0x8fffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffff8f8f8f8f8f8f8f8f8f8f8f8fedededffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffff9494948d8d8d8d8d8d8d8d8dfefefeffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff9191 ,
                                        0x91919191919191919191eeeeeeffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffefefe7c7c7c7474747474747777 ,
                                        0x77f7f7f7ffffffffffffffffffdadadaffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffff929292929292929292929292eeeeeeffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffefefe6e6e6e636363636363676767f2f2f2ffffffffffffffffffa7a7a7 ,
                                        0xc9c9c9fefefeffffffffffffffffffffffffffffffffffffffffffffffff9393 ,
                                        0x93939393939393939393eeeeeeffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffffff9f9f9f8f8f8f8f8f8f9f9 ,
                                        0xf9fefefeffffffffffffffffff9f9f9fa6a6a6bebebefafafaffffffffffffff ,
                                        0xffffffffffffffffffffffffffff959595959595959595959595eeeeeeffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff999999 ,
                                        0x9c9c9ca6a6a6b6b6b6f4f4f4ffffffffffffffffffffffffffffffffffff9696 ,
                                        0x96969696969696969696efefefffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffdddddddededededededededededededfdfdfdfdfdfdfdfdfdfdfdfe0e0 ,
                                        0xe0e0e0e0e0e0e0e0e0e0e0e0e09898989898989b9b9ba4a4a4b1b1b1ecececff ,
                                        0xffffffffffffffffffffffffffff979797979797979797979797efefefffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffafafafb1b1b1b1b1b1b1b1b1b1 ,
                                        0xb1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1b1adadad9d9d9d ,
                                        0x9898989898989c9c9ca4a4a4adadade1e1e1ffffffffffffffffffffffff9898 ,
                                        0x98989898989898989898efefefffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffa1a1a1a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0a0 ,
                                        0xa0a0a0a0a0a0a0a0a0a09f9f9f9b9b9b9a9a9a9a9a9a9a9a9a9d9d9da4a4a4ac ,
                                        0xacacd4d4d4ffffffffffffffffff9a9a9a9a9a9a9a9a9a9a9a9aefefefffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffff9e9e9e9b9b9b9b9b9b9b9b9b9b ,
                                        0x9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b9b ,
                                        0x9b9b9b9b9b9b9b9b9b9b9b9b9c9c9c9797978d8d8dc6c6c6ffffffffffff9b9b ,
                                        0x9b9b9b9b9b9b9b9b9b9befefefffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfefefe9999999696969696969696969696969696969696969696969696969696 ,
                                        0x969696969696969696969797979b9b9b9c9c9c9c9c9c9c9c9c9a9a9a8e8e8e76 ,
                                        0x76767d7d7de5e5e5ffffffffffff9c9c9c9c9c9c9c9c9c9c9c9cefefefffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffff7f7f77c7c7c74747474747474747474 ,
                                        0x74747474747474747474747474747474747474747474747474747b7b7b989898 ,
                                        0x9e9e9e9d9d9d9a9a9a8b8b8b737373868686f0f0f0ffffffffffffffffff9e9e ,
                                        0x9e9e9e9e9e9e9e9e9e9ef0f0f0ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xf4f4f47a7a7a7272727272727272727272727272727272727272727272727272 ,
                                        0x727272727272727373737b7b7b9797979e9e9e999999888888717171979797f9 ,
                                        0xf9f9ffffffffffffffffffffffff9f9f9f9f9f9f9f9f9f9f9f9ff0f0f0ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffefefefefefefefefefefefefe ,
                                        0xfefefefefefefefefefefefefefefefefefefefefefefefefefefefefea3a3a3 ,
                                        0x989898848484717171aeaeaefdfdfdffffffffffffffffffffffffffffffa0a0 ,
                                        0xa0a0a0a0a0a0a0a0a0a0f0f0f0ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffff9d9d9d808080747474c6c6c6ffffffffffffff ,
                                        0xffffffffffffffffffffffffffffa2a2a2a2a2a2a2a2a2a2a2a2f0f0f0ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff828282 ,
                                        0x797979dadadaffffffffffffffffffffffffffffffffffffffffffffffffa3a3 ,
                                        0xa3a3a3a3a3a3a3a3a3a3f1f1f1ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffb8b8b8b8b8b8b8b8b8b7b7 ,
                                        0xb7ffffffffffffffffffffffff8f8f8feaeaeaffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffa4a4a4a4a4a4a4a4a4a4a4a4f1f1f1ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffaeaeaea9a9a9a9a9a9a9a9a9fffffffffffffffffffffffff8f8f8 ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffa5a5 ,
                                        0xa5a5a5a5a5a5a5a5a5a5f1f1f1ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffacacaca7a7a7a7a7a7a7a7 ,
                                        0xa7ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffa7a7a7a7a7a7a7a7a7a7a7a7f1f1f1ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffadadada8a8a8a8a8a8a8a8a8ffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffa8a8 ,
                                        0xa8a8a8a8a8a8a8a8a8a8f1f1f1ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffaeaeaea9a9a9a9a9a9a9a9 ,
                                        0xa9ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffa9a9a9a9a9a9a9a9a9a9a9a9f2f2f2ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffb0b0b0abababababababababffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffabab ,
                                        0xababababababababababf2f2f2ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffb1b1b1acacacacacacacac ,
                                        0xacffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffacacacacacacacacacacacacf2f2f2ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffb2b2b2adadadadadadaeaeaef0f0f0ffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffff0f0f0aeae ,
                                        0xaeadadadadadadadadadf2f2f2ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffb3b3b3afafafafafafb0b0 ,
                                        0xb0b6b6b6bfbfbfc1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1c1 ,
                                        0xc1c1c1c1c1c1c1c1c0c0c0b8b8b8b1b1b1afafafafafafaeaeaef2f2f2ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffa6a6a6a8a8a8a9a9a9a9a9a9aaaaaaacacacadadadadadadadadad ,
                                        0xadadadadadadadadadadadadadadadadadadadadadadadadacacacabababa9a9 ,
                                        0xa9a9a9a9a8a8a8a3a3a3f2f2f2ffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffefefe8787878282828484848484 ,
                                        0x8484848484848484848484848484848484848484848484848484848484848484 ,
                                        0x8484848484848484848484848484848484848484828282818181f3f3f3ffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffbababa7b7b7b777777777777777777777777777777777777777777 ,
                                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                                        0x777777777a7a7aabababffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffffffffffffffffffffffffffffffffefefefefefefefe ,
                                        0xfefefefefefefefefefefefefefefefefefefefefefefefefefefefefefefefe ,
                                        0xfefefefefefefefefefefefefefefefefefefefefefefeffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff
                                    End
                                    FontName ="Calibri"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Exit the application"
                                    Picture ="ic_menu_exit.bmp"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120
                                    GridlineStyleLeft =1
                                    GridlineStyleTop =1
                                    GridlineStyleRight =1
                                    GridlineStyleBottom =1

                                    LayoutCachedLeft =9180
                                    LayoutCachedTop =3840
                                    LayoutCachedWidth =9966
                                    LayoutCachedHeight =4626
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =8289145
                                    BackThemeColorIndex =4
                                    BorderColor =8289145
                                    BorderThemeColorIndex =4
                                    HoverColor =4819194
                                    HoverThemeColorIndex =5
                                    HoverTint =80.0
                                    PressedColor =6644321
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                    Overlaps =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    CanGrow = NotDefault
                                    CanShrink = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    TextAlign =2
                                    BackStyle =0
                                    Left =1140
                                    Top =4380
                                    Width =7860
                                    Height =234
                                    TabIndex =5
                                    ForeColor =9870754
                                    Name ="tbxLinkPath"
                                    StatusBarText ="Currently linked back-end database file"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1140
                                    LayoutCachedTop =4380
                                    LayoutCachedWidth =9000
                                    LayoutCachedHeight =4614
                                End
                                Begin Label
                                    FontItalic = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3885
                                    Top =4200
                                    Width =2355
                                    Height =210
                                    ForeColor =9870754
                                    Name ="lblLinkPath"
                                    Caption ="U s i n g   B a c k e n d ......."
                                    FontName ="Calibri"
                                    LayoutCachedLeft =3885
                                    LayoutCachedTop =4200
                                    LayoutCachedWidth =6240
                                    LayoutCachedHeight =4410
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =180
                                    Top =2399
                                    Width =786
                                    Height =786
                                    FontSize =12
                                    FontWeight =700
                                    TabIndex =6
                                    ForeColor =0
                                    Name ="btnDashboard"
                                    Caption ="Setup"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x280000002000000020000000010008000000000002040000120b0000120b0000 ,
                                        0x0001000000000000000000000101010002020200030303000404040005050500 ,
                                        0x060606000707070008080800090909000a0a0a000b0b0b000c0c0c000d0d0d00 ,
                                        0x0e0e0e000f0f0f00101010001111110012121200131313001414140015151500 ,
                                        0x161616001717170018181800191919001a1a1a001b1b1b001c1c1c001d1d1d00 ,
                                        0x1e1e1e001f1f1f00202020002121210022222200232323002424240025252500 ,
                                        0x262626002727270028282800292929002a2a2a002b2b2b002c2c2c002d2d2d00 ,
                                        0x2e2e2e002f2f2f00303030003131310032323200333333003434340035353500 ,
                                        0x363636003737370038383800393939003a3a3a003b3b3b003c3c3c003d3d3d00 ,
                                        0x3e3e3e003f3f3f00404040004141410042424200434343004444440045454500 ,
                                        0x464646004747470048484800494949004a4a4a004b4b4b004c4c4c004d4d4d00 ,
                                        0x4e4e4e004f4f4f00505050005151510052525200535353005454540055555500 ,
                                        0x565656005757570058585800595959005a5a5a005b5b5b005c5c5c005d5d5d00 ,
                                        0x5e5e5e005f5f5f00606060006161610062626200636363006464640065656500 ,
                                        0x666666006767670068686800696969006a6a6a006b6b6b006c6c6c006d6d6d00 ,
                                        0x6e6e6e006f6f6f00707070007171710072727200737373007474740075757500 ,
                                        0x767676007777770078787800797979007a7a7a007b7b7b007c7c7c007d7d7d00 ,
                                        0x7e7e7e007f7f7f00808080008181810082828200838383008484840085858500 ,
                                        0x868686008787870088888800898989008a8a8a008b8b8b008c8c8c008d8d8d00 ,
                                        0x8e8e8e008f8f8f00909090009191910092929200939393009494940095959500 ,
                                        0x969696009797970098989800999999009a9a9a009b9b9b009c9c9c009d9d9d00 ,
                                        0x9e9e9e009f9f9f00a0a0a000a1a1a100a2a2a200a3a3a300a4a4a400a5a5a500 ,
                                        0xa6a6a600a7a7a700a8a8a800a9a9a900aaaaaa00ababab00acacac00adadad00 ,
                                        0xaeaeae00afafaf00b0b0b000b1b1b100b2b2b200b3b3b300b4b4b400b5b5b500 ,
                                        0xb6b6b600b7b7b700b8b8b800b9b9b900bababa00bbbbbb00bcbcbc00bdbdbd00 ,
                                        0xbebebe00bfbfbf00c0c0c000c1c1c100c2c2c200c3c3c300c4c4c400c5c5c500 ,
                                        0xc6c6c600c7c7c700c8c8c800c9c9c900cacaca00cbcbcb00cccccc00cdcdcd00 ,
                                        0xcecece00cfcfcf00d0d0d000d1d1d100d2d2d200d3d3d300d4d4d400d5d5d500 ,
                                        0xd6d6d600d7d7d700d8d8d800d9d9d900dadada00dbdbdb00dcdcdc00dddddd00 ,
                                        0xdedede00dfdfdf00e0e0e000e1e1e100e2e2e200e3e3e300e4e4e400e5e5e500 ,
                                        0xe6e6e600e7e7e700e8e8e800e9e9e900eaeaea00ebebeb00ececec00ededed00 ,
                                        0xeeeeee00efefef00f0f0f000f1f1f100f2f2f200f3f3f300f4f4f400f5f5f500 ,
                                        0xf6f6f600f7f7f700f8f8f800f9f9f900fafafa00fbfbfb00fcfcfc00fdfdfd00 ,
                                        0xfefefe00ffffff00ffffffffffffffffffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffeeb9b2 ,
                                        0xc6edfcffffffffffffffffffffffffffffffffffffffffffffffffd0b9d1d5b9 ,
                                        0xd3c2b1ade8fffffffffffffffffffffffffffffffffffbc961618ec3ebf0d9b9 ,
                                        0xdad2cbb7dae7fcfffffffffffffffffffffffffffb8961675c5f5fbef2f1d9b8 ,
                                        0xdad2ccb6dfe4ecfbfffffffffffffffffac4848eab9372725c5f64c1f2f0d9b8 ,
                                        0xdad2ccb6e3eaf0fdfffffffffffff0968ca1819c819073725c5f63c1f3f0dab8 ,
                                        0xdad3ccb7e8f4fefffffff47b4a85a1b6afaf819b899175725c5f63c1f3f0d9b8 ,
                                        0xdad3cbb7edf9ffffd1675e7543575dc1b0af829b899176725c5f63c1f4f0d9b8 ,
                                        0xdad2ccb7f1f6fdff71858989445d64c3afaf819b899b7a725c5f64c5f5f1d9b8 ,
                                        0xdad2cbb8fbffffff628e8889475d64c6b0af819b899b7d725c5f64c4f7f1dab9 ,
                                        0xdad2cbb7ffffffff628f8888485e64d0afb0829c899c7f715c5f63c5f8f1dab9 ,
                                        0xdad2cbb9ffffffff62968988475d65d6b0b0889c89a685725c5f63c5f8f3daba ,
                                        0xdad3cbbcffffffff62968888485e64d8b2b0879c89a58a725c5f63c8faf6d4b9 ,
                                        0xdad3ccb8ffffffff62988889485d64d8b7b0879c89a58e725e5f62c8f5e5beb4 ,
                                        0xb2b7c0b7ffffffff629a8888485e64d8bcaf879c89a696725f5e62badceef2f1 ,
                                        0xe8d1b9b3ffffffff62a18888475d64d9c5b0889c89b1aa725f5f63626a92d3dd ,
                                        0xccd2eeffffffffff62a58988485d64d8cdaf879c89b1b8725f5f625f5b6efdfe ,
                                        0xffffffffffffffff62aa8889485d64d8d7b0879c88b0c4735e5f625f5b6efeff ,
                                        0xffffffffffffffff64b98889475d64e2e0b3879c88b0c57a5f5f635f5b6efdff ,
                                        0xffffffffffffffff64bf8989485d64e2e1c1859288b0d38c5f5f635f5c6efdff ,
                                        0xffffffffffffffff64c38a89475d66dfccb7948778b1d3a45a5c635f5b6efdff ,
                                        0xffffffffffffffff64c38f88485d74a5c6dbc6b693b1b1777678625c576bfaff ,
                                        0xffffffffffffffff64ce9788485d554c509bbeb1ad6c6da7bfb09a7f635deaff ,
                                        0xffffffffffffffff64ce9e89475d554d4a8cfffffffffab879726c95d3ffffff ,
                                        0xffffffffffffffff64cea888485d554d4a63ffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff64ceb188485d554d4b72ffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff64cecc8a485d554d4a72ffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff64ced283465c554d4a73ffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff6f9e6d7178624a49456affffffffffffffffffffffffffff ,
                                        0xffffffffffffffff70628ea5ae9e845d496effffffffffffffffffffffffffff ,
                                        0xfffffffffffffffffffdc2774d5ba4f1ffffffffffffffffffffffffffffffff ,
                                        0xffffffffffffffff0000
                                    End
                                    FontName ="Calibri"
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Progress Dashboard"
                                    Picture ="3d Graph - Grayscale.bmp"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =2399
                                    LayoutCachedWidth =966
                                    LayoutCachedHeight =3185
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Shape =1
                                    Gradient =12
                                    BackColor =8289145
                                    BackThemeColorIndex =4
                                    BorderColor =8289145
                                    BorderThemeColorIndex =4
                                    HoverColor =65280
                                    PressedColor =6644321
                                    PressedThemeColorIndex =4
                                    PressedShade =80.0
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =23
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    CanGrow = NotDefault
                                    CanShrink = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    TextAlign =1
                                    BackStyle =0
                                    Left =180
                                    Top =4980
                                    Width =1512
                                    Height =234
                                    TabIndex =7
                                    ForeColor =9870754
                                    Name ="tbxVersionFE"
                                    StatusBarText ="Currently linked back-end database version"
                                    FontName ="Calibri"
                                    ConditionalFormat = Begin
                                        0x01000000a0000000010000000100000000000000000000001f00000001000000 ,
                                        0xff7d7d00ffffff00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x5b00740062007800560065007200730069006f006e00460045005d003c003e00 ,
                                        0x5b00740062007800560065007200730069006f006e00460045005d0000000000
                                    End

                                    LayoutCachedLeft =180
                                    LayoutCachedTop =4980
                                    LayoutCachedWidth =1692
                                    LayoutCachedHeight =5214
                                    ConditionalFormat14 = Begin
                                        0x010001000000010000000000000001000000ff7d7d00ffffff001e0000005b00 ,
                                        0x740062007800560065007200730069006f006e00460045005d003c003e005b00 ,
                                        0x740062007800560065007200730069006f006e00460045005d00000000000000 ,
                                        0x000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Locked = NotDefault
                                    CanGrow = NotDefault
                                    CanShrink = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =223
                                    TextAlign =3
                                    BackStyle =0
                                    Left =8520
                                    Top =4980
                                    Width =1512
                                    Height =234
                                    TabIndex =8
                                    ForeColor =9870754
                                    Name ="tbxVersionBE"
                                    StatusBarText ="Currently linked back-end database version"
                                    FontName ="Calibri"
                                    ConditionalFormat = Begin
                                        0x01000000bc000000010000000100000000000000000000002d00000001000000 ,
                                        0xff7d7d00ffffff00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x4e007a0028005b00740062007800560065007200730069006f006e0046004500 ,
                                        0x5d002c002700270029003c003e004e007a0028005b0074006200780056006500 ,
                                        0x7200730069006f006e00420045005d002c0027002700290000000000
                                    End

                                    LayoutCachedLeft =8520
                                    LayoutCachedTop =4980
                                    LayoutCachedWidth =10032
                                    LayoutCachedHeight =5214
                                    ConditionalFormat14 = Begin
                                        0x010001000000010000000000000001000000ff7d7d00ffffff002c0000004e00 ,
                                        0x7a0028005b00740062007800560065007200730069006f006e00460045005d00 ,
                                        0x2c002700270029003c003e004e007a0028005b00740062007800560065007200 ,
                                        0x730069006f006e00420045005d002c0027002700290000000000000000000000 ,
                                        0x0000000000000000000000
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =180
                                    Top =4800
                                    Width =907
                                    Height =210
                                    FontWeight =500
                                    ForeColor =9870754
                                    Name ="lblVersionFE"
                                    Caption ="FE Version"
                                    LayoutCachedLeft =180
                                    LayoutCachedTop =4800
                                    LayoutCachedWidth =1087
                                    LayoutCachedHeight =5010
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =9185
                                    Top =4800
                                    Width =847
                                    Height =210
                                    FontWeight =500
                                    ForeColor =9870754
                                    Name ="lblVersionBE"
                                    Caption ="BE Version"
                                    LayoutCachedLeft =9185
                                    LayoutCachedTop =4800
                                    LayoutCachedWidth =10032
                                    LayoutCachedHeight =5010
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =75
                            Top =2220
                            Width =10065
                            Height =3285
                            Name ="pagDefaults"
                            Caption =" Defaults"
                            LayoutCachedLeft =75
                            LayoutCachedTop =2220
                            LayoutCachedWidth =10140
                            LayoutCachedHeight =5505
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin OptionGroup
                                    BackStyle =1
                                    OverlapFlags =255
                                    Left =4920
                                    Top =2940
                                    Width =4980
                                    Height =1560
                                    BackColor =16709608
                                    BorderColor =10921638
                                    Name ="rctPaths"

                                    LayoutCachedLeft =4920
                                    LayoutCachedTop =2940
                                    LayoutCachedWidth =9900
                                    LayoutCachedHeight =4500
                                    BorderThemeColorIndex =1
                                    BorderShade =65.0
                                    Begin
                                        Begin Label
                                            BackStyle =1
                                            OverlapFlags =255
                                            TextAlign =2
                                            Left =5160
                                            Top =2820
                                            Width =1680
                                            Height =300
                                            FontWeight =600
                                            TopMargin =58
                                            Name ="lblPaths"
                                            Caption ="Directory Paths"
                                            LayoutCachedLeft =5160
                                            LayoutCachedTop =2820
                                            LayoutCachedWidth =6840
                                            LayoutCachedHeight =3120
                                            BackThemeColorIndex =1
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =120
                                    Top =2520
                                    Width =9840
                                    Height =2103
                                    Name ="boxDefaultPane"
                                    LayoutCachedLeft =120
                                    LayoutCachedTop =2520
                                    LayoutCachedWidth =9960
                                    LayoutCachedHeight =4623
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =8700
                                    Top =2400
                                    Width =900
                                    Height =309
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =1
                                    ForeColor =0
                                    Name ="btnChangeDefaults"
                                    Caption ="Change"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"
                                    ControlTipText ="Change the defaults"
                                    LeftPadding =60
                                    RightPadding =75
                                    BottomPadding =120

                                    LayoutCachedLeft =8700
                                    LayoutCachedTop =2400
                                    LayoutCachedWidth =9600
                                    LayoutCachedHeight =2709
                                    ForeThemeColorIndex =0
                                    UseTheme =1
                                    Gradient =12
                                    BackColor =0
                                    BackThemeColorIndex =0
                                    BorderThemeColorIndex =0
                                    HoverColor =65280
                                    PressedColor =0
                                    PressedThemeColorIndex =0
                                    PressedShade =80.0
                                    HoverForeColor =0
                                    HoverForeThemeColorIndex =0
                                    PressedForeColor =0
                                    PressedForeThemeColorIndex =0
                                    Shadow =-1
                                    QuickStyle =22
                                    QuickStyleMask =-1
                                    WebImagePaddingTop =1
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1650
                                    Top =3060
                                    Width =3000
                                    TabIndex =2
                                    Name ="cUser"
                                    ControlSource ="User_name"
                                    FontName ="Arial"

                                    LayoutCachedLeft =1650
                                    LayoutCachedTop =3060
                                    LayoutCachedWidth =4650
                                    LayoutCachedHeight =3300
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =930
                                            Top =3060
                                            Width =663
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            Name ="lblUser"
                                            Caption ="User"
                                            FontName ="Arial"
                                            LayoutCachedLeft =930
                                            LayoutCachedTop =3060
                                            LayoutCachedWidth =1593
                                            LayoutCachedHeight =3312
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1650
                                    Top =3735
                                    Width =1095
                                    TabIndex =3
                                    Name ="cPanel"
                                    ControlSource ="Panel"
                                    FontName ="Arial"

                                    LayoutCachedLeft =1650
                                    LayoutCachedTop =3735
                                    LayoutCachedWidth =2745
                                    LayoutCachedHeight =3975
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =930
                                            Top =3735
                                            Width =654
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            Name ="lblPanel"
                                            Caption ="Panel"
                                            FontName ="Arial"
                                            LayoutCachedLeft =930
                                            LayoutCachedTop =3735
                                            LayoutCachedWidth =1584
                                            LayoutCachedHeight =3987
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1650
                                    Top =4065
                                    Width =3000
                                    TabIndex =4
                                    Name ="cProtocol"
                                    ControlSource ="Protocol_Name"
                                    FontName ="Arial"

                                    LayoutCachedLeft =1650
                                    LayoutCachedTop =4065
                                    LayoutCachedWidth =4650
                                    LayoutCachedHeight =4305
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =810
                                            Top =4065
                                            Width =795
                                            Height =255
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            Name ="lblActivity"
                                            Caption ="Protocol"
                                            FontName ="Arial"
                                            LayoutCachedLeft =810
                                            LayoutCachedTop =4065
                                            LayoutCachedWidth =1605
                                            LayoutCachedHeight =4320
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =1680
                                    Top =4770
                                    ColumnWidth =2208
                                    TabIndex =5
                                    Name ="chkBackupOnStartup"
                                    ControlSource ="Backup_prompt_startup"
                                    StatusBarText ="Whether or not the application prompts for backups upon startup"

                                    LayoutCachedLeft =1680
                                    LayoutCachedTop =4770
                                    LayoutCachedWidth =1940
                                    LayoutCachedHeight =5010
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =1910
                                            Top =4740
                                            Width =2532
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            Name ="lblBackupOnStartup"
                                            Caption ="Prompt for backup on startup"
                                            FontName ="Arial"
                                            LayoutCachedLeft =1910
                                            LayoutCachedTop =4740
                                            LayoutCachedWidth =4442
                                            LayoutCachedHeight =4992
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =1680
                                    Top =5130
                                    ColumnWidth =1908
                                    TabIndex =6
                                    Name ="chkBackupOnExit"
                                    ControlSource ="Backup_prompt_exit"
                                    StatusBarText ="Whether or not the application prompts for backups upon exiting"

                                    LayoutCachedLeft =1680
                                    LayoutCachedTop =5130
                                    LayoutCachedWidth =1940
                                    LayoutCachedHeight =5370
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =1910
                                            Top =5100
                                            Width =2244
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            Name ="lblBackupOnExit"
                                            Caption ="Prompt for backup on exit"
                                            FontName ="Arial"
                                            LayoutCachedLeft =1910
                                            LayoutCachedTop =5100
                                            LayoutCachedWidth =4154
                                            LayoutCachedHeight =5352
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =5100
                                    Top =4772
                                    TabIndex =7
                                    Name ="chkCompactBEOnExit"
                                    ControlSource ="Compact_be_exit"
                                    StatusBarText ="Whether or not the application compacts the back-end db upon exiting"

                                    LayoutCachedLeft =5100
                                    LayoutCachedTop =4772
                                    LayoutCachedWidth =5360
                                    LayoutCachedHeight =5012
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5328
                                            Top =4740
                                            Width =2376
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            Name ="lblCompactBEOnExit"
                                            Caption ="Compact back-end on exit"
                                            FontName ="Arial"
                                            LayoutCachedLeft =5328
                                            LayoutCachedTop =4740
                                            LayoutCachedWidth =7704
                                            LayoutCachedHeight =4992
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =5100
                                    Top =5132
                                    TabIndex =8
                                    Name ="chkVerifyOnStartup"
                                    ControlSource ="Verify_links_startup"
                                    StatusBarText ="Whether or not the application verifies table connections upon startup"

                                    LayoutCachedLeft =5100
                                    LayoutCachedTop =5132
                                    LayoutCachedWidth =5360
                                    LayoutCachedHeight =5372
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =5328
                                            Top =5100
                                            Width =2376
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            Name ="lblVerifyOnStartup"
                                            Caption ="Verify table links on startup"
                                            FontName ="Arial"
                                            LayoutCachedLeft =5328
                                            LayoutCachedTop =5100
                                            LayoutCachedWidth =7704
                                            LayoutCachedHeight =5352
                                        End
                                    End
                                End
                                Begin ComboBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OldBorderStyle =0
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =1650
                                    Top =3390
                                    Width =3000
                                    Height =255
                                    TabIndex =9
                                    Name ="cEntry_Role"
                                    ControlSource ="Entry_Role"
                                    RowSourceType ="Table/Query"
                                    RowSource ="PRIMARY;SECONDARY;SINGLE"
                                    FontName ="Arial"

                                    LayoutCachedLeft =1650
                                    LayoutCachedTop =3390
                                    LayoutCachedWidth =4650
                                    LayoutCachedHeight =3645
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =240
                                            Top =3390
                                            Width =1365
                                            Height =255
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            Name ="lblProject"
                                            Caption ="Data Entry Role"
                                            FontName ="Arial"
                                            LayoutCachedLeft =240
                                            LayoutCachedTop =3390
                                            LayoutCachedWidth =1605
                                            LayoutCachedHeight =3645
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3555
                                    Top =3735
                                    Width =1095
                                    TabIndex =10
                                    Name ="cTimeframe"
                                    ControlSource ="Timeframe"
                                    FontName ="Arial"

                                    LayoutCachedLeft =3555
                                    LayoutCachedTop =3735
                                    LayoutCachedWidth =4650
                                    LayoutCachedHeight =3975
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =2820
                                            Top =3735
                                            Width =654
                                            Height =252
                                            FontSize =9
                                            FontWeight =700
                                            BackColor =11056034
                                            Name ="lblTimeframe"
                                            Caption ="Year"
                                            FontName ="Arial"
                                            LayoutCachedLeft =2820
                                            LayoutCachedTop =3735
                                            LayoutCachedWidth =3474
                                            LayoutCachedHeight =3987
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =5760
                                    Top =3360
                                    Width =3180
                                    Height =255
                                    TabIndex =11
                                    Name ="tbxRoot"
                                    ControlSource ="Root_Path"

                                    LayoutCachedLeft =5760
                                    LayoutCachedTop =3360
                                    LayoutCachedWidth =8940
                                    LayoutCachedHeight =3615
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =4980
                                            Top =3360
                                            Width =615
                                            Height =240
                                            FontWeight =700
                                            Name ="lblRoot"
                                            Caption ="Root"
                                            LayoutCachedLeft =4980
                                            LayoutCachedTop =3360
                                            LayoutCachedWidth =5595
                                            LayoutCachedHeight =3600
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =6720
                                    Top =3720
                                    Width =3060
                                    Height =255
                                    TabIndex =12
                                    Name ="tbxData"
                                    ControlSource ="Data_Path"

                                    LayoutCachedLeft =6720
                                    LayoutCachedTop =3720
                                    LayoutCachedWidth =9780
                                    LayoutCachedHeight =3975
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =4980
                                            Top =3720
                                            Width =615
                                            Height =240
                                            FontWeight =700
                                            Name ="lblData"
                                            Caption ="Data"
                                            LayoutCachedLeft =4980
                                            LayoutCachedTop =3720
                                            LayoutCachedWidth =5595
                                            LayoutCachedHeight =3960
                                        End
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =6720
                                    Top =4140
                                    Width =3060
                                    Height =255
                                    TabIndex =13
                                    Name ="tbxPhoto"
                                    ControlSource ="Photo_Path"

                                    LayoutCachedLeft =6720
                                    LayoutCachedTop =4140
                                    LayoutCachedWidth =9780
                                    LayoutCachedHeight =4395
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =4980
                                            Top =4140
                                            Width =615
                                            Height =240
                                            FontWeight =700
                                            Name ="lblPhoto"
                                            Caption ="Photo"
                                            LayoutCachedLeft =4980
                                            LayoutCachedTop =4140
                                            LayoutCachedWidth =5595
                                            LayoutCachedHeight =4380
                                        End
                                    End
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5760
                                    Top =3720
                                    Width =780
                                    Height =300
                                    Name ="lblDataRoot"
                                    Caption ="Root"
                                    ControlTipText ="Root path of the data directory"
                                    LayoutCachedLeft =5760
                                    LayoutCachedTop =3720
                                    LayoutCachedWidth =6540
                                    LayoutCachedHeight =4020
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =5760
                                    Top =4140
                                    Width =780
                                    Height =300
                                    Name ="lblPhotoRoot"
                                    Caption ="Root"
                                    ControlTipText ="Root path of the photo directory"
                                    LayoutCachedLeft =5760
                                    LayoutCachedTop =4140
                                    LayoutCachedWidth =6540
                                    LayoutCachedHeight =4440
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =75
                            Top =2220
                            Width =10065
                            Height =3285
                            Name ="pagAbout"
                            Caption =" About"
                            LayoutCachedLeft =75
                            LayoutCachedTop =2220
                            LayoutCachedWidth =10140
                            LayoutCachedHeight =5505
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =7725
                                    Top =2550
                                    Width =1980
                                    Height =324
                                    FontSize =9
                                    FontWeight =700
                                    Name ="btnReleaseHistory"
                                    Caption ="View release history"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"

                                    LayoutCachedLeft =7725
                                    LayoutCachedTop =2550
                                    LayoutCachedWidth =9705
                                    LayoutCachedHeight =2874
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =7725
                                    Top =3030
                                    Width =1980
                                    Height =324
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =1
                                    Name ="btnReportBug"
                                    Caption ="Report a bug"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"

                                    LayoutCachedLeft =7725
                                    LayoutCachedTop =3030
                                    LayoutCachedWidth =9705
                                    LayoutCachedHeight =3354
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =7725
                                    Top =3510
                                    Width =1980
                                    Height =594
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =2
                                    Name ="btnViewMetadata"
                                    Caption ="View DB Metadata/Purpose"
                                    OnClick ="[Event Procedure]"
                                    FontName ="Arial"

                                    LayoutCachedLeft =7725
                                    LayoutCachedTop =3510
                                    LayoutCachedWidth =9705
                                    LayoutCachedHeight =4104
                                    WebImagePaddingLeft =2
                                    WebImagePaddingTop =2
                                    WebImagePaddingRight =1
                                    WebImagePaddingBottom =1
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =375
                                    Top =2550
                                    Width =2820
                                    Height =270
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =3
                                    Name ="tbxVersion"
                                    ControlSource ="Release_ID"
                                    FontName ="Arial"

                                    LayoutCachedLeft =375
                                    LayoutCachedTop =2550
                                    LayoutCachedWidth =3195
                                    LayoutCachedHeight =2820
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =735
                                    Top =2910
                                    Width =2280
                                    Height =270
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =4
                                    BackColor =11056034
                                    Name ="tbxAuthorName"
                                    FontName ="Arial"

                                    LayoutCachedLeft =735
                                    LayoutCachedTop =2910
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =3180
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =735
                                    Top =3230
                                    Width =2280
                                    Height =270
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =5
                                    BackColor =11056034
                                    Name ="tbxAuthorOrg"
                                    FontName ="Arial"

                                    LayoutCachedLeft =735
                                    LayoutCachedTop =3230
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =3500
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    SpecialEffect =0
                                    OverlapFlags =247
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =735
                                    Top =3550
                                    Width =2280
                                    Height =270
                                    FontSize =9
                                    FontWeight =700
                                    TabIndex =6
                                    BackColor =11056034
                                    Name ="tbxAuthorPhone"
                                    FontName ="Arial"

                                    LayoutCachedLeft =735
                                    LayoutCachedTop =3550
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =3820
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =735
                                    Top =3870
                                    Width =2280
                                    Height =285
                                    FontSize =9
                                    FontWeight =700
                                    BackColor =11056034
                                    Name ="lblAuthorEmail"
                                    FontName ="Arial"
                                    LayoutCachedLeft =735
                                    LayoutCachedTop =3870
                                    LayoutCachedWidth =3015
                                    LayoutCachedHeight =4155
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =223
                    Left =240
                    Top =180
                    Width =6360
                    Height =780
                    FontSize =12
                    FontWeight =700
                    ForeColor =16777215
                    Name ="lblNetwork"
                    Caption ="[I&&M Network Name]\015\012Inventory and Monitoring Program"
                    FontName ="Tahoma"
                    OnDblClick ="[Event Procedure]"
                    ShortcutMenuBar ="Double-click to open website"
                    LayoutCachedLeft =240
                    LayoutCachedTop =180
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =960
                End
                Begin Image
                    Left =8400
                    Top =60
                    Width =1560
                    Height =1980
                    Name ="imgNPS_Logo"
                    PictureData = Begin
                        0x0e00000000000000010000006c00000000000000000000005200000071000000 ,
                        0x0000000000000000550a00006f0d000020454d46000001002033010011000000 ,
                        0x0100000000000000000000000000000000050000000400009a01000036010000 ,
                        0x00000000000000000000000090410600f0ba0400460000001c99000010990000 ,
                        0x474449430100008000030000ca227e1000000000f89800000100090000037c4c ,
                        0x000000004e4c00000000050000000c0282006400040000000301080005000000 ,
                        0x0b0200000000050000000c028200640005000000070103000000050000000902 ,
                        0x00000000050000000102ffffff004e4c0000430f2000cc000000820064000000 ,
                        0x0000820064000000000028000000640000008200000001001800000000005898 ,
                        0x0000130b0000130b000000000000000000001d1d1d0d0d0d1010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x10101010100e0e0e0e0e0e101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010100d0d0d1d1d1d1010 ,
                        0x1002020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020101010000000102000e140f0c120d000000030101020202 ,
                        0x0202020201030202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202010302020200020202030102020202020202020202 ,
                        0x03010402020002020303030002020201030100020604030d181025422b3a6141 ,
                        0x355d412b47331421190002030102000201030002020202020202020203010202 ,
                        0x0200020302010302030102020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020203 ,
                        0x0303010101020202010101040202000202030204000202030100030100080f0a ,
                        0x203a2a3b5c413d6045395d453b5e443c5f4439594019271b0200000201030101 ,
                        0x0103030301010105030302020202020201010105030302020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202000301020202040202020202000101020202040202000202020202020202 ,
                        0x0000000604031f32233c6448427150486c4e446a4c3f634b43664b3f65494069 ,
                        0x4a40614615221a01010102010302020203030303030303010005030202020200 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202010101030303010101010402000301 ,
                        0x0101010202020100040001002236292d4f3740684c476e4e446d4d466f4f4870 ,
                        0x544a6c4e3f654740664a4463483f684c396047131f1302000300020001040201 ,
                        0x0101050204010200030402020103020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202010101030303 ,
                        0x0002020202020601020101010103030500010305051f36273c624445714d436e ,
                        0x4d4170504975564b76554b7556477050476a4f45644943664c3c60483f65493b ,
                        0x5c41182319000100030101000202000202030204030002020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020002020101010303030300020401030200000207062642 ,
                        0x2b436f52446f4e446e4f4974534777534e7b5a4b7c5c4e7d5c4d795a4b735742 ,
                        0x684a45684d4063494063483e62443c5d42152317020000020000020202040103 ,
                        0x0302040003010202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020002020303030202020002 ,
                        0x020003010b140a29453245714d456b4d436c4d4471504a77564c7b5b4f806057 ,
                        0x8a65548664588969527f5e4c7556476d4f456b4d45684d4063483e64463b5e43 ,
                        0x253b29080e090200000003010004020201030202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202021010101010 ,
                        0x1002020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020300 ,
                        0x020303030003020202020200001523183f6149426e4a436d4e4770514872534d ,
                        0x7857507c5d527b5f517d5e53815d538664508260547f5e5781624c7657497253 ,
                        0x466c4e40664a3f63453863423763442949311422170202020401030402020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020402020202020203010402 ,
                        0x01020202010303020202040103000202040300000001131a1333563b3760413d ,
                        0x604543664c466c504d73574d73574c725650765a547e5f517d5e588767598665 ,
                        0x5787635787634e7d5c4b7758496f513f68493e64463d60453b61433a63433862 ,
                        0x431b302104010301010102020200020202020200030101010104010301020002 ,
                        0x0103020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0204020202020200020204020204020202020204020202030102020200010011 ,
                        0x1c14385e423e6446456b4f4e725450765a51795d537c60517b5c4e78594f795a ,
                        0x527c5d517d5e547f5e57846357835f5281605883624f78594770514167493a5f ,
                        0x45375a4034573c2f5839315a3a33593b172a1b04040400010104010301040202 ,
                        0x0103040103000200000202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020201010102030101030304010305030302020202 ,
                        0x0301020301020103020000171f18426d4c3c62443f6849466c504871524c7556 ,
                        0x4e78594f785c517b5c507a5b4f795a527c5d547f5e5183615281605183615582 ,
                        0x61547f5e527b5c47705143684e3e6146395e4434573c3056382b55323056380c ,
                        0x130e050002030303010101010101030303010101020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202030103040202 ,
                        0x0202010101020301000202010303030101000003030303253f2d447150446c50 ,
                        0x477051466f50477051497155476f53456d51476d51476d51466c50496f514a73 ,
                        0x574b775a4e7d5d4e805e527f5e4d7f5d517c5b4c78594a7354426b4c39624338 ,
                        0x573c2b50363154392c55361f3927020001020202020202010402000101040103 ,
                        0x0101010100020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202040202030402000301000202000003060907172519 ,
                        0x324a383d6f4d426f4e41714d447752477251416d4e406a4b3e664a3d65493f65 ,
                        0x493f674b3f674b40694a40694a446a4c486e504a73544a735346785649785747 ,
                        0x76554571524671503d6744395f4130553b2b53372d5337264f300f2013040202 ,
                        0x0202020002020401030202020202020303030202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202040202020103020103020202000301000101 ,
                        0x040103020001192b1e416c4b3f6a493f68483e6a4b41724c447352497c574371 ,
                        0x4d426d4c3f68484067473e6446396243376041376041345d3d335e3d3d634542 ,
                        0x664842684a4670514675554976554678563d6a49376a44366642365c3e30593a ,
                        0x30563a2e57382d5133101e130201030401030401030303030002020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020103 ,
                        0x0201030203010002020502040203010a150b33593d3b64443862433963443862 ,
                        0x433c65463e6a4d41724c3e754e41764f3e734c3e714c396c4735684335684336 ,
                        0x66425d80664c76573766453c68443b68473e68493b6d4b416e4d3d704a658d71 ,
                        0x5e8a6b3a6d483e6948345d3d305a3b2d55392b57332c4f35080e090400010303 ,
                        0x0301010102030100020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202021010101010 ,
                        0x1002020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202000301020202040202020301000202020003060e0734513a3962 ,
                        0x43375f43375d3f375a3f355f403762413761423a6d483d754c3c7149366b4435 ,
                        0x6a432e633c2f633e31643f2c5c38bed1c2a5c2ab3367423a6c443a6f483b6b47 ,
                        0x366846396645325f3e99b7a4c3d9c736664233603f316343336640326140305a ,
                        0x3b2f593621372505050500010000020004000501010102020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020002030002020203010202020201 ,
                        0x030000000a0d0b35603f3966453a69483b70493c6a463c684436614037674337 ,
                        0x6a45306740326b45376d4834674131613d2e5d3c2c5b3a295c36cbdcd1f0f7f4 ,
                        0x3c6f49346d46386b453a6f483669433a6642305e3a517e5dd0ded23e69482f5c ,
                        0x3b33613d315f3b305c3f345a3e2b59352e5738243e270c1c1104040400020004 ,
                        0x0202020202020202020202020202020202020202010101000202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020402 ,
                        0x020202020202020102060003010400000a13092f5e3d366642386a483d704b3c ,
                        0x724d396a4436694736694431634145765031613d2d603b2e5e3a2a5632204d2c ,
                        0x1746255a8662f8fbf9fffffe84aa8e245c312d653c366b44346d463267403367 ,
                        0x42386943c2d2c75c82662a53372e5738325e3a335f3b305f3e2e5b3a2d55392f ,
                        0x5a392c4f35080b09020001020103020202020202020202010101020202010101 ,
                        0x0303030202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202050204010101050302040000040202010604243d2932 ,
                        0x603c35633f396847396e47396e473c74492e67402358315f8567e9efea5a7d63 ,
                        0x1b4c26214c2b275031466f50708975dce3defdfffffefefed7e4dc2e633c225a ,
                        0x2f2761382f65403265402a633c2c613adbe5d993ad9b25512c2e5c382e5e3a30 ,
                        0x613b33613d2f593a2a5a362e5c382e57381e3925040705020000020202020202 ,
                        0x0202020202020303030101010202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020200020001030302020202 ,
                        0x0202020202213d2933603f346140346440386943386d46386c47346d466f977b ,
                        0x9db8a4d9e6deffffffd4e5d8486e52537c5cbcd5c1fafffdfffefffffefdffff ,
                        0xfffffffefefefe98b59e99b79e78a0842d603b1b5b2d2f6540a3bfacf3fbf453 ,
                        0x7e5d2756302e57372d593a2d59352b54352c59382d5a39315c3b2b54352e5737 ,
                        0x1b3720060a050202020303030202020101010202020303030202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020402010402020202020000010c130c285131325e3a335e3d305b3a33623c ,
                        0x34623e30643f2f623cbcd2c0fefffdfffefffffffffefdfff6f5f7fafcfcffff ,
                        0xfefffffffffffffffffefefefefffffefefefefffffefffefffffeffcbd8d094 ,
                        0xad99ccd9cbfffefec5dacb4076533163412a56372a53372f56362b54352b5133 ,
                        0x294f332a55342955362851312d593519271c0000000202020402020303030202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020103010303020301040003172a192e5e3a ,
                        0x2e5a3b2f5c3b325b3c2e5d3c3162422b603944754fddede2fdfffefffffefdff ,
                        0xfefdfffefefefefffefefffffffffffffefefefbfffffdfffffefdfffdffffff ,
                        0xfdfffffffefefffdfffffeffffffffffffffffff96b69da2c2af598862285c37 ,
                        0x2957332a53342b51332c4a31274d31274d2f295135274d2f2b56351526190000 ,
                        0x0002020203010104020202020203030302020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202010101 ,
                        0x0202020a0e09244326295c362d593a2a54352b55362e5737315d3e2a56327499 ,
                        0x7ffffffffffffffffffffffffffffdfdfffffffefefefffffffefffdfdffffff ,
                        0xfefffdfffffdfffffffffffffefffffffffdfffffcfffdfefdfffffffcffffff ,
                        0xc5d7ca70987c50805c2860372d623b2a5333284b30254b2d294f33264a2c2649 ,
                        0x2e2b4e33275231152a1b00030100010003030302020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202021010101010 ,
                        0x1002020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202040202030101080f0c1c40282b55362759372d593a2a55342852 ,
                        0x332d5636285430275030afc6b7ffffffccd7cdf2f8f3fffefffffffffbfdfdff ,
                        0xfffffefefefffffffffefefffffffffffffffffffffffffffffefffffffffeff ,
                        0xfdfffefffdfffcfefefbfffffafbf978a2835a83632654302a59392e58352754 ,
                        0x332a5334234c2d254b2f284e30264f30284e32264f30161d1800020002020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202040202020202 ,
                        0x0003010202020402020203010002020402020607050b1b101c3a2126522e2852 ,
                        0x33295a3a2d5e38305c382d56362a4d33224b2c345d3ddae2dbfffffcbccdc0eb ,
                        0xf0eefdfffffffdfffffdfffcfefefffffffffefffdfffefffffffdffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffefdfffffffefffffffeb5c8b96089 ,
                        0x692251302c58392b58372f5d392a5a362b5e392f5a392d5637295232294c312a ,
                        0x5236284e32132416040103010002000301010303010002040202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020402020202020202020203010202020402020202020202020a11 ,
                        0x0c26462e244d2d285132275433285c3730603c31643e2c5d372b573325553132 ,
                        0x6442e0e7e2fbfffeffffffcbd9cd819e85f4fbf8fffffffffffefffefffcfefe ,
                        0xfdfffcfffefffffefffffffffffffefffffffffefffffefffefdfffffffffffe ,
                        0xfffffffcfffffff4fbf859796123512d2953342b56352b563532633d3067402f ,
                        0x603a2b5434275132264f2f285132264f30284e30112215080104020202050002 ,
                        0x0303030003010202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020402020202020201030202020202 ,
                        0x02020202020300000000162d1e214926224b2623492d2752312e5a36305a372d ,
                        0x613c2e613c2e5e3a2652332c5732cfdad2fcfefeffffffa8bbac3f6248f5fbf6 ,
                        0xfefefefdfffffffdfffffefffffffffffefdfffffefefefefffefffffeffffff ,
                        0xfffefefefdfffffdfffffdfffffffefffdfffea4bcaa2b5635254e2e274d2f29 ,
                        0x56352756352e5a3530603c34674232623e2f5d392b5434284b30294f31264e32 ,
                        0x265130111b0f0101010202020401030202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0204020102010303020402020200030101020005050517351c22452b23462b27 ,
                        0x4a2f2a50322c5a362e613b2e613c2e633c2f5e382d5b3721502f5f8569e7eeeb ,
                        0xffffffe2e9e2cbd9cdfffefffffffffcfffdfbfffffbfffffdfffffdfffffdff ,
                        0xfffefffdfffffefefffdfdfffffdfffffffffffffffffdfffedbe3dc7a96822c ,
                        0x5b3a2556302752312c56372952332a5334305b3a32633d31613d34644031613d ,
                        0x2c5637274b2d214a2b264a2c254b2d1e3f2a0408030200000202020203010202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202020002020402010002020300020202020202020300020b ,
                        0x190e224a2721442a1f462c244a2c294c312a52362f5c3b2e5738305a3b2e5839 ,
                        0x325e3a2e5d37214a2a426b4c93af9b9ebaa386a18dc4d4c9fdfffffffeffffff ,
                        0xfffefffdfbfffefffefffffffffffffffffefffffefff8fbf9dce7dfa5bba977 ,
                        0x947d5c826631643e21512d275031274e2e274d2f274b2d2750302b5133315b3c ,
                        0x33613d32603c35644333633f2c56372a50322a4d33244a2c254b2d2c56371117 ,
                        0x1200000002030102010302020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020201010102020200030102 ,
                        0x03010103030300020606061c39221f432520412623492d224b2c294f332a4b30 ,
                        0x2e54362e57372d56372c55392e58352b5c362a5435214c2b1f48281f45291b42 ,
                        0x222f5636c8d6cbfffdfffffdfffffefffffffffbfffefdfffffdfffedce7dfa0 ,
                        0xb5a6597960285132244f2e1d48271e412621412820442c20472d234729244a2e ,
                        0x224c2d28502d2750312954332d5335315c3b30633e35633f2f5936294f332a4e ,
                        0x30264e32214b2c274c2c09140c00000002020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202010302020200030101010100000007180b1e4426234c2d224b2b ,
                        0x204d2c244c30264c30274a2f274d312b51332e54382d5b37295135254b2d254b ,
                        0x2f24472c244c30294c312952331e462a2d50367f9884b2c3b6b7c2b8b7c5bab6 ,
                        0xc1b99cb2a078957e33593d1842231841222b54352651302b5a392e5c38285233 ,
                        0x29533426522e254e2e1f4829234c2d22422923462b2954332a5334294f31264c ,
                        0x2e2c52342b5635294f3121472b21452721472b224d2c1e4224070d0802000304 ,
                        0x0103020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202021010101010 ,
                        0x1002020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202020003010201030401030202020604040e190f ,
                        0x1b412522452a1e4428204027214429244a2e234c2d274b2d234c2d284f2f274b ,
                        0x2d26492f224b2c28482f23472921472b24462720492a20462a24472d22482a19 ,
                        0x3f21123b1c123d1c143c20133b1f193f211840241e462a23492d254b2d254b2d ,
                        0x22482a254b2d20492a264a2c25492b24472c2243282043291e41272041261d43 ,
                        0x251e41261f402520432925462b244929234c2d23492b244b2b2043281e442621 ,
                        0x402b20492912291a050505020301020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x0202040202030303000101020202030204020103020202020301020103040202 ,
                        0x000202020103132916214b2c1a43231f41231f43251d40261d43252047272446 ,
                        0x2821442923492d22452b244a2c21472922482c2144291f42271e412622452a22 ,
                        0x46281e442821442922452a21442922452a22482c23462b22452a25482d1f4626 ,
                        0x21442923492b21472b2347292046282449292046282047272249292043291e44 ,
                        0x282043281f40251b4123203e251e3e251f442a20462a20472723482825492b1f ,
                        0x45292241262043281c42261b42281f43251e3e25060907010200020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202020202010101030402010402010101020202 ,
                        0x0402020402020202020202020001020301011b3b221b44251d40261e40221b3d ,
                        0x251a3d221c40221e41261e3e252144292043282043292046281e44262043281d ,
                        0x40251d40252043281c3f241f43251d40251d40252142272144292043281f4227 ,
                        0x1f422722452a2145272142272147292043292345271f422724452a1e422a2547 ,
                        0x281c43291d45291f42271e41261c41271d40261e42241e42242044261e412620 ,
                        0x43281c41271e43291c3f242043281c45261c42241f41231d40251941251d4026 ,
                        0x0508060003010202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202000101040202 ,
                        0x0202020103030303030503030101010402020202020402020504060e1f141a40 ,
                        0x241c3f25183d231c42241b3e231c42241b3e231c3e201e40221c3f241c3f241e ,
                        0x43231b3e231d40251d40251d40251f40251c3d221e3f241c3d221f40251f3f26 ,
                        0x2040272041261d40251e41261f42271e41261d40262042241e41261b41251c42 ,
                        0x261c43291b442522452a1f42281e43231f41231f3f261c411f1b3e241e412619 ,
                        0x3f211a3d222041261e40221e40221d41231b3f21193d1f18391e173a1f204328 ,
                        0x1e47281c3d2217311a162c19080b090503030202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x0202020202020101010402020201030201030301010203010004030002030203 ,
                        0x010502040f22111d43271f40251d41231f4025173f231e4224183d231a3d221b ,
                        0x3d1f193b23183a221c3d221b3e241f42271b3e232041261d3e231d3e231d3d24 ,
                        0x1d3e231c3d221d3d241d3d241c3c231c3c231e3f241e3f241c3f241d3e231c40 ,
                        0x281c3f241f3e231e3f241c3f241d41231d43271f41231d40251a3d231d3d241c ,
                        0x3f24193f211f40251c3b201a3d221c3922183b20113a1b123b1c153b1f1c3f25 ,
                        0x22482c325b3c345d3e16311d142b1c0c19110201030405030403050101010202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020202020101010202020202020503 ,
                        0x020202020001020400050304020c1c111d42281b3f211d3d241b3d1f1e40221c ,
                        0x3c231b3e23193f231a3c1e1c3b20183a221a3c241b3c211c3d221b3b221d3e23 ,
                        0x1c3d221d3e231c3c231c3c231c3c231c3c231d3b221b3b221a3a211c3c231d3d ,
                        0x241a3b201a3a211d3d241d3e232042241c3e261d3d241a3d231b3e231e3d221a ,
                        0x3d221b3b1c18391e18391e1b3a1d193b1d1b3b221c3d22183b21193f232c4c33 ,
                        0x45634a5b79606c8b70849e86a8b8a7d0d8cec5c9c320211f0001000001000100 ,
                        0x0200010001020002020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020202 ,
                        0x020303030402020202020002030001010402010603000e1b0d1c41211c3b1e1c ,
                        0x3c23183b201d3d25173d1f1d3b221939201a3a211939211b3821193a1f18391e ,
                        0x1a3b201939201939201c3c231b3c211b3c211c3c231c3c231c3c231b3b221d3b ,
                        0x221a3a211b3b221c3c231c3c231a3a211b3b221c3c231b3c211b3b221a3b2018 ,
                        0x391e183a1b1638191c3f1d1b3e241a3f2b1f43321e42322146361e42341a3f25 ,
                        0x1a3b201a3d225164517484739aa297bdbdb7cac5c2d0cbc8d5d0cfdbd6d5e4e1 ,
                        0xdc8f8d8c00000000030106010302020200010102010302020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100202020202020202020202 ,
                        0x0202020202020202020202020202020203000202010300010205030303010002 ,
                        0x090420403b2145391b3a1d193a1f1b3a1f163b211a3a22193c211b3d1f193920 ,
                        0x1c3a211a391e1a3b20163a1c16391e1a3a221939201939201d3e231c3d221e3f ,
                        0x241e3e251d3d241c3c231c3d221a3a211b3b221b3b221b3b221b3b221c3c231a ,
                        0x3a211b3d1f1a3b201d3f27234642294a532d4f662b5774315b78365c7f3b5f87 ,
                        0x385c8a355a92385991224a3f1b3a1f1d3b221b3c211e3e2527442d2f4b343b55 ,
                        0x3e455c466378628396839aa99ba5a9a32f2a2b00000003000200010002020203 ,
                        0x0002020202010101020202030303020202020202020202010101020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202021010101010 ,
                        0x1002020202020202020202020202020202020202020202030100020202010301 ,
                        0x0200010500060201000004172a322848792847741e3d3419371e1c3a2118381f ,
                        0x1c3a211a3a21183b211c3b20193921183b21183a1c1c3d2f21403918391e1a3a ,
                        0x211a3b201a3b201e3f241f40251f3f261d3d241f3d241b3c211f40251d40261b ,
                        0x3b221c3d221c3c231c3c23183b21193c211a3b263a5b823d63993f5f943d60a0 ,
                        0x3d66a4476aac456bab4169a44068a23f659f385b8d34546b1a41271e44281b40 ,
                        0x261b3e241a3b2018391e12351b0f341a14371d193a1f193e24203f2a1624180d ,
                        0x1e11101e120e201309110a030402020202010101030303030303020202020202 ,
                        0x0303030101010202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x02020401030201030202020300020503030000000b0e1628457227497e28457e ,
                        0x25486a1c402a193a1f18391e1b372318391e193c221c3d22153b1f1b38211638 ,
                        0x191f3d3e23445719381b183b2123442f264a3c1c452522482a22422923462b1d ,
                        0x3f2120412c2d5958224a2e1f42282c533323492b214a2a1c3c2319381b254637 ,
                        0x3c619343669e3f6399456795476fa04a6fa34b6ea04c6c9d476d9d446ca0456d ,
                        0xa83e669b2b554e2547282446281f40251d3e231c3b26193c211a391e193a1f18 ,
                        0x381f16382017361b1c3c23173d21193c211e3d221a372014321918381f1c3b2c ,
                        0x1126170400000002020602010002030202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202000301040201000201000203010200050400 ,
                        0x192c4d2a4f872b518b2b4e8d294c90204151193f211b3a1d19361c203a2d1b3e ,
                        0x241c4526193d2719371e17391a203f403052881c44321c3f1d255654365b932a ,
                        0x4e5622482a1f45271e44261f3f202d51513c5c91274c4425572d28522f1b4123 ,
                        0x23492d1d40251d43272c5b53406b9c456aa44670a5476c9e486ca2486fa34d73 ,
                        0xa94e76ab486fa34b6c9e47679c416698366174245430264f2f1c44281d3d241f ,
                        0x3c221b3a1f2044341c40301b381e1c3c2320473817371e1b39201a3b2016391f ,
                        0x1a391e1b4026284c5c2a506e1832220307020201030102000203010002020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202020202020202 ,
                        0x0002030202020200000e131c2d4b7a30559132578f30578e2f5595325184234a ,
                        0x3a1a3e202043392745561a41211e44261b39201b392017391a20423b35589c2c ,
                        0x51731f4e3835648a3e64a4355c6b1d49242547281d3d241c4127416a8b466dab ,
                        0x355a5e265430275231244d2d1f442a1a3e202347393762894069a04464994366 ,
                        0x9e40689d476da7456da2426a9b44699b446596486da940669c41669e3d5e9029 ,
                        0x574b244b2b1f452923492d1b422819391a2149441f42381837182042312c4e65 ,
                        0x1c3f24294f67244b4d1b3c21183c1e2748582a4d8c3252872d516312201f0000 ,
                        0x0001030405020404010302020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x020202020202030303040103030402010101040202172d4935558a3152833257 ,
                        0x8f33589233589232578f2b4c5b1f401e2e53672b4b741e44261d40251a3b201c ,
                        0x38211a391a1d433d31599a2f558b3058753862a53f64a0365d73254d31214325 ,
                        0x1e42242b524a4a6fa34b75aa3e637d284b31234d2a254a301e40281b3a1d254f ,
                        0x443b628f3d61973e65993a60963f64983b5d923b60943c5f973b60983d649b3a ,
                        0x659e3a629c3b619b3b609e3363751f4724193b1c2a505029494f193a18356179 ,
                        0x305668163514264b47375b912b4a4b2f53812f527a1a3a2725483e2f53833354 ,
                        0x852e4b78314e7526406410161b02000000030102020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x020202020202020202020202020202020202020202010206040202040100131f ,
                        0x312e4f812d497834538a34588e375a9a325a95325e9b305172264d3e3b65a02f ,
                        0x57881e43291d4424193c2118391e1a3a1b22484a2f5798325990315690345a90 ,
                        0x3c6298426baa335a631a391c1c43293e66834769a43f6ba1406aa52e5553234f ,
                        0x2b1f442a1c3c231c3e1f2d565f38598a3d5d92375f94375f90375b9135598f32 ,
                        0x588e36599130538b305b9432578f355999365b9737609e335b8c2142331f3b1d ,
                        0x3258782d536b284f3f436ba03b5e8a1b43272f5b683a60a0335983365c92385a ,
                        0x95264c58274f54375c9636578933527f334b6f33527922365508040902030102 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020002 ,
                        0x0303000202000004070c233a602546772e4d822d5088325793325892305b9a34 ,
                        0x5e9f2f557f2e5778355fa02f53831d442b1d42221c3d2216391e16371c26464c ,
                        0x2e539130519035548b2e57883758903a609a365b8d1a402a234b393d689b3c65 ,
                        0xa33c649e3862a3345d84244d321d46271d3d1e1f462d32567e345796395b9137 ,
                        0x5a9234558d33538e3351882d4d822b50882f4d882a4d852a4d8c2f518c2f548e ,
                        0x304f8634548f2448501c4022355e8b35598137566f44659d466ba3395e7a3c64 ,
                        0x953e61993d5f943b5d8b375889385987395d8b3b5d92396094345281334f7836 ,
                        0x4f77314f781c2943030505050001020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100202020202020202020202 ,
                        0x020202020202020202020402010301000302000f192a22437128457228497a2a ,
                        0x4f872e508b32538b31568e3456913456923156922d54922b49801e3f311d4123 ,
                        0x1a3d221939211a371d2747542648832c4b802c4e83314f862f4f802b4f853553 ,
                        0x8a2549492e556b33559030569030528e3153892e50852449411a3b201b3a1923 ,
                        0x454430558f3154863554873151863053853352892d4c7f2a4f89294c84284a86 ,
                        0x294b802348802b497a2a4b832b4a812e51892b517422474533558333588c3559 ,
                        0x893d5e8b3c5e8c3f60923a5f91375d93365a883859863b5e903b5d93385e943a ,
                        0x5c9136578531537e2c48712a446c28466f2d4c790d1721030100020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202021010101010 ,
                        0x100202020202020202020202020202020202020202020202020101010301001b ,
                        0x31542443762342752646812c4c872c4f8e2d52902f53932e51902c518b2b4b86 ,
                        0x284b832649811f433d1e40211a4024173b1d1a391e1f404f27467d27467d2847 ,
                        0x7e29477829467928497a2d4e802c476c2e4f81294b813150832c4f812b48812a ,
                        0x4c812549591a3a2119381b2349552f538930508136578f365a962f568d32568c ,
                        0x304f862f4f8a2b4d832a497c27467b27467b26457c27467b284b8327457e2e4c ,
                        0x832f4e852b4f85315284305182344f8133548235538232538136548534578f31 ,
                        0x5688395b91385c92395b9038598a3352853353882c4c752d45732a436f2f4978 ,
                        0x131c2a0000010202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x02020202020501000002021d30562846772142702242732642782546782b4980 ,
                        0x2b467f25487a27477c2745762342752340731e3d3a1f42201a3b20153a201839 ,
                        0x1e25415223457b25437a22457d23447c25447726457a24447528457828497b23 ,
                        0x477d27457c25477d294b812446812647791f3f2c2242372c50802d4e8633558b ,
                        0x34568b33558b33558a2f50823350832e4f7d2c4d7f28467d2540782442712341 ,
                        0x7024427326457c2548802a4b832b498227487a2b4c7a29497e2d4d7e2e4c7b29 ,
                        0x47762c49752b4d822d4d823454892e508532598d34538a35598f31548c2e5085 ,
                        0x2c4b7e2b4470273e6b264370172c4807090a0501000102000503030001010203 ,
                        0x0104020100020304020202020202020202020210101010101002020202020202 ,
                        0x02020202020202020202020202020101010402010101011c31572241761f3d6e ,
                        0x223d701f407125447924447926477f23437e23417a21437f1e41802342791b3a ,
                        0x391b40201c3c231b382118371c203c4d263f7724427324416e27406a243f7126 ,
                        0x4179233f7525426f29416f28417323437424407625417024467c2648841e4641 ,
                        0x244b5a294c842c508631538938568f315486345489365789345684304f822b4b ,
                        0x8028487926477523427924457d23427f244378224273234475254880294b8627 ,
                        0x47822b498226487e27467928447328427128467527457629487b2d4c832d4d82 ,
                        0x3052882e4f872e4d80284a802d4573264271224371213f7028406e1c2a470305 ,
                        0x0602010500020202010302010300020204020204020102020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202040202030101 ,
                        0x00030226395e28477a1a3970183b6d234073214176243e6d1f3c4323423b2243 ,
                        0x461f4246203f401e413d1a3b26193d1f17371e1b39201a391c1d3a49233f7b22 ,
                        0x3c601a392a1a3a291b3e3420415125447723488224468223437e25447b254576 ,
                        0x294a82294b8725488823456323486e2a467c314f862f507e3252833452813654 ,
                        0x853153813355833453862f517f2b4c7d27457e24437a26437c22427326437629 ,
                        0x47782a45772744772444792748802749852b4b8625457a26437624406f264174 ,
                        0x25437a26477925477d2c4a81294c842a4c8727497f26487d24467b27416f233f ,
                        0x6e1f3c691d3c692642780a131d03010104020104020203030300020201010103 ,
                        0x0303020202020202020202101010101010020202020202020202020202020202 ,
                        0x020202020202040202000203000001818489bdc7ce989fa8707d93576581596c ,
                        0x87526685324e701e44461c43231e41262143241e411f1c3a21193920193a1f19 ,
                        0x3a1f1a391e1e3b4a203d691b3a3319381b1a391e1a381b1537191e3e43223f6c ,
                        0x25458028477e25417725407223457b23447c254580294885294b872849812b4d ,
                        0x882b508431538833559034568b345583335182284c7c2d4b7a26457825447725 ,
                        0x40722943722742741e3c6b1e3e69243f7123406c203c6b263e72224277254477 ,
                        0x27457c22417425417729437227427425407322427324406f2442731f3c6f1e3d ,
                        0x70264473234275203f7225406c1a36651a2b4c0f1d30060c1302000005030202 ,
                        0x0202040201020202010304010101020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020003010301010f141d455f8e5574 ,
                        0xab899aafdad8d8faf6f1fffafcdbdad663786f2445371e412622452a1e41271b ,
                        0x3e231b3b221d3a231b3920193a1f1939201738231839241c3b201b40261c4022 ,
                        0x203f22183b21163b1b18372e25416323457b23417c254477203f72213d6c243c ,
                        0x6a2144762948852a4c882d4f85294c8b2a4d8528477c2a4c7731507d2e47712c ,
                        0x446e26466f28497729487b2b4a7f25477d24468122437b26447b223e741f3c6f ,
                        0x203b6d253c6a223c6a2442732544791f3f70234172243f722243712341722641 ,
                        0x73253c6922385c203961223b65253b6b2540731e3c6d1e40768194b79da0a800 ,
                        0x0000040404030303000203000203040201020202020103050303020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020002 ,
                        0x030200011a283a3f68ad3660a539608c3b5751607965758b7947644a1a3d2227 ,
                        0x4b2d274e2e295030254c2c1d4129183e221a3b201b39201939201939201a3b20 ,
                        0x1c42242144291d48271c41271e43291b3e241c3b1e1d3b1e1b39261f3c4a233f ,
                        0x5e224174234277204075243f7124427326406f29426e2543722843752b487b2a ,
                        0x487f28497b27497f2d4f7d2d4a7d2b487b24447527487a22437526457c234476 ,
                        0x234073213d6c213f6e213f70234170203e6d26406f233e76243f77254073233e ,
                        0x6a243c70203a681c355f152e5a1c355d223b651f3c69253f6e1e3c6b14366c44 ,
                        0x6194acb8cafcfbf7dfdddc353f50141e2f050302000000000203040201030204 ,
                        0x0101010402020202020202020202021010101010100202020202020202020202 ,
                        0x02020202020202020202030204000000263a5d426eb5446eb9355e7e18401717 ,
                        0x431c183e201c45252750312c58342f5839294f33254b2d2144291c3c231a3a21 ,
                        0x1d3b22173a201939201d3e231b47281e4525193f231d40251b3f211b3c211c3c ,
                        0x231c3c231d3c211d3c1f1d3d25243a5d273f73223f72273f6d233f6e24427323 ,
                        0x427724417423417222417623437429467338548324427327457624437826457a ,
                        0x24457723427722407126417325426f24406f21406d213b692239692e45722140 ,
                        0x6d213b6a233a6a243d67263a631f3a66183365394d768897aab6bfc93c4f741e ,
                        0x345e13305c3a5177959fb0e1e0e2fdfaf5e3e2de5d718a334d7b334f78121a27 ,
                        0x0200000202020203010100020002020202020202020202020202021010101010 ,
                        0x10020202020202020202020202020202020202020202050100020202304d8642 ,
                        0x6fc03b6cba3d67b2345f8634617c39618437669a2b5561264b2b2e5436264f30 ,
                        0x21442a1e3f241b39201d3c1f1b3821193920193a1f1b38213153353157391b3b ,
                        0x2216381a203e251e412621442a1d41231d43272147291d3f211e3e2d1d39621e ,
                        0x396c243e6c1f3e6b233d6b233e70223e6d233e6a253f6e234172253c698193aa ,
                        0x5b6e9115346923426f25407225407322406f23417227427425407222406f1e39 ,
                        0x65162d5b586a89adb5bc1f3d66233c68233d6b223c6a1c3b6e223b657e8aa6ea ,
                        0xeff2fdffffd1d1d11d375c2c45718b9cb7eef1f6fffefffdffffefeeea6a7993 ,
                        0x2b4a7d354e7a38568531486e090e110502000103040302000202020003010202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x02020001010a060b3454853d63a33e68b5426cb73d6cbd456ebd4271c73d70b9 ,
                        0x2f594e21472921472922452a1e3f241f41292245371c3b20183b211c3b20173a ,
                        0x2013391b96ac99eef5f0dbe3dc708b771d4123284c2e26503128513227533428 ,
                        0x4c2e274b2d1b4425475c71364d7d1633601e37631a34621d39681f3d6e223f6c ,
                        0x273e6e234073203d6a3c5478b0bbc953678a173566263f71203f6c2541702540 ,
                        0x7220416f233f6e233e701937666a7c99f0f3f1708196223a68273f69223c6a23 ,
                        0x3d6c17325e92a0b6fffffffefdfffffffba9b0b98d98acd5d9def8f7f9fdffff ,
                        0xfdffffdfe1e16978922e4b7836538037548138517b35507c334b6f0a0d110200 ,
                        0x0102030100020302030102020202020202020210101010101002020202020202 ,
                        0x020202020202020202020202020201010106080934598b4068b03f68b13d6ab4 ,
                        0x3d6bb93d6cbd3d6ec437699d24562c1f4f2b2c5c32254f2c1d421a20463a3d71 ,
                        0xb323473b17381d1b39201a3a2115392129472e627564a2b2a1b5c2b435553c24 ,
                        0x472c284e32285131264e32274d2f21452d163b1bbecdbfe9eaf47b8aaa3c5577 ,
                        0x516686344f741e3b62203862274170203d6a2943711c3a71576e8edddde3566a ,
                        0x89213761203d6a233e70243d6f233f6e26406e1b3865597195fffdfdd8dadb77 ,
                        0x89a0919fb1253d611f3967163661617495fffdfdfffffeadbccce4e7efffffff ,
                        0xf9fbfcc4d1e1eff2f6ffffffe0dcdb576f933459933e5d9439588b395786344f ,
                        0x7b34517e3c5684151e2b00000103030302020204020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202040000121b2f ,
                        0x3d62a63d68b1406bb43e6cba426ec1406fc1406bc23f6cc739668b2b54392a52 ,
                        0x402d58552951501f432b2e596821463618371c1a3b2014371d1b3a1b183b2114 ,
                        0x381a183d1d2c4932274d31264f302d56372a51312b55362d5935204b2a183e22 ,
                        0xa2b3a6fffffefffbfff9f9f9fffdfff1f2f6b7c0ce324b77183562647798344c ,
                        0x76182f5f7485a0eef0f0d1d9e0233d651d3864233e6a243f6b223e6d27406a15 ,
                        0x315a8c9aadfbffffd2d4d5f9f7f775839519305d18346a394f79d4dadffdffff ,
                        0xd0d7e6bdcadafcfffbffffffe0e7f0e6ecf1fdffffd1d6d95d7aa6375fa03f6a ,
                        0xad4167a73c5d953a5b89395682375584324f7c212f4203010000010004020200 ,
                        0x0203020202020202020202101010101010020202020202020202020202020202 ,
                        0x020202010101000000283e624674c13d6bb93f6dba3f6bb23e65913e6486446f ,
                        0xae4170c13f6fc1406cb33d6aae436dba406aaf3e66a129515619381b1d3f211b ,
                        0x3a1f2241322a4f45173d2123492d2751322851322857372c59382b55322e5938 ,
                        0x295233274a3026492e1b3a1f2a482fcdd4cffffeffffffffc4cddba9b3c5dee7 ,
                        0xebdddfe93a527663728cbdc6cf3d567e7b899ffdfffffdffff66779218315d22 ,
                        0x3a641f3c691d3b6a223b65162f5b76849bfffffefefbf7ebecea3a4f6e17305c ,
                        0x41577b9caac6c4cfe38da3cca4b8d7fdfffffffffffffefefdfffff1efee9fa9 ,
                        0xba4f6fa03c67aa426cb1456aae4168ad3a619f38568739538137548136538029 ,
                        0x3e5d010002010002040103010101020202020202020202101010101010020202 ,
                        0x02020202020202020202020202020201010100020023365b406fb9426ec14371 ,
                        0xc9325c6f193813173c1a2549393264884470c4396688315e6939658e4372cd44 ,
                        0x73ce4173c62249391f482814351a2b565f4675a11b3b1c204126254c2c2a5834 ,
                        0x2e563a2d53372e57372e59382c5535295334224b2c193a1f15361ba4b7a4fdff ,
                        0xfffffffebdc5d61d39681932648794aeeeeff3657697b9c0cfe7ebf07b8ea3e0 ,
                        0xe3e8fffdffd4d7df25416a213a62243e6c253e6a2039652a4069c2cadbffffff ,
                        0xfffeff99a4ac122e5719315f879db994a9c94b6fab4a6ca7eff0f4fffefffffd ,
                        0xfff8f8f2b5bbc07086a94167a1476dad476fb0456bab4164a44266a63e639f3b ,
                        0x5c9439588b3855813d5a87263f61020404030101010101000302020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020101010000 ,
                        0x00293b64446fb8426db6416fc33b68b12d5c711d442a163613244b3b3455651c ,
                        0x3a2116371421432b3764864473b13f6d9c2046341b42221736152d55534f79ae ,
                        0x1d3e231e3b211f3f26214729274e2e2958372c5536284e301f4e2e234d2a1b41 ,
                        0x231b3f211e3c1f466548e4e8e3fefefeffffffced7e55d76981c3e73ced4e1fc ,
                        0xfefedbdee2d1d8e1eceff3edf0f8b0bbcfffffff73859c0a2857203660283e68 ,
                        0x14315e7c8ca3fffffffafffefffffeb0b9c63a50738690a8b0bacc95a8cd859a ,
                        0xbacad4e5fdfffffffffff2efea8ea2b442699d456da84d74b24c74af4873b644 ,
                        0x6eb1446dac3d639d3e609b39568339588d3b598a3856873e5b88151e2c000000 ,
                        0x0202020101010202020202020202021010101010100202020202020202020202 ,
                        0x020202020202020000000c0f143d62a04271b5446eb34170ba416fbd4474c638 ,
                        0x657325502b1b4522183a1c1c3e201c4022173a20193a1f1e422a203f2415381e ,
                        0x173a1f1737182c54534a73931a3d221a3b201b3b221b3e241e42241c42261d3e ,
                        0x2920422a20412627533c2654472448421c412d0f351f92a19dfffffbffffffff ,
                        0xfffcfffeffa4b1cb5f769cf6fbfafffefffbfafcd5dbe0c7ced78b9ebfbcc6d8 ,
                        0xf8f8f87b8ba82a4167132c56182e587f8fa6ffffffeff4f5c4cdd7e0e5eef6f8 ,
                        0xf9d0d7e0c3d0de9bb0d0fcfbfdfffefffcf9f4c6ccd17186a64b70ac4f7abd52 ,
                        0x7bb94f78b74b75b64d72b0486eae416aa9385b933c55813a5a8f39598a38588d ,
                        0x37548038528021344f0304020102000202020202020202020202021010101010 ,
                        0x1002020202020202020202020202020202020200000017243a4973b84875b946 ,
                        0x76be4473bd406fc14172c83a647b234f28204e241d4a291d44241e47281d4026 ,
                        0x1a361f15361b14381a18381f1e382015391b244433224531183b211d3e291c3b ,
                        0x1e163b211a3a211a3c1e1b3b221a3c1e19381b2a514f4975b44b76bf436ea743 ,
                        0x6d9a4771ac7993b8c8d1d5fdf8f9fefffdfffffeafb7d4839abadde3eeffffff ,
                        0xd7dce56b7e9fa7b1c2b9c5ddeef1f9fffeffd6dee55d75930f2a5c667895ffff ,
                        0xffb0bbd11b3668c8d0ddfdffffebeef2f0f3f7eef1f6ffffffe2e7e68b9bb24f ,
                        0x74a64e7fbd5480c05c82bc5682bf517dbc5079b74c73b1436aa83f629a3f5d8c ,
                        0x3c5a8936538039538137578838547d3a5784344d77070a0f0102000401030202 ,
                        0x0202020202020210101010101000020302020201040203000200020200010008 ,
                        0x06063356884c75be4673b64c76bb4976c04476c24273c33b6aa73363972f5d75 ,
                        0x254e2e214a2a173f231c3d22193920193a1f1b3a2517392116391e1a3a211837 ,
                        0x161e3d2e3b607c3a596e1d3d24193920193c221b3f211739211b39201a391e1c ,
                        0x3c23426886517dbd4f79be5079c24a77bb426fb3466da47a8fafc1c7cefcf7f4 ,
                        0xffffffe9ecf4becce3e3e7f2fffefffbfdfecad2dfeceff7fffffefffefffbff ,
                        0xfffffefe95a2b875889dfdffffd8dee9b2bfcffffffefffefffffffffefdffff ,
                        0xffffc8caca617a9c517bc05282c45883bc5f87c25985c25783c05984c3547cbd ,
                        0x527abb4971ac4a70b13e639f3a5b883d5a863856873956833b5d92395a92354f ,
                        0x7d0a0f1205010001000202020202020202020210101010101003010002010302 ,
                        0x02020201030403000001001725414674b44b75bc4f79ba4d77ba4a76bd4979c1 ,
                        0x4775c24473c53f70d64173d73059621a391e17371e18381f19381d335e7f3966 ,
                        0x881b3e2418381f16391e193d25476e94597cbe365967193b1d18391e193b2319 ,
                        0x3c21193c221c4022193a1f16381a325864517ab9527cb74e78b94f7bc2507cbc ,
                        0x4d7abe4373bb4b73ae7792adc5cbd0d5d9ded3d8dbd8dadbe5e4e0d8dcd7f0f1 ,
                        0xeffffefdfffffefffffffefefefffffefdfcfff5f7f8fefffdfffffffffefdfa ,
                        0xfffffbfffffffefffdfaf2b7c2c65779a74d75b05581c15d88c16086c05f87c1 ,
                        0x608acb5a89c75884c35982c1517dba517bbc4e78b94572b64873bc4570bf3c67 ,
                        0xb04267a54164a83c5e9931496d02030703010104020102020202020202020210 ,
                        0x10100f100e0304020103040101010002020000010d11163554874a78c64976b9 ,
                        0x4e7aba547ec14b77b6487cbf497dd04a7bd7466fad3560752145341a371d193c ,
                        0x22173a1f1c391f234c45284e481b3b2319392119361f1e432f507aa4355c6417 ,
                        0x37181537191a3d231f42271c45261b44251f432520422416341b2c4f4b527dbc ,
                        0x507dc05480c04e7fc35282c4517ec14f7ec24e7dc14b77c44e77b6557eb5567c ,
                        0xb65a7db55378b05176b06586b4bdc3cafdfdf7fbfffffffffffdfcfefdfffffe ,
                        0xfefefffffffcfefefffffffefefeffffffe0e2e2788ba64f73b34b79b95380c3 ,
                        0x547fc25b87c75989c95b87c65d86c45c86c15581c05784c8527fc25280c04f7e ,
                        0xc84c7bcd4977be4371bf436ebd3e6ab7395fa73f619c1a283b00000102020204 ,
                        0x0201020202020202020202101010100f11010101040202040201000000090a0e ,
                        0x3250814369aa3f6aa94c76bd4e79b8507dc1507dc6507bba3e6787385c6e2343 ,
                        0x3015351622463a1c41311a391e1b3c211c392219381d16371c17371e19392019 ,
                        0x3b1d163a1c1838191634112e5458365e711d41231b41232145271d43271a3f25 ,
                        0x19381726463b486b974c77b6527ebe507abf527fc35481c5507dc0527ec5557f ,
                        0xc45180c44e80c84e7cc34e81ca4d7bc2497ccc497ece4377c44c74b58e9db7ec ,
                        0xe9e1fffffffdfffffffefdfcfefefefffdfffdfffbfbfffffefffefefea6b2be ,
                        0x3a6cb44676c44b78bb527ac25079b75280c75986c35686c65482c25680c35680 ,
                        0xc55582c65081c74c7dc74c78c54878c04778c44172be4071bd3e6bb43a58873d ,
                        0x5b8c15212d00000003030300020302020202020202020210101010110f000203 ,
                        0x0202020003010301010d13183d5f9a3e63a1456fb44974b74972b0537bbc5581 ,
                        0xc8446c881b40181e481f1e47271b3e232a50502f596c193b1d1838201942231c ,
                        0x47261a40241f44241d3d241c382118371c1c3c293e677e5585d347749a1d4724 ,
                        0x244e2b1f4b261e47221e44202042373962894d75c04b78bc4e79c24d7cc64e7f ,
                        0xc5517cc5537fc65380c95180c45083cc5082ca5380c3527ec55183cf5281cb54 ,
                        0x81cb4f83c64f7ec84774bd6883a5bcc4cbe9eae0f6f7f5fffffefffffffefdff ,
                        0xfcfefefffffcc8c9d35877aa416bb64373bb4b77c44b78bc4d7cc6507bbe4d81 ,
                        0xc3527cc14c7cbe507dc0527dc04e7bbf4f78c14b7abe4a7cc84474c24471ba45 ,
                        0x6eb73e6cb93f68b13a598c3d5b8c162031020000040202030204020202020202 ,
                        0x020202101010101010040202040202020301000001090f16365c9d4367a7446b ,
                        0xaf4771b44979bb4e79b85480c74f7faf2d5e4221532b264f2f204d2c193d1f1c ,
                        0x3e261c39221c39221c3b201d3e231f4325264b2b194528183a1b2345345478a0 ,
                        0x5f90d65989d12a4c4517391a1d4025244f443765703a668f406eae4671ba4574 ,
                        0xbe4476be4676be4b76bf4976bf4c7ac14d7abe4d7ac3507dba517ec2537fc652 ,
                        0x80c75481c55082ce5382c64e80c8507fc9507dc05082ca4776ba4b75bc5778b0 ,
                        0x768eb2d7d8dcfffffffdfffffffffbb8bfc24b71ab3966b7416bb84172be4172 ,
                        0xbe4578c84976c04774bd4d77be4977be4a75b44a77ba4a77ba4b78c24878c047 ,
                        0x74be446eb14068a33f69aa3f67af4068a93e67a63b5b9034547f121d25020000 ,
                        0x0202020003010202020202020202021010101010100101010201030000001316 ,
                        0x1b3c506f3c5b903d629c3e67a5436db04673b74e78bb517db35983c4406d7b22 ,
                        0x4e2722502c214b2c193f2316391f1a3a21203d231d3f211a3c241a3c241d3e23 ,
                        0x1c402218371c335757618dca507fb237606319391a18371c325569416aaf416d ,
                        0xba416cbb426cb7426ebb426db64371be4370b44771b64473bd4774be4a73bc4a ,
                        0x77c04974bd4b79c04c79bd4f79be507cc34f7fc7507dc0527ebe507fc34f7cc5 ,
                        0x4d7dc54e7cca4876c33e6fb93a6bbb4f71a7b2bbc8d9d7cd95a1b3416aa93c69 ,
                        0xb33f6dbb3d6ebe416bb6406ebc3f6cb54570bf4272c04670b54775bc4774bd48 ,
                        0x72b74873b64372bc4973c0446cb43861983c60963c629c3a63a23e68af3f6aad ,
                        0x3d619d395a88192334030303020000000202020202020202020202101010120f ,
                        0x11000202020202000000282d367291c8395e9a395e963d63a3416dad4a71b54a ,
                        0x74b9456e952449352a4f3f1c44281e41261d42221d3b1e1f452726502d225130 ,
                        0x22472d2140251d40261f4025203e25193c2222442c2c5546274b3d1938171938 ,
                        0x1b284a494368a0416eb7436ab8416bb03e6cb33e6bb5446cb73e6db7416eb841 ,
                        0x6bb6426fb9416eb2436db24372b64471b54673bd4770b54875b94a75b84a77ba ,
                        0x4b79c0497ac64875b84d78c14a76bd4c77c04477bd4273bf3f6ebf396ec53f70 ,
                        0xba496eb23c68b53c6bbd3e6cba3d6cb63d6dbf416cb5426dbc406bba3f6fb742 ,
                        0x6eb53f6dba416db4416dba3e6ebc416ebf3e69b23a5f993b5c8a3855823a5784 ,
                        0x3a5b933c619b395fa03a63a23a5f993c598c39527e35465b1f25300200000202 ,
                        0x0202020202020210101010110f0100020304020201002a2d32799cd43e64ac39 ,
                        0x5b963b5c943c619f436dae4c74bf3a607e1737181536141d40261c3c231f4337 ,
                        0x27504b274b3526502d2c5736254b2f254b2d25492b27482d2043291d40252140 ,
                        0x2b436c8d557ebd395d6d1d3b22294c423f6196416fb63e69a83c6baf416db43d ,
                        0x6ab43e6bb43f6bb2436ab83f6ab3406aad3e69ac3966aa3963b03a64af3764a7 ,
                        0x416bb0416bac3d67ae3e6ab13c6db93d6bb84270be4472b93e70bc3d67b24071 ,
                        0xc14070c23566b43868b03d6ab43c6ab84069b82e65b63065bc376bbe416cbd3d ,
                        0x6cbd3167bc3566bc416ec53f69b43d69b03e66b13162ae2f5a9d345b923b5b90 ,
                        0x324f7b284977314d7c304a78345184385c923c5d953859913a5a8b3b5c8a3957 ,
                        0x883c557d606e853e4443000000020202000202101010100f1104020200010200 ,
                        0x000334373b8ba8d53c66ab385a8f3e5d943e64a54270b74771b2476c9e3c5e76 ,
                        0x3c657e4573a24570a14a6d9833585c204732234826264c302f593a2e5938305b ,
                        0x3a2a5a36244c30203f2421452f406a7d426e8b39616625472f2147232f544c33 ,
                        0x57573858634367ad4068b03f6cb63d6cbd426cb73f6aa93e69b83e68af486aa6 ,
                        0x8ea6caa7bcdca8bee1809bc73b6ab45278b997b2de94adcf96b0d59db7e5809f ,
                        0xd24675b988a6d791acde4472b94776c09cafd5809fd2396ebe3e68b34f79be97 ,
                        0xb1e09cb5df5882c73968ba4a75be98b3df87a3d93e69b23969b74b75ba7f9bca ,
                        0xa6bcd8adbce391a5c84668a47891b99aabc690a2bf8fa1be879ab54f6892314e ,
                        0x7a35548135538233537e3d5d9237588643587e778699373a4200000002020210 ,
                        0x10100d0f0f05030201010100000127221f6b7c974670b73b65aa3a61983d629e ,
                        0x4063a54365a0456da84e77c04f7ccd4c79bd4876ac325d661f46201f45272043 ,
                        0x29274b2d2d583738634230563a30593a2d5a39244d2e20432821462c25472923 ,
                        0x47292e54383864402b4d2f264724294c3e3d648b4a70b14270b73e6ebc4168ad ,
                        0x4067ac3f69b03b60a44567a2839dcb6b86b8758ebae9eff67f9ed3587cb8fdff ,
                        0xfed0ddf3678aca6e8bbe6182ba4872bddbe5f6ebeff43a6abc94b2e3fcfefe8c ,
                        0xa7da376ac03464b6a3bbdffffffefaf9fdc2d4f13368bf4e76befbfdfddee6f3 ,
                        0x3364b0799ad2fcfdffdfe9f38aa4d27392c5819dc64d72b0c9d7edfbfdfe7993 ,
                        0xb7647ea663799c445e8c36547d37537c37507c36517d38578436527b334d7233 ,
                        0x4c747781931313130000011010101010100301010503020001010000013d5170 ,
                        0x4d7ac33f6ec03e64a43d61a13c619f3e64a44168ac4771b64677bd3461761e43 ,
                        0x23204a27214828274f2c2755312e59382f5b3c2d56362e5c382e5c382a543127 ,
                        0x50301f4228244929315739315838335a3a315a3b34573d345a3c2b4e2c2c523c ,
                        0x2f5647325b463d66973e66b13e65a93c65a43f65a6395fa02f599c5679b99eb1 ,
                        0xdcfcfbfd8199c35278aefcfffdced5e95e83bd658dc84f76ba446db6dce3f2f2 ,
                        0xf8fd9ab6e5ebf3fae5edf44b75c23b65b2446fb8e9effaccd6eeaec3dffdffff ,
                        0x597dbd476fb7fbfbfbdbe3f43766bac0cfe9fefdff7196d02f5db13866b33965 ,
                        0xb23963aec7d4eafefcfc7389ac5f769c5e759b39527c34517e3753823a568538 ,
                        0x528035527936538034547d37537c8794aa1414140000001010100f0e12000202 ,
                        0x0202020000002a2b29697fa24166aa3a62aa3d68ab3f65a541639e3f639f3f65 ,
                        0xa6406aad456fb4335a69244e3d264e43234b32294f3d2c5b412c5a302b5d3534 ,
                        0x6245355e4226513026492f20482c21442a20472d34583a35633f3861413d6945 ,
                        0x3b6445375e3e355b3d2750312345271a3f1f2d515b436bb63c69b23f6bab3d64 ,
                        0xa23762a1a0b9d9f9fefdfdfffedfe8f54167a75a7bb3fcfefefefdfff9fdfefe ,
                        0xfffa9db0d53661a4d6deeffffffef2f2f8fcfbfdcad8ef4974bd3362b67494c9 ,
                        0xfcfefe8eafdd6f90cffffffe688dc9436fbcfafdffdae1f03b68b1d2dcedf7f9 ,
                        0xfa5681c43567b93f6ebf3e6dbe3364aec3d0e6fffffff8fafafefdf9e0e3f13d ,
                        0x5c8939527e35517a3653863859873b568239538132517e3d57854d5a70141315 ,
                        0x0102000f0f0f1010100002020201030200002225296884b33a64a93e68ad3c68 ,
                        0xaf3d67ae3c65aa4066a74269ae4069ae406cb94370b94571b8436fbc3e68a342 ,
                        0x6db6386481355c65396388406ca73f6391284a391f41232042242747342e5349 ,
                        0x2b4f2b315d3835603f3a6344375f433a6344375d3f2f593620492a20411f264b ,
                        0x433e68ab406bba3d6ab43966b05c80c0fffeffeaeef96e94ce3e6cb9325da65e ,
                        0x80bbffffffc5d1e34a6dac4d71ad3f68a64169aadce5efe7f0f4426eb57999ce ,
                        0xffffff9fb4d42d57a2abc0dbfffffe6a8abf4670b7f2fafa6e91d04372c4fdfc ,
                        0xfedbe3f43668bbb0c5ebffffff8ba8d43264b63464b2375fa7365b97c5d0e4ff ,
                        0xfcfe617fb0436ba64d75bd3b64a93d5c933754813554813856853c5581374f7d ,
                        0x37548138517d485c7b5a5f68000000101010101010020202030101191b1b6171 ,
                        0x8846679f39609e3d639d3b5f9f3861a03d69b03f69ae3e6baf426cb7416fbd3e ,
                        0x6dc13d6dc54270bd3e67943965843258643053563b66633e656d325a55244830 ,
                        0x23432a2442291e4126234a312c52542d4f49325b3b3566403b6542325e3f2f61 ,
                        0x3729522c254a2a1f3e2131546e416bb2416bb23f6cb63a69b35179c1e9effaf1 ,
                        0xf4f99cb4d894aedd456cb05e82b8fffeffdee7f49db5d9a1b5d47992ba436aae ,
                        0xdae1f2fefdffa9bfe3d3ddeffffffe95abd44369a9eef0f8e2eaf1496fb03765 ,
                        0xb3c3d2eca5bde14576c6fbfcffdfe8f6356ac16289c7eef0fafffcffaec3df9b ,
                        0xb4dcbbc7e34d6fa4c5d1e3fffeffafc2e7a1bae29cb0d94973ba3e6dbe4064a0 ,
                        0x37547b36507f37558438548338547d3653804c658f3a43510000001010101210 ,
                        0x100002000000015760696c8cb73e5f9e41609f395e9c395fa03b64a33e64a541 ,
                        0x66b03e68b3426cb73a6cb8416cbd3c698e3257551e421e20482c355c64305747 ,
                        0x2c522e294e2828482926492e24482a22452b32585c26493b2f585b2a524d2249 ,
                        0x292a52362a5032285136315957305853325a6d3a65983e6ab13d6cb63d6dbb3e ,
                        0x6ebc436bbe3a6ab86a8ecebdd1eaeaf1fadae7f74a74c1547abbc4d5f0d2daeb ,
                        0xd1dae8dee6f7a2b7dd416bb6adc0e3d1deeed3e0f0cfdcf6a0b8e24772bb6184 ,
                        0xc4d6ddf09ab4d83963a83663ac7998c5b4c4e14675bfc8d6edb4c3e33a6bb539 ,
                        0x67bb6187c7b6c7e8e4ecf9eaf0f7bdccdf47649797acc8d0dceed4ddebd2d8dd ,
                        0xc2ceda567aba3b69b6406cb3385b8d3b548037507c37537c344f7b345485506a ,
                        0x92454d5a0200000e10101210100200011315167a91b14869a83c62a338609b3b ,
                        0x609e3d629c3c619d3f63a93e68af3e6bb43c6bbc406dbe406ebb335b671f4d22 ,
                        0x2b55322b583733623b335c3d365d4e305650304f3a264c3025482d1f45273969 ,
                        0x812b4d4c1b3a191c3e1f2246282041261f3d243055694069b83c6cb4406bba41 ,
                        0x6cbb426cb93e6bb53e6ebc3d6cbd406bba3f6cb63e6ab73967b53e6cba3e6dbf ,
                        0x3f6dbb3f70bc3d69b63c66b13b6cb83f6bb83e6dc14270be3d6bb93b6abc406b ,
                        0xbc3d6cbd3766ba396aba416dba3e68b33e68ad4168ad3c67aa3d66af426bb43e ,
                        0x68b33d6bb93c6cba4170c1406fc03c68bc3a69bb416dba3c609c395a8c37588a ,
                        0x38588d34538635527e365687395c9b4068a93d67aa3b64a94066a73c5b903755 ,
                        0x7e3a578337558436528157729767707902000010110f0c11100102000d0f1079 ,
                        0x93c2446bb53b62a73e65aa3c66a73c63a13d62a03e63a13a619f3c66ab406bba ,
                        0x3f6dbb3f6ebf406bb42d585528542f32593931573b3555423c697e4371a1365f ,
                        0x612b4a2d29482d2244263c6984305857204a271f43251f42271b3b231b40203c ,
                        0x6390436ab54069b23c69b33e6bb5406cb9416bb63f69b43f6cb6406cb93f6dba ,
                        0x3a6bb73d6bb83c6db93c6db93d6abb3d6cbd3f69b63c65ae3765b23d6ab43f6d ,
                        0xba3e6ebc3c6dbb3666b83769bc3c6dbb3c6cbe3d6bb93767b93a69ba406bba3a ,
                        0x64b13663ad3e68b33966af3864b13a6bb9386ec13868c03b6dc04170c1396aba ,
                        0x3565b33a619f3c5b923956833852803854833755843a5b8d375a923a5d954066 ,
                        0xa63e64a43b619b3b5b903755863a55873b5786375284577196575f6c00000012 ,
                        0x10101210100200011314128fa7d1426caf3a62a34166aa3d62a0395f953b6098 ,
                        0x3f65a63e68ad3f6cb5406ebc3f6dbb436fbc375d752d534d264c302f53352e55 ,
                        0x3b2f583834583a35593b2f54322c4e3023462b214227335c5f274b3d21452722 ,
                        0x4e2a2347291d412320432f3e66a14269b7406ab13f6bb23e69b23c6cb43e6bb5 ,
                        0x426cb3436bb33e6bb53c6db93e69b8406dbe396ab83f6ab94069b23a6db63d6b ,
                        0xb95f82c16f90cf4970ba3d68b7386abc507ac57697d65984c73869b93d6dbf3f ,
                        0x6ebf698ecc7095cf4a78c66d8fca678cc63f6ebf6188c67190c34b76c5557ec7 ,
                        0x7694d5507cc3366bc25786ca799ad26786c33a5a8f35527f3855823755843956 ,
                        0x823b56883d5b94365a90395a8c395f9939619b3a5c9139548639558437558439 ,
                        0x54863c59806973852b2b2b0c0e0f0f0f0f020000151313859fc43c66a93c64a5 ,
                        0x3c619f3b5d983b5d993a5d9d3b5ea03963a43f6ab33e6cb93d6bb8426ec1335b ,
                        0x771d3f1a264e32315c3b365f40395e44395c41345c40305638294c3223422d1e ,
                        0x42241f3c251c3a211739211b3a1d1e47271d3e231d3d2a3861884369b7426cb9 ,
                        0x3d67ac436aaf3f69b03e6db74068b33c67b03e6db73f6bb83c6bbc3b6dbf3e6f ,
                        0xbb3e6ab73c6db93e6bbc3c69bab8c9e4ffffff5e88c93566b63b69b75a81cbff ,
                        0xfcfdd0dff23d6cbd3867bb678dcefcfefed5e0f4527dc6e4eaf7e5ebf84170c6 ,
                        0xd0dceef7faff4977c594b2e1fdfffe7e9fd74c7ac7e7f0faffffff7493c6365e ,
                        0x9f3d5d923b56823754813a56853855823554813a598c3a598e3d5d8e3c5c913f ,
                        0x5d963b5a8f385685395683365281355384677c981e22230f0f0f0e1011000100 ,
                        0x0a090b6278a14169aa3a60a03c609c3c5e933c5e943c5d953d60983b609c3e65 ,
                        0xaa426cb73f6cb63f6dbb3e6baf29514526502d3159363663423e68493b634032 ,
                        0x583c2a50322246282041261a4024193d1f1a3a2118391e214741234933163c20 ,
                        0x16371c1e423139617e4069a8456bab426cb73e6bb5456cba4172be3d6cb6416b ,
                        0xb23f69b43f6bb8416eb8446eb5416cbb3a6cb8406bb43e6abebacbe6fdffff67 ,
                        0x8ed33066bb3364b43a69babecce8fffffedbe4f2d8e5f5e5ebf8fffeff8faddc ,
                        0x3c6ec0e5eef8deeaf63768bedae1f0b7c9e82d63c199b3e1fefefe809cdcb4c9 ,
                        0xe8fefdff9fb8e43665b63a67ab3c5e993a578a3a58873453803953813b55833a ,
                        0x55813852803455823a5b8c3c5b8e3a588735548136537f3450793756833a4d6e ,
                        0x34373c0e0f0d1010100000005f626a5871993e639d3a5f99395c943d5c913c5a ,
                        0x933c5e943d619d3d65a63d67ac416bb6416dba3f6abb3d6ba12e5a3d2d5c3636 ,
                        0x64403c6948335f3b2e55403f71ad3b5e781d42221e40211a391c315450254338 ,
                        0x1939203460712f5260193b1c1b41251b381f1838152f56644467ab4167a73f67 ,
                        0xaf416bae4069ae3d69b03d66ab3e65af3c69b3426bb4416aaf416bb6406ab73d ,
                        0x6ab33c68afb7c8ddfffeffebf1f8e1e8f19bb5da3566bc7898cdffffffd2ddf1 ,
                        0x99b2def6f9fef3f6fb517ec24671c2e2eafbf5f8fdccd7edf2f6fb688fcd2e60 ,
                        0xb895b2dffffffccedceffdfffec3ceea416bb63d6cb04265a54267a5426caf43 ,
                        0x6db43e649e3b5c8a3a57833a55813858893b58853956833b598a3b5a8f365483 ,
                        0x3a5581354f7d334f783d53768b949e110f0f0d100e0000005560745c7daf3c61 ,
                        0x9f3d619d385a95375a993b5e9d3b609e385e9f3b65a83f6bb23e6bb43d6eb43d ,
                        0x6bb83b6aa72a544225502b26523528522f23472f3b66913d6fa42a4f471f3e23 ,
                        0x193e2416351a3d6472355a5e18371a305a61254d421c3c23224b2c2146261c3f ,
                        0x25365d8a3f6db43c6bb54269b33b64a93d65a64265a73c62a23d66a54066a73e ,
                        0x66a73d68b13e6bb4416bb83d6ab4426cb7bacbe5ffffffa1bae2c5d4eefffffe ,
                        0x88a0d43866b3eaecf7ccdaec547bbff9fbfcb7c9e83766b04874c1e0e9f3f9fb ,
                        0xfcc9d7f3e5eff97fa0d23364b29ab0d9fcfdfff9fefffcfeff7791b932528d3b ,
                        0x5f953c619b3d60a23961a23e69a84368ac4164a3395e984064a04364923f5f90 ,
                        0x38588937568339568337527e38537f395683334e80445c865b67790c0e0e100d ,
                        0x0f2025267284ad3f639f3a6198375fa03c63a13b609c3d60a239619b3d66a53e ,
                        0x67ac3e68af416bb83e69b8416cbb3e6dbf3c67b02d5d57255128264f29234f2b ,
                        0x3258522144291a3a1b25493931534d133517446f84477089183514173b231c3e ,
                        0x261c3f242648291b44242a4e583e67ac3f67af4068b34168ac3a62a34065a341 ,
                        0x64a33e63a13c619d3b64a23a63a83b65a83e68ad4168ac4168ac3d66abb9cced ,
                        0xfcfbfd4672b97390c3ffffff9eb4d72f5c9fa0b0d5f4f8f9b4c5e0ffffff708f ,
                        0xbc32599d486fb3e4eaf7e0e9f73264b7d1ddefeaeef93b6bc396afe1ffffffa2 ,
                        0xb7d7eef3f6e1e9f04f6da63958953d5f953c5c913c60963e63a14166aa4169b1 ,
                        0x4268b04268a943619c3e5d903a588738578438537f38507e36517d34547d344e ,
                        0x7c3d5c833d4a600f0f0f0e0d0f7f899a5277af3c69ac3d66ab3e62a23e61a13c ,
                        0x649f3e629e3d5f9a3b62a03b63ab3b67ae3d69b63f6dba3f6cbd3d6ebe4170c6 ,
                        0x3963922c4e4d3259623a6481244f421e40211b3c212b544d3d6675173617497a ,
                        0x94507fab2644312040272140251e3e251d3d241e3d203054644164ad3e6aaa3d ,
                        0x6aae3f67af3e66ae3d68ab4266a63d629e3a5f9d3b5f9f3d629e4065a33c65a3 ,
                        0x3d67a83f6bb23a69babbcae4fffeffe3e9f6f2f6fbfaf7f96e88b62e56905c79 ,
                        0xacf9f9f9fffffedadfee42659d375a99466dabe2e8f5fffffedbe4edfffdfcc8 ,
                        0xd7f13968bc97b2e5fffffc7694c37c97cafdfffeced7e443639837568b395683 ,
                        0x3c57893a5e943e639d4166a23e66a73e60a23b5f9b3d5e8f39588534547d3b52 ,
                        0x8036527b35507c36538036517d3e5f8c2b35470c0e0e0c0e0f4950595074b042 ,
                        0x6db63d66af3e66ae4064a43d629a3c60963a60963d5d983d61a13e62a23963a6 ,
                        0x3f66aa3d6cb0406ab5416db43d6bb8406ab53f69b4456fbc2c5147204f212043 ,
                        0x21406e90436d901a39185681ac5580ab275041284f3628513125502f264c2e21 ,
                        0x44291d4327385f8c3963a63e61a53e62a23b669f3d60a23c62a33d64a23b609e ,
                        0x3b629941649c4164a33d64a83d67aa4369b13f69b07999cea6bdddaabde29cb6 ,
                        0xde6385c03a5d9f3c619b3e6096899ec4a8bad17694bd365d9b3c64a5426ab290 ,
                        0xaad2a4b9d4adc0e398b1dd5079be366abd6e91d0a6bbe15e85c93863ac88a3d5 ,
                        0xb2c5e66280b13554873858893858893b58843856853656873b55833c59853a58 ,
                        0x873455833957883c5d8f3a588939568336537f35527e365380435d854a57670f ,
                        0x0d0d0e0f0d12161b5472ab416eb83b69b03e68ab3c6cb43c67b03e65b042639b ,
                        0x375c903b60943a5c973f63a34368a63e66a73f69b03e6bb4416eb8456ec33f6f ,
                        0xc14070c2365f862e566238627f5284d0416d8c1838194c788f406b741b3d1e27 ,
                        0x50312a58342b54352c5236254c2c22462e3a5e8e426ab23e639f3f62a23f64a2 ,
                        0x4165ab4169a43b65a64164a64164a64063a73e62a23c68a83c69ac3c67a63e68 ,
                        0xab3a63ac3566b23762b13764ad3a67b14063a73a5f973f5d9632569231548c36 ,
                        0x5a963d63a33c64a53d64a83964ad3768b83664b13666b8396bbd3f6ebf3a6cbe ,
                        0x3968b93c6bbc3c6cc43565b33562ab3b61a23b60943c598c3a58893b598a3c57 ,
                        0x893856853d57863d588a3d5c8f3c5d95405e973a5f973b5b903b568937578239 ,
                        0x55843a59863e5275484e55100e0e100e0e1d23286586be3f68b14068b34368b2 ,
                        0x406bb43d6ab43d6bb83f6dbb3e67b03a5f9b3b5b8c3c5e9a405d9a385d954165 ,
                        0xa5456aae3f69b0406ebc416ebf3d6dbf4473c94775c34779b55181c933575714 ,
                        0x371521452d25462b23462b2c53332c55352c56372b5631284f2f274b3d3e64a4 ,
                        0x3f67af406aad3d67a83760a93763aa3e6aa73a67ab355fa6375a9c3761a23761 ,
                        0xa43861a64062a83e65aa3e6ab13560af3264b03f6dba3c6bbd3964b53b61a139 ,
                        0x62993f64983d5f9b345b9f3e65a94369aa3e67a63b65ac3360b13160b13564b6 ,
                        0x3c6cbe3b6dbf3367ba3565b73c6dbd4070be3367ba3362b63869b93e68ad345c ,
                        0xa4345a9a39598e3956833554812f517f3251883b5c9431579131599433569631 ,
                        0x538f385a8f3e598b3956833653803756834c60831213170e0e0e0e0e0e272c2b ,
                        0x7e9ccb3b68b13c6ab83f68b73f6cb63d6bb83e6ab13d6bab3d69b63e6cb93f67 ,
                        0xaf3b639e3c5b903a5c923b62a03f6ab3406cb93e6fbd416fbc4170c43d65a026 ,
                        0x4b412042242e51471d3c2718381f1c41211d4a292851322a53343158482d5539 ,
                        0x2e5535234828264f4a4069a83e64a43f66ab4b70aa92a8cc708fc23762ab5176 ,
                        0xb499acd295aed84e79b87c9ed493acd63f64a23761a2446aab94a7d292a6cf43 ,
                        0x6eb74f74b094a8c7708aaf3d5c93365691657cac9cb2d56387c33865ae385fa3 ,
                        0x6483b8a8b7d7b7c6e085a5d6456fbc406db77a99ce92add9476eb83963a87e9d ,
                        0xd0a2b9e67090c5486eae93a6cb8099c13e63a13a599036558c7f95b892a7c34b ,
                        0x6594738db19dafce9eb2d1a0b3ce5874a3345489365582365483385784566887 ,
                        0x2121270e0e0e0c0d113b3b3b90a3c83c64ac3c69b23e68b33f6dbb4269b73c68 ,
                        0xb53c63a73e68ad3d6ab33d6dbf3d6eba3d6bb8426ab23f66aa436dba406ebb40 ,
                        0x6fc03e6cc04270c83e6bae2a4b471f433217361917381d1a3a221c3e1f275030 ,
                        0x2a59332e5b4a4671b037605b2750301f4424385e81446fb8406cb33b68b25c81 ,
                        0xc5fbfdfea6bbe1345faea0b8dcfffefffffcfd6185c58ba8d5fefffd5c7eb948 ,
                        0x71af7896c5ffffffcbd9ec345d9c667fb1fefefeb4bfda365991375fa097add6 ,
                        0xfffffe7e98bc2d56958aa2ccffffffe8eef9d8e4f6fdffff809bce3361aebcc7 ,
                        0xe5eef4f9496cac5778b7f0f6fbfffeffb5c6e13c66b1d7e1f2f4f8f96287c54c ,
                        0x75ba6382b9f2f5faeaedf2415b8ac0cadcfdffffd7dbe6e5eaf36b87b02e5387 ,
                        0x3d5484385582385784495e8467707e0c0e0f0d0e0c0706024759764f73b33c66 ,
                        0xa93e6cb93968b23d6bb8426cb73d6ab33c62a24067b23b6ab43f6ab33c6ab842 ,
                        0x6cb93a69ba406dbe3e6ec03f6dc13b6ece3a67922b56532249391d442f1f3c25 ,
                        0x1a3b261a391e1a3a211a3d2220411f355c6b4471bb3b678c27512e22461a234b ,
                        0x39426aab436eb13f6ab3547abaf9fbfba0b6df4d7cc6ecf3fcfffefffbfdfe65 ,
                        0x8ac84e77c0eff1fbeef4ffe7eefdf3f7fcfffeff7b97c02f54986483b6fdffff ,
                        0xb2c2d93a5e9e31548c8ea1c4fcfffd7c8fb2456295f0f4f5e9eff4547ab03963 ,
                        0xa6b3c1d8dce4f13762a5b4c3ddeff4f74268a9acbfdafffffcffffffbdcae430 ,
                        0x5eac97add7ffffffedf3fae7f0f9eff2f7fefefe9aaec72b4a7dc7d0daebeef6 ,
                        0x42649f4267a13d5e96395b9038578a3a53853954863e5e8f4f648334373c100f ,
                        0x1102000044536d557fc63764a7375a923d5f9a3d6cbd3d6fc23d6ab43a6bb93f ,
                        0x629a3f69ac3c6bbf3f6ec0406cb3436cb1406eb53e6fbf3b6fc23f6eb838658b ,
                        0x33628e294c5019371a244c3a22462e183b201a371e1d3d2a3d5e78496ea2436e ,
                        0xb74171c33e6b9e3461832f5a81436aae436cb53b68b9547ec5f8fcfda7badfc7 ,
                        0xd8eddee5f4a5bde1fdffff6784bd325e9e9fb5d9d9e0f390abd7ffffffcfdbed ,
                        0x3f6cb03256926381b0fffffeb0bfd935518730549090a8d2ffffff7c94c25c76 ,
                        0xa4fffefebfcde4315ba0305392738cb4eef3f64c6ca1aec1d6eaedf591a6c6fb ,
                        0xfbfba2b7d6ecf0fbbccbde345c9d496ba1d8e0edeef4f98ba7caf2f4fceaeef9 ,
                        0x4f74b2365893c8cfe3e8edec2b486f2f4b743b59904263a23f629a3c5e933c58 ,
                        0x8e36547d384e711e2129110f0e19191957657b3f5e8b3c5b923e63a13e64a53d ,
                        0x66ab3f69b43c6cba3d6bb93f6dc13e6cb9406cb33c69b3436db43d6fb7446dbc ,
                        0x3e6ebc4173c6335b741e40212b514b24483a203c1e2e5662284f40224e291c45 ,
                        0x1f20482c436b9f476fba406bae3f66ab4168ad4069ae3f68ad3f68ad3f68ad3a ,
                        0x65a85679b1f2f5faeceff7fdfffe7a95c7839dcbffffff6685b8325ca15375b0 ,
                        0xd9e1ee97accbfffeff7a97be355898315ca55e83bdfffeffb1c1de315b9e3356 ,
                        0x988ea0c5fdffff7992b44e6c9bffffffd3dce6395b902e589b95aed6e8eff241 ,
                        0x639fb1c0dafbfdfef3f7fcd0d7e84d6ca1e8eef5bdcbde385c9c305da18ea8cc ,
                        0xffffff93aad7fcfdffa4badd305494375487c4cedfe7eef13657853d588a3858 ,
                        0x833654833a57843953823956833955844c668e2e333c0f0d0d3339406276993d ,
                        0x60983b5d983a5d9c3c619d3d66a54067ac3f6ab93f6eb83e6dbe3f66ab3c6db3 ,
                        0x406ab5426cb3406ab14170b44371b84775c3376486254c26244e2b2449291d45 ,
                        0x2c335f6c2347392448382c58511b401e2f5861426eb5436bb3426db63e6eb63f ,
                        0x69b43d66ab3e64a43f64a23b609e5476acf6f3f5fffeffcbdaea3b5f958ca0c3 ,
                        0xfffffe6581b0355a94476699f5f5fbffffffe5e8f047669d3d629e90aacfc4d5 ,
                        0xf0ffffffe0e7f0abbedf5776b38da5c9fffffe8194b7315284b1bdd9fffffeba ,
                        0xc8da9bacc7fafbffa2b7d72f5798afbed8fffffffefffd6f88b047669bf3f6fa ,
                        0xbcc9df3357933b5b9d5077aeedf2fbfafffdfafcfd638ccb3d67b43b5f95cbd4 ,
                        0xe2eeedef3952843755863657853c54823855823957863753823553825c769b3b ,
                        0x3e46373b406174953e598b3b5b963b5b963e629e4064a44065a13e66ae3e6db1 ,
                        0x406ebc4071c74072c4416eb84670b3446eb3446baf456db54475bb487bc4416e ,
                        0xa723502f2658362d5c352e5a362a4f2d1e40212f58613e7297224a2124494544 ,
                        0x71bb406bb4406ab13e68af446eb53f67af3f64a83c609c375a924c6a99bdc7d8 ,
                        0xc9d0df7185a42a4a7f788eb2ccd4e558729a335486425e94b2c1d4d7e4f48ba8 ,
                        0xcf365f9d4367a7aebad2cdd6dfc7d5eccbdaf4d4e0ec6382b57691c4d0d7ea6e ,
                        0x85b232578f45649b98adcddee6f3e8eef5bbc7d94b73b43665af92a8d1d3ddef ,
                        0xabbcd73e5e9349699eb8c5db94a8c13c5d954166a4406aaba4b8e1dae4f6a9bf ,
                        0xe34072c44070c23d6ab39bb3d1b3bfd139527c38547d39527e38517d3956893b ,
                        0x568239527e35527e5d719440434b292f364967903d5c913c5b92395f993b5e9d ,
                        0x3c5b983c619f4064a44165a54271b54474c24276c9477bc84c76c14c76b94e76 ,
                        0xb74974b34d77b84f79be517bc0325d5a264f2f2b5135254b2f25432a20402728 ,
                        0x4d3d315e62204420294e564468a8426cb14269ae3f68ad406aaf416caf3f66aa ,
                        0x4168ad436cb53d67aa385fa33b5e9d3c5d95416097365a903658933c63a7456a ,
                        0xa84063a33e63a740629e395e963d5f9b3d619d3a5e9a3960a43b60aa3b63ab3b ,
                        0x61a13d62a03b5e96395d93395c943c5d953b5d99345e993961963f619c375ea2 ,
                        0x3a64a53f67a83f68b13d68ab3e69b2426ab24064a04167a7406cab476eb34870 ,
                        0xb84476be4377c34776c84676c84477cd4077c84778ce4174c4416eb84167a13b ,
                        0x59883c54823e5b873c55813c557f38517b324f7b586f8f3d41462a32394f689a ,
                        0x3b5992385a8f3c5c973d619d3d629e3f5fa03a5f993b639e4771b84b76bf4a79 ,
                        0xc3497dd0507ec54e77b54f7bb85179b3517ab1537fc65381c8456f922750302e ,
                        0x5a352f5838294f33264c301f4325254c2c1f401e315972406db73d6bb23f6cb6 ,
                        0x436db8446db6406ab1416bb2416ab3406ab1426fb3426bb44167a83f69ac4168 ,
                        0xad3e69ac4167a73e62983e5d923a5e94405f943e61993d6098435e963b5d923a ,
                        0x5c973a629d3c619d3f5f943a5d953a5c923a5b8d3756893957863a5a8b395b91 ,
                        0x3c5e943b609a40609b3e61a03c629c4267ab4169b1436db23d6fb74771b64777 ,
                        0xbf476fb04772b14873b24c77ba527cc3507cbc4d7ec4507bc44d7dcb4c7ed14b ,
                        0x7ece4b7dcf4c7bcd4876bd4370b346679f3e5d923a588139527e385483304d79 ,
                        0x6075944c4e5640434b5b7bac395c9439588d3b5d8b3b5c8d3b5a913c5d8f3c59 ,
                        0x8c3a60964369a94d77ba4d7ec45585cd5a87d05885c8557fba5884c15780be59 ,
                        0x84c75b86c94e78a7244b2b25522b2c55352d5734284e3024472c21452722483c ,
                        0x4267a5436faf426aab426bb0426eb53d6ab3426db6426bb4436db4446eb5416f ,
                        0xb64972bb4973b84673b74973b6426dac456aa84368a6416aa9456aa6395f953c ,
                        0x5a8b3e59853c5c853b598a3c5a8b405b8d395a8b3b5a8d3c588e3e5887365483 ,
                        0x3a55813957883a56853a59863d5e903b5b903c5d8f3f6098405f9c4066a6436e ,
                        0xb1456fb2446fb84770b5496da3476ca05274a95576ae4e76a75678ad5378b057 ,
                        0x7fba5881bf5985c05b8ace5589d65486d24f83d04d7ec44c76bd456eac43659a ,
                        0x425f8c3758863c5a893a55876b81a54b51584a4c576781a93b5c943f60983d5e ,
                        0x903d5b8a3a55813855823655823e5c8d476dad4f77b2527dbc5883c25d89c85f ,
                        0x89c46186c25f88c65c87c65b84c25e85c35986c947719b3662632c57362c5833 ,
                        0x2c5b352e59381f4626274f4e4573c03f6fb74472bf3f69aa3f64a24167a8446b ,
                        0xb04970b44973b64773ba4b74b94874bb4774bd4573ba4775c24376c64977c542 ,
                        0x73bd4973b84771b24477bd4875be4c76c34875be466fb8446ca7436398415c8f ,
                        0x3d58843e57813d58843857843a54823555803856853755863753823c57833b5c ,
                        0x8d365a903f619c4466a23e629e4068a2466ca64c73aa4b71ab5072a85072a750 ,
                        0x76a65378aa5779a7587fb35d82b45c86bb5c85bc5e86c05b85c05a87ca5c88d5 ,
                        0x5382c6517bbc4a6dac3c5e8c38567f38537f38537f314f80687ea24c5257474b ,
                        0x5061779a33527f3755863b56883656873a5b893856853b5b8c3d6092496aa250 ,
                        0x76b0527fb85e87c55d8bcb608ccb6288c85c87c06087c55a86c15887c55c87ca ,
                        0x5881ca4f7cc6426c91335d5c2c514127502a2450212c514d4570b3486fb34370 ,
                        0xb44671ba4770b9456fb2446dac4971b24c77ba4976b94c74b54975b54c77ba49 ,
                        0x71ab4971ac4a73b24e74b54874b44978bc4f79be5179c1527cbd4f7cc54d80c6 ,
                        0x507dc65380c9527fc84e78bd4a73aa42618e3b58853d56803a54823a56853a54 ,
                        0x823a57843855813954803953813c5a8b3f5f943c5f9741629443649243669849 ,
                        0x6b99486d9f4a70a0526f9c50729d5576a45679a15a7aa55b79a25c7daa6081b3 ,
                        0x5e84b45b85ba6087bb618bc65e8bd44e7ebe486e9e4061923e5e8f3e5c8b3d55 ,
                        0x835368875059673c3d413d434a5d779c3655883a56853b578d3b5a8f3555863b ,
                        0x5c8e4166a44368a4476da74e76b15980be5c8aca5d8ed45f89ca5984bd567bb3 ,
                        0x5f84be5c88c85985c25682c1557fc4507bc44879c54978ca456fb02f5966325c ,
                        0x69426aa44876bd4774b84771b4446fb84470b7436db44771b64873b25079b84e ,
                        0x76b74f75af5075af5277b3507ab54c75ac4e77b54a75b45078b94f79b4537fbe ,
                        0x5b86c95b87c75b87c65787c75786ca5884cb5986c95180ca5381ce517cbf4a6e ,
                        0xa444618e3e5a833b57803a54823953813754813655823d5b8c3c588e3b588b3e ,
                        0x5c8d40619243659b476c9e4a6fa34c70a04c6d9b4f6e9b4f6f9a526f96547196 ,
                        0x536f91526e91526f965577a2587caa567cac5377a75577a55076b04e79bc4e75 ,
                        0xb33e60953d5c893f5a7f3a4456282b331210100e0e0e4b4e56647a9e3253853a ,
                        0x588738578c385b933d619d3e64a53e69a84169aa4b70a84f7bba5281c55981bb ,
                        0x5783b95480bd5480bd567eb84f78af537fbe5681c0527ebe4f7bbb4e7aba4874 ,
                        0xc14572bc4470bd426ebb4872bd4a76bd446eaf476dad476ca8426aa5446ba942 ,
                        0x67a34368a44a72ad507cb75080c2547abb4f77b24f78af527ab45079b84d7ab7 ,
                        0x5681c4557fc05381c15780be5d82ba6186be6189c3618ec7638fcc608ac55e88 ,
                        0xc35f89c45c85c35a84bf537ab147699e4060913f5b8a39578039557e37527e3a ,
                        0x57843c5a893859873c5c8d3f5e933f5e95446499456694476899436698456792 ,
                        0x4b6b96486a954869904d6a8f4d688d56739a567295536e90546e925b769b6681 ,
                        0xa34f6480576e8e54719d4e6b984f69913f54734e5c6f171a1e0000000402010e ,
                        0x1010393a3e60748d496a983c5e993a5e9e3a60a03f65a54066a73f6aad456cb0 ,
                        0x4b74b34873ac4a70aa4c71ab4a6fa74f74ac4d75aa496ea6496ea84a6fab4a70 ,
                        0xa64970a74b70ac4d71b14570b3446fb2416bae4169aa4069a84970b5476fb045 ,
                        0x6ea5466ba545669e4368a0456ea54567a24466a14871a85277b5507cbc4f78b7 ,
                        0x5278b85179b45079b75178b65177ad5278ae557fba5b85ba5d84bb6089c0618b ,
                        0xc0658dc2658dc2658fca618ec76591ce6792d1608ece5a89cd5480c74d76b542 ,
                        0x69a03e60963f5d94405c923856853758863c5b8e3b5a91395b913e61933e6193 ,
                        0x40609143608d45619045638c46648d43618a4963885d73965a697c212b32363f ,
                        0x493b45562e37442d333a61676e3b3b3b606262363e4532394251596044474b05 ,
                        0x01000101010303030202021010100d100e3b3e46717f926681ad4e71b54169b1 ,
                        0x3c66a93e68ad416eb7426cb34469a74468a4476ca44564974565964065973d5e ,
                        0x8c4866954e6c9d47699e4b6b9c5171a64c6ca14165a14167a13f68a64067ac41 ,
                        0x67a73f67a2436dae4872b34870b1496dad476caa4469a3476caa456bab466ca6 ,
                        0x4a72ad4b79b94d79b9507abf4a73b24d74b24c76b14d76b44d75b04b73ae5277 ,
                        0xb15a7fb96088bd5e87be6089c05e86b76085b9688ebe648ec3618dc2618ec760 ,
                        0x8fcd5f8ed25b8ace4f7cc04b75ba4370b33f659f39598a405b8d3f5b913d5b92 ,
                        0x3b5a8d39598a3856853858833d5a863d5982405d893f5c883c57835b72985b68 ,
                        0x761318191d1b1a02000100000100000101010100000000000104000000000000 ,
                        0x000100000000000001000203000200020200020302020212100f10110f000001 ,
                        0x06070b484e5575849e5b78ab466fb43a67b04064a43d68a73d639d3e5f913f60 ,
                        0x8e405f8c3c57834c658f5e74974a5b756170838a98ae94a1b76a789443567741 ,
                        0x62903a5f993b609c3e64a43d67aa416aaf416bac416cab426db0416aa9416aa9 ,
                        0x4368a240629d41649c42639542649a4770af4872b54c77ba436ead466cac486f ,
                        0xb4466fae4671b04a76b55075b14e78b3537fbf5781bc5a7fb7587cb2547aaa58 ,
                        0x7eae5f84b85f84bc6086c05c85bc5b83bd547dbc507abd4b75b64668a4406199 ,
                        0x3e5f973b5d933a5a8f3856853956833855823956823e57833956833f5a864658 ,
                        0x756c7688525b651d23281b191900000101010102030102020200030100030102 ,
                        0x0202000200020301030303010002020202030303000101020301020301020202 ,
                        0x0202021010100e101101020002020200000003020053575c899dc04b6ead3960 ,
                        0x9e3e5e933b5d923d5c913a588947638c53647938404d636b722325260202020e ,
                        0x1011100e0d0101012124294f5f763e58873e5a893f5e953e629e4066a73f67a8 ,
                        0x3e69a8446db64371be4169b14067ab3c67a64269a73e649e3a609643639e466d ,
                        0xb1436dae446db2436dae426aab4467a74469a3466ba9446eaf4973b84a79b74d ,
                        0x77b24f75af5075ad5075ad5173a85075a7567eaf557cb05379af4f75af4d75b0 ,
                        0x4a73b1446eaf4066a63d619d3d5e9639588b3757883653803953813b54803b55 ,
                        0x84375284395788566b8b5861650806060c0d0b00000001010101030302020202 ,
                        0x0202030101040103020103020301030303010101020103020103040202010101 ,
                        0x030303020202020202040202020202100f1110110f0201030203010202020200 ,
                        0x010b080322292c424d6b667dad4e6a993e5a89395480485d7c57606d40424300 ,
                        0x00000000040200000101010000000200000301000000003e4a5c4866953b5988 ,
                        0x3a58893c5d8b43659b4467a94167a73f64a24265a43e609b3d5e9640609b4164 ,
                        0xa34064a03f63993c5a933c5e934063a23c65a34168ac4167a73c62a23e639f41 ,
                        0x639941629a4166a24a74b74670b34c71af476aa2426696466a984a6ea44a6da5 ,
                        0x4b6ea6486d9f466b9d42689e4266a23d62a03b5e963b60943b5d93405e973956 ,
                        0x8937537c34517836547d3b56823654834a61873f475405030200000000020204 ,
                        0x0201030303020103050204020202000202020202020202020103010101010303 ,
                        0x0001010203010203010203010101010002030002030202020202020f0f0f1010 ,
                        0x1002010302020200020200020304010302000011161531353a1920235b697b51 ,
                        0x5b6c222627060507020001020202040202040201000202000202020202000102 ,
                        0x010002576474496a973957863c56843d578d3d5f9a3962a13d65ad3f65a64066 ,
                        0xa6406aad3e6aa73e639f3d5e963b5b8c3c5c8d3e5e8f3d5a8d3d5a873b5b8c3b ,
                        0x5d924061993e61993a5f993c60963d5f954160973d60983f66a4416aa94169a4 ,
                        0x3f649c41639843659a44689e42659d44669b4364953e5e8f41609547659e405c ,
                        0x9238578a3b5a8d3d5c8f3654833954803d57853b56883954803656814c5f8065 ,
                        0x686d000000000202010402020103020202020301020202020202040202050303 ,
                        0x0202020202020402010202020202020202020202020002030202020101010203 ,
                        0x0102020202020210101010101002020202020202020202020202020202020202 ,
                        0x02020101010101010b0b0b080808000000020202020202020202000200040201 ,
                        0x0002030402020402020202020000015c667747608a34528136568b3d5c913e5c ,
                        0x953f65a5436cb14166aa4168a64066a74066a64268a84166a43e639d3a5b8d3f ,
                        0x57853856853c57833a57833857843a57833b59883d5c8f3c5b8e375c903e5c8d ,
                        0x3f5f9440629d4065a33f659b4261983d5c934060953d62963a61953e60953d5c ,
                        0x8f3a598637558438528139568339538139538136538037558639578635527e34 ,
                        0x4e7c38507e3c5887354660060403010200020103020202020301020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202020202010101010101000000020202030303 ,
                        0x020202010101000202030303010101050303000105000202000000596374465f ,
                        0x873a57843b59883b588538598b3c5d953e609b3d619d3d629a3e639d3f66a43d ,
                        0x659f3e629841609744669c456aa643649641619242608f3d5a863e5a893c5783 ,
                        0x3c57833957863e5b883c5b8839588b3a5e9a3e67a63e64a43d61973f6098395a ,
                        0x923d619d4665a23e5e9339578639568333537e36538037537c324f7b35527e38 ,
                        0x537f38558237568336537f39527c354f7d3958853a4a671c1c1c000001000101 ,
                        0x0303030001010202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202020202020202020202020202020202 ,
                        0x0101010202020303030101010202020303030002030201030202020402010001 ,
                        0x010301000907073e4b5b3f598135527f3755863b57863a55873a598c3c5c913e ,
                        0x5c8b3e5e8f3d5e8f3f5b84425d8943659b4769a44970ae4872b34e7cc94e7fc9 ,
                        0x4a78bf4b72b6466da44b6ba04465964166984160933f5e8b415f8e4368a44469 ,
                        0xa73f619d3e639b3a5d953f619c37598f37548039527c3555803854833b528237 ,
                        0x567d384f7f3c568536528135548138577e3f598139548034538033507733507d ,
                        0x485f7f40444f0806050301000102000503030202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x0202020202020202020101010202020303030202020202020303030101010202 ,
                        0x020402020202020002030202020200003f434e586c8b36537f37507c39538237 ,
                        0x58863958853c57893a578a3c5c8d415c8e4461883f618f44669b4a6fa3527ab5 ,
                        0x527ec55083d34d85d45384d25587d35587d35384d24f80ca4f7ec2507abb4c77 ,
                        0xba4c79bd4a73b2446ba9446eb1426dac426cad406baa3b68ab4167a83f66b03b ,
                        0x62a03858893a54823a537d365380375382365380375481375481395682355481 ,
                        0x34517e37537c39527e32507937517936455f474f602c30350a0c0d0003010202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202030303010002040201020202020103020000080b0f62708739 ,
                        0x527a35517a31517c3655823957863858893a5a853c567e3f588244618d4b70a4 ,
                        0x547bb95481c55888ca5784c75a89d35889cd6190da6390d3618ed1608dd05b89 ,
                        0xc95d8dcf588bd45186cf5481c4507abb517dc44d7abe4b75b64973b4456eb344 ,
                        0x6fb23f6eb23f6ab33f6ab3416eb7406dbe3c66b33a5f993957863a567f36557c ,
                        0x33517a37517f37527e37547b38517b364e78344f7b334e7a324f7636537f3b54 ,
                        0x7c4457784c586419181a02020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020202020202020202 ,
                        0x0202020201010102020203030302020202020202020200020204020200010102 ,
                        0x00002321205b6576344d753f5c8937527e324e71354f743d59883f5e91426391 ,
                        0x4b6f9f4e73ab567ab0577eb55d8cd05e8dd15e8dc56190ce5f8bcb668cc2648f ,
                        0xc86b96cf6792c56995cb6c99d66793d3628ecd618ed2548bd45485d15381c854 ,
                        0x7ec34c7ec64c7fcf4a74bf4b72b64073c9406ebb436aae406dbe3d71c44174c4 ,
                        0x4072ca3b6ebe3865af3a63a84268a93e65a93e68ab4068a94063a23b5d983d5a ,
                        0x8d3555863b56883d5c8f3e5f91475f7d42475000010002020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100202020202020202020202 ,
                        0x0202020202020202020202020202020202020203030302020202020203030302 ,
                        0x02020203010201030103040000002e323392a2b9345281315284385a903d5c91 ,
                        0x3956893956833d5b8a456aa65783ca5a89cd618dcd618ecb628ecb618bcc618d ,
                        0xc85e86ba6388ba5e84b46488b8688fbc698ec06c94c56d96cd6592cb6591d15c ,
                        0x8dd15885c85781c6507ebe527ec5507dc14d7bc24a78c54676c44272c43c72c7 ,
                        0x4271c34270bd4171c33f6fc7396bc5416ec9396dc73a6dcd3c6fcd3e6fc53c6b ,
                        0xbd3c6bbc3e6bb53b66a93c649f3e5d923855823654834d6180454b5806040301 ,
                        0x0002020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202021010101010 ,
                        0x1002020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202020201030202020102000301010303032c353e ,
                        0x566b8a8aa4cc6c8ec45176b04168a6416bac4573ba4a78bf5683c65987c75d8a ,
                        0xcd638fce6088c36389bf5b82b65c7fb15c7dab5c7fab6285b16181aa6585b063 ,
                        0x88ba678fc96592d55b8ed4588ad25580cf5082ca5383cb507ec54e7bc44b79c0 ,
                        0x4678c04674c14374c44270be3f71c43f71c43e70c3396ec8386ec53a71c63e6d ,
                        0xc13d6cbd3a68b53c69b33f69b03b64a33c619d4366a83b65a8395d9938588958 ,
                        0x6d8c616266050303000001020301020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020210101010101002020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020303030202020101011c1c1c4343433038497484a16f8ebb4c72ac426d ,
                        0xac476dad567fbe5884c15f89c45e8bc85c8cce5d89c85d84bb5678ad5778a657 ,
                        0x7aa25777a25a7eae6085b96288be618dcd5b8bd35787cf5481cb517ec74e7bbf ,
                        0x4b7ac4497ac84776c04977be4473bd4575c34271bb416fbc406cbf4871c04374 ,
                        0xc03e6fbd426cb73e6ab13d67ac3e64ac4165a53d5e903d5a8d3c5d954268a93d ,
                        0x66a53d64a23d5e964e63835a62690c090b000001000202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020210101010101002020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020101010303030202020101010000000200 ,
                        0x0103010134393c424d6156719d496da94e78b34e75ac527db65580c35a84c957 ,
                        0x81c25780bf537ab14f76aa5176aa557db2547eb95884bf5682bf5580bf5281c5 ,
                        0x4b7cc84a7cc84778c84676c44773ba4873bc4271bb4470bd4572bc3f6fbd426f ,
                        0xb93b6bb93d70c03c6cba3b69b63b67ae3f66a43b61a13c609c3e63973c5e943b ,
                        0x5b903b5a8d3f629a3d61a13b5e9d395d993951753b444e050604000100010101 ,
                        0x0202020402020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020210 ,
                        0x1010101010020202020202020202020202010101020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020101010202020202020202 ,
                        0x020202020202020202020003010202020000011312165a606b66788f54719853 ,
                        0x72a5597eba4974b74b76b54c75b44f78b74a72ad4c75b44c78b84e75b35077b5 ,
                        0x4d76b54c77ba4977be4b75ba4d77bc4776c04374c24473c5446ec14471bb4270 ,
                        0xbd4370ba446fbe4071bf416fc34271c24171bf4470bd3e6eb63d69b03c62a33d ,
                        0x5f9b3c5b923b598a3d5a8d3a5a8f39588d39588f4b70ae4c74b5526da04c5a70 ,
                        0x0000010200000002020104020301010503020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202101010101010020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0203030303030301010102020202020201010101010102030102020200020201 ,
                        0x02000200002c2b2d20242f6e7b8b4e60776b89b25070a54a6ca8446ba9446ea9 ,
                        0x4870ab4368a64369a34470af4771b44870b14571b84773c04370ba4170c13f71 ,
                        0xc3416fbc3f70bc3c6ec13b6cc23e6fc53b6ec44071c73c71c83d6ec43e6cc440 ,
                        0x6fc53f6fbd406ab74064aa3b609a35558a3b5b9039588d445e8c667aa35d7091 ,
                        0x6676938795a79299a834373b0200000401030202020101010302060002030202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202101010101010020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020203030302020202020202020203030302020202020202 ,
                        0x02020003010202020300020002030002020001000000000101011c1d213e4149 ,
                        0x515d6971859e4a638b475e844b658d6186c05479b34167a73f69ac426bb0426a ,
                        0xab3d63a34168ac3c66a93d66af416bb84169b4406ab73b69b6406db63d69b039 ,
                        0x69b73568c43970c53d72c3406dbe436ebd426fb93b68ac446ba95d7cb3516a94 ,
                        0x6d7d9439424b4d4f593f44450001000701020602010402020201030303030102 ,
                        0x0005030201020000030102020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202021010101010100202020202020202020202020202020101010101010202 ,
                        0x0202020202020202020202020202020202020202020201010102020203030302 ,
                        0x0202010101020202020202020202000203010303050303040201040201000301 ,
                        0x0203010402010301010001000e0e0e0d0b0a0d08090401032a2d316878894354 ,
                        0x6f8099c17896c55a79ae486aa5476fb0476eb34469ad3b62a73c65aa436cb54a ,
                        0x74b95176ba4a6aa54a6ba35376ae5577b34b72b73e69b23b64ad4066a74c6eaa ,
                        0x637eb166799c6270873c42470d0c0e0608090101010102000302040002020002 ,
                        0x0200020300020200020000020202020204020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202021010101010100101010202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202010101020202050303020202 ,
                        0x0202020101010201030002030002020202020002030203010000000001020000 ,
                        0x00010200040000010200010101282b2f2b2e362127327a889e7183a03f506a46 ,
                        0x597a7691c36f8fc03f567c46567b7c8fb05761720404042527318f9eae62768f ,
                        0x6785b65c84c55f81af6d7b9161666f1619170000000200000300020101010402 ,
                        0x0104020204010304020102010303010106020104020102020202020202020202 ,
                        0x0103020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020101011010101d1d ,
                        0x1d0e0e0e1010101010100f0f0f11111110101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101111110e101013111012100f1210100e10100d10 ,
                        0x0e11111114100f0e10100d0f10121010100f11100f1110110f0e0e0e0e0f0d0e ,
                        0x0e0e0c0e0f0e0e0e100e0e333131585c5d47484c1312140e0f0d0c0e0e0e0e0e ,
                        0x100e0e0e0f0d0e0d0f0e0d0f3d40453d4552343a410e0f0d0c0e0f0f0e100e10 ,
                        0x101311100f0f0f0e1010100f1110101010101010101010110f10110f0e10100f ,
                        0x0e101110120d0f0f0e1010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x100f0f0f0d0d0d1d1d1d030000000000110000000c000000080000000b000000 ,
                        0x1000000053000000720000000900000010000000530000007200000009000000 ,
                        0x1000000064000000820000000a00000010000000000000000000000009000000 ,
                        0x100000006400000082000000150000000c00000003000000180000000c000000 ,
                        0x00000000190000000c000000ffffff0051000000d09800000000000000000000 ,
                        0x5200000071000000000000000000000000000000000000006400000082000000 ,
                        0x50000000280000007800000058980000000000002000cc006400000082000000 ,
                        0x280000006400000082000000010018000000000058980000130b0000130b0000 ,
                        0x00000000000000001d1d1d0d0d0d101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010100e0e0e0e0e ,
                        0x0e10101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x10101010101010101010101010100d0d0d1d1d1d101010020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020101 ,
                        0x010000000102000e140f0c120d00000003010102020202020202010302020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0201030202020002020203010202020202020202020203010402020002020303 ,
                        0x030002020201030100020604030d181025422b3a6141355d412b473314211900 ,
                        0x0203010200020103000202020202020202020301020202000203020103020301 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020303030101010202020101 ,
                        0x01040202000202030204000202030100030100080f0a203a2a3b5c413d604539 ,
                        0x5d453b5e443c5f4439594019271b020000020103010101030303010101050303 ,
                        0x0202020202020101010503030202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020003010202020402 ,
                        0x020202020001010202020402020002020202020202020000000604031f32233c ,
                        0x6448427150486c4e446a4c3f634b43664b3f654940694a40614615221a010101 ,
                        0x0201030202020303030303030301000503020202020002020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020201010103030301010101040200030101010102020201000400 ,
                        0x01002236292d4f3740684c476e4e446d4d466f4f4870544a6c4e3f654740664a ,
                        0x4463483f684c396047131f130200030002000104020101010502040102000304 ,
                        0x0202010302020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020201010103030300020202020206010201 ,
                        0x01010103030500010305051f36273c624445714d436e4d4170504975564b7655 ,
                        0x4b7556477050476a4f45644943664c3c60483f65493b5c411823190001000301 ,
                        0x0100020200020203020403000202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020200 ,
                        0x020201010103030303000204010302000002070626422b436f52446f4e446e4f ,
                        0x4974534777534e7b5a4b7c5c4e7d5c4d795a4b735742684a45684d4063494063 ,
                        0x483e62443c5d4215231702000002000002020204010303020400030102020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020002020303030202020002020003010b140a294532 ,
                        0x45714d456b4d436c4d4471504a77564c7b5b4f8060578a65548664588969527f ,
                        0x5e4c7556476d4f456b4d45684d4063483e64463b5e43253b29080e0902000000 ,
                        0x0301000402020103020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202101010101010020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202030002030303000302020202 ,
                        0x0200001523183f6149426e4a436d4e4770514872534d7857507c5d527b5f517d ,
                        0x5e53815d538664508260547f5e5781624c7657497253466c4e40664a3f634538 ,
                        0x6342376344294931142217020202040103040202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202040202020202020301040201020202010303020202 ,
                        0x040103000202040300000001131a1333563b3760413d604543664c466c504d73 ,
                        0x574d73574c725650765a547e5f517d5e5887675986655787635787634e7d5c4b ,
                        0x7758496f513f68493e64463d60453b61433a63433862431b3021040103010101 ,
                        0x0202020002020202020003010101010401030102000201030202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202040202020202000202 ,
                        0x040202040202020202040202020301020202000100111c14385e423e6446456b ,
                        0x4f4e725450765a51795d537c60517b5c4e78594f795a527c5d517d5e547f5e57 ,
                        0x846357835f5281605883624f78594770514167493a5f45375a4034573c2f5839 ,
                        0x315a3a33593b172a1b0404040001010401030104020201030401030002000002 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020101010203010103030401030503030202020203010203010201030200 ,
                        0x00171f18426d4c3c62443f6849466c504871524c75564e78594f785c517b5c50 ,
                        0x7a5b4f795a527c5d547f5e518361528160518361558261547f5e527b5c477051 ,
                        0x43684e3e6146395e4434573c3056382b55323056380c130e0500020303030101 ,
                        0x0101010103030301010102020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020203010304020202020101010203010002 ,
                        0x02010303030101000003030303253f2d447150446c50477051466f5047705149 ,
                        0x7155476f53456d51476d51476d51466c50496f514a73574b775a4e7d5d4e805e ,
                        0x527f5e4d7f5d517c5b4c78594a7354426b4c39624338573c2b50363154392c55 ,
                        0x361f392702000102020202020201040200010104010301010101000202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02040202030402000301000202000003060907172519324a383d6f4d426f4e41 ,
                        0x714d447752477251416d4e406a4b3e664a3d65493f65493f674b3f674b40694a ,
                        0x40694a446a4c486e504a73544a73534678564978574776554571524671503d67 ,
                        0x44395f4130553b2b53372d5337264f300f201304020202020200020204010302 ,
                        0x0202020202030303020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202040202020103020103020202000301000101040103020001192b1e41 ,
                        0x6c4b3f6a493f68483e6a4b41724c447352497c5743714d426d4c3f6848406747 ,
                        0x3e6446396243376041376041345d3d335e3d3d634542664842684a4670514675 ,
                        0x554976554678563d6a49376a44366642365c3e30593a30563a2e57382d513310 ,
                        0x1e13020103040103040103030303000202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202010302010302030100020205 ,
                        0x02040203010a150b33593d3b64443862433963443862433c65463e6a4d41724c ,
                        0x3e754e41764f3e734c3e714c396c473568433568433666425d80664c76573766 ,
                        0x453c68443b68473e68493b6d4b416e4d3d704a658d715e8a6b3a6d483e694834 ,
                        0x5d3d305a3b2d55392b57332c4f35080e09040001030303010101020301000202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202101010101010020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020200030102 ,
                        0x0202040202020301000202020003060e0734513a396243375f43375d3f375a3f ,
                        0x355f403762413761423a6d483d754c3c7149366b44356a432e633c2f633e3164 ,
                        0x3f2c5c38bed1c2a5c2ab3367423a6c443a6f483b6b47366846396645325f3e99 ,
                        0xb7a4c3d9c736664233603f316343336640326140305a3b2f5936213725050505 ,
                        0x0001000002000400050101010202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020002030002020203010202020201030000000a0d0b35603f ,
                        0x3966453a69483b70493c6a463c6844366140376743376a45306740326b45376d ,
                        0x4834674131613d2e5d3c2c5b3a295c36cbdcd1f0f7f43c6f49346d46386b453a ,
                        0x6f483669433a6642305e3a517e5dd0ded23e69482f5c3b33613d315f3b305c3f ,
                        0x345a3e2b59352e5738243e270c1c110404040002000402020202020202020202 ,
                        0x0202020202020202020201010100020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202040202020202020202010206 ,
                        0x0003010400000a13092f5e3d366642386a483d704b3c724d396a443669473669 ,
                        0x4431634145765031613d2d603b2e5e3a2a5632204d2c1746255a8662f8fbf9ff ,
                        0xfffe84aa8e245c312d653c366b44346d46326740336742386943c2d2c75c8266 ,
                        0x2a53372e5738325e3a335f3b305f3e2e5b3a2d55392f5a392c4f35080b090200 ,
                        0x0102010302020202020202020201010102020201010103030302020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x050204010101050302040000040202010604243d2932603c35633f396847396e ,
                        0x47396e473c74492e67402358315f8567e9efea5a7d631b4c26214c2b27503146 ,
                        0x6f50708975dce3defdfffffefefed7e4dc2e633c225a2f2761382f6540326540 ,
                        0x2a633c2c613adbe5d993ad9b25512c2e5c382e5e3a30613b33613d2f593a2a5a ,
                        0x362e5c382e57381e392504070502000002020202020202020202020203030301 ,
                        0x0101020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202020202000200010303020202020202020202213d293360 ,
                        0x3f346140346440386943386d46386c47346d466f977b9db8a4d9e6deffffffd4 ,
                        0xe5d8486e52537c5cbcd5c1fafffdfffefffffefdfffffffffffefefefe98b59e ,
                        0x99b79e78a0842d603b1b5b2d2f6540a3bfacf3fbf4537e5d2756302e57372d59 ,
                        0x3a2d59352b54352c59382d5a39315c3b2b54352e57371b3720060a0502020203 ,
                        0x0303020202010101020202030303020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020402010402020202 ,
                        0x020000010c130c285131325e3a335e3d305b3a33623c34623e30643f2f623cbc ,
                        0xd2c0fefffdfffefffffffffefdfff6f5f7fafcfcfffffefffffffffffffffffe ,
                        0xfefefefffffefefefefffffefffefffffeffcbd8d094ad99ccd9cbfffefec5da ,
                        0xcb4076533163412a56372a53372f56362b54352b5133294f332a553429553628 ,
                        0x51312d593519271c000000020202040202030303020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020103010303020301040003172a192e5e3a2e5a3b2f5c3b325b3c2e ,
                        0x5d3c3162422b603944754fddede2fdfffefffffefdfffefdfffefefefefffefe ,
                        0xfffffffffffffefefefbfffffdfffffefdfffdfffffffdfffffffefefffdffff ,
                        0xfeffffffffffffffffff96b69da2c2af598862285c372957332a53342b51332c ,
                        0x4a31274d31274d2f295135274d2f2b5635152619000000020202030101040202 ,
                        0x0202020303030202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020202020202020202020202020202020101010202020a0e0924432629 ,
                        0x5c362d593a2a54352b55362e5737315d3e2a563274997fffffffffffffffffff ,
                        0xfffffffffdfdfffffffefefefffffffefffdfdfffffffefffdfffffdffffffff ,
                        0xfffffefffffffffdfffffcfffdfefdfffffffcffffffc5d7ca70987c50805c28 ,
                        0x60372d623b2a5333284b30254b2d294f33264a2c26492e2b4e33275231152a1b ,
                        0x0003010001000303030202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202101010101010020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020204020203 ,
                        0x0101080f0c1c40282b55362759372d593a2a55342852332d5636285430275030 ,
                        0xafc6b7ffffffccd7cdf2f8f3fffefffffffffbfdfdfffffffefefefffffffffe ,
                        0xfefffffffffffffffffffffffffffffefffffffffefffdfffefffdfffcfefefb ,
                        0xfffffafbf978a2835a83632654302a59392e58352754332a5334234c2d254b2f ,
                        0x284e30264f30284e32264f30161d180002000202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020204020202020200030102020204020202 ,
                        0x03010002020402020607050b1b101c3a2126522e285233295a3a2d5e38305c38 ,
                        0x2d56362a4d33224b2c345d3ddae2dbfffffcbccdc0ebf0eefdfffffffdfffffd ,
                        0xfffcfefefffffffffefffdfffefffffffdffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffefdfffffffefffffffeb5c8b96089692251302c58392b5837 ,
                        0x2f5d392a5a362b5e392f5a392d5637295232294c312a5236284e321324160401 ,
                        0x0301000200030101030301000204020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020204020202 ,
                        0x02020202020203010202020402020202020202020a110c26462e244d2d285132 ,
                        0x275433285c3730603c31643e2c5d372b5733255531326442e0e7e2fbfffeffff ,
                        0xffcbd9cd819e85f4fbf8fffffffffffefffefffcfefefdfffcfffefffffeffff ,
                        0xfffffffffefffffffffefffffefffefdfffffffffffefffffffcfffffff4fbf8 ,
                        0x59796123512d2953342b56352b563532633d3067402f603a2b5434275132264f ,
                        0x2f285132264f30284e3011221508010402020205000203030300030102020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202040202020202020103020202020202020202020300000000 ,
                        0x162d1e214926224b2623492d2752312e5a36305a372d613c2e613c2e5e3a2652 ,
                        0x332c5732cfdad2fcfefeffffffa8bbac3f6248f5fbf6fefefefdfffffffdffff ,
                        0xfefffffffffffefdfffffefefefefffefffffefffffffffefefefdfffffdffff ,
                        0xfdfffffffefffdfffea4bcaa2b5635254e2e274d2f2956352756352e5a353060 ,
                        0x3c34674232623e2f5d392b5434284b30294f31264e32265130111b0f01010102 ,
                        0x0202040103020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202040201020103030204 ,
                        0x02020200030101020005050517351c22452b23462b274a2f2a50322c5a362e61 ,
                        0x3b2e613c2e633c2f5e382d5b3721502f5f8569e7eeebffffffe2e9e2cbd9cdff ,
                        0xfefffffffffcfffdfbfffffbfffffdfffffdfffffdfffffefffdfffffefefffd ,
                        0xfdfffffdfffffffffffffffffdfffedbe3dc7a96822c5b3a2556302752312c56 ,
                        0x372952332a5334305b3a32633d31613d34644031613d2c5637274b2d214a2b26 ,
                        0x4a2c254b2d1e3f2a040803020000020202020301020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0002020402010002020300020202020202020300020b190e224a2721442a1f46 ,
                        0x2c244a2c294c312a52362f5c3b2e5738305a3b2e5839325e3a2e5d37214a2a42 ,
                        0x6b4c93af9b9ebaa386a18dc4d4c9fdfffffffefffffffffefffdfbfffefffeff ,
                        0xfffffffffffffffefffffefff8fbf9dce7dfa5bba977947d5c826631643e2151 ,
                        0x2d275031274e2e274d2f274b2d2750302b5133315b3c33613d32603c35644333 ,
                        0x633f2c56372a50322a4d33244a2c254b2d2c5637111712000000020301020103 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020101010202020003010203010103030300020606 ,
                        0x061c39221f432520412623492d224b2c294f332a4b302e54362e57372d56372c ,
                        0x55392e58352b5c362a5435214c2b1f48281f45291b42222f5636c8d6cbfffdff ,
                        0xfffdfffffefffffffffbfffefdfffffdfffedce7dfa0b5a6597960285132244f ,
                        0x2e1d48271e412621412820442c20472d234729244a2e224c2d28502d27503129 ,
                        0x54332d5335315c3b30633e35633f2f5936294f332a4e30264e32214b2c274c2c ,
                        0x09140c0000000202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020201030202 ,
                        0x0200030101010100000007180b1e4426234c2d224b2b204d2c244c30264c3027 ,
                        0x4a2f274d312b51332e54382d5b37295135254b2d254b2f24472c244c30294c31 ,
                        0x2952331e462a2d50367f9884b2c3b6b7c2b8b7c5bab6c1b99cb2a078957e3359 ,
                        0x3d1842231841222b54352651302b5a392e5c3828523329533426522e254e2e1f ,
                        0x4829234c2d22422923462b2954332a5334294f31264c2e2c52342b5635294f31 ,
                        0x21472b21452721472b224d2c1e4224070d080200030401030202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202101010101010020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020003010201030401030202020604040e190f1b412522452a1e442820 ,
                        0x4027214429244a2e234c2d274b2d234c2d284f2f274b2d26492f224b2c28482f ,
                        0x23472921472b24462720492a20462a24472d22482a193f21123b1c123d1c143c ,
                        0x20133b1f193f211840241e462a23492d254b2d254b2d22482a254b2d20492a26 ,
                        0x4a2c25492b24472c2243282043291e41272041261d43251e41261f4025204329 ,
                        0x25462b244929234c2d23492b244b2b2043281e442621402b20492912291a0505 ,
                        0x0502030102020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020402020303030001 ,
                        0x0102020203020402010302020202030102010304020200020202010313291621 ,
                        0x4b2c1a43231f41231f43251d40261d432520472724462821442923492d22452b ,
                        0x244a2c21472922482c2144291f42271e412622452a2246281e44282144292245 ,
                        0x2a21442922452a22482c23462b22452a25482d1f462621442923492b21472b23 ,
                        0x47292046282449292046282047272249292043291e44282043281f40251b4123 ,
                        0x203e251e3e251f442a20462a20472723482825492b1f45292241262043281c42 ,
                        0x261b42281f43251e3e2506090701020002020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020202020201010103040201040201010102020204020204020202020202 ,
                        0x02020001020301011b3b221b44251d40261e40221b3d251a3d221c40221e4126 ,
                        0x1e3e252144292043282043292046281e44262043281d40251d40252043281c3f ,
                        0x241f43251d40251d40252142272144292043281f42271f422722452a21452721 ,
                        0x42272147292043292345271f422724452a1e422a2547281c43291d45291f4227 ,
                        0x1e41261c41271d40261e42241e42242044261e41262043281c41271e43291c3f ,
                        0x242043281c45261c42241f41231d40251941251d402605080600030102020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x0202020202020202020202020202020200010104020202020201030303030305 ,
                        0x03030101010402020202020402020504060e1f141a40241c3f25183d231c4224 ,
                        0x1b3e231c42241b3e231c3e201e40221c3f241c3f241e43231b3e231d40251d40 ,
                        0x251d40251f40251c3d221e3f241c3d221f40251f3f262040272041261d40251e ,
                        0x41261f42271e41261d40262042241e41261b41251c42261c43291b442522452a ,
                        0x1f42281e43231f41231f3f261c411f1b3e241e4126193f211a3d222041261e40 ,
                        0x221e40221d41231b3f21193d1f18391e173a1f2043281e47281c3d2217311a16 ,
                        0x2c19080b09050303020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020201010104 ,
                        0x02020201030201030301010203010004030002030203010502040f22111d4327 ,
                        0x1f40251d41231f4025173f231e4224183d231a3d221b3d1f193b23183a221c3d ,
                        0x221b3e241f42271b3e232041261d3e231d3e231d3d241d3e231c3d221d3d241d ,
                        0x3d241c3c231c3c231e3f241e3f241c3f241d3e231c40281c3f241f3e231e3f24 ,
                        0x1c3f241d41231d43271f41231d40251a3d231d3d241c3f24193f211f40251c3b ,
                        0x201a3d221c3922183b20113a1b123b1c153b1f1c3f2522482c325b3c345d3e16 ,
                        0x311d142b1c0c1911020103040503040305010101020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202010101020202020202050302020202000102040005 ,
                        0x0304020c1c111d42281b3f211d3d241b3d1f1e40221c3c231b3e23193f231a3c ,
                        0x1e1c3b20183a221a3c241b3c211c3d221b3b221d3e231c3d221d3e231c3c231c ,
                        0x3c231c3c231c3c231d3b221b3b221a3a211c3c231d3d241a3b201a3a211d3d24 ,
                        0x1d3e232042241c3e261d3d241a3d231b3e231e3d221a3d221b3b1c18391e1839 ,
                        0x1e1b3a1d193b1d1b3b221c3d22183b21193f232c4c3345634a5b79606c8b7084 ,
                        0x9e86a8b8a7d0d8cec5c9c320211f000100000100010002000100010200020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x0202020202020202020202020202020202020202020202030303040202020202 ,
                        0x0002030001010402010603000e1b0d1c41211c3b1e1c3c23183b201d3d25173d ,
                        0x1f1d3b221939201a3a211939211b3821193a1f18391e1a3b201939201939201c ,
                        0x3c231b3c211b3c211c3c231c3c231c3c231b3b221d3b221a3a211b3b221c3c23 ,
                        0x1c3c231a3a211b3b221c3c231b3c211b3b221a3b2018391e183a1b1638191c3f ,
                        0x1d1b3e241a3f2b1f43321e42322146361e42341a3f251a3b201a3d2251645174 ,
                        0x84739aa297bdbdb7cac5c2d0cbc8d5d0cfdbd6d5e4e1dc8f8d8c000000000301 ,
                        0x0601030202020001010201030202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x02020202020203000202010300010205030303010002090420403b2145391b3a ,
                        0x1d193a1f1b3a1f163b211a3a22193c211b3d1f1939201c3a211a391e1a3b2016 ,
                        0x3a1c16391e1a3a221939201939201d3e231c3d221e3f241e3e251d3d241c3c23 ,
                        0x1c3d221a3a211b3b221b3b221b3b221b3b221c3c231a3a211b3d1f1a3b201d3f ,
                        0x27234642294a532d4f662b5774315b78365c7f3b5f87385c8a355a9238599122 ,
                        0x4a3f1b3a1f1d3b221b3c211e3e2527442d2f4b343b553e455c46637862839683 ,
                        0x9aa99ba5a9a32f2a2b0000000300020001000202020300020202020101010202 ,
                        0x0203030302020202020202020201010102020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202101010101010020202020202020202 ,
                        0x0202020202020202020202020203010002020201030102000105000602010000 ,
                        0x04172a322848792847741e3d3419371e1c3a2118381f1c3a211a3a21183b211c ,
                        0x3b20193921183b21183a1c1c3d2f21403918391e1a3a211a3b201a3b201e3f24 ,
                        0x1f40251f3f261d3d241f3d241b3c211f40251d40261b3b221c3d221c3c231c3c ,
                        0x23183b21193c211a3b263a5b823d63993f5f943d60a03d66a4476aac456bab41 ,
                        0x69a44068a23f659f385b8d34546b1a41271e44281b40261b3e241a3b2018391e ,
                        0x12351b0f341a14371d193a1f193e24203f2a1624180d1e11101e120e20130911 ,
                        0x0a03040202020201010103030303030302020202020203030301010102020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020401030201030202 ,
                        0x020300020503030000000b0e1628457227497e28457e25486a1c402a193a1f18 ,
                        0x391e1b372318391e193c221c3d22153b1f1b38211638191f3d3e23445719381b ,
                        0x183b2123442f264a3c1c452522482a22422923462b1d3f2120412c2d5958224a ,
                        0x2e1f42282c533323492b214a2a1c3c2319381b2546373c619343669e3f639945 ,
                        0x6795476fa04a6fa34b6ea04c6c9d476d9d446ca0456da83e669b2b554e254728 ,
                        0x2446281f40251d3e231c3b26193c211a391e193a1f18381f16382017361b1c3c ,
                        0x23173d21193c211e3d221a372014321918381f1c3b2c11261704000000020206 ,
                        0x0201000203020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x02020202000301040201000201000203010200050400192c4d2a4f872b518b2b ,
                        0x4e8d294c90204151193f211b3a1d19361c203a2d1b3e241c4526193d2719371e ,
                        0x17391a203f403052881c44321c3f1d255654365b932a4e5622482a1f45271e44 ,
                        0x261f3f202d51513c5c91274c4425572d28522f1b412323492d1d40251d43272c ,
                        0x5b53406b9c456aa44670a5476c9e486ca2486fa34d73a94e76ab486fa34b6c9e ,
                        0x47679c416698366174245430264f2f1c44281d3d241f3c221b3a1f2044341c40 ,
                        0x301b381e1c3c2320473817371e1b39201a3b2016391f1a391e1b4026284c5c2a ,
                        0x506e183222030702020103010200020301000202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x020202020202020202020202020202020202020202020002030202020200000e ,
                        0x131c2d4b7a30559132578f30578e2f5595325184234a3a1a3e20204339274556 ,
                        0x1a41211e44261b39201b392017391a20423b35589c2c51731f4e3835648a3e64 ,
                        0xa4355c6b1d49242547281d3d241c4127416a8b466dab355a5e26543027523124 ,
                        0x4d2d1f442a1a3e202347393762894069a044649943669e40689d476da7456da2 ,
                        0x426a9b44699b446596486da940669c41669e3d5e9029574b244b2b1f45292349 ,
                        0x2d1b422819391a2149441f42381837182042312c4e651c3f24294f67244b4d1b ,
                        0x3c21183c1e2748582a4d8c3252872d516312201f000000010304050204040103 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020203030304 ,
                        0x0103030402010101040202172d4935558a31528332578f33589233589232578f ,
                        0x2b4c5b1f401e2e53672b4b741e44261d40251a3b201c38211a391a1d433d3159 ,
                        0x9a2f558b3058753862a53f64a0365d73254d312143251e42242b524a4a6fa34b ,
                        0x75aa3e637d284b31234d2a254a301e40281b3a1d254f443b628f3d61973e6599 ,
                        0x3a60963f64983b5d923b60943c5f973b60983d649b3a659e3a629c3b619b3b60 ,
                        0x9e3363751f4724193b1c2a505029494f193a18356179305668163514264b4737 ,
                        0x5b912b4a4b2f53812f527a1a3a2725483e2f53833354852e4b78314e75264064 ,
                        0x10161b0200000003010202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202010206040202040100131f312e4f812d497834538a ,
                        0x34588e375a9a325a95325e9b305172264d3e3b65a02f57881e43291d4424193c ,
                        0x2118391e1a3a1b22484a2f5798325990315690345a903c6298426baa335a631a ,
                        0x391c1c43293e66834769a43f6ba1406aa52e5553234f2b1f442a1c3c231c3e1f ,
                        0x2d565f38598a3d5d92375f94375f90375b9135598f32588e36599130538b305b ,
                        0x9432578f355999365b9737609e335b8c2142331f3b1d3258782d536b284f3f43 ,
                        0x6ba03b5e8a1b43272f5b683a60a0335983365c92385a95264c58274f54375c96 ,
                        0x36578933527f334b6f3352792236550804090203010202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x020202020202020202020202020202020202020200020303000202000004070c ,
                        0x233a602546772e4d822d5088325793325892305b9a345e9f2f557f2e5778355f ,
                        0xa02f53831d442b1d42221c3d2216391e16371c26464c2e539130519035548b2e ,
                        0x57883758903a609a365b8d1a402a234b393d689b3c65a33c649e3862a3345d84 ,
                        0x244d321d46271d3d1e1f462d32567e345796395b91375a9234558d33538e3351 ,
                        0x882d4d822b50882f4d882a4d852a4d8c2f518c2f548e304f8634548f2448501c ,
                        0x4022355e8b35598137566f44659d466ba3395e7a3c64953e61993d5f943b5d8b ,
                        0x375889385987395d8b3b5d92396094345281334f78364f77314f781c29430305 ,
                        0x0505000102020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x0402010301000302000f192a22437128457228497a2a4f872e508b32538b3156 ,
                        0x8e3456913456923156922d54922b49801e3f311d41231a3d221939211a371d27 ,
                        0x47542648832c4b802c4e83314f862f4f802b4f8535538a2549492e556b335590 ,
                        0x30569030528e3153892e50852449411a3b201b3a1923454430558f3154863554 ,
                        0x873151863053853352892d4c7f2a4f89294c84284a86294b802348802b497a2a ,
                        0x4b832b4a812e51892b517422474533558333588c3559893d5e8b3c5e8c3f6092 ,
                        0x3a5f91375d93365a883859863b5e903b5d93385e943a5c9136578531537e2c48 ,
                        0x712a446c28466f2d4c790d172103010002020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202101010101010020202020202020202 ,
                        0x0202020202020202020202020202020101010301001b31542443762342752646 ,
                        0x812c4c872c4f8e2d52902f53932e51902c518b2b4b86284b832649811f433d1e ,
                        0x40211a4024173b1d1a391e1f404f27467d27467d28477e29477829467928497a ,
                        0x2d4e802c476c2e4f81294b813150832c4f812b48812a4c812549591a3a211938 ,
                        0x1b2349552f538930508136578f365a962f568d32568c304f862f4f8a2b4d832a ,
                        0x497c27467b27467b26457c27467b284b8327457e2e4c832f4e852b4f85315284 ,
                        0x305182344f8133548235538232538136548534578f315688395b91385c92395b ,
                        0x9038598a3352853353882c4c752d45732a436f2f4978131c2a00000102020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020202020501000002 ,
                        0x021d30562846772142702242732642782546782b49802b467f25487a27477c27 ,
                        0x45762342752340731e3d3a1f42201a3b20153a2018391e25415223457b25437a ,
                        0x22457d23447c25447726457a24447528457828497b23477d27457c25477d294b ,
                        0x812446812647791f3f2c2242372c50802d4e8633558b34568b33558b33558a2f ,
                        0x50823350832e4f7d2c4d7f28467d25407824427123417024427326457c254880 ,
                        0x2a4b832b498227487a2b4c7a29497e2d4d7e2e4c7b2947762c49752b4d822d4d ,
                        0x823454892e508532598d34538a35598f31548c2e50852c4b7e2b4470273e6b26 ,
                        0x4370172c4807090a050100010200050303000101020301040201000203040202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x020202020101010402010101011c31572241761f3d6e223d701f407125447924 ,
                        0x447926477f23437e23417a21437f1e41802342791b3a391b40201c3c231b3821 ,
                        0x18371c203c4d263f7724427324416e27406a243f71264179233f7525426f2941 ,
                        0x6f28417323437424407625417024467c2648841e4641244b5a294c842c508631 ,
                        0x538938568f315486345489365789345684304f822b4b80284879264775234279 ,
                        0x24457d23427f244378224273234475254880294b862747822b498226487e2746 ,
                        0x7928447328427128467527457629487b2d4c832d4d823052882e4f872e4d8028 ,
                        0x4a802d4573264271224371213f7028406e1c2a47030506020105000202020103 ,
                        0x0201030002020402020402010202020202020202021010101010100202020202 ,
                        0x0202020202020202020202020202020204020203010100030226395e28477a1a ,
                        0x3970183b6d234073214176243e6d1f3c4323423b2243461f4246203f401e413d ,
                        0x1a3b26193d1f17371e1b39201a391c1d3a49233f7b223c601a392a1a3a291b3e ,
                        0x3420415125447723488224468223437e25447b254576294a82294b8725488823 ,
                        0x456323486e2a467c314f862f507e325283345281365485315381335583345386 ,
                        0x2f517f2b4c7d27457e24437a26437c2242732643762947782a45772744772444 ,
                        0x792748802749852b4b8625457a26437624406f26417425437a26477925477d2c ,
                        0x4a81294c842a4c8727497f26487d24467b27416f233f6e1f3c691d3c69264278 ,
                        0x0a131d0301010402010402020303030002020101010303030202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020204020200 ,
                        0x0203000001818489bdc7ce989fa8707d93576581596c87526685324e701e4446 ,
                        0x1c43231e41262143241e411f1c3a21193920193a1f193a1f1a391e1e3b4a203d ,
                        0x691b3a3319381b1a391e1a381b1537191e3e43223f6c25458028477e25417725 ,
                        0x407223457b23447c254580294885294b872849812b4d882b5084315388335590 ,
                        0x34568b345583335182284c7c2d4b7a2645782544772540722943722742741e3c ,
                        0x6b1e3e69243f7123406c203c6b263e7222427725447727457c22417425417729 ,
                        0x437227427425407322427324406f2442731f3c6f1e3d70264473234275203f72 ,
                        0x25406c1a36651a2b4c0f1d30060c130200000503020202020402010202020103 ,
                        0x0401010102020202020202020210101010101002020202020202020202020202 ,
                        0x02020202020202020003010301010f141d455f8e5574ab899aafdad8d8faf6f1 ,
                        0xfffafcdbdad663786f2445371e412622452a1e41271b3e231b3b221d3a231b39 ,
                        0x20193a1f1939201738231839241c3b201b40261c4022203f22183b21163b1b18 ,
                        0x372e25416323457b23417c254477203f72213d6c243c6a2144762948852a4c88 ,
                        0x2d4f85294c8b2a4d8528477c2a4c7731507d2e47712c446e26466f2849772948 ,
                        0x7b2b4a7f25477d24468122437b26447b223e741f3c6f203b6d253c6a223c6a24 ,
                        0x42732544791f3f70234172243f72224371234172264173253c6922385c203961 ,
                        0x223b65253b6b2540731e3c6d1e40768194b79da0a80000000404040303030002 ,
                        0x0300020304020102020202010305030302020202020202020210101010101002 ,
                        0x02020202020202020202020202020202020202020002030200011a283a3f68ad ,
                        0x3660a539608c3b5751607965758b7947644a1a3d22274b2d274e2e295030254c ,
                        0x2c1d4129183e221a3b201b39201939201939201a3b201c42242144291d48271c ,
                        0x41271e43291b3e241c3b1e1d3b1e1b39261f3c4a233f5e224174234277204075 ,
                        0x243f7124427326406f29426e2543722843752b487b2a487f28497b27497f2d4f ,
                        0x7d2d4a7d2b487b24447527487a22437526457c234476234073213d6c213f6e21 ,
                        0x3f70234170203e6d26406f233e76243f77254073233e6a243c70203a681c355f ,
                        0x152e5a1c355d223b651f3c69253f6e1e3c6b14366c446194acb8cafcfbf7dfdd ,
                        0xdc353f50141e2f05030200000000020304020103020401010104020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x030204000000263a5d426eb5446eb9355e7e18401717431c183e201c45252750 ,
                        0x312c58342f5839294f33254b2d2144291c3c231a3a211d3b22173a201939201d ,
                        0x3e231b47281e4525193f231d40251b3f211b3c211c3c231c3c231d3c211d3c1f ,
                        0x1d3d25243a5d273f73223f72273f6d233f6e2442732342772441742341722241 ,
                        0x7623437429467338548324427327457624437826457a24457723427722407126 ,
                        0x417325426f24406f21406d213b692239692e457221406d213b6a233a6a243d67 ,
                        0x263a631f3a66183365394d768897aab6bfc93c4f741e345e13305c3a5177959f ,
                        0xb0e1e0e2fdfaf5e3e2de5d718a334d7b334f78121a2702000002020202030101 ,
                        0x0002000202020202020202020202020202101010101010020202020202020202 ,
                        0x020202020202020202020202050100020202304d86426fc03b6cba3d67b2345f ,
                        0x8634617c39618437669a2b5561264b2b2e5436264f3021442a1e3f241b39201d ,
                        0x3c1f1b3821193920193a1f1b38213153353157391b3b2216381a203e251e4126 ,
                        0x21442a1d41231d43272147291d3f211e3e2d1d39621e396c243e6c1f3e6b233d ,
                        0x6b233e70223e6d233e6a253f6e234172253c698193aa5b6e9115346923426f25 ,
                        0x407225407322406f23417227427425407222406f1e3965162d5b586a89adb5bc ,
                        0x1f3d66233c68233d6b223c6a1c3b6e223b657e8aa6eaeff2fdffffd1d1d11d37 ,
                        0x5c2c45718b9cb7eef1f6fffefffdffffefeeea6a79932b4a7d354e7a38568531 ,
                        0x486e090e11050200010304030200020202000301020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020001010a060b3454 ,
                        0x853d63a33e68b5426cb73d6cbd456ebd4271c73d70b92f594e21472921472922 ,
                        0x452a1e3f241f41292245371c3b20183b211c3b20173a2013391b96ac99eef5f0 ,
                        0xdbe3dc708b771d4123284c2e265031285132275334284c2e274b2d1b4425475c ,
                        0x71364d7d1633601e37631a34621d39681f3d6e223f6c273e6e234073203d6a3c ,
                        0x5478b0bbc953678a173566263f71203f6c25417025407220416f233f6e233e70 ,
                        0x1937666a7c99f0f3f1708196223a68273f69223c6a233d6c17325e92a0b6ffff ,
                        0xfffefdfffffffba9b0b98d98acd5d9def8f7f9fdfffffdffffdfe1e16978922e ,
                        0x4b7836538037548138517b35507c334b6f0a0d11020001020301000203020301 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020201010106080934598b4068b03f68b13d6ab43d6bb93d6cbd3d6ec437 ,
                        0x699d24562c1f4f2b2c5c32254f2c1d421a20463a3d71b323473b17381d1b3920 ,
                        0x1a3a2115392129472e627564a2b2a1b5c2b435553c24472c284e32285131264e ,
                        0x32274d2f21452d163b1bbecdbfe9eaf47b8aaa3c5577516686344f741e3b6220 ,
                        0x3862274170203d6a2943711c3a71576e8edddde3566a89213761203d6a233e70 ,
                        0x243d6f233f6e26406e1b3865597195fffdfdd8dadb7789a0919fb1253d611f39 ,
                        0x67163661617495fffdfdfffffeadbccce4e7effffffff9fbfcc4d1e1eff2f6ff ,
                        0xffffe0dcdb576f933459933e5d9439588b395786344f7b34517e3c5684151e2b ,
                        0x0000010303030202020402020202020202020202021010101010100202020202 ,
                        0x02020202020202020202020202020202040000121b2f3d62a63d68b1406bb43e ,
                        0x6cba426ec1406fc1406bc23f6cc739668b2b54392a52402d58552951501f432b ,
                        0x2e596821463618371c1a3b2014371d1b3a1b183b2114381a183d1d2c4932274d ,
                        0x31264f302d56372a51312b55362d5935204b2a183e22a2b3a6fffffefffbfff9 ,
                        0xf9f9fffdfff1f2f6b7c0ce324b77183562647798344c76182f5f7485a0eef0f0 ,
                        0xd1d9e0233d651d3864233e6a243f6b223e6d27406a15315a8c9aadfbffffd2d4 ,
                        0xd5f9f7f775839519305d18346a394f79d4dadffdffffd0d7e6bdcadafcfffbff ,
                        0xffffe0e7f0e6ecf1fdffffd1d6d95d7aa6375fa03f6aad4167a73c5d953a5b89 ,
                        0x395682375584324f7c212f420301000001000402020002030202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020201010100000028 ,
                        0x3e624674c13d6bb93f6dba3f6bb23e65913e6486446fae4170c13f6fc1406cb3 ,
                        0x3d6aae436dba406aaf3e66a129515619381b1d3f211b3a1f2241322a4f45173d ,
                        0x2123492d2751322851322857372c59382b55322e5938295233274a3026492e1b ,
                        0x3a1f2a482fcdd4cffffeffffffffc4cddba9b3c5dee7ebdddfe93a527663728c ,
                        0xbdc6cf3d567e7b899ffdfffffdffff66779218315d223a641f3c691d3b6a223b ,
                        0x65162f5b76849bfffffefefbf7ebecea3a4f6e17305c41577b9caac6c4cfe38d ,
                        0xa3cca4b8d7fdfffffffffffffefefdfffff1efee9fa9ba4f6fa03c67aa426cb1 ,
                        0x456aae4168ad3a619f385687395381375481365380293e5d0100020100020401 ,
                        0x0301010102020202020202020210101010101002020202020202020202020202 ,
                        0x020202020201010100020023365b406fb9426ec14371c9325c6f193813173c1a ,
                        0x2549393264884470c4396688315e6939658e4372cd4473ce4173c62249391f48 ,
                        0x2814351a2b565f4675a11b3b1c204126254c2c2a58342e563a2d53372e57372e ,
                        0x59382c5535295334224b2c193a1f15361ba4b7a4fdfffffffffebdc5d61d3968 ,
                        0x1932648794aeeeeff3657697b9c0cfe7ebf07b8ea3e0e3e8fffdffd4d7df2541 ,
                        0x6a213a62243e6c253e6a2039652a4069c2cadbfffffffffeff99a4ac122e5719 ,
                        0x315f879db994a9c94b6fab4a6ca7eff0f4fffefffffdfff8f8f2b5bbc07086a9 ,
                        0x4167a1476dad476fb0456bab4164a44266a63e639f3b5c9439588b3855813d5a ,
                        0x87263f6102040403010101010100030202020202020202020210101010101002 ,
                        0x0202020202020202020202020202020202010101000000293b64446fb8426db6 ,
                        0x416fc33b68b12d5c711d442a163613244b3b3455651c3a2116371421432b3764 ,
                        0x864473b13f6d9c2046341b42221736152d55534f79ae1d3e231e3b211f3f2621 ,
                        0x4729274e2e2958372c5536284e301f4e2e234d2a1b41231b3f211e3c1f466548 ,
                        0xe4e8e3fefefeffffffced7e55d76981c3e73ced4e1fcfefedbdee2d1d8e1ecef ,
                        0xf3edf0f8b0bbcfffffff73859c0a2857203660283e6814315e7c8ca3fffffffa ,
                        0xfffefffffeb0b9c63a50738690a8b0bacc95a8cd859abacad4e5fdffffffffff ,
                        0xf2efea8ea2b442699d456da84d74b24c74af4873b6446eb1446dac3d639d3e60 ,
                        0x9b39568339588d3b598a3856873e5b88151e2c00000002020201010102020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202000000 ,
                        0x0c0f143d62a04271b5446eb34170ba416fbd4474c638657325502b1b4522183a ,
                        0x1c1c3e201c4022173a20193a1f1e422a203f2415381e173a1f1737182c54534a ,
                        0x73931a3d221a3b201b3b221b3e241e42241c42261d3e2920422a20412627533c ,
                        0x2654472448421c412d0f351f92a19dfffffbfffffffffffcfffeffa4b1cb5f76 ,
                        0x9cf6fbfafffefffbfafcd5dbe0c7ced78b9ebfbcc6d8f8f8f87b8ba82a416713 ,
                        0x2c56182e587f8fa6ffffffeff4f5c4cdd7e0e5eef6f8f9d0d7e0c3d0de9bb0d0 ,
                        0xfcfbfdfffefffcf9f4c6ccd17186a64b70ac4f7abd527bb94f78b74b75b64d72 ,
                        0xb0486eae416aa9385b933c55813a5a8f39598a38588d37548038528021344f03 ,
                        0x0402010200020202020202020202020202101010101010020202020202020202 ,
                        0x02020202020202020200000017243a4973b84875b94676be4473bd406fc14172 ,
                        0xc83a647b234f28204e241d4a291d44241e47281d40261a361f15361b14381a18 ,
                        0x381f1e382015391b244433224531183b211d3e291c3b1e163b211a3a211a3c1e ,
                        0x1b3b221a3c1e19381b2a514f4975b44b76bf436ea7436d9a4771ac7993b8c8d1 ,
                        0xd5fdf8f9fefffdfffffeafb7d4839abadde3eeffffffd7dce56b7e9fa7b1c2b9 ,
                        0xc5ddeef1f9fffeffd6dee55d75930f2a5c667895ffffffb0bbd11b3668c8d0dd ,
                        0xfdffffebeef2f0f3f7eef1f6ffffffe2e7e68b9bb24f74a64e7fbd5480c05c82 ,
                        0xbc5682bf517dbc5079b74c73b1436aa83f629a3f5d8c3c5a8936538039538137 ,
                        0x578838547d3a5784344d77070a0f010200040103020202020202020202101010 ,
                        0x1010100002030202020104020300020002020001000806063356884c75be4673 ,
                        0xb64c76bb4976c04476c24273c33b6aa73363972f5d75254e2e214a2a173f231c ,
                        0x3d22193920193a1f1b3a2517392116391e1a3a211837161e3d2e3b607c3a596e ,
                        0x1d3d24193920193c221b3f211739211b39201a391e1c3c23426886517dbd4f79 ,
                        0xbe5079c24a77bb426fb3466da47a8fafc1c7cefcf7f4ffffffe9ecf4becce3e3 ,
                        0xe7f2fffefffbfdfecad2dfeceff7fffffefffefffbfffffffefe95a2b875889d ,
                        0xfdffffd8dee9b2bfcffffffefffefffffffffefdffffffffc8caca617a9c517b ,
                        0xc05282c45883bc5f87c25985c25783c05984c3547cbd527abb4971ac4a70b13e ,
                        0x639f3a5b883d5a863856873956833b5d92395a92354f7d0a0f12050100010002 ,
                        0x0202020202020202021010101010100301000201030202020201030403000001 ,
                        0x001725414674b44b75bc4f79ba4d77ba4a76bd4979c14775c24473c53f70d641 ,
                        0x73d73059621a391e17371e18381f19381d335e7f3966881b3e2418381f16391e ,
                        0x193d25476e94597cbe365967193b1d18391e193b23193c21193c221c4022193a ,
                        0x1f16381a325864517ab9527cb74e78b94f7bc2507cbc4d7abe4373bb4b73ae77 ,
                        0x92adc5cbd0d5d9ded3d8dbd8dadbe5e4e0d8dcd7f0f1effffefdfffffeffffff ,
                        0xfefefefffffefdfcfff5f7f8fefffdfffffffffefdfafffffbfffffffefffdfa ,
                        0xf2b7c2c65779a74d75b05581c15d88c16086c05f87c1608acb5a89c75884c359 ,
                        0x82c1517dba517bbc4e78b94572b64873bc4570bf3c67b04267a54164a83c5e99 ,
                        0x31496d0203070301010402010202020202020202021010100f100e0304020103 ,
                        0x040101010002020000010d11163554874a78c64976b94e7aba547ec14b77b648 ,
                        0x7cbf497dd04a7bd7466fad3560752145341a371d193c22173a1f1c391f234c45 ,
                        0x284e481b3b2319392119361f1e432f507aa4355c641737181537191a3d231f42 ,
                        0x271c45261b44251f432520422416341b2c4f4b527dbc507dc05480c04e7fc352 ,
                        0x82c4517ec14f7ec24e7dc14b77c44e77b6557eb5567cb65a7db55378b05176b0 ,
                        0x6586b4bdc3cafdfdf7fbfffffffffffdfcfefdfffffefefefffffffcfefeffff ,
                        0xfffefefeffffffe0e2e2788ba64f73b34b79b95380c3547fc25b87c75989c95b ,
                        0x87c65d86c45c86c15581c05784c8527fc25280c04f7ec84c7bcd4977be4371bf ,
                        0x436ebd3e6ab7395fa73f619c1a283b0000010202020402010202020202020202 ,
                        0x02101010100f11010101040202040201000000090a0e3250814369aa3f6aa94c ,
                        0x76bd4e79b8507dc1507dc6507bba3e6787385c6e23433015351622463a1c4131 ,
                        0x1a391e1b3c211c392219381d16371c17371e193920193b1d163a1c1838191634 ,
                        0x112e5458365e711d41231b41232145271d43271a3f2519381726463b486b974c ,
                        0x77b6527ebe507abf527fc35481c5507dc0527ec5557fc45180c44e80c84e7cc3 ,
                        0x4e81ca4d7bc2497ccc497ece4377c44c74b58e9db7ece9e1fffffffdfffffffe ,
                        0xfdfcfefefefffdfffdfffbfbfffffefffefefea6b2be3a6cb44676c44b78bb52 ,
                        0x7ac25079b75280c75986c35686c65482c25680c35680c55582c65081c74c7dc7 ,
                        0x4c78c54878c04778c44172be4071bd3e6bb43a58873d5b8c15212d0000000303 ,
                        0x0300020302020202020202020210101010110f0002030202020003010301010d ,
                        0x13183d5f9a3e63a1456fb44974b74972b0537bbc5581c8446c881b40181e481f ,
                        0x1e47271b3e232a50502f596c193b1d1838201942231c47261a40241f44241d3d ,
                        0x241c382118371c1c3c293e677e5585d347749a1d4724244e2b1f4b261e47221e ,
                        0x44202042373962894d75c04b78bc4e79c24d7cc64e7fc5517cc5537fc65380c9 ,
                        0x5180c45083cc5082ca5380c3527ec55183cf5281cb5481cb4f83c64f7ec84774 ,
                        0xbd6883a5bcc4cbe9eae0f6f7f5fffffefffffffefdfffcfefefffffcc8c9d358 ,
                        0x77aa416bb64373bb4b77c44b78bc4d7cc6507bbe4d81c3527cc14c7cbe507dc0 ,
                        0x527dc04e7bbf4f78c14b7abe4a7cc84474c24471ba456eb73e6cb93f68b13a59 ,
                        0x8c3d5b8c16203102000004020203020402020202020202020210101010101004 ,
                        0x0202040202020301000001090f16365c9d4367a7446baf4771b44979bb4e79b8 ,
                        0x5480c74f7faf2d5e4221532b264f2f204d2c193d1f1c3e261c39221c39221c3b ,
                        0x201d3e231f4325264b2b194528183a1b2345345478a05f90d65989d12a4c4517 ,
                        0x391a1d4025244f443765703a668f406eae4671ba4574be4476be4676be4b76bf ,
                        0x4976bf4c7ac14d7abe4d7ac3507dba517ec2537fc65280c75481c55082ce5382 ,
                        0xc64e80c8507fc9507dc05082ca4776ba4b75bc5778b0768eb2d7d8dcfffffffd ,
                        0xfffffffffbb8bfc24b71ab3966b7416bb84172be4172be4578c84976c04774bd ,
                        0x4d77be4977be4a75b44a77ba4a77ba4b78c24878c04774be446eb14068a33f69 ,
                        0xaa3f67af4068a93e67a63b5b9034547f121d2502000002020200030102020202 ,
                        0x020202020210101010101001010102010300000013161b3c506f3c5b903d629c ,
                        0x3e67a5436db04673b74e78bb517db35983c4406d7b224e2722502c214b2c193f ,
                        0x2316391f1a3a21203d231d3f211a3c241a3c241d3e231c402218371c33575761 ,
                        0x8dca507fb237606319391a18371c325569416aaf416dba416cbb426cb7426ebb ,
                        0x426db64371be4370b44771b64473bd4774be4a73bc4a77c04974bd4b79c04c79 ,
                        0xbd4f79be507cc34f7fc7507dc0527ebe507fc34f7cc54d7dc54e7cca4876c33e ,
                        0x6fb93a6bbb4f71a7b2bbc8d9d7cd95a1b3416aa93c69b33f6dbb3d6ebe416bb6 ,
                        0x406ebc3f6cb54570bf4272c04670b54775bc4774bd4872b74873b64372bc4973 ,
                        0xc0446cb43861983c60963c629c3a63a23e68af3f6aad3d619d395a8819233403 ,
                        0x0303020000000202020202020202020202101010120f11000202020202000000 ,
                        0x282d367291c8395e9a395e963d63a3416dad4a71b54a74b9456e952449352a4f ,
                        0x3f1c44281e41261d42221d3b1e1f452726502d22513022472d2140251d40261f ,
                        0x4025203e25193c2222442c2c5546274b3d19381719381b284a494368a0416eb7 ,
                        0x436ab8416bb03e6cb33e6bb5446cb73e6db7416eb8416bb6426fb9416eb2436d ,
                        0xb24372b64471b54673bd4770b54875b94a75b84a77ba4b79c0497ac64875b84d ,
                        0x78c14a76bd4c77c04477bd4273bf3f6ebf396ec53f70ba496eb23c68b53c6bbd ,
                        0x3e6cba3d6cb63d6dbf416cb5426dbc406bba3f6fb7426eb53f6dba416db4416d ,
                        0xba3e6ebc416ebf3e69b23a5f993b5c8a3855823a57843a5b933c619b395fa03a ,
                        0x63a23a5f993c598c39527e35465b1f2530020000020202020202020202101010 ,
                        0x10110f0100020304020201002a2d32799cd43e64ac395b963b5c943c619f436d ,
                        0xae4c74bf3a607e1737181536141d40261c3c231f433727504b274b3526502d2c ,
                        0x5736254b2f254b2d25492b27482d2043291d402521402b436c8d557ebd395d6d ,
                        0x1d3b22294c423f6196416fb63e69a83c6baf416db43d6ab43e6bb43f6bb2436a ,
                        0xb83f6ab3406aad3e69ac3966aa3963b03a64af3764a7416bb0416bac3d67ae3e ,
                        0x6ab13c6db93d6bb84270be4472b93e70bc3d67b24071c14070c23566b43868b0 ,
                        0x3d6ab43c6ab84069b82e65b63065bc376bbe416cbd3d6cbd3167bc3566bc416e ,
                        0xc53f69b43d69b03e66b13162ae2f5a9d345b923b5b90324f7b284977314d7c30 ,
                        0x4a78345184385c923c5d953859913a5a8b3b5c8a3957883c557d606e853e4443 ,
                        0x000000020202000202101010100f1104020200010200000334373b8ba8d53c66 ,
                        0xab385a8f3e5d943e64a54270b74771b2476c9e3c5e763c657e4573a24570a14a ,
                        0x6d9833585c204732234826264c302f593a2e5938305b3a2a5a36244c30203f24 ,
                        0x21452f406a7d426e8b39616625472f2147232f544c3357573858634367ad4068 ,
                        0xb03f6cb63d6cbd426cb73f6aa93e69b83e68af486aa68ea6caa7bcdca8bee180 ,
                        0x9bc73b6ab45278b997b2de94adcf96b0d59db7e5809fd24675b988a6d791acde ,
                        0x4472b94776c09cafd5809fd2396ebe3e68b34f79be97b1e09cb5df5882c73968 ,
                        0xba4a75be98b3df87a3d93e69b23969b74b75ba7f9bcaa6bcd8adbce391a5c846 ,
                        0x68a47891b99aabc690a2bf8fa1be879ab54f6892314e7a35548135538233537e ,
                        0x3d5d9237588643587e778699373a420000000202021010100d0f0f0503020101 ,
                        0x0100000127221f6b7c974670b73b65aa3a61983d629e4063a54365a0456da84e ,
                        0x77c04f7ccd4c79bd4876ac325d661f46201f4527204329274b2d2d5837386342 ,
                        0x30563a30593a2d5a39244d2e20432821462c2547292347292e54383864402b4d ,
                        0x2f264724294c3e3d648b4a70b14270b73e6ebc4168ad4067ac3f69b03b60a445 ,
                        0x67a2839dcb6b86b8758ebae9eff67f9ed3587cb8fdfffed0ddf3678aca6e8bbe ,
                        0x6182ba4872bddbe5f6ebeff43a6abc94b2e3fcfefe8ca7da376ac03464b6a3bb ,
                        0xdffffffefaf9fdc2d4f13368bf4e76befbfdfddee6f33364b0799ad2fcfdffdf ,
                        0xe9f38aa4d27392c5819dc64d72b0c9d7edfbfdfe7993b7647ea663799c445e8c ,
                        0x36547d37537c37507c36517d38578436527b334d72334c747781931313130000 ,
                        0x011010101010100301010503020001010000013d51704d7ac33f6ec03e64a43d ,
                        0x61a13c619f3e64a44168ac4771b64677bd3461761e4323204a27214828274f2c ,
                        0x2755312e59382f5b3c2d56362e5c382e5c382a54312750301f42282449293157 ,
                        0x39315838335a3a315a3b34573d345a3c2b4e2c2c523c2f5647325b463d66973e ,
                        0x66b13e65a93c65a43f65a6395fa02f599c5679b99eb1dcfcfbfd8199c35278ae ,
                        0xfcfffdced5e95e83bd658dc84f76ba446db6dce3f2f2f8fd9ab6e5ebf3fae5ed ,
                        0xf44b75c23b65b2446fb8e9effaccd6eeaec3dffdffff597dbd476fb7fbfbfbdb ,
                        0xe3f43766bac0cfe9fefdff7196d02f5db13866b33965b23963aec7d4eafefcfc ,
                        0x7389ac5f769c5e759b39527c34517e3753823a56853852803552793653803454 ,
                        0x7d37537c8794aa1414140000001010100f0e120002020202020000002a2b2969 ,
                        0x7fa24166aa3a62aa3d68ab3f65a541639e3f639f3f65a6406aad456fb4335a69 ,
                        0x244e3d264e43234b32294f3d2c5b412c5a302b5d35346245355e422651302649 ,
                        0x2f20482c21442a20472d34583a35633f3861413d69453b6445375e3e355b3d27 ,
                        0x50312345271a3f1f2d515b436bb63c69b23f6bab3d64a23762a1a0b9d9f9fefd ,
                        0xfdfffedfe8f54167a75a7bb3fcfefefefdfff9fdfefefffa9db0d53661a4d6de ,
                        0xeffffffef2f2f8fcfbfdcad8ef4974bd3362b67494c9fcfefe8eafdd6f90cfff ,
                        0xfffe688dc9436fbcfafdffdae1f03b68b1d2dcedf7f9fa5681c43567b93f6ebf ,
                        0x3e6dbe3364aec3d0e6fffffff8fafafefdf9e0e3f13d5c8939527e35517a3653 ,
                        0x863859873b568239538132517e3d57854d5a701413150102000f0f0f10101000 ,
                        0x02020201030200002225296884b33a64a93e68ad3c68af3d67ae3c65aa4066a7 ,
                        0x4269ae4069ae406cb94370b94571b8436fbc3e68a3426db6386481355c653963 ,
                        0x88406ca73f6391284a391f41232042242747342e53492b4f2b315d3835603f3a ,
                        0x6344375f433a6344375d3f2f593620492a20411f264b433e68ab406bba3d6ab4 ,
                        0x3966b05c80c0fffeffeaeef96e94ce3e6cb9325da65e80bbffffffc5d1e34a6d ,
                        0xac4d71ad3f68a64169aadce5efe7f0f4426eb57999ceffffff9fb4d42d57a2ab ,
                        0xc0dbfffffe6a8abf4670b7f2fafa6e91d04372c4fdfcfedbe3f43668bbb0c5eb ,
                        0xffffff8ba8d43264b63464b2375fa7365b97c5d0e4fffcfe617fb0436ba64d75 ,
                        0xbd3b64a93d5c933754813554813856853c5581374f7d37548138517d485c7b5a ,
                        0x5f68000000101010101010020202030101191b1b61718846679f39609e3d639d ,
                        0x3b5f9f3861a03d69b03f69ae3e6baf426cb7416fbd3e6dc13d6dc54270bd3e67 ,
                        0x943965843258643053563b66633e656d325a5524483023432a2442291e412623 ,
                        0x4a312c52542d4f49325b3b3566403b6542325e3f2f613729522c254a2a1f3e21 ,
                        0x31546e416bb2416bb23f6cb63a69b35179c1e9effaf1f4f99cb4d894aedd456c ,
                        0xb05e82b8fffeffdee7f49db5d9a1b5d47992ba436aaedae1f2fefdffa9bfe3d3 ,
                        0xddeffffffe95abd44369a9eef0f8e2eaf1496fb03765b3c3d2eca5bde14576c6 ,
                        0xfbfcffdfe8f6356ac16289c7eef0fafffcffaec3df9bb4dcbbc7e34d6fa4c5d1 ,
                        0xe3fffeffafc2e7a1bae29cb0d94973ba3e6dbe4064a037547b36507f37558438 ,
                        0x548338547d3653804c658f3a4351000000101010121010000200000001576069 ,
                        0x6c8cb73e5f9e41609f395e9c395fa03b64a33e64a54166b03e68b3426cb73a6c ,
                        0xb8416cbd3c698e3257551e421e20482c355c643057472c522e294e2828482926 ,
                        0x492e24482a22452b32585c26493b2f585b2a524d2249292a52362a5032285136 ,
                        0x315957305853325a6d3a65983e6ab13d6cb63d6dbb3e6ebc436bbe3a6ab86a8e ,
                        0xcebdd1eaeaf1fadae7f74a74c1547abbc4d5f0d2daebd1dae8dee6f7a2b7dd41 ,
                        0x6bb6adc0e3d1deeed3e0f0cfdcf6a0b8e24772bb6184c4d6ddf09ab4d83963a8 ,
                        0x3663ac7998c5b4c4e14675bfc8d6edb4c3e33a6bb53967bb6187c7b6c7e8e4ec ,
                        0xf9eaf0f7bdccdf47649797acc8d0dceed4ddebd2d8ddc2ceda567aba3b69b640 ,
                        0x6cb3385b8d3b548037507c37537c344f7b345485506a92454d5a0200000e1010 ,
                        0x1210100200011315167a91b14869a83c62a338609b3b609e3d629c3c619d3f63 ,
                        0xa93e68af3e6bb43c6bbc406dbe406ebb335b671f4d222b55322b583733623b33 ,
                        0x5c3d365d4e305650304f3a264c3025482d1f45273969812b4d4c1b3a191c3e1f ,
                        0x2246282041261f3d243055694069b83c6cb4406bba416cbb426cb93e6bb53e6e ,
                        0xbc3d6cbd406bba3f6cb63e6ab73967b53e6cba3e6dbf3f6dbb3f70bc3d69b63c ,
                        0x66b13b6cb83f6bb83e6dc14270be3d6bb93b6abc406bbc3d6cbd3766ba396aba ,
                        0x416dba3e68b33e68ad4168ad3c67aa3d66af426bb43e68b33d6bb93c6cba4170 ,
                        0xc1406fc03c68bc3a69bb416dba3c609c395a8c37588a38588d34538635527e36 ,
                        0x5687395c9b4068a93d67aa3b64a94066a73c5b9037557e3a5783375584365281 ,
                        0x57729767707902000010110f0c11100102000d0f107993c2446bb53b62a73e65 ,
                        0xaa3c66a73c63a13d62a03e63a13a619f3c66ab406bba3f6dbb3f6ebf406bb42d ,
                        0x585528542f32593931573b3555423c697e4371a1365f612b4a2d29482d224426 ,
                        0x3c6984305857204a271f43251f42271b3b231b40203c6390436ab54069b23c69 ,
                        0xb33e6bb5406cb9416bb63f69b43f6cb6406cb93f6dba3a6bb73d6bb83c6db93c ,
                        0x6db93d6abb3d6cbd3f69b63c65ae3765b23d6ab43f6dba3e6ebc3c6dbb3666b8 ,
                        0x3769bc3c6dbb3c6cbe3d6bb93767b93a69ba406bba3a64b13663ad3e68b33966 ,
                        0xaf3864b13a6bb9386ec13868c03b6dc04170c1396aba3565b33a619f3c5b9239 ,
                        0x56833852803854833755843a5b8d375a923a5d954066a63e64a43b619b3b5b90 ,
                        0x3755863a55873b5786375284577196575f6c0000001210101210100200011314 ,
                        0x128fa7d1426caf3a62a34166aa3d62a0395f953b60983f65a63e68ad3f6cb540 ,
                        0x6ebc3f6dbb436fbc375d752d534d264c302f53352e553b2f583834583a35593b ,
                        0x2f54322c4e3023462b214227335c5f274b3d214527224e2a2347291d41232043 ,
                        0x2f3e66a14269b7406ab13f6bb23e69b23c6cb43e6bb5426cb3436bb33e6bb53c ,
                        0x6db93e69b8406dbe396ab83f6ab94069b23a6db63d6bb95f82c16f90cf4970ba ,
                        0x3d68b7386abc507ac57697d65984c73869b93d6dbf3f6ebf698ecc7095cf4a78 ,
                        0xc66d8fca678cc63f6ebf6188c67190c34b76c5557ec77694d5507cc3366bc257 ,
                        0x86ca799ad26786c33a5a8f35527f3855823755843956823b56883d5b94365a90 ,
                        0x395a8c395f9939619b3a5c913954863955843755843954863c59806973852b2b ,
                        0x2b0c0e0f0f0f0f020000151313859fc43c66a93c64a53c619f3b5d983b5d993a ,
                        0x5d9d3b5ea03963a43f6ab33e6cb93d6bb8426ec1335b771d3f1a264e32315c3b ,
                        0x365f40395e44395c41345c40305638294c3223422d1e42241f3c251c3a211739 ,
                        0x211b3a1d1e47271d3e231d3d2a3861884369b7426cb93d67ac436aaf3f69b03e ,
                        0x6db74068b33c67b03e6db73f6bb83c6bbc3b6dbf3e6fbb3e6ab73c6db93e6bbc ,
                        0x3c69bab8c9e4ffffff5e88c93566b63b69b75a81cbfffcfdd0dff23d6cbd3867 ,
                        0xbb678dcefcfefed5e0f4527dc6e4eaf7e5ebf84170c6d0dceef7faff4977c594 ,
                        0xb2e1fdfffe7e9fd74c7ac7e7f0faffffff7493c6365e9f3d5d923b5682375481 ,
                        0x3a56853855823554813a598c3a598e3d5d8e3c5c913f5d963b5a8f3856853956 ,
                        0x83365281355384677c981e22230f0f0f0e10110001000a090b6278a14169aa3a ,
                        0x60a03c609c3c5e933c5e943c5d953d60983b609c3e65aa426cb73f6cb63f6dbb ,
                        0x3e6baf29514526502d3159363663423e68493b634032583c2a50322246282041 ,
                        0x261a4024193d1f1a3a2118391e214741234933163c2016371c1e423139617e40 ,
                        0x69a8456bab426cb73e6bb5456cba4172be3d6cb6416bb23f69b43f6bb8416eb8 ,
                        0x446eb5416cbb3a6cb8406bb43e6abebacbe6fdffff678ed33066bb3364b43a69 ,
                        0xbabecce8fffffedbe4f2d8e5f5e5ebf8fffeff8faddc3c6ec0e5eef8deeaf637 ,
                        0x68bedae1f0b7c9e82d63c199b3e1fefefe809cdcb4c9e8fefdff9fb8e43665b6 ,
                        0x3a67ab3c5e993a578a3a58873453803953813b55833a55813852803455823a5b ,
                        0x8c3c5b8e3a588735548136537f3450793756833a4d6e34373c0e0f0d10101000 ,
                        0x00005f626a5871993e639d3a5f99395c943d5c913c5a933c5e943d619d3d65a6 ,
                        0x3d67ac416bb6416dba3f6abb3d6ba12e5a3d2d5c363664403c6948335f3b2e55 ,
                        0x403f71ad3b5e781d42221e40211a391c3154502543381939203460712f526019 ,
                        0x3b1c1b41251b381f1838152f56644467ab4167a73f67af416bae4069ae3d69b0 ,
                        0x3d66ab3e65af3c69b3426bb4416aaf416bb6406ab73d6ab33c68afb7c8ddfffe ,
                        0xffebf1f8e1e8f19bb5da3566bc7898cdffffffd2ddf199b2def6f9fef3f6fb51 ,
                        0x7ec24671c2e2eafbf5f8fdccd7edf2f6fb688fcd2e60b895b2dffffffccedcef ,
                        0xfdfffec3ceea416bb63d6cb04265a54267a5426caf436db43e649e3b5c8a3a57 ,
                        0x833a55813858893b58853956833b598a3b5a8f3654833a5581354f7d334f783d ,
                        0x53768b949e110f0f0d100e0000005560745c7daf3c619f3d619d385a95375a99 ,
                        0x3b5e9d3b609e385e9f3b65a83f6bb23e6bb43d6eb43d6bb83b6aa72a54422550 ,
                        0x2b26523528522f23472f3b66913d6fa42a4f471f3e23193e2416351a3d647235 ,
                        0x5a5e18371a305a61254d421c3c23224b2c2146261c3f25365d8a3f6db43c6bb5 ,
                        0x4269b33b64a93d65a64265a73c62a23d66a54066a73e66a73d68b13e6bb4416b ,
                        0xb83d6ab4426cb7bacbe5ffffffa1bae2c5d4eefffffe88a0d43866b3eaecf7cc ,
                        0xdaec547bbff9fbfcb7c9e83766b04874c1e0e9f3f9fbfcc9d7f3e5eff97fa0d2 ,
                        0x3364b29ab0d9fcfdfff9fefffcfeff7791b932528d3b5f953c619b3d60a23961 ,
                        0xa23e69a84368ac4164a3395e984064a04364923f5f9038588937568339568337 ,
                        0x527e38537f395683334e80445c865b67790c0e0e100d0f2025267284ad3f639f ,
                        0x3a6198375fa03c63a13b609c3d60a239619b3d66a53e67ac3e68af416bb83e69 ,
                        0xb8416cbb3e6dbf3c67b02d5d57255128264f29234f2b3258522144291a3a1b25 ,
                        0x493931534d133517446f84477089183514173b231c3e261c3f242648291b4424 ,
                        0x2a4e583e67ac3f67af4068b34168ac3a62a34065a34164a33e63a13c619d3b64 ,
                        0xa23a63a83b65a83e68ad4168ac4168ac3d66abb9ccedfcfbfd4672b97390c3ff ,
                        0xffff9eb4d72f5c9fa0b0d5f4f8f9b4c5e0ffffff708fbc32599d486fb3e4eaf7 ,
                        0xe0e9f73264b7d1ddefeaeef93b6bc396afe1ffffffa2b7d7eef3f6e1e9f04f6d ,
                        0xa63958953d5f953c5c913c60963e63a14166aa4169b14268b04268a943619c3e ,
                        0x5d903a588738578438537f38507e36517d34547d344e7c3d5c833d4a600f0f0f ,
                        0x0e0d0f7f899a5277af3c69ac3d66ab3e62a23e61a13c649f3e629e3d5f9a3b62 ,
                        0xa03b63ab3b67ae3d69b63f6dba3f6cbd3d6ebe4170c63963922c4e4d3259623a ,
                        0x6481244f421e40211b3c212b544d3d6675173617497a94507fab264431204027 ,
                        0x2140251e3e251d3d241e3d203054644164ad3e6aaa3d6aae3f67af3e66ae3d68 ,
                        0xab4266a63d629e3a5f9d3b5f9f3d629e4065a33c65a33d67a83f6bb23a69babb ,
                        0xcae4fffeffe3e9f6f2f6fbfaf7f96e88b62e56905c79acf9f9f9fffffedadfee ,
                        0x42659d375a99466dabe2e8f5fffffedbe4edfffdfcc8d7f13968bc97b2e5ffff ,
                        0xfc7694c37c97cafdfffeced7e443639837568b3956833c57893a5e943e639d41 ,
                        0x66a23e66a73e60a23b5f9b3d5e8f39588534547d3b528036527b35507c365380 ,
                        0x36517d3e5f8c2b35470c0e0e0c0e0f4950595074b0426db63d66af3e66ae4064 ,
                        0xa43d629a3c60963a60963d5d983d61a13e62a23963a63f66aa3d6cb0406ab541 ,
                        0x6db43d6bb8406ab53f69b4456fbc2c5147204f21204321406e90436d901a3918 ,
                        0x5681ac5580ab275041284f3628513125502f264c2e2144291d4327385f8c3963 ,
                        0xa63e61a53e62a23b669f3d60a23c62a33d64a23b609e3b629941649c4164a33d ,
                        0x64a83d67aa4369b13f69b07999cea6bdddaabde29cb6de6385c03a5d9f3c619b ,
                        0x3e6096899ec4a8bad17694bd365d9b3c64a5426ab290aad2a4b9d4adc0e398b1 ,
                        0xdd5079be366abd6e91d0a6bbe15e85c93863ac88a3d5b2c5e66280b135548738 ,
                        0x58893858893b58843856853656873b55833c59853a58873455833957883c5d8f ,
                        0x3a588939568336537f35527e365380435d854a57670f0d0d0e0f0d12161b5472 ,
                        0xab416eb83b69b03e68ab3c6cb43c67b03e65b042639b375c903b60943a5c973f ,
                        0x63a34368a63e66a73f69b03e6bb4416eb8456ec33f6fc14070c2365f862e5662 ,
                        0x38627f5284d0416d8c1838194c788f406b741b3d1e2750312a58342b54352c52 ,
                        0x36254c2c22462e3a5e8e426ab23e639f3f62a23f64a24165ab4169a43b65a641 ,
                        0x64a64164a64063a73e62a23c68a83c69ac3c67a63e68ab3a63ac3566b23762b1 ,
                        0x3764ad3a67b14063a73a5f973f5d9632569231548c365a963d63a33c64a53d64 ,
                        0xa83964ad3768b83664b13666b8396bbd3f6ebf3a6cbe3968b93c6bbc3c6cc435 ,
                        0x65b33562ab3b61a23b60943c598c3a58893b598a3c57893856853d57863d588a ,
                        0x3d5c8f3c5d95405e973a5f973b5b903b56893757823955843a59863e5275484e ,
                        0x55100e0e100e0e1d23286586be3f68b14068b34368b2406bb43d6ab43d6bb83f ,
                        0x6dbb3e67b03a5f9b3b5b8c3c5e9a405d9a385d954165a5456aae3f69b0406ebc ,
                        0x416ebf3d6dbf4473c94775c34779b55181c933575714371521452d25462b2346 ,
                        0x2b2c53332c55352c56372b5631284f2f274b3d3e64a43f67af406aad3d67a837 ,
                        0x60a93763aa3e6aa73a67ab355fa6375a9c3761a23761a43861a64062a83e65aa ,
                        0x3e6ab13560af3264b03f6dba3c6bbd3964b53b61a13962993f64983d5f9b345b ,
                        0x9f3e65a94369aa3e67a63b65ac3360b13160b13564b63c6cbe3b6dbf3367ba35 ,
                        0x65b73c6dbd4070be3367ba3362b63869b93e68ad345ca4345a9a39598e395683 ,
                        0x3554812f517f3251883b5c9431579131599433569631538f385a8f3e598b3956 ,
                        0x833653803756834c60831213170e0e0e0e0e0e272c2b7e9ccb3b68b13c6ab83f ,
                        0x68b73f6cb63d6bb83e6ab13d6bab3d69b63e6cb93f67af3b639e3c5b903a5c92 ,
                        0x3b62a03f6ab3406cb93e6fbd416fbc4170c43d65a0264b412042242e51471d3c ,
                        0x2718381f1c41211d4a292851322a53343158482d55392e5535234828264f4a40 ,
                        0x69a83e64a43f66ab4b70aa92a8cc708fc23762ab5176b499acd295aed84e79b8 ,
                        0x7c9ed493acd63f64a23761a2446aab94a7d292a6cf436eb74f74b094a8c7708a ,
                        0xaf3d5c93365691657cac9cb2d56387c33865ae385fa36483b8a8b7d7b7c6e085 ,
                        0xa5d6456fbc406db77a99ce92add9476eb83963a87e9dd0a2b9e67090c5486eae ,
                        0x93a6cb8099c13e63a13a599036558c7f95b892a7c34b6594738db19dafce9eb2 ,
                        0xd1a0b3ce5874a33454893655823654833857845668872121270e0e0e0c0d113b ,
                        0x3b3b90a3c83c64ac3c69b23e68b33f6dbb4269b73c68b53c63a73e68ad3d6ab3 ,
                        0x3d6dbf3d6eba3d6bb8426ab23f66aa436dba406ebb406fc03e6cc04270c83e6b ,
                        0xae2a4b471f433217361917381d1a3a221c3e1f2750302a59332e5b4a4671b037 ,
                        0x605b2750301f4424385e81446fb8406cb33b68b25c81c5fbfdfea6bbe1345fae ,
                        0xa0b8dcfffefffffcfd6185c58ba8d5fefffd5c7eb94871af7896c5ffffffcbd9 ,
                        0xec345d9c667fb1fefefeb4bfda365991375fa097add6fffffe7e98bc2d56958a ,
                        0xa2ccffffffe8eef9d8e4f6fdffff809bce3361aebcc7e5eef4f9496cac5778b7 ,
                        0xf0f6fbfffeffb5c6e13c66b1d7e1f2f4f8f96287c54c75ba6382b9f2f5faeaed ,
                        0xf2415b8ac0cadcfdffffd7dbe6e5eaf36b87b02e53873d548438558238578449 ,
                        0x5e8467707e0c0e0f0d0e0c0706024759764f73b33c66a93e6cb93968b23d6bb8 ,
                        0x426cb73d6ab33c62a24067b23b6ab43f6ab33c6ab8426cb93a69ba406dbe3e6e ,
                        0xc03f6dc13b6ece3a67922b56532249391d442f1f3c251a3b261a391e1a3a211a ,
                        0x3d2220411f355c6b4471bb3b678c27512e22461a234b39426aab436eb13f6ab3 ,
                        0x547abaf9fbfba0b6df4d7cc6ecf3fcfffefffbfdfe658ac84e77c0eff1fbeef4 ,
                        0xffe7eefdf3f7fcfffeff7b97c02f54986483b6fdffffb2c2d93a5e9e31548c8e ,
                        0xa1c4fcfffd7c8fb2456295f0f4f5e9eff4547ab03963a6b3c1d8dce4f13762a5 ,
                        0xb4c3ddeff4f74268a9acbfdafffffcffffffbdcae4305eac97add7ffffffedf3 ,
                        0xfae7f0f9eff2f7fefefe9aaec72b4a7dc7d0daebeef642649f4267a13d5e9639 ,
                        0x5b9038578a3a53853954863e5e8f4f648334373c100f1102000044536d557fc6 ,
                        0x3764a7375a923d5f9a3d6cbd3d6fc23d6ab43a6bb93f629a3f69ac3c6bbf3f6e ,
                        0xc0406cb3436cb1406eb53e6fbf3b6fc23f6eb838658b33628e294c5019371a24 ,
                        0x4c3a22462e183b201a371e1d3d2a3d5e78496ea2436eb74171c33e6b9e346183 ,
                        0x2f5a81436aae436cb53b68b9547ec5f8fcfda7badfc7d8eddee5f4a5bde1fdff ,
                        0xff6784bd325e9e9fb5d9d9e0f390abd7ffffffcfdbed3f6cb03256926381b0ff ,
                        0xfffeb0bfd935518730549090a8d2ffffff7c94c25c76a4fffefebfcde4315ba0 ,
                        0x305392738cb4eef3f64c6ca1aec1d6eaedf591a6c6fbfbfba2b7d6ecf0fbbccb ,
                        0xde345c9d496ba1d8e0edeef4f98ba7caf2f4fceaeef94f74b2365893c8cfe3e8 ,
                        0xedec2b486f2f4b743b59904263a23f629a3c5e933c588e36547d384e711e2129 ,
                        0x110f0e19191957657b3f5e8b3c5b923e63a13e64a53d66ab3f69b43c6cba3d6b ,
                        0xb93f6dc13e6cb9406cb33c69b3436db43d6fb7446dbc3e6ebc4173c6335b741e ,
                        0x40212b514b24483a203c1e2e5662284f40224e291c451f20482c436b9f476fba ,
                        0x406bae3f66ab4168ad4069ae3f68ad3f68ad3f68ad3a65a85679b1f2f5faecef ,
                        0xf7fdfffe7a95c7839dcbffffff6685b8325ca15375b0d9e1ee97accbfffeff7a ,
                        0x97be355898315ca55e83bdfffeffb1c1de315b9e3356988ea0c5fdffff7992b4 ,
                        0x4e6c9bffffffd3dce6395b902e589b95aed6e8eff241639fb1c0dafbfdfef3f7 ,
                        0xfcd0d7e84d6ca1e8eef5bdcbde385c9c305da18ea8ccffffff93aad7fcfdffa4 ,
                        0xbadd305494375487c4cedfe7eef13657853d588a3858833654833a5784395382 ,
                        0x3956833955844c668e2e333c0f0d0d3339406276993d60983b5d983a5d9c3c61 ,
                        0x9d3d66a54067ac3f6ab93f6eb83e6dbe3f66ab3c6db3406ab5426cb3406ab141 ,
                        0x70b44371b84775c3376486254c26244e2b2449291d452c335f6c234739244838 ,
                        0x2c58511b401e2f5861426eb5436bb3426db63e6eb63f69b43d66ab3e64a43f64 ,
                        0xa23b609e5476acf6f3f5fffeffcbdaea3b5f958ca0c3fffffe6581b0355a9447 ,
                        0x6699f5f5fbffffffe5e8f047669d3d629e90aacfc4d5f0ffffffe0e7f0abbedf ,
                        0x5776b38da5c9fffffe8194b7315284b1bdd9fffffebac8da9bacc7fafbffa2b7 ,
                        0xd72f5798afbed8fffffffefffd6f88b047669bf3f6fabcc9df3357933b5b9d50 ,
                        0x77aeedf2fbfafffdfafcfd638ccb3d67b43b5f95cbd4e2eeedef395284375586 ,
                        0x3657853c54823855823957863753823553825c769b3b3e46373b406174953e59 ,
                        0x8b3b5b963b5b963e629e4064a44065a13e66ae3e6db1406ebc4071c74072c441 ,
                        0x6eb84670b3446eb3446baf456db54475bb487bc4416ea723502f2658362d5c35 ,
                        0x2e5a362a4f2d1e40212f58613e7297224a212449454471bb406bb4406ab13e68 ,
                        0xaf446eb53f67af3f64a83c609c375a924c6a99bdc7d8c9d0df7185a42a4a7f78 ,
                        0x8eb2ccd4e558729a335486425e94b2c1d4d7e4f48ba8cf365f9d4367a7aebad2 ,
                        0xcdd6dfc7d5eccbdaf4d4e0ec6382b57691c4d0d7ea6e85b232578f45649b98ad ,
                        0xcddee6f3e8eef5bbc7d94b73b43665af92a8d1d3ddefabbcd73e5e9349699eb8 ,
                        0xc5db94a8c13c5d954166a4406aaba4b8e1dae4f6a9bfe34072c44070c23d6ab3 ,
                        0x9bb3d1b3bfd139527c38547d39527e38517d3956893b568239527e35527e5d71 ,
                        0x9440434b292f364967903d5c913c5b92395f993b5e9d3c5b983c619f4064a441 ,
                        0x65a54271b54474c24276c9477bc84c76c14c76b94e76b74974b34d77b84f79be ,
                        0x517bc0325d5a264f2f2b5135254b2f25432a204027284d3d315e62204420294e ,
                        0x564468a8426cb14269ae3f68ad406aaf416caf3f66aa4168ad436cb53d67aa38 ,
                        0x5fa33b5e9d3c5d95416097365a903658933c63a7456aa84063a33e63a740629e ,
                        0x395e963d5f9b3d619d3a5e9a3960a43b60aa3b63ab3b61a13d62a03b5e96395d ,
                        0x93395c943c5d953b5d99345e993961963f619c375ea23a64a53f67a83f68b13d ,
                        0x68ab3e69b2426ab24064a04167a7406cab476eb34870b84476be4377c34776c8 ,
                        0x4676c84477cd4077c84778ce4174c4416eb84167a13b59883c54823e5b873c55 ,
                        0x813c557f38517b324f7b586f8f3d41462a32394f689a3b5992385a8f3c5c973d ,
                        0x619d3d629e3f5fa03a5f993b639e4771b84b76bf4a79c3497dd0507ec54e77b5 ,
                        0x4f7bb85179b3517ab1537fc65381c8456f922750302e5a352f5838294f33264c ,
                        0x301f4325254c2c1f401e315972406db73d6bb23f6cb6436db8446db6406ab141 ,
                        0x6bb2416ab3406ab1426fb3426bb44167a83f69ac4168ad3e69ac4167a73e6298 ,
                        0x3e5d923a5e94405f943e61993d6098435e963b5d923a5c973a629d3c619d3f5f ,
                        0x943a5d953a5c923a5b8d3756893957863a5a8b395b913c5e943b609a40609b3e ,
                        0x61a03c629c4267ab4169b1436db23d6fb74771b64777bf476fb04772b14873b2 ,
                        0x4c77ba527cc3507cbc4d7ec4507bc44d7dcb4c7ed14b7ece4b7dcf4c7bcd4876 ,
                        0xbd4370b346679f3e5d923a588139527e385483304d796075944c4e5640434b5b ,
                        0x7bac395c9439588d3b5d8b3b5c8d3b5a913c5d8f3c598c3a60964369a94d77ba ,
                        0x4d7ec45585cd5a87d05885c8557fba5884c15780be5984c75b86c94e78a7244b ,
                        0x2b25522b2c55352d5734284e3024472c21452722483c4267a5436faf426aab42 ,
                        0x6bb0426eb53d6ab3426db6426bb4436db4446eb5416fb64972bb4973b84673b7 ,
                        0x4973b6426dac456aa84368a6416aa9456aa6395f953c5a8b3e59853c5c853b59 ,
                        0x8a3c5a8b405b8d395a8b3b5a8d3c588e3e58873654833a55813957883a56853a ,
                        0x59863d5e903b5b903c5d8f3f6098405f9c4066a6436eb1456fb2446fb84770b5 ,
                        0x496da3476ca05274a95576ae4e76a75678ad5378b0577fba5881bf5985c05b8a ,
                        0xce5589d65486d24f83d04d7ec44c76bd456eac43659a425f8c3758863c5a893a ,
                        0x55876b81a54b51584a4c576781a93b5c943f60983d5e903d5b8a3a5581385582 ,
                        0x3655823e5c8d476dad4f77b2527dbc5883c25d89c85f89c46186c25f88c65c87 ,
                        0xc65b84c25e85c35986c947719b3662632c57362c58332c5b352e59381f462627 ,
                        0x4f4e4573c03f6fb74472bf3f69aa3f64a24167a8446bb04970b44973b64773ba ,
                        0x4b74b94874bb4774bd4573ba4775c24376c64977c54273bd4973b84771b24477 ,
                        0xbd4875be4c76c34875be466fb8446ca7436398415c8f3d58843e57813d588438 ,
                        0x57843a54823555803856853755863753823c57833b5c8d365a903f619c4466a2 ,
                        0x3e629e4068a2466ca64c73aa4b71ab5072a85072a75076a65378aa5779a7587f ,
                        0xb35d82b45c86bb5c85bc5e86c05b85c05a87ca5c88d55382c6517bbc4a6dac3c ,
                        0x5e8c38567f38537f38537f314f80687ea24c5257474b5061779a33527f375586 ,
                        0x3b56883656873a5b893856853b5b8c3d6092496aa25076b0527fb85e87c55d8b ,
                        0xcb608ccb6288c85c87c06087c55a86c15887c55c87ca5881ca4f7cc6426c9133 ,
                        0x5d5c2c514127502a2450212c514d4570b3486fb34370b44671ba4770b9456fb2 ,
                        0x446dac4971b24c77ba4976b94c74b54975b54c77ba4971ab4971ac4a73b24e74 ,
                        0xb54874b44978bc4f79be5179c1527cbd4f7cc54d80c6507dc65380c9527fc84e ,
                        0x78bd4a73aa42618e3b58853d56803a54823a56853a54823a5784385581395480 ,
                        0x3953813c5a8b3f5f943c5f97416294436492436698496b99486d9f4a70a0526f ,
                        0x9c50729d5576a45679a15a7aa55b79a25c7daa6081b35e84b45b85ba6087bb61 ,
                        0x8bc65e8bd44e7ebe486e9e4061923e5e8f3e5c8b3d55835368875059673c3d41 ,
                        0x3d434a5d779c3655883a56853b578d3b5a8f3555863b5c8e4166a44368a4476d ,
                        0xa74e76b15980be5c8aca5d8ed45f89ca5984bd567bb35f84be5c88c85985c256 ,
                        0x82c1557fc4507bc44879c54978ca456fb02f5966325c69426aa44876bd4774b8 ,
                        0x4771b4446fb84470b7436db44771b64873b25079b84e76b74f75af5075af5277 ,
                        0xb3507ab54c75ac4e77b54a75b45078b94f79b4537fbe5b86c95b87c75b87c657 ,
                        0x87c75786ca5884cb5986c95180ca5381ce517cbf4a6ea444618e3e5a833b5780 ,
                        0x3a54823953813754813655823d5b8c3c588e3b588b3e5c8d40619243659b476c ,
                        0x9e4a6fa34c70a04c6d9b4f6e9b4f6f9a526f96547196536f91526e91526f9655 ,
                        0x77a2587caa567cac5377a75577a55076b04e79bc4e75b33e60953d5c893f5a7f ,
                        0x3a4456282b331210100e0e0e4b4e56647a9e3253853a588738578c385b933d61 ,
                        0x9d3e64a53e69a84169aa4b70a84f7bba5281c55981bb5783b95480bd5480bd56 ,
                        0x7eb84f78af537fbe5681c0527ebe4f7bbb4e7aba4874c14572bc4470bd426ebb ,
                        0x4872bd4a76bd446eaf476dad476ca8426aa5446ba94267a34368a44a72ad507c ,
                        0xb75080c2547abb4f77b24f78af527ab45079b84d7ab75681c4557fc05381c157 ,
                        0x80be5d82ba6186be6189c3618ec7638fcc608ac55e88c35f89c45c85c35a84bf ,
                        0x537ab147699e4060913f5b8a39578039557e37527e3a57843c5a893859873c5c ,
                        0x8d3f5e933f5e954464994566944768994366984567924b6b96486a954869904d ,
                        0x6a8f4d688d56739a567295536e90546e925b769b6681a34f6480576e8e54719d ,
                        0x4e6b984f69913f54734e5c6f171a1e0000000402010e1010393a3e60748d496a ,
                        0x983c5e993a5e9e3a60a03f65a54066a73f6aad456cb04b74b34873ac4a70aa4c ,
                        0x71ab4a6fa74f74ac4d75aa496ea6496ea84a6fab4a70a64970a74b70ac4d71b1 ,
                        0x4570b3446fb2416bae4169aa4069a84970b5476fb0456ea5466ba545669e4368 ,
                        0xa0456ea54567a24466a14871a85277b5507cbc4f78b75278b85179b45079b751 ,
                        0x78b65177ad5278ae557fba5b85ba5d84bb6089c0618bc0658dc2658dc2658fca ,
                        0x618ec76591ce6792d1608ece5a89cd5480c74d76b54269a03e60963f5d94405c ,
                        0x923856853758863c5b8e3b5a91395b913e61933e619340609143608d45619045 ,
                        0x638c46648d43618a4963885d73965a697c212b32363f493b45562e37442d333a ,
                        0x61676e3b3b3b606262363e4532394251596044474b0501000101010303030202 ,
                        0x021010100d100e3b3e46717f926681ad4e71b54169b13c66a93e68ad416eb742 ,
                        0x6cb34469a74468a4476ca44564974565964065973d5e8c4866954e6c9d47699e ,
                        0x4b6b9c5171a64c6ca14165a14167a13f68a64067ac4167a73f67a2436dae4872 ,
                        0xb34870b1496dad476caa4469a3476caa456bab466ca64a72ad4b79b94d79b950 ,
                        0x7abf4a73b24d74b24c76b14d76b44d75b04b73ae5277b15a7fb96088bd5e87be ,
                        0x6089c05e86b76085b9688ebe648ec3618dc2618ec7608fcd5f8ed25b8ace4f7c ,
                        0xc04b75ba4370b33f659f39598a405b8d3f5b913d5b923b5a8d39598a38568538 ,
                        0x58833d5a863d5982405d893f5c883c57835b72985b68761318191d1b1a020001 ,
                        0x0000010000010101010000000000010400000000000000010000000000000100 ,
                        0x0203000200020200020302020212100f10110f00000106070b484e5575849e5b ,
                        0x78ab466fb43a67b04064a43d68a73d639d3e5f913f608e405f8c3c57834c658f ,
                        0x5e74974a5b756170838a98ae94a1b76a78944356774162903a5f993b609c3e64 ,
                        0xa43d67aa416aaf416bac416cab426db0416aa9416aa94368a240629d41649c42 ,
                        0x639542649a4770af4872b54c77ba436ead466cac486fb4466fae4671b04a76b5 ,
                        0x5075b14e78b3537fbf5781bc5a7fb7587cb2547aaa587eae5f84b85f84bc6086 ,
                        0xc05c85bc5b83bd547dbc507abd4b75b64668a44061993e5f973b5d933a5a8f38 ,
                        0x56853956833855823956823e57833956833f5a864658756c7688525b651d2328 ,
                        0x1b19190000010101010203010202020003010003010202020002000203010303 ,
                        0x030100020202020303030001010203010203010202020202021010100e101101 ,
                        0x020002020200000003020053575c899dc04b6ead39609e3e5e933b5d923d5c91 ,
                        0x3a588947638c53647938404d636b722325260202020e1011100e0d0101012124 ,
                        0x294f5f763e58873e5a893f5e953e629e4066a73f67a83e69a8446db64371be41 ,
                        0x69b14067ab3c67a64269a73e649e3a609643639e466db1436dae446db2436dae ,
                        0x426aab4467a74469a3466ba9446eaf4973b84a79b74d77b24f75af5075ad5075 ,
                        0xad5173a85075a7567eaf557cb05379af4f75af4d75b04a73b1446eaf4066a63d ,
                        0x619d3d5e9639588b3757883653803953813b54803b5584375284395788566b8b ,
                        0x5861650806060c0d0b0000000101010103030202020202020301010401030201 ,
                        0x0302030103030301010102010302010304020201010103030302020202020204 ,
                        0x0202020202100f1110110f0201030203010202020200010b080322292c424d6b ,
                        0x667dad4e6a993e5a89395480485d7c57606d4042430000000000040200000101 ,
                        0x010000000200000301000000003e4a5c4866953b59883a58893c5d8b43659b44 ,
                        0x67a94167a73f64a24265a43e609b3d5e9640609b4164a34064a03f63993c5a93 ,
                        0x3c5e934063a23c65a34168ac4167a73c62a23e639f41639941629a4166a24a74 ,
                        0xb74670b34c71af476aa2426696466a984a6ea44a6da54b6ea6486d9f466b9d42 ,
                        0x689e4266a23d62a03b5e963b60943b5d93405e9739568937537c34517836547d ,
                        0x3b56823654834a61873f47540503020000000002020402010303030201030502 ,
                        0x0402020200020202020202020202010301010101030300010102030102030102 ,
                        0x03010101010002030002030202020202020f0f0f101010020103020202000202 ,
                        0x00020304010302000011161531353a1920235b697b515b6c2226270605070200 ,
                        0x01020202040202040201000202000202020202000102010002576474496a9739 ,
                        0x57863c56843d578d3d5f9a3962a13d65ad3f65a64066a6406aad3e6aa73e639f ,
                        0x3d5e963b5b8c3c5c8d3e5e8f3d5a8d3d5a873b5b8c3b5d924061993e61993a5f ,
                        0x993c60963d5f954160973d60983f66a4416aa94169a43f649c41639843659a44 ,
                        0x689e42659d44669b4364953e5e8f41609547659e405c9238578a3b5a8d3d5c8f ,
                        0x3654833954803d57853b56883954803656814c5f8065686d0000000002020104 ,
                        0x0202010302020202030102020202020204020205030302020202020204020102 ,
                        0x0202020202020202020202000203020202010101020301020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020101010101010b0b ,
                        0x0b08080800000002020202020202020200020004020100020304020204020202 ,
                        0x02020000015c667747608a34528136568b3d5c913e5c953f65a5436cb14166aa ,
                        0x4168a64066a74066a64268a84166a43e639d3a5b8d3f57853856853c57833a57 ,
                        0x833857843a57833b59883d5c8f3c5b8e375c903e5c8d3f5f9440629d4065a33f ,
                        0x659b4261983d5c934060953d62963a61953e60953d5c8f3a5986375584385281 ,
                        0x39568339538139538136538037558639578635527e344e7c38507e3c58873546 ,
                        0x6006040301020002010302020202030102020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020202020201010101010100000002020203030302020201010100020203 ,
                        0x0303010101050303000105000202000000596374465f873a57843b59883b5885 ,
                        0x38598b3c5d953e609b3d619d3d629a3e639d3f66a43d659f3e62984160974466 ,
                        0x9c456aa643649641619242608f3d5a863e5a893c57833c57833957863e5b883c ,
                        0x5b8839588b3a5e9a3e67a63e64a43d61973f6098395a923d619d4665a23e5e93 ,
                        0x39578639568333537e36538037537c324f7b35527e38537f3855823756833653 ,
                        0x7f39527c354f7d3958853a4a671c1c1c00000100010103030300010102020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x0202020202020202020202020202020202020202020201010102020203030301 ,
                        0x01010202020303030002030201030202020402010001010301000907073e4b5b ,
                        0x3f598135527f3755863b57863a55873a598c3c5c913e5c8b3e5e8f3d5e8f3f5b ,
                        0x84425d8943659b4769a44970ae4872b34e7cc94e7fc94a78bf4b72b6466da44b ,
                        0x6ba04465964166984160933f5e8b415f8e4368a44469a73f619d3e639b3a5d95 ,
                        0x3f619c37598f37548039527c3555803854833b528237567d384f7f3c56853652 ,
                        0x8135548138577e3f598139548034538033507733507d485f7f40444f08060503 ,
                        0x0100010200050303020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020202020201 ,
                        0x0101020202030303020202020202030303010101020202040202020202000203 ,
                        0x0202020200003f434e586c8b36537f37507c3953823758863958853c57893a57 ,
                        0x8a3c5c8d415c8e4461883f618f44669b4a6fa3527ab5527ec55083d34d85d453 ,
                        0x84d25587d35587d35384d24f80ca4f7ec2507abb4c77ba4c79bd4a73b2446ba9 ,
                        0x446eb1426dac426cad406baa3b68ab4167a83f66b03b62a03858893a54823a53 ,
                        0x7d36538037538236538037548137548139568235548134517e37537c39527e32 ,
                        0x507937517936455f474f602c30350a0c0d000301020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202030303 ,
                        0x010002040201020202020103020000080b0f62708739527a35517a31517c3655 ,
                        0x823957863858893a5a853c567e3f588244618d4b70a4547bb95481c55888ca57 ,
                        0x84c75a89d35889cd6190da6390d3618ed1608dd05b89c95d8dcf588bd45186cf ,
                        0x5481c4507abb517dc44d7abe4b75b64973b4456eb3446fb23f6eb23f6ab33f6a ,
                        0xb3416eb7406dbe3c66b33a5f993957863a567f36557c33517a37517f37527e37 ,
                        0x547b38517b364e78344f7b334e7a324f7636537f3b547c4457784c586419181a ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x0202020202020202020202020202020202020202020202020202010101020202 ,
                        0x0303030202020202020202020002020402020001010200002321205b6576344d ,
                        0x753f5c8937527e324e71354f743d59883f5e914263914b6f9f4e73ab567ab057 ,
                        0x7eb55d8cd05e8dd15e8dc56190ce5f8bcb668cc2648fc86b96cf6792c56995cb ,
                        0x6c99d66793d3628ecd618ed2548bd45485d15381c8547ec34c7ec64c7fcf4a74 ,
                        0xbf4b72b64073c9406ebb436aae406dbe3d71c44174c44072ca3b6ebe3865af3a ,
                        0x63a84268a93e65a93e68ab4068a94063a23b5d983d5a8d3555863b56883d5c8f ,
                        0x3e5f91475f7d4247500001000202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010020202020202020202020202020202020202020202 ,
                        0x0202020202020202020303030202020202020303030202020203010201030103 ,
                        0x040000002e323392a2b9345281315284385a903d5c913956893956833d5b8a45 ,
                        0x6aa65783ca5a89cd618dcd618ecb628ecb618bcc618dc85e86ba6388ba5e84b4 ,
                        0x6488b8688fbc698ec06c94c56d96cd6592cb6591d15c8dd15885c85781c6507e ,
                        0xbe527ec5507dc14d7bc24a78c54676c44272c43c72c74271c34270bd4171c33f ,
                        0x6fc7396bc5416ec9396dc73a6dcd3c6fcd3e6fc53c6bbd3c6bbc3e6bb53b66a9 ,
                        0x3c649f3e5d923855823654834d6180454b580604030100020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202101010101010020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020201030202020102000301010303032c353e566b8a8aa4cc6c8ec451 ,
                        0x76b04168a6416bac4573ba4a78bf5683c65987c75d8acd638fce6088c36389bf ,
                        0x5b82b65c7fb15c7dab5c7fab6285b16181aa6585b06388ba678fc96592d55b8e ,
                        0xd4588ad25580cf5082ca5383cb507ec54e7bc44b79c04678c04674c14374c442 ,
                        0x70be3f71c43f71c43e70c3396ec8386ec53a71c63e6dc13d6cbd3a68b53c69b3 ,
                        0x3f69b03b64a33c619d4366a83b65a8395d99385889586d8c6162660503030000 ,
                        0x0102030102020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202101010 ,
                        0x1010100202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020203030302020201 ,
                        0x01011c1c1c4343433038497484a16f8ebb4c72ac426dac476dad567fbe5884c1 ,
                        0x5f89c45e8bc85c8cce5d89c85d84bb5678ad5778a6577aa25777a25a7eae6085 ,
                        0xb96288be618dcd5b8bd35787cf5481cb517ec74e7bbf4b7ac4497ac84776c049 ,
                        0x77be4473bd4575c34271bb416fbc406cbf4871c04374c03e6fbd426cb73e6ab1 ,
                        0x3d67ac3e64ac4165a53d5e903d5a8d3c5d954268a93d66a53d64a23d5e964e63 ,
                        0x835a62690c090b00000100020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202021010101010100202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x020202020201010103030302020201010100000002000103010134393c424d61 ,
                        0x56719d496da94e78b34e75ac527db65580c35a84c95781c25780bf537ab14f76 ,
                        0xaa5176aa557db2547eb95884bf5682bf5580bf5281c54b7cc84a7cc84778c846 ,
                        0x76c44773ba4873bc4271bb4470bd4572bc3f6fbd426fb93b6bb93d70c03c6cba ,
                        0x3b69b63b67ae3f66a43b61a13c609c3e63973c5e943b5b903b5a8d3f629a3d61 ,
                        0xa13b5e9d395d993951753b444e05060400010001010102020204020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202021010101010100202020202 ,
                        0x0202020202020201010102020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202010101020202020202020202020202020202020202 ,
                        0x0003010202020000011312165a606b66788f5471985372a5597eba4974b74b76 ,
                        0xb54c75b44f78b74a72ad4c75b44c78b84e75b35077b54d76b54c77ba4977be4b ,
                        0x75ba4d77bc4776c04374c24473c5446ec14471bb4270bd4370ba446fbe4071bf ,
                        0x416fc34271c24171bf4470bd3e6eb63d69b03c62a33d5f9b3c5b923b598a3d5a ,
                        0x8d3a5a8f39588d39588f4b70ae4c74b5526da04c5a7000000102000000020201 ,
                        0x0402030101050302020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0210101010101002020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202030303030303010101 ,
                        0x0202020202020101010101010203010202020002020102000200002c2b2d2024 ,
                        0x2f6e7b8b4e60776b89b25070a54a6ca8446ba9446ea94870ab4368a64369a344 ,
                        0x70af4771b44870b14571b84773c04370ba4170c13f71c3416fbc3f70bc3c6ec1 ,
                        0x3b6cc23e6fc53b6ec44071c73c71c83d6ec43e6cc4406fc53f6fbd406ab74064 ,
                        0xaa3b609a35558a3b5b9039588d445e8c667aa35d70916676938795a79299a834 ,
                        0x373b020000040103020202010101030206000203020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020210101010101002020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0303030202020202020202020303030202020202020202020003010202020300 ,
                        0x020002030002020001000000000101011c1d213e4149515d6971859e4a638b47 ,
                        0x5e844b658d6186c05479b34167a73f69ac426bb0426aab3d63a34168ac3c66a9 ,
                        0x3d66af416bb84169b4406ab73b69b6406db63d69b03969b73568c43970c53d72 ,
                        0xc3406dbe436ebd426fb93b68ac446ba95d7cb3516a946d7d9439424b4d4f593f ,
                        0x4445000100070102060201040202020103030303010200050302010200000301 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020210101010101002 ,
                        0x0202020202020202020202020202010101010101020202020202020202020202 ,
                        0x0202020202020202020202020101010202020303030202020101010202020202 ,
                        0x0202020200020301030305030304020104020100030102030104020103010100 ,
                        0x01000e0e0e0d0b0a0d08090401032a2d3168788943546f8099c17896c55a79ae ,
                        0x486aa5476fb0476eb34469ad3b62a73c65aa436cb54a74b95176ba4a6aa54a6b ,
                        0xa35376ae5577b34b72b73e69b23b64ad4066a74c6eaa637eb166799c6270873c ,
                        0x42470d0c0e060809010101010200030204000202000202000203000202000200 ,
                        0x0002020202020402020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202101010101010010101020202020202020202020202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x0202020202020202020201010102020205030302020202020201010102010300 ,
                        0x0203000202020202000203020301000000000102000000010200040000010200 ,
                        0x010101282b2f2b2e362127327a889e7183a03f506a46597a7691c36f8fc03f56 ,
                        0x7c46567b7c8fb05761720404042527318f9eae62768f6785b65c84c55f81af6d ,
                        0x7b9161666f161917000000020000030002010101040201040202040103040201 ,
                        0x0201030301010602010402010202020202020202020201030202020202020202 ,
                        0x0202020202020202020202020202020202020202020202020202020202020202 ,
                        0x02020202020202020202020202020101011010101d1d1d0e0e0e101010101010 ,
                        0x0f0f0f1111111010101010101010101010101010101010101010101010101010 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x10101111110e101013111012100f1210100e10100d100e11111114100f0e1010 ,
                        0x0d0f10121010100f11100f1110110f0e0e0e0e0f0d0e0e0e0c0e0f0e0e0e100e ,
                        0x0e333131585c5d47484c1312140e0f0d0c0e0e0e0e0e100e0e0e0f0d0e0d0f0e ,
                        0x0d0f3d40453d4552343a410e0f0d0c0e0f0f0e100e10101311100f0f0f0e1010 ,
                        0x100f1110101010101010101010110f10110f0e10100f0e101110120d0f0f0e10 ,
                        0x1010101010101010101010101010101010101010101010101010101010101010 ,
                        0x10101010101010101010101010101010101010101010100f0f0f0d0d0d1d1d1d ,
                        0x250000000c00000007000080250000000c00000000000080300000000c000000 ,
                        0x0f0000804b0000001000000000000000050000000e0000001400000000000000 ,
                        0x1000000014000000
                    End
                    Picture ="NPS_Arrowhead_Small.jpg"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8400
                    LayoutCachedTop =60
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =2040
                    TabIndex =1
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =1
                    Left =240
                    Top =960
                    Width =6360
                    Height =540
                    FontSize =20
                    FontWeight =700
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Forest Vegetation Monitoring"
                    FontName ="Tahoma"
                    LayoutCachedLeft =240
                    LayoutCachedTop =960
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =1500
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =11056034
            Name ="FormFooter"
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
' FORM:         frm_Switchboard
' Level:        Form module
' Version:      1.06
'
' Description:  Standard module - main screen of the user interface, viewed at startup
' Data source:  tsys_App_Defaults
' Data access:  edit only, no additions, moving between records or deletions
' Pages:        pagMain, pagDefaults, pagAbout
' Functions:    none
' References:   fxnMakeBackup, fxnFileExists, fxnDeleteFile
'
' Description:  Form related functions & procedures for application
' Requires:     -
'
' Source/date:  John R. Boetsch, May 24, 2006
' Adapted:      Bonnie Campbell, April 22, 2018
' Revisions:    Simon Kingston, Sept. 2006 - 1.00 - added lookup for release information to look at tsys_App_Releases
'               ML/GS - unknown   - 1.01 - initial version updates
'               BLC   - 4/22/2018 - 1.02 - added documentation, error handling
'               BLC   - 10/22/2018 - 1.03 - updated Exit_Procedure > Exit_Handler, revised Browse functionality
'               BLC   - 10/23/2018 - 1.04 - updated to display BE version
'               BLC   - 1/30/2019  - 1.05 - added SetDbVersions
'               BLC   - 4/17/2019  - 1.06 - updated create to open PseudoEventList and EventAdd2 for event creation
'               BLC   - 8/16/2019  - 1.07 - add ADMIN
' =================================

' ---------------------------------
' SUB:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, April 22, 2018
' Adapted:      -
' Revisions:
'   BLC - 4/22/2018 - initial version
'   BLC - 10/23/2018 - revised to display BE version
'   BLC - 1/30/2019 - set db versions via SetDbVersions in Form_Load (if used here SetDbVersions returns NULL error)
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler

    Dim strCaption As String

    ' Set the application font to more closely match the forms.
    ' Useful in cases where the subforms use tables directly
    Application.SetOption "Default Font Name", "Arial"
    Application.SetOption "Default Font Size", 9

    ' Set the table-driven caption of the switchboard
    strCaption = Nz(DLookup("[Database_title]", "tsys_App_Releases", "[Release_ID] = '" _
        & Me!Release_ID & "'"), "")
    Me.Caption = strCaption
    
'    SetDbVersions
    'Set current user ID
    SetTempVar "CurrentUserID", ""

    'get user full name
    Dim sysInfo As Object
    Dim oUser As Object
    
'    Set sysInfo = CreateObject("ADSystemInfo")
'    If Not IsNull(sysInfo) Then
'        Set oUser = GetObject("LDAP://" & sysInfo.UserName & "")
'        If Not IsNothing(oUser) Then
'            Debug.Print "Display Name: "; Tab(20); oUser.Get("DisplayName")
'        End If
'    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3270 'property not found
            'continue on w/o error
            Resume Next
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_App_Releases)"
        Case 2001   ' Field name in DLookup improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_App_Releases)"
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_App_Releases)"
        Case -2147023541 'no connection to LDAP so can't get username
            'continue on w/o error
            Resume Next
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
                "Error encountered (#" & Err.Number & " - Form_Open[frm_Switchboard])"
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
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
'   BLC - 1/30/2019 - set db version labels via SetDbVersions
' ---------------------------------
Private Sub Form_Load()
    On Error GoTo Err_Handler

    Dim strLinkPath As String
    Dim strSQL As String
    Dim rstReleaseInfo As DAO.Recordset
    Dim varVersion As Variant
    Dim varAuthor As Variant
    Dim varAuthorOrg As Variant
    Dim varAuthorPhone As Variant
    Dim varAuthorEmail As Variant
    Dim varAppTitle As Variant
    Dim varAuthorOrgName As Variant
    Dim varLastReleaseDate As Variant
    Const cstrDefaultAppTitle As String = "No Application Title"

    ' Set the current back-end database path control according to the system table
    'strLinkPath = "Using.... " & Nz(DLookup("[Link_file_path]", "tsys_Link_Files", "[Link_type] = 'Back-end data'"), "")
    strLinkPath = Nz(DLookup("[Link_file_path]", "tsys_Link_Files", "[Link_type] = 'Back-end data'"), "")
    Me!tbxLinkPath = strLinkPath

    '10/23/2018 update
    SetTempVar "BEfilepath", strLinkPath

    'get application release information and fill in the appropriate text boxes
    strSQL = "SELECT TOP 1 * FROM tsys_App_Releases"
    strSQL = strSQL & " ORDER BY Release_date DESC;"

    Set rstReleaseInfo = CurrentDb.OpenRecordset(strSQL, dbOpenForwardOnly)
    
    With rstReleaseInfo
        If Not (.EOF And .BOF) Then
            varVersion = ("Version " + !Version_number + " ")
            If Not IsNothing(!Release_date) Then
                varVersion = varVersion & ("(" & !Release_date & ")")
            End If
            varAuthor = "by " + !Release_by
            varAuthorOrg = !Author_org
            varAuthorPhone = !Author_phone
            varAuthorEmail = !Author_email
            varAuthorOrgName = !Author_org_name
            varAppTitle = !Database_title
        End If
    End With
    
    varVersion = NothingZ(varVersion, "Version unknown")
    varAuthor = NothingZ(varAuthor, "Author unknown")
    varAuthorOrg = NothingZ(varAuthorOrg, "Organization unknown")
    varAuthorPhone = NothingZ(varAuthorPhone, "Author phone unknown")
    varAppTitle = NothingZ(varAppTitle, cstrDefaultAppTitle)
    varAuthorOrgName = NothingZ(varAuthorOrgName, "")

    Me.Caption = varAppTitle
    Me!lblNetwork.Caption = varAuthorOrgName & vbCrLf & "Inventory and Monitoring Program"

    If IsNothing(varAuthorEmail) Then
        varAuthorEmail = "Author email unknown"
        Me!lblAuthorEmail.forecolor = vbWhite
        Me!lblAuthorEmail.FontUnderline = False
        Me!lblAuthorEmail.HyperlinkAddress = ""
    Else
        Me!lblAuthorEmail.forecolor = vbBlue
        Me!lblAuthorEmail.FontUnderline = True
        Me!lblAuthorEmail.HyperlinkAddress = "mailto:" + varAuthorEmail
    End If
        
    Me!tbxVersion = varVersion
    Me!tbxAuthorName = varAuthor
    Me!tbxAuthorOrg = varAuthorOrg
    Me!tbxAuthorPhone = varAuthorPhone
    Me!lblAuthorEmail.Caption = varAuthorEmail

    'set FE/BE db version textboxes
    SetDbVersions

Exit_Handler:
    On Error Resume Next
    rstReleaseInfo.Close
    Set rstReleaseInfo = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Files)"
        Case 2001   ' Field name in DLookup improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_Link_Files)"
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Files)"
        Case Else
            MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' SUB:          lblNPS_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub lblNPS_DblClick(Cancel As Integer)
    On Error GoTo Err_Handler

    ' Upon clicking the NPS label, open the website
    DoCmd.Hourglass True
    Application.FollowHyperlink "http://www.nps.gov", , True

Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          imgNPS_DblClick
' Description:  image double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub imgNPS_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    ' Upon clicking the NPS arrowhead, open the website
    DoCmd.Hourglass True
    Application.FollowHyperlink "http://www.nps.gov", , True

Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - imgNPS_DblClick[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          lblNetwork_DblClick
' Description:  label double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub lblNetwork_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    Dim varAuthorUnitCode As Variant
    Dim varLastReleaseDate As Variant
    
    varLastReleaseDate = DMax("[Release_date]", "tsys_App_Releases")
    If Not IsNull(varLastReleaseDate) Then
        varAuthorUnitCode = LCase(DLookup("[Author_org]", "tsys_App_Releases", "[Release_date]=#" & varLastReleaseDate & "#"))
    End If

    If Not IsNothing(varAuthorUnitCode) Then
        DoCmd.Hourglass True
        If IsNetwork(varAuthorUnitCode) Then
            ' Upon clicking the network name, open the website
            Application.FollowHyperlink "http://www1.nature.nps.gov/im/units/" & varAuthorUnitCode & "/index.htm", , True
        Else
            'if it's not a network code, assume it's a park code
            Application.FollowHyperlink "http://www.nps.gov/" & varAuthorUnitCode, , True
        End If
    End If

Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - lblNetwork_DblClick[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxLinkPath_DblClick
' Description:  textbox double click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub tbxLinkPath_DblClick(Cancel As Integer)
On Error GoTo Err_Handler

    ' Upon clicking the current link path, reconnect back end tables
    DoCmd.OpenForm "frm_Connect_Tables"

Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxLinkPath_DblClick[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnEvents_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling, revise frm_Event_Add > EventAdd
'   BLC - 11/9/2018  - add pseudoevent check
'   BLC - 4/17/2019  - revise to open EventAdd2 for event creation
' ---------------------------------
Private Sub btnAddEvent_Click()
On Error GoTo Err_Handler
    
    ' Proceed to add event if the database is connected
    If fxnVerifyLinks() Then
        Me!Activity = "enter"
        
        'check if there are pseudoevents first
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Set db = CurrDb
        
        Set rs = db.OpenRecordset("qFrm_PseudoEvents", dbOpenDynaset)
'        If Not (rs.BOF And rs.EOF) Then
'            rs.MoveLast
'            rs.MoveFirst
'        End If
        If rs.RecordCount > 0 Then
            DoCmd.OpenForm "PseudoEventList"
        Else
            DoCmd.OpenForm "EventAdd2" '"frm_Event_Add"
        End If
    Else
        MsgBox "The database must be connected first", vbOKOnly, "Data Tables Not Connected"
    End If

Exit_Handler:
    DoCmd.Hourglass False
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAddEvent_DblClick[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnGateway_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnGateway_Click()
On Error GoTo Err_Handler

    ' Proceed to gateway if the database is connected
    If fxnVerifyLinks() Then
        Me!Activity = "enter"
        'DoCmd.Close , , acSaveNo
        'DoCmd.OpenForm "frm_Set_Defaults", , , , , , 1
        DoCmd.OpenForm "frm_Data_Gateway"
    Else
        MsgBox "The database must be connected first", vbOKOnly, "Data Tables Not Connected"
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnGateway_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnViewMetadata_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnViewMetadata_Click()
On Error GoTo Err_Handler
    
    If Not (Not IsNothing(fxnGetLocalMetadataFileName) Or fxnNPSDataStoreMetadataExists Or fxnDBPurposeExists) Then
        MsgBox "No metadata or purpose was entered for this database."
    Else
        DoCmd.OpenForm "frm_Metadata_display", , , , acFormReadOnly, acDialog
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewMetadata_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnExit_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnExit_Click()

'Updated: Simon Kingston, 2/26/2007 - Set up to compact multiple back-end files
Dim booLinked As Boolean
Dim rstLinkedFiles As DAO.Recordset
Dim strOrigDbName As String
Dim strNewDbName As String
Dim strSuffix As String
Dim intCount As Integer
Dim booSuccess As Boolean

On Error GoTo Err_Handler

    booLinked = fxnVerifyLinks()
    
    ' Prompt for backups, depending on system default settings
    If booLinked And Me!chkBackupOnExit Then fxnMakeBackup
    Me!Activity = Null
    
    ' Compact and repair back-end databases prior to exit, depending on
    '   default settings and on whether there is a valid link to the database
    If booLinked And Me!chkCompactBEOnExit And IsNull(Me!tbxLinkPath) = False Then
        
        Set rstLinkedFiles = CurrentDb.OpenRecordset("SELECT Link_file_path FROM tsys_Link_Files WHERE Backup;", dbOpenForwardOnly)
    
        Do Until rstLinkedFiles.EOF
            strOrigDbName = rstLinkedFiles!Link_file_path
            ' Don't do anything if the link path string is empty or isn't an mdb file
            If Right(strOrigDbName, 4) = ".mdb" Then
                intCount = 0
                ' If needed, loop through alternate temporary names until an unused name is found
                Do
                    intCount = intCount + 1
                    strSuffix = "_" & CStr(intCount) & ".mdb"
                    strNewDbName = Left(strOrigDbName, Len(strOrigDbName) - 4) & strSuffix
                Loop Until fxnFileExists(strNewDbName) = False
                
                booSuccess = True 'initialize the success flag
                DBEngine.CompactDatabase strOrigDbName, strNewDbName
                'if compaction was successful, then attempt to delete original
                If booSuccess Then
                    ' If successful deleting the original, uncompacted file the rename the compacted file
                    '   to the original name
                    If fxnDeleteFile(strOrigDbName) Then Name strNewDbName As strOrigDbName
                End If
            End If
            
            rstLinkedFiles.MoveNext
        Loop
    End If

    ' Compact the front-end db upon closing if the database is connected and
    '   if the verify tables on startup is not set (otherwise slower performance)
    ' Does not work with Access 2007
    If booLinked And Me!chkVerifyOnStartup = False Then
        CommandBars("Menu Bar").Controls("Tools"). _
            Controls("Database utilities"). _
            Controls("Compact and repair database...").accDoDefaultAction
    End If
    
    ' Close the application
    DoCmd.Quit acQuitSaveNone

Exit_Handler:
    On Error Resume Next
    rstLinkedFiles.Close
    Set rstLinkedFiles = Nothing
    Exit Sub
    
Err_Handler:
    booSuccess = False
    Select Case Err.Number
      Case 3356, 70
            ' The back-end database is already open when trying to compact ...
            DoCmd.Quit acQuitSaveNone
            Resume Next
      
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnExit_Click[frm_Switchboard])" & _
            "Exiting main menu..."
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' SUB:          btnReview_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub cmdReview_Click()
On Error GoTo Err_Handler

    ' Proceed to review and edit data if the database is connected
    If fxnVerifyLinks() Then
        Me!Activity = "review"
        DoCmd.Close , , acSaveNo
        DoCmd.OpenForm "frm_Set_Defaults", , , , , , 2
    Else
        MsgBox "The database must be connected first", vbOKOnly, "Data Tables Not Connected"
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReview_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnQA_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnQA_Click()
On Error GoTo Err_Handler

    ' Perform data validation if the database is connected
    If fxnVerifyLinks() = False Then
        MsgBox "The database must be connected first", vbOKOnly, "Data Tables Not Connected"
    Else
        Me!Activity = "validate"
        ' Make sure the user name isn't null
        If IsNull(Me!cUser) = False Then
            ' Prompt the user to confirm the current user name
            If MsgBox("Current user:  " & Me!cUser, vbYesNo, "Please verify user name") = vbYes Then
                DoCmd.OpenForm "frm_QA_Tool", , , , , , 3
            Else    ' Open the defaults window to change user
                DoCmd.Close , , acSaveNo
                DoCmd.OpenForm "frm_Set_Defaults", , , , , , 3
            End If
        Else    ' Open the defaults window to change user
            DoCmd.Close , , acSaveNo
            DoCmd.OpenForm "frm_Set_Defaults", , , , , , 3
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnViewMetadata_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnLookups_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnLookups_Click()
On Error GoTo Err_Handler

    ' Review and edit lookup tables if the database is connected
    If fxnVerifyLinks() = False Then
        MsgBox "The database must be connected first", vbOKOnly, "Data Tables Not Connected"
    Else
        Me!Activity = "review"
        DoCmd.OpenForm "frm_Lookups"
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnLookups_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDbWindow_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnDbWindow_Click()
On Error GoTo Err_Handler

    ' Show the database window.  To re-hide: DoCmd.RunCommand acCmdWindowHide
    DoCmd.SelectObject acForm, "", True

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDbWindow_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnBackup_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnBackup_Click()
On Error GoTo Err_Handler

    ' Start the database backup function
    If fxnVerifyLinks() Then
        fxnMakeBackup
    Else
        MsgBox "Cannot create a backup until the database connection is fixed", _
            vbExclamation, "Data Tables Not Connected"
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBackup_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnReconnect_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnReconnect_Click()
On Error GoTo Err_Handler

    ' Reconnect back end tables
    DoCmd.OpenForm "frm_Connect_Tables"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReconnect_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' PAGE NAME:    Application Defaults (pagDefaults)
' Description:  system defaults for the run-time environment
' Bound ctls:   various fields for displaying default values
' Unbound ctls: cmdChangeDefaults - opens a popup for changing default values
' Subforms:     none
' =================================
' ---------------------------------
' SUB:          btnChangeDefaults_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnChangeDefaults_Click()
On Error GoTo Err_Handler

    ' Perform data validation if the database is connected
    If fxnVerifyLinks() = False Then
        MsgBox "The database must be connected first", vbOKOnly, "Data Tables Not Connected"
    Else
        ' Change application defaults in a popup window.  Closing the switchboard
        '   first avoids data write errors upon exit that may occur if edits are made
        '   directly in the form
        DoCmd.Close , , acSaveNo
        DoCmd.OpenForm "frm_Set_Defaults", , , , , , 4
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnChangeDefaults_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' =================================
' PAGE NAME:    Database Information (pagAbout)
' Description:  database development and release information
' Unbound ctls: cmdReleaseHistory, cmdReportBug
' Subforms:     none
' =================================
' ---------------------------------
' SUB:          btnReleaseHistory_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnReleaseHistory_Click()
On Error GoTo Err_Handler

    ' View the release history form
    DoCmd.OpenForm "frm_App_Releases", , , , acFormReadOnly

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReleaseHistory_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnReportBug_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnReportBug_Click()
    On Error GoTo Err_Handler

    ' View the release history form
    DoCmd.OpenForm "frm_App_Releases", , , , acFormEdit
    
    '5/23/2011 Renenabled the bug reporting button for NCRN.  Expect to only uses in the Main DB.
    '
    'Dim strMessage As String
    ' Report an application bug - used to open the subform but better to track this info
    '   centrally than in distributed applications
    'strMessage = "Please call or email the developer with details of the bug.  The following information "
    'strMessage = strMessage & "is helpful when reporting a bug:" & vbCrLf
    'strMessage = strMessage & vbTab & "- application name" & vbCrLf
    'strMessage = strMessage & vbTab & "- application version" & vbCrLf
    'strMessage = strMessage & vbTab & "- name of the form/report you were on when the bug happened" & vbCrLf
    'strMessage = strMessage & vbTab & "- action, if any, you took right before the bug occurred" & vbCrLf
    'strMessage = strMessage & vbTab & "- screen capture of any error messages"
    'MsgBox strMessage, , "Report a bug"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnReportBug_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPlants_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnPlants_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Plants"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnPlants_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnTags_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnTags_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Tags"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnTags_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnAppend_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnAppend_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Switchboard"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAppend_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDataSummary_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnDataSummary_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Data_Summary_Basic"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDataSummary_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' SUB:          btnDataQA_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnDataQA_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Data_QA"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDataQA_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnXX_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub Command85_Click()
On Error GoTo Err_Command85_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Manage_Links"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Command85_Click:
    Exit Sub

Err_Command85_Click:
    MsgBox Err.Description
    Resume Exit_Command85_Click
    
End Sub

' ---------------------------------
' SUB:          btnUtilities_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnUtilities_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Utilities"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnUtilities_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnDashboard_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 10/22/2018 - update documentation, error handling
' ---------------------------------
Private Sub btnDashboard_Click()
On Error GoTo Err_Handler

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Dashboard"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnDashboard_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetDbVersions
' Description:  sets database version displays
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  B. Campbell, January 30, 2019
' Adapted:      -
' Revisions:
'   BLC - 1/30/2019  - initial version
' ---------------------------------
Public Sub SetDbVersions()
On Error GoTo Err_Handler

    'set Db BE Version
    Dim beDb As DAO.Database
    Set beDb = OpenDatabase(TempVars("BEfilepath"))
    'Debug.Print beDb.Properties("Db Version")
    SetTempVar "Db BE Version", CStr(beDb.Properties("Db Version"))

    'get front & back-end versions
    tbxVersionFE = Nz(CurrDb.Properties("Db Version"), "-")
    'tbxVersionBE = Nz(CurrDb.Properties("Db BE Version"), "-")
    tbxVersionBE = Nz(TempVars("Db BE Version"), "-")
    
    'hide BE label for now
    lblVersionBE.visible = IIf(Len(tbxVersionBE) > 0, True, False)

'    Debug.Print "lp=" & TempVars("BEfilepath")

'    Debug.Print "Db BE: " & Me.tbxVersionBE & " FE: " & Me.tbxVersionFE

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetDbVersions[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub
