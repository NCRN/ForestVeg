Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    AutoCenter = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =14400
    DatasheetFontHeight =9
    ItemSuffix =45
    Left =12270
    Top =5310
    Right =27210
    Bottom =14055
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='{75F868EE-CC52-42CC-84DE-0E6963E99CA7}'"
    RecSrcDt = Begin
        0xdca6db037508e340
    End
    RecordSource ="tbl_Locations"
    Caption =" Locations"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
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
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
            BorderLineStyle =0
            BorderColor =12632256
        End
        Begin Tab
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            CanGrow = NotDefault
            Height =8760
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =3180
                    Top =600
                    Width =1020
                    FontSize =9
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboUnit_Code"
                    ControlSource ="Unit_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"Unit"
                        " Code\" ORDER BY Enum_Code; "
                    ColumnWidths ="720;5040"
                    StatusBarText ="NPS Unit code"
                    FontName ="Calibri"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =600
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2220
                            Top =600
                            Width =870
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblUnitCode"
                            Caption ="Unit Code"
                            FontName ="Calibri"
                            LayoutCachedLeft =2220
                            LayoutCachedTop =600
                            LayoutCachedWidth =3090
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9840
                    Top =600
                    Width =1080
                    FontSize =9
                    TabIndex =3
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtX_Coord"
                    ControlSource ="X_Coord"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =9840
                    LayoutCachedTop =600
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8820
                            Top =600
                            Width =930
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblX_Coord"
                            Caption ="UTM East"
                            FontName ="Calibri"
                            LayoutCachedLeft =8820
                            LayoutCachedTop =600
                            LayoutCachedWidth =9750
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9840
                    Top =900
                    Width =1080
                    FontSize =9
                    TabIndex =4
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtY_Coord"
                    ControlSource ="Y_Coord"
                    StatusBarText ="M. Y coordinate (Y_Coord)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =9840
                    LayoutCachedTop =900
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8820
                            Top =900
                            Width =930
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblY_Coord"
                            Caption ="UTM North"
                            FontName ="Calibri"
                            LayoutCachedLeft =8820
                            LayoutCachedTop =900
                            LayoutCachedWidth =9750
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =600
                    Width =2040
                    Height =480
                    FontSize =18
                    FontWeight =700
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtPlot_Name"
                    ControlSource ="Plot_Name"
                    StatusBarText ="M. Name of the location (Loc_Name)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =60
                    LayoutCachedTop =600
                    LayoutCachedWidth =2100
                    LayoutCachedHeight =1080
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =3
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7176
                    Top =8025
                    Width =2013
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="Unique identifier for each sample location"
                    FontName ="Calibri"

                    LayoutCachedLeft =7176
                    LayoutCachedTop =8025
                    LayoutCachedWidth =9189
                    LayoutCachedHeight =8265
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextAlign =3
                            Left =6210
                            Top =8025
                            Width =840
                            Height =228
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblLocation_ID"
                            Caption ="Location ID"
                            FontName ="Calibri"
                            LayoutCachedLeft =6210
                            LayoutCachedTop =8025
                            LayoutCachedWidth =7050
                            LayoutCachedHeight =8253
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =8400
                    Height =300
                    FontWeight =700
                    TabIndex =6
                    Name ="cmdDelete"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =8400
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =8700
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =2580
                    Top =1200
                    Width =1620
                    Height =255
                    BackColor =-2147483633
                    ForeColor =1279872587
                    Name ="lblLink_to_Google_Maps"
                    Caption ="Show on Google Maps"
                    FontName ="Calibri"
                    HyperlinkAddress ="http://maps.google.com/maps?q=CHOH-0847@39.4730699,-77.7929945&iwloc=A&t=h"
                    LayoutCachedLeft =2580
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =1455
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12270
                    Top =600
                    Width =1170
                    FontSize =9
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtLat"
                    ControlSource ="Lat_WGS84"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =12270
                    LayoutCachedTop =600
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11220
                            Top =600
                            Width =960
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblLat"
                            Caption ="WGS84 Lat"
                            FontName ="Calibri"
                            LayoutCachedLeft =11220
                            LayoutCachedTop =600
                            LayoutCachedWidth =12180
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12270
                    Top =900
                    Width =1170
                    FontSize =9
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtLon"
                    ControlSource ="Lon_WGS84"
                    StatusBarText ="M. Y coordinate (Y_Coord)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =12270
                    LayoutCachedTop =900
                    LayoutCachedWidth =13440
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11220
                            Top =900
                            Width =960
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblLon"
                            Caption ="WGS84 Lon"
                            FontName ="Calibri"
                            LayoutCachedLeft =11220
                            LayoutCachedTop =900
                            LayoutCachedWidth =12180
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =14400
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="lblLocationHeader"
                    Caption ="Vegetation Sampling Plots"
                    FontName ="Calibri"
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =540
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =5520
                    Top =600
                    Width =1020
                    FontSize =9
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboAdmin_Unit"
                    ControlSource ="Admin_Unit_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"Unit"
                        " Code\" ORDER BY Enum_Code; "
                    ColumnWidths ="720;5040"
                    StatusBarText ="NPS Unit code"
                    FontName ="Calibri"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =600
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4500
                            Top =600
                            Width =960
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblAdmin_Unit"
                            Caption ="Admin Unit"
                            FontName ="Calibri"
                            LayoutCachedLeft =4500
                            LayoutCachedTop =600
                            LayoutCachedWidth =5460
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3180
                    Top =900
                    Width =1020
                    FontSize =9
                    TabIndex =10
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtPanel"
                    ControlSource ="Panel"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =3180
                    LayoutCachedTop =900
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2400
                            Top =900
                            Width =690
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblPanel"
                            Caption ="Panel"
                            FontName ="Calibri"
                            LayoutCachedLeft =2400
                            LayoutCachedTop =900
                            LayoutCachedWidth =3090
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5520
                    Top =900
                    Width =1020
                    FontSize =9
                    TabIndex =11
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtFrame"
                    ControlSource ="Frame"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =900
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4740
                            Top =900
                            Width =690
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblFrame"
                            Caption ="Frame"
                            FontName ="Calibri"
                            LayoutCachedLeft =4740
                            LayoutCachedTop =900
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7560
                    Top =900
                    Width =1020
                    FontSize =9
                    TabIndex =12
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtGRTS_Order"
                    ControlSource ="GRTS_Order"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    FontName ="Calibri"
                    Tag ="<data>"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =900
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =1140
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6780
                            Top =900
                            Width =690
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblGRTS"
                            Caption ="GRTS"
                            FontName ="Calibri"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =900
                            LayoutCachedWidth =7470
                            LayoutCachedHeight =1140
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =7560
                    Top =600
                    Width =1020
                    FontSize =9
                    TabIndex =13
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboSubunit"
                    ControlSource ="Subunit_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"Unit"
                        " Code\" ORDER BY Enum_Code; "
                    ColumnWidths ="720;5040"
                    StatusBarText ="NPS Unit code"
                    FontName ="Calibri"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =600
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6660
                            Top =600
                            Width =840
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="lblSubunit"
                            Caption ="Subunit"
                            FontName ="Calibri"
                            LayoutCachedLeft =6660
                            LayoutCachedTop =600
                            LayoutCachedWidth =7500
                            LayoutCachedHeight =840
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =13380
                    Top =120
                    Width =720
                    Height =300
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =13380
                    LayoutCachedTop =120
                    LayoutCachedWidth =14100
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =120
                    Top =120
                    Width =720
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    ForeColor =0
                    Name ="cmdBrowse_Edit"
                    Caption ="Edit"
                    FontName ="Calibri"
                    ControlTipText ="Close the data entry form"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =840
                    LayoutCachedHeight =420
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10080
                    Top =1200
                    Width =1080
                    FontSize =9
                    TabIndex =15
                    Name ="txtDatum_Zone"
                    ControlSource ="=[Datum] & \", \" & [UTM_Zone]"
                    FontName ="Calibri"

                    LayoutCachedLeft =10080
                    LayoutCachedTop =1200
                    LayoutCachedWidth =11160
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =8820
                            Top =1200
                            Width =1170
                            Height =240
                            FontSize =9
                            Name ="Label37"
                            Caption ="Datum, Zone"
                            FontName ="Calibri"
                            LayoutCachedLeft =8820
                            LayoutCachedTop =1200
                            LayoutCachedWidth =9990
                            LayoutCachedHeight =1440
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5520
                    Top =1200
                    Width =1020
                    FontSize =9
                    TabIndex =16
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtSlope"
                    ControlSource ="Slope"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    OnExit ="[Event Procedure]"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    Tag ="<data>"

                    LayoutCachedLeft =5520
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6540
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4740
                            Top =1200
                            Width =690
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label39"
                            Caption ="Slope"
                            FontName ="Calibri"
                            LayoutCachedLeft =4740
                            LayoutCachedTop =1200
                            LayoutCachedWidth =5430
                            LayoutCachedHeight =1440
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7560
                    Top =1200
                    Width =1020
                    FontSize =9
                    TabIndex =17
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="txtAspect"
                    ControlSource ="Aspect"
                    StatusBarText ="M. X coordinate (X_Coord)"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    Tag ="<data>"

                    LayoutCachedLeft =7560
                    LayoutCachedTop =1200
                    LayoutCachedWidth =8580
                    LayoutCachedHeight =1440
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6780
                            Top =1200
                            Width =690
                            Height =240
                            FontSize =9
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label41"
                            Caption ="Aspect"
                            FontName ="Calibri"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =1200
                            LayoutCachedWidth =7470
                            LayoutCachedHeight =1440
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =247
                    Left =120
                    Top =1500
                    Width =14175
                    Height =6810
                    TabIndex =18
                    Name ="TabCtl42"
                    FontName ="Franklin Gothic Medium"
                    GridlineColor =10921638

                    LayoutCachedLeft =120
                    LayoutCachedTop =1500
                    LayoutCachedWidth =14295
                    LayoutCachedHeight =8310
                    BackColor =14277081
                    BorderColor =9277327
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    ForeColor =4210752
                    Begin
                        Begin Page
                            OverlapFlags =119
                            Left =195
                            Top =2010
                            Width =14030
                            Height =6223
                            BorderColor =10921638
                            Name ="Page43"
                            Caption ="Plot Information"
                            GridlineColor =10921638
                            LayoutCachedLeft =195
                            LayoutCachedTop =2010
                            LayoutCachedWidth =14225
                            LayoutCachedHeight =8233
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                            Begin
                                Begin TextBox
                                    EnterKeyBehavior = NotDefault
                                    ScrollBars =2
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =3305
                                    Top =2345
                                    Width =10920
                                    Height =558
                                    FontSize =9
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="txtLocation_Notes"
                                    ControlSource ="Location_Notes"
                                    StatusBarText ="MA. General notes on the location (Loc_Notes)"
                                    FontName ="Calibri"
                                    Tag ="<data>"

                                    LayoutCachedLeft =3305
                                    LayoutCachedTop =2345
                                    LayoutCachedWidth =14225
                                    LayoutCachedHeight =2903
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =1
                                            Left =3305
                                            Top =2090
                                            Width =2220
                                            Height =225
                                            FontSize =9
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblLocation_Notes"
                                            Caption ="Location Notes"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =3305
                                            LayoutCachedTop =2090
                                            LayoutCachedWidth =5525
                                            LayoutCachedHeight =2315
                                        End
                                    End
                                End
                                Begin TextBox
                                    SpecialEffect =0
                                    OldBorderStyle =1
                                    OverlapFlags =247
                                    TextAlign =2
                                    BackStyle =0
                                    IMESentenceMode =3
                                    Left =1805
                                    Top =2345
                                    FontSize =9
                                    BackColor =-2147483643
                                    ForeColor =-2147483640
                                    Name ="txtDate_Established"
                                    ControlSource ="Install_Date"
                                    StatusBarText ="MA. Date of entry or last change (Upd_Date)"
                                    DefaultValue ="=Now()"
                                    FontName ="Calibri"

                                    LayoutCachedLeft =1805
                                    LayoutCachedTop =2345
                                    LayoutCachedWidth =3245
                                    LayoutCachedHeight =2585
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =275
                                            Top =2345
                                            Width =1485
                                            Height =240
                                            FontSize =9
                                            BackColor =-2147483633
                                            ForeColor =-2147483630
                                            Name ="lblDate_Established"
                                            Caption ="Date Established"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =275
                                            LayoutCachedTop =2345
                                            LayoutCachedWidth =1760
                                            LayoutCachedHeight =2585
                                        End
                                    End
                                End
                                Begin ComboBox
                                    OverlapFlags =247
                                    TextAlign =2
                                    IMESentenceMode =3
                                    ColumnCount =2
                                    ListWidth =8640
                                    Left =1805
                                    Top =2660
                                    FontSize =9
                                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                                    Name ="Location_Status"
                                    ControlSource ="Location_Status"
                                    RowSourceType ="Table/Query"
                                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                                        "umerations WHERE (((tlu_Enumerations.Enum_Group)=\"Location Status\")) ORDER BY "
                                        "tlu_Enumerations.Sort_Order; "
                                    ColumnWidths ="1080;7560"
                                    StatusBarText ="Status of the sample location"
                                    FontName ="Calibri"
                                    AllowValueListEdits =1

                                    LayoutCachedLeft =1805
                                    LayoutCachedTop =2660
                                    LayoutCachedWidth =3245
                                    LayoutCachedHeight =2900
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =290
                                            Top =2660
                                            Width =1470
                                            Height =240
                                            FontSize =9
                                            Name ="Label34"
                                            Caption ="Plot Status"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =290
                                            LayoutCachedTop =2660
                                            LayoutCachedWidth =1760
                                            LayoutCachedHeight =2900
                                        End
                                    End
                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =275
                                    Top =3223
                                    Width =7380
                                    Height =5010
                                    Name ="fsub_Tags"
                                    SourceObject ="Form.fsub_Tags"
                                    LinkChildFields ="Location_ID"
                                    LinkMasterFields ="Location_ID"

                                    LayoutCachedLeft =275
                                    LayoutCachedTop =3223
                                    LayoutCachedWidth =7655
                                    LayoutCachedHeight =8233
                                    Begin
                                        Begin Label
                                            OverlapFlags =255
                                            Left =275
                                            Top =2983
                                            Width =4980
                                            Height =240
                                            FontSize =9
                                            Name ="lblfsub_Tags"
                                            Caption ="Tagged Trees and Saplings"
                                            FontName ="Calibri"
                                            LayoutCachedLeft =275
                                            LayoutCachedTop =2983
                                            LayoutCachedWidth =5255
                                            LayoutCachedHeight =3223
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =195
                            Top =2010
                            Width =14025
                            Height =6225
                            BorderColor =10921638
                            Name ="Page44"
                            Caption ="Landcover Information"
                            GridlineColor =10921638
                            LayoutCachedLeft =195
                            LayoutCachedTop =2010
                            LayoutCachedWidth =14220
                            LayoutCachedHeight =8235
                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =2
                            WebImagePaddingBottom =2
                        End
                    End
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
' Description:  Locations entry form
' Data source:  tbl_Locations
' Data access:  edit, add, delete
' Pages:        none
' Functions:    none
' References:   fxnGUIDGen
' Source/date:  Simon Kingston, Sept. 2006.   Adapted to NCRN needs by Mark Lehman May, 2010.
' Revisions:    <name, date, desc - add lines as you go>
' =================================


Private Sub cmdClose_Click()
DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Error_Handler

DoCmd.RunCommand acCmdDeleteRecord
DoCmd.Close acForm, Me.Name

MsgBox "Record deleted successfully", , "Record Deleted"

Exit_Handler:
    Exit Sub

Error_Handler:
    Select Case Err.Number
        Case 2046 'command not available
            MsgBox "Unable to delete record.", vbExclamation, "Cannot Delete Record"
            Resume Exit_Handler
        Case 2501 'user canceled delete
            MsgBox "Delete canceled", , "Delete Canceled"
            Resume Exit_Handler
        Case 3200 'related records
            MsgBox "There are related records that prevent this record from being deleted.  Delete all related records first and then delete this record.", vbInformation, "Cannot Delete Record"
            Resume Exit_Handler
        Case Else
            MsgBox Err.Number & " - " & Err.Description, vbCritical, "Error - Form: " & Me.Name & " - cmdDelete_Click"
            Resume Exit_Handler
    End Select

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'check to see if a primary key is needed and add it (used for string GUIDs)
If fxnFormCheck(Me) Then
    'Removed the field for date below.  May want to reinstate.
    'Me!txtDate_Updated = Now()
    If Me.NewRecord Then
        If GetDataType("tbl_Locations", "Location_ID") = dbText Then
            Me!Location_ID = fxnGUIDGen
        End If
    End If
Else
    Cancel = True
End If

End Sub

Private Sub Form_Close()
'update control as necessary on calling form to reflect new location values
fxnUpdateControl Me.OpenArgs
If IsLoaded("frm_Data_Gateway") Then
    Forms!frm_Data_Gateway.Requery
End If
End Sub

Private Sub Form_Current()
    If Me!txtLon = "" Or IsNull(Me!txtLon) Then
        lblLink_to_Google_Maps.HyperlinkAddress = "http://maps.google.com"
    Else
        lblLink_to_Google_Maps.HyperlinkAddress = "http://maps.google.com/maps?q=" & Me!txtPlot_Name & "@" & Me!txtLat & "," & Me!txtLon & "&iwloc=A&t=h"
    End If
End Sub

Private Sub txtAspect_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
    'This routine requires the presence of the Keypad_Utils module.
    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'The two lines below should be changed to reflect the name of the keypad to open
    '    and the name of the control to be updated.
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "txtAspect"
    'The lines below should not usually be edited.
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
    Exit Sub
Err_cmdOpenKeyPad_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenKeyPad_Click
End Sub

Private Sub txtDate_Established_AfterUpdate()
    Me.fsub_Tags.Requery
End Sub


Private Sub txtSlope_Click()
On Error GoTo Err_cmdOpenKeyPad_Click
    'This routine requires the presence of the Keypad_Utils module.
    Dim strKeypadFormName As String
    Dim strControlToUpdate As String
    Dim frmFormToUpdate As Form
    
    'The two lines below should be changed to reflect the name of the keypad to open
    '    and the name of the control to be updated.
    strKeypadFormName = "frm_Pad_Num"
    strControlToUpdate = "txtSlope"
    'The lines below should not usually be edited.
    Set frmFormToUpdate = Me
    Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_cmdOpenKeyPad_Click:
    Exit Sub
Err_cmdOpenKeyPad_Click:
    MsgBox Err.Description
    Resume Exit_cmdOpenKeyPad_Click
End Sub

Private Sub txtSlope_Exit(Cancel As Integer)
If Me!txtSlope = 0 Then
    Me!txtAspect = "N/A"
    Me!txtAspect.Locked = True
Else
    Me!txtAspect.Locked = False

End If

End Sub
