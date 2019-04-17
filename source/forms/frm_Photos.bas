Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14820
    DatasheetFontHeight =10
    ItemSuffix =225
    Left =1575
    Right =16320
    Bottom =13425
    DatasheetGridlinesColor =12632256
    Filter ="[Location_ID]='{5B8AA79D-FDEE-4CA5-AC93-EA82744EDCEF}'"
    OrderBy ="[tbl_Locations].[Plot_Name], [tbl_Locations].[Panel]"
    RecSrcDt = Begin
        0x6851775efa47e440
    End
    RecordSource ="tbl_Locations"
    Caption ="NCRN Sampling Location"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            BorderThemeColorIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
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
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Page
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =1635
            BackColor =15527148
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Width =14820
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =0
                    ForeColor =16777215
                    Name ="lblEvent_Form_Header"
                    Caption ="Vegetation Sampling Photographs"
                    FontName ="Calibri"
                    LayoutCachedWidth =14820
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =13740
                    Top =120
                    Width =900
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdClose"
                    Caption ="Close"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Close the data entry form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13740
                    LayoutCachedTop =120
                    LayoutCachedWidth =14640
                    LayoutCachedHeight =450
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =7775995
                    HoverThemeColorIndex =5
                    HoverTint =60.0
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
                Begin CommandButton
                    OverlapFlags =93
                    Left =13740
                    Top =600
                    Width =900
                    Height =660
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdEditLocation"
                    Caption ="Edit Location"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Edit the current location record."
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =13740
                    LayoutCachedTop =600
                    LayoutCachedWidth =14640
                    LayoutCachedHeight =1260
                    ForeThemeColorIndex =0
                    UseTheme =255
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =9226162
                    HoverThemeColorIndex =7
                    HoverTint =60.0
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =180
                    Top =540
                    Width =2583
                    Height =360
                    ColumnWidth =1440
                    ColumnOrder =0
                    FontSize =18
                    FontWeight =700
                    Name ="txtPlot_Name"
                    ControlSource ="Plot_Name"
                    StatusBarText ="Unique identifier for each sample location"
                    FontName ="Calibri"

                    LayoutCachedLeft =180
                    LayoutCachedTop =540
                    LayoutCachedWidth =2763
                    LayoutCachedHeight =900
                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =13620
                    Top =1320
                    Width =1005
                    ColumnOrder =1
                    TabIndex =4
                    Name ="txtLocation_ID"
                    ControlSource ="Location_ID"
                    StatusBarText ="M. Link to tbl_Locations (Loc_ID)"

                    LayoutCachedLeft =13620
                    LayoutCachedTop =1320
                    LayoutCachedWidth =14625
                    LayoutCachedHeight =1560
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =2820
                    Top =600
                    Width =1860
                    Height =300
                    FontSize =10
                    ForeColor =1279872587
                    Name ="lblLink_to_Google_Maps"
                    Caption ="Show on Google Maps"
                    FontName ="Calibri"
                    HyperlinkAddress ="http://maps.google.com/maps?q=ANTI-0082@39.4771777,-77.7146946&iwloc=A&t=h"
                    LayoutCachedLeft =2820
                    LayoutCachedTop =600
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =900
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =85
                    Left =7440
                    Top =600
                    Width =1680
                    Height =300
                    FontSize =10
                    ForeColor =1279872587
                    Name ="lblLink_To_Plot_Photos"
                    Caption ="Explore Plot Photos"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"
                    LayoutCachedLeft =7440
                    LayoutCachedTop =600
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =900
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    ListRows =4
                    Left =180
                    Top =1020
                    Width =2160
                    Height =480
                    ColumnOrder =2
                    FontSize =18
                    FontWeight =700
                    TabIndex =1
                    ColumnInfo ="\"\";\"ddddd\";\"8\";\"8\""
                    Name ="cboEvent_Date1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Events.Event_Date FROM tbl_Locations INNER JOIN tbl_Events ON tbl_Loc"
                        "ations.Location_ID = tbl_Events.Location_ID WHERE (((tbl_Locations.Location_ID)="
                        "[Forms]![frm_Photos]![txtLocation_ID])) ORDER BY tbl_Events.Event_Date;"
                    ColumnWidths ="2880"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    Format ="Short Date"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2340
                    LayoutCachedHeight =1500
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9300
                    Top =600
                    Width =3963
                    Height =300
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="txtXY"
                    StatusBarText ="Unique identifier for each sample location"
                    FontName ="Calibri"

                    LayoutCachedLeft =9300
                    LayoutCachedTop =600
                    LayoutCachedWidth =13263
                    LayoutCachedHeight =900
                End
                Begin ComboBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    ListRows =4
                    Left =10860
                    Top =1020
                    Width =2160
                    Height =480
                    ColumnOrder =4
                    FontSize =18
                    FontWeight =700
                    TabIndex =6
                    ColumnInfo ="\"\";\"ddddd\";\"8\";\"8\""
                    Name ="cboEvent_Date2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Events.Event_Date FROM tbl_Locations INNER JOIN tbl_Events ON tbl_Loc"
                        "ations.Location_ID = tbl_Events.Location_ID WHERE (((tbl_Locations.Location_ID)="
                        "[Forms]![frm_Photos]![txtLocation_ID])) ORDER BY tbl_Events.Event_Date;"
                    ColumnWidths ="2880"
                    StatusBarText ="M. Starting date for the event (Start_Date)"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"
                    Format ="Short Date"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =1020
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =1500
                End
                Begin Line
                    BorderWidth =3
                    OverlapFlags =85
                    Top =1620
                    Width =14760
                    Name ="Line179"
                    GridlineColor =10921638
                    LayoutCachedTop =1620
                    LayoutCachedWidth =14760
                    LayoutCachedHeight =1620
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =20760
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =2400
                    Top =11595
                    Width =3840
                    Height =2880
                    Name ="img_060h"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =11595
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =14475
                    TabIndex =6
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =2400
                    Top =14595
                    Width =3840
                    Height =2880
                    Name ="img_180h"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =14595
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =17475
                    TabIndex =7
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =2400
                    Top =17595
                    Width =3840
                    Height =2880
                    Name ="img_300h"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =17595
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =20475
                    TabIndex =8
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2400
                    Top =12615
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_060h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =12615
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =13005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2400
                    Top =15615
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_180h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =15615
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =16005
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2400
                    Top =18615
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_300h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =18615
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =19005
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =2400
                    Top =2415
                    Width =3840
                    Height =2880
                    Name ="img_360h"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =2415
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =5295
                    TabIndex =9
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =2400
                    Top =5415
                    Width =3840
                    Height =2880
                    Name ="img_120h"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =5415
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =8295
                    TabIndex =10
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =2400
                    Top =8415
                    Width =3840
                    Height =2880
                    Name ="img_240h"

                    LayoutCachedLeft =2400
                    LayoutCachedTop =8415
                    LayoutCachedWidth =6240
                    LayoutCachedHeight =11295
                    TabIndex =11
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2400
                    Top =3435
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_360h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =3435
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =3825
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2400
                    Top =6435
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_120h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =6435
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =6825
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =2400
                    Top =9435
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_240h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =2400
                    LayoutCachedTop =9435
                    LayoutCachedWidth =6180
                    LayoutCachedHeight =9825
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13800
                    Top =2340
                    Width =600
                    Height =8895
                    FontSize =20
                    FontWeight =700
                    BackColor =14281957
                    Name ="Label168"
                    Caption ="Looking OUT From Center"
                    FontName ="Calibri"
                    LayoutCachedLeft =13800
                    LayoutCachedTop =2340
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =11235
                    BackThemeColorIndex =7
                    BackTint =20.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13200
                    Top =2355
                    Width =480
                    Height =2880
                    FontSize =16
                    FontWeight =700
                    BackColor =15921906
                    Name ="Label169"
                    Caption ="360"
                    FontName ="Calibri"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =2355
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =5235
                    BackThemeColorIndex =1
                    BackShade =95.0
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =120
                    Top =120
                    Width =2880
                    Height =2160
                    Name ="img_plotcenter"

                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =2280
                    TabIndex =12
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =1140
                    Width =2880
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_plotcenter"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1140
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =1530
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13800
                    Top =11580
                    Width =600
                    Height =8895
                    FontSize =20
                    FontWeight =700
                    BackColor =15459034
                    Name ="Label159"
                    Caption ="Looking IN From Perimeter"
                    FontName ="Calibri"
                    LayoutCachedLeft =13800
                    LayoutCachedTop =11580
                    LayoutCachedWidth =14400
                    LayoutCachedHeight =20475
                    BackThemeColorIndex =9
                    BackTint =20.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13200
                    Top =5355
                    Width =480
                    Height =2880
                    FontSize =16
                    FontWeight =700
                    BackColor =15921906
                    Name ="Label160"
                    Caption ="120"
                    FontName ="Calibri"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =5355
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =8235
                    BackThemeColorIndex =1
                    BackShade =95.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13200
                    Top =8355
                    Width =480
                    Height =2880
                    FontSize =16
                    FontWeight =700
                    BackColor =15921906
                    Name ="Label161"
                    Caption ="240"
                    FontName ="Calibri"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =8355
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =11235
                    BackThemeColorIndex =1
                    BackShade =95.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13200
                    Top =11595
                    Width =480
                    Height =2880
                    FontSize =16
                    FontWeight =700
                    BackColor =15921906
                    Name ="Label162"
                    Caption ="60"
                    FontName ="Calibri"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =11595
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =14475
                    BackThemeColorIndex =1
                    BackShade =95.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13200
                    Top =14595
                    Width =480
                    Height =2880
                    FontSize =16
                    FontWeight =700
                    BackColor =15921906
                    Name ="Label163"
                    Caption ="180"
                    FontName ="Calibri"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =14595
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =17475
                    BackThemeColorIndex =1
                    BackShade =95.0
                End
                Begin Label
                    Vertical = NotDefault
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =13200
                    Top =17595
                    Width =480
                    Height =2880
                    FontSize =16
                    FontWeight =700
                    BackColor =15921906
                    Name ="Label164"
                    Caption ="300"
                    FontName ="Calibri"
                    LayoutCachedLeft =13200
                    LayoutCachedTop =17595
                    LayoutCachedWidth =13680
                    LayoutCachedHeight =20475
                    BackThemeColorIndex =1
                    BackShade =95.0
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =120
                    Top =2415
                    Width =2160
                    Height =2880
                    Name ="img_360v"

                    LayoutCachedLeft =120
                    LayoutCachedTop =2415
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =5295
                    TabIndex =13
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =3435
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_360v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =120
                    LayoutCachedTop =3435
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =3825
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =120
                    Top =5415
                    Width =2160
                    Height =2880
                    Name ="img_120v"

                    LayoutCachedLeft =120
                    LayoutCachedTop =5415
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =8295
                    TabIndex =14
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =6435
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_120v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =120
                    LayoutCachedTop =6435
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =6825
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =120
                    Top =8415
                    Width =2160
                    Height =2880
                    Name ="img_240v"

                    LayoutCachedLeft =120
                    LayoutCachedTop =8415
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =11295
                    TabIndex =15
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =9435
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_240v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =120
                    LayoutCachedTop =9435
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =9825
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =120
                    Top =17595
                    Width =2160
                    Height =2880
                    Name ="img_300v"

                    LayoutCachedLeft =120
                    LayoutCachedTop =17595
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =20475
                    TabIndex =16
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =18615
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_300v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =120
                    LayoutCachedTop =18615
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =19005
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =120
                    Top =14595
                    Width =2160
                    Height =2880
                    Name ="img_180v"

                    LayoutCachedLeft =120
                    LayoutCachedTop =14595
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =17475
                    TabIndex =17
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =15615
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_180v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =120
                    LayoutCachedTop =15615
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =16005
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =120
                    Top =11595
                    Width =2160
                    Height =2880
                    Name ="img_060v"

                    LayoutCachedLeft =120
                    LayoutCachedTop =11595
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =14475
                    TabIndex =18
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =12615
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl_060v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =120
                    LayoutCachedTop =12615
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =13005
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =10140
                    Top =120
                    Width =2880
                    Height =2160
                    Name ="img2_plotcenter"

                    LayoutCachedLeft =10140
                    LayoutCachedTop =120
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =2280
                    TabIndex =19
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10140
                    Top =1140
                    Width =2880
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_plotcenter"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =10140
                    LayoutCachedTop =1140
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =1530
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =10860
                    Top =2415
                    Width =2160
                    Height =2880
                    Name ="img2_360v"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =2415
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =5295
                    TabIndex =20
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10860
                    Top =3435
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_360v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =10860
                    LayoutCachedTop =3435
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =3825
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =10860
                    Top =5415
                    Width =2160
                    Height =2880
                    Name ="img2_120v"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =5415
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =8295
                    TabIndex =21
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10860
                    Top =6435
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_120v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =10860
                    LayoutCachedTop =6435
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =6825
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =10860
                    Top =8415
                    Width =2160
                    Height =2880
                    Name ="img2_240v"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =8415
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =11295
                    TabIndex =22
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10860
                    Top =9435
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_240v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =10860
                    LayoutCachedTop =9435
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =9825
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =10860
                    Top =17595
                    Width =2160
                    Height =2880
                    Name ="img2_300v"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =17595
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =20475
                    TabIndex =23
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10860
                    Top =18615
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_300v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =10860
                    LayoutCachedTop =18615
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =19005
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =10860
                    Top =14595
                    Width =2160
                    Height =2880
                    Name ="img2_180v"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =14595
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =17475
                    TabIndex =24
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10860
                    Top =15615
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_180v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =10860
                    LayoutCachedTop =15615
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =16005
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =10860
                    Top =11595
                    Width =2160
                    Height =2880
                    Name ="img2_060v"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =11595
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =14475
                    TabIndex =25
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =10860
                    Top =12615
                    Width =2160
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_060v"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =10860
                    LayoutCachedTop =12615
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =13005
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =6900
                    Top =11595
                    Width =3840
                    Height =2880
                    Name ="img2_060h"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =11595
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =14475
                    TabIndex =26
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =6900
                    Top =14594
                    Width =3840
                    Height =2880
                    Name ="img2_180h"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =14594
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =17474
                    TabIndex =27
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =6900
                    Top =17594
                    Width =3840
                    Height =2880
                    Name ="img2_300h"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =17594
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =20474
                    TabIndex =28
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6900
                    Top =12600
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_060h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =12600
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =12990
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6900
                    Top =15600
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_180h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =15600
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =15990
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6900
                    Top =18600
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_300h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =18600
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =18990
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =6900
                    Top =2415
                    Width =3840
                    Height =2880
                    Name ="img2_360h"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =2415
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =5295
                    TabIndex =29
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =6900
                    Top =5414
                    Width =3840
                    Height =2880
                    Name ="img2_120h"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =5414
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =8294
                    TabIndex =30
                End
                Begin Image
                    OldBorderStyle =1
                    SizeMode =3
                    PictureType =1
                    Left =6900
                    Top =8415
                    Width =3840
                    Height =2880
                    Name ="img2_240h"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =8415
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =11295
                    TabIndex =31
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6900
                    Top =3420
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_360h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =3420
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =3810
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6900
                    Top =6420
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_120h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =6420
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =6810
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =6900
                    Top =9420
                    Width =3780
                    Height =390
                    FontWeight =700
                    ForeColor =1643706
                    Name ="lbl2_240h"
                    Caption ="New Record - No Image"
                    LayoutCachedLeft =6900
                    LayoutCachedTop =9420
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =9810
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7620
                    Top =120
                    Width =786
                    Height =786
                    FontSize =12
                    FontWeight =700
                    Name ="cmdMatch"
                    Caption ="Setup"
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
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffe5e5e5cececec1c1c1b3b3b3a5a5a5 ,
                        0x9797978989897c7c7c6e6e6e6363636e6e6e7f7f7f909090a1a1a1b1b1b1cbcb ,
                        0xcbfafafaffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffff2f2f2bdbdbd8585854e4e ,
                        0x4e33333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333333d3d3d747474b0b0b0f8f8f8ffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffff9f9f9b9 ,
                        0xb9b9696969353535333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333545454d8d8d8ffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffff9f9f9 ,
                        0xdfdfdfc2c2c2a5a5a58585854242423333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333393939 ,
                        0xeaeaeaffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333bbbbbbffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333393939 ,
                        0xf1f1f1ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333616161fefefeffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x757575f8f8f8ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333333333333333333333333333333333336e6e6effffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333333333ebebebffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333333333333333333333333333333333333c3c3cfbfbfbffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333b2b2b2ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333333333333333333333333333338d8d8dffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x434343dadadaffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333333333333333333333333333333333334b4b4bfdfdfdffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333333333ddddddffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333333333353535efefefffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff838383 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x343434959595ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffdfdfdd5d5d591919133333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333464646d1d1d1ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffa3a3a33333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333333333333333333333333333333c3c3c ,
                        0xf9f9f9ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffff6f6f6f33333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333d5d5d5ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffe1e1e13838383333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0xb7b7b7ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffa8a8a833 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333373737dededeffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffff7979793333333333333333333333333333 ,
                        0x333333333333333333333333333333333333338e8e8edbdbdbf2f2f2e1e1e1ca ,
                        0xcacab3b3b39d9d9d8585856e6e6e575757414141363636404040656565c5c5c5 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7 ,
                        0xf7f75d5d5d333333333333333333333333333333333333333333333333333333 ,
                        0x575757fefefeffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffececec4b4b4b3333333333333333 ,
                        0x33333333333333333333333333333333696969ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffdcdcdc3e3e3e333333333333333333333333333333333333333333 ,
                        0x555555ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffc7c7c73636363333 ,
                        0x33333333333333333333333333333333333333e3e3e3ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffaeaeae333333333333333333333333333333333333 ,
                        0x3333339a9a9affffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff8f8f ,
                        0x8f333333333333333333333333333333333333494949fcfcfcffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffafafa555555333333333333333333333333 ,
                        0x333333333333bdbdbdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffd1d1d1343434333333333333333333333333333333686868ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff878787333333333333333333 ,
                        0x333333333333333333e2e2e2ffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffff4f4f43e3e3e333333333333333333333333333333bbbbbbffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffff858585333333333333 ,
                        0x3333333333333333339f9f9fffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffd4d4d4333333333333333333333333333333b5b5b5ffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffff555555333333 ,
                        0x333333333333353535efefefffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffb5b5b5333333333333333333969696ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffefefe868686 ,
                        0x636363b0b0b0ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffff
                    End
                    FontName ="Calibri"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Yes, looks like a match"
                    Picture ="Hands-Thumbs-up-icon 64.bmp"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =7620
                    LayoutCachedTop =120
                    LayoutCachedWidth =8406
                    LayoutCachedHeight =906
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =6731160
                    HoverThemeColorIndex =7
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
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =8580
                    Top =120
                    Width =786
                    Height =786
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdNoMatch"
                    Caption ="Exit"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000030000000300000000100180000000000001b0000c40e0000c40e0000 ,
                        0x0000000000000000ffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffefefe868686636363b0b0b0ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffb5b5b5333333 ,
                        0x333333333333969696ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffff555555333333333333333333353535efefefffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffd4d4d4333333333333 ,
                        0x333333333333333333b5b5b5ffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff8585853333333333333333333333333333339f9f9fffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffff4f4f43e3e3e333333333333 ,
                        0x333333333333333333bbbbbbffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff878787333333333333333333333333333333333333e2e2e2ffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffd1d1d1343434333333333333333333 ,
                        0x333333333333686868ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffafa ,
                        0xfa555555333333333333333333333333333333333333bdbdbdffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffff8f8f8f333333333333333333333333333333 ,
                        0x333333494949fcfcfcffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffaeaeae3333 ,
                        0x333333333333333333333333333333333333339a9a9affffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc7c7c7363636333333333333333333333333333333333333 ,
                        0x333333e3e3e3ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffdcdcdc3e3e3e3333333333 ,
                        0x33333333333333333333333333333333555555ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffececec4b4b4b333333333333333333333333333333333333333333333333 ,
                        0x696969ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffff7f7f75d5d5d3333333333333333333333 ,
                        0x33333333333333333333333333333333575757fefefeffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff79 ,
                        0x7979333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333338e8e8edbdbdbf2f2f2e1e1e1cacacab3b3b39d9d9d8585856e6e6e5757 ,
                        0x57414141363636404040656565c5c5c5ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffa8a8a83333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333373737 ,
                        0xdededeffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffe1e1e138383833 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333b7b7b7ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffff6f6f6f3333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0xd5d5d5ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffa3a3a333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333333333333333333333333c3c3cf9f9f9ffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffffffffdfdfd ,
                        0xd5d5d59191913333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333464646 ,
                        0xd1d1d1ffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff83838333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333343434959595ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333333353535efefefffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333333333333333ddddddffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333334b4b4bfdfdfdffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333434343dadadaffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x8d8d8dffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333333333b2b2b2ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333c3c3cfbfbfbffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333333333333333ebebebffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333336e6e6effffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333333333757575f8f8f8ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333616161 ,
                        0xfefefeffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff7f7f7f33333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333393939f1f1f1ffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff7f7f7f ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0xbbbbbbffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffff9f9f9dfdfdfc2c2c2a5a5a585858542424233 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x33333333333333333333333333393939eaeaeaffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffff9f9f9b9b9b96969693535353333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333545454d8d8d8 ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffff2f2f2bdbdbd8585854e4e4e333333333333333333333333333333 ,
                        0x3333333333333333333333333333333333333333333333333333333333333333 ,
                        0x333d3d3d747474b0b0b0f8f8f8ffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffe5e5e5cececec1c1c1b3b3b3a5a5a59797978989897c7c7c6e6e6e6363636e ,
                        0x6e6e7f7f7f909090a1a1a1b1b1b1cbcbcbfafafaffffffffffffffffffffffff ,
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
                    ControlTipText ="No, some photos are not the same scene"
                    Picture ="Hands-Thumbs-down-icon (1).bmp"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1

                    LayoutCachedLeft =8580
                    LayoutCachedTop =120
                    LayoutCachedWidth =9366
                    LayoutCachedHeight =906
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
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =3780
                    Top =60
                    Width =3660
                    Height =960
                    FontSize =13
                    FontWeight =700
                    Name ="Label219"
                    Caption ="Do these two years of photos seem to be appropriately matched and a complete set"
                        "?"
                    FontName ="Calibri"
                    LayoutCachedLeft =3780
                    LayoutCachedTop =60
                    LayoutCachedWidth =7440
                    LayoutCachedHeight =1020
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =7620
                    Top =1020
                    Width =780
                    FontWeight =700
                    TabIndex =2
                    Name ="txtMatch_Votes"
                    ControlSource ="=DCount(\"[tbl_QA_Photos]![QA_Photo_ID]\",\"[tbl_QA_Photos]\",\"[tbl_QA_Photos]!"
                        "[Error_Detected]=False AND  [tbl_QA_Photos]![Location_ID] =  [txtLocation_ID]\")"
                        " & \" Votes\""
                    FontName ="Calibri"

                    LayoutCachedLeft =7620
                    LayoutCachedTop =1020
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =1260
                End
                Begin TextBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8580
                    Top =1020
                    Width =780
                    FontWeight =700
                    TabIndex =3
                    Name ="txtNoMatch_Votes"
                    ControlSource ="=DCount(\"[tbl_QA_Photos]![QA_Photo_ID]\",\"[tbl_QA_Photos]\",\"[tbl_QA_Photos]!"
                        "[Error_Detected]=True AND [tbl_QA_Photos]![Error_Addressed]= False AND [tbl_QA_P"
                        "hotos]![Location_ID] =  [txtLocation_ID]\") & \" Votes\""
                    FontName ="Calibri"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =1020
                    LayoutCachedWidth =9360
                    LayoutCachedHeight =1260
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =7800
                    Top =1380
                    Width =576
                    Height =576
                    TabIndex =4
                    Name ="cmdNext_Record"
                    Caption ="Command223"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dddddddddddddddddddddddddddddddd00000ddddd00000d ,
                        0x0f000ddddd0f000d0f000ddddd0f000d0000000d0000000d00f000000f00000d ,
                        0x00f000d00f00000d00f000d00f00000dd0000000000000dddd0f000d0f000ddd ,
                        0xdd00000d00000dddddd000ddd000ddddddd0f0ddd0f0ddddddd000ddd000dddd ,
                        0xdddddddddddddddd000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Find Record"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =10
                        Begin
                            Action ="OnError"
                            Argument ="0"
                        End
                        Begin
                            Action ="GoToControl"
                            Argument ="=[Screen].[PreviousControl].[Name]"
                        End
                        Begin
                            Action ="ClearMacroError"
                        End
                        Begin
                            Action ="RunCommand"
                            Argument ="30"
                        End
                        Begin
                            Condition ="[MacroError]<>0"
                            Action ="MsgBox"
                            Argument ="=[MacroError].[Description]"
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =7800
                    LayoutCachedTop =1380
                    LayoutCachedWidth =8376
                    LayoutCachedHeight =1956
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =8700
                    Top =1380
                    Width =570
                    Height =570
                    TabIndex =5
                    Name ="Command224"
                    Caption ="Command224"
                    PictureData = Begin
                        0x2800000020000000200000000100180000000000000c0000c40e0000c40e0000 ,
                        0x0000000000000000fffffffffffffffffffefefefdfdfdfdfdfdfbfbfbfbfbfb ,
                        0xfafafaf9f9f9f8f8f8f7f7f7f6f6f6f6f6f6f6f6f6f5f5f5f5f5f5f6f6f6f6f6 ,
                        0xf6f6f6f6f7f7f7f8f8f8f9f9f9f9f9f9fafafafbfbfbfdfdfdfdfdfdfefefeff ,
                        0xfffffffffffffffffefefeefefefe0e0e0d6d6d6cfcfcfcacacac6c6c6c2c2c2 ,
                        0xbfbfbfbdbdbdbbbbbbb9b9b9b8b8b8b8b8b8b8b8b8b7b7b7b7b7b7b8b8b8b8b8 ,
                        0xb8b8b8b8b9b9b9bbbbbbbdbdbdbfbfbfc2c2c2c5c5c5c9c9c9cfcfcfd6d6d6df ,
                        0xdfdfedededfdfdfdfdfdfde6e6e6d1d1d1c7c7c7bfbfbfbbbbbbb7b7b7b4b4b4 ,
                        0xb2b2b2b0b0b0afafafaeaeaeaeaeaea6a8aa949ba38a949f89949e939aa2a6a8 ,
                        0xaaadadadaeaeaeafafafb0b0b0b2b2b2b4b4b4b7b7b7babababfbfbfc6c6c6d0 ,
                        0xd0d0e3e3e3fbfbfbfffffffffffffdfdfdfbfbfbf9f9f9f7f7f7f5f5f5f3f3f3 ,
                        0xf1f1f1eeefefd7dcdf93a6b85176992b5a8828578527558226538026517d2751 ,
                        0x7c4c6b8b8d9cadd5d8dbedeeeef0f0f1f3f3f3f5f5f5f7f7f7f9f9f9fbfbfbfd ,
                        0xfdfdfefefeffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xe7ecf17b9cb92c62932a5f8f295d8d295b8a2859882857852755832754802652 ,
                        0x7e25507b254e79254d777088a1e3e7ebffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffa6bfd5 ,
                        0x336c9e2c65972b63952b61922a5f90295d8d295c8b285a882858862756832754 ,
                        0x8126527e26507c254f79244d772a50789bacbfffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffdfefe81a7c72e6b9f ,
                        0x2d699d2c679a2c65982b63952c65983073ab337eba3583c03582c0337bb62f6e ,
                        0xa4285a8826537f26517d254f7a254d77244b75728ba5fbfbfcffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffff81a9ca2f6fa52e6da3 ,
                        0x2e6ba02d6a9e2e6ea43582c3326dd22b5aa3243c4d2337442432c72632c62c5a ,
                        0xae3171a1337db7295d8d26517d25507b254e78244c76718aa5fefefeffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffa7c5dc3073ab2f71a82f6fa6 ,
                        0x2e6ea3337cb7359cc32748c5211ec12436aa212d3621343d232ec52223c22440 ,
                        0x7222313c264864316ccf2f6fa626527e25507b254e79244c7698aabeffffffff ,
                        0xffffffffffffffffffffffffffffffffffe2ecf4387cb33176ae3074ab3072a9 ,
                        0x3481be31a3b32baf9b2a8fb11f28c31a70d8149ab910b3d211b2ee139be61a7d ,
                        0x97203b47253f9f2325c12b4fc93176af26527e26507c254f7929517ad9e0e7ff ,
                        0xffffffffffffffffffffffffffffffffff75a6ce327ab33178b13176ae337eb9 ,
                        0x2a4ac62877b92aaf9a1fbcbf11c0f50dcbf70ec7f70dccf70dcdf70ccef707df ,
                        0xf60eccee1873d92223c02540932a55752e6da326537f26517c254f7a637f9eff ,
                        0xffffffffffffffffffffffffffffe1ecf5357fba337cb7327ab4327ab42f62c4 ,
                        0x2017bf2018bf1ba4db11c0f811bef814afec1b8bcb17a7f818a5f919a5f914b7 ,
                        0xf70dcdf707e2f61793cf21303b212f3a2e6692295d8d26538026517d26517cd6 ,
                        0xdee6ffffffffffffffffffffffff96bedd3480bc337eba337cb7347eb9233c57 ,
                        0x22329e195ed512b9f812baf812b6fa2693d42e6ba01995db16a9fa18a6f919a3 ,
                        0xf91aa3f810c2f707e1f617798c24416f253cb63071ae27558327548026527e84 ,
                        0x9cb5ffffffffffffffffffffffff5c9cce3582bf3481bd347fba2c5e861e2225 ,
                        0x1d2d3213b6ef12b7f90fbcfb10b9fb17b1f52679b1169fe515adfa16aafa17a7 ,
                        0xfa19a4f91aa3f80dccf70dcff52030c32221bd2c5abc295c8b27568327548145 ,
                        0x6b91fffffffffffffffffff8fbfd3888c63585c23583c03480bb27489421325a ,
                        0x196e8515acf90ebdfb0dc0fc0ebdfc0fbafb18a2e11d7ab114a9f014aefa16ab ,
                        0xfa1c99e4276ba317a2de07e1f6187bd92336a9284d932c679a28588627568428 ,
                        0x5681f0f3f6ffffffffffffddebf63789c93687c63585c3337eb9211ebc1f12bd ,
                        0x1686e514b1f80ac8fc0bc5fc0cc1fc0ebefc1b87be26517c244d771b7bb41b92 ,
                        0xd42b81bf3481bd1e91da0dcdf711a5c1202b3321313e2e6ca1295b8a28598728 ,
                        0x5785d1dbe5ffffffffffffcfe3f3388bcc3789ca3688c7327bb61f19bf1f18bf ,
                        0x1697ed14affa08cdfd09cafd0ac6fc0fb9f22d699d2d689c295d8d254e782292 ,
                        0xd116adf716abf816aafa0dcbf70ebfda202f382030392e6ca0295d8d295b8a29 ,
                        0x5a88c3d1ddffffffffffffd0e4f43b90d0378ccd378aca327ab41f2b2f1e2b2d ,
                        0x169bd114affa06d1fd09c8f709cafc16a5da3480bc347fbc2f6ea428588613aa ,
                        0xea12b4fb13b1fb14aefa0ec7f70fbaef2128bf2228bf2d6aa12a5f90295d8d2c ,
                        0x5e8dc4d2dfffffffffffffddecf83e94d4388ed0388ccd327bb51f262d1d1e1e ,
                        0x1688b415adf90cc2ee2c6a9e2673a521a7df26a7e4388dcf3176ae1f85bb0fbb ,
                        0xfb11b8fb12b5fb13b2fa0ec9f711a0e7201aba2020b82d689c2b61932a60902e ,
                        0x6392d2dde7fffffffffffff7fbfd4398d83990d3388ed13480bd243f7d20299b ,
                        0x1968d51a9bfa0fbdf72b98d112aedb06d0fd08cdfd0fc0f41bb0e9219eda0dbf ,
                        0xfc0fbcfc10b9fb10bbfa0ccef7187f9822376525418f2c679a2b64962b629431 ,
                        0x6796eff3f7ffffffffffffffffff65ade13a93d63991d4388ccd2745a61e0fbe ,
                        0x1c1dc215aef613b0fa03d7fe03d7fe04d4fe06d1fd07cdfd09cafc2a81b911a9 ,
                        0xdf0dc0fc0ebdfb0fc2f80fc7f21d353d1f252a264a692d689b2c66992c64974f ,
                        0x7fa8ffffffffffffffffffffffff9ccbed3f98da3a93d73991d42f6eaa2126b3 ,
                        0x20306a18687e199ff90fbafa02d8fe03d8fe04d5fe05d2fd0ac6f53279b3237d ,
                        0xb00bc4fc0fc0f910c3f81770d822339921364c2b61902d6a9f2d699c2f699c8c ,
                        0xacc8ffffffffffffffffffffffffe0effa4aa1e03a95da3a94d7378aca254867 ,
                        0x1d1d1d1c1d1d168dc419a0f914affa0ac5fd03d7fe04d6fe05d3fe16bbef13b5 ,
                        0xef10bdf812bcf818aee01e16ba1f15b925419f2e6ba02e6da22e6b9f3a76a6d5 ,
                        0xe1ebffffffffffffffffffffffffffffff81bfeb3f9ade3b96db3a94d83279b3 ,
                        0x213345202f771e10be185fd815b2f31a9bfa15aef914adfa13b3f913b5f815af ,
                        0xf912bbf71cbcc428ac942774b4222ba82c66992f71a82f6fa5316fa46e9cc0ff ,
                        0xffffffffffffffffffffffffffffffffffe0effa53a9e63b98de3b96db3a94d9 ,
                        0x2f70a7212aa71e0fbe212ba01b2d3018708b168ce8169df115a5db1594bf1872 ,
                        0xd81c25c12889ab29ad94288e8d2b64953175ad3073ab2f71a84582b3d8e4eeff ,
                        0xffffffffffffffffffffffffffffffffffffffffaad5f34da6e43b99de3b97dc ,
                        0x3a95d93175ae2338972030491d1d1d202f541e0fbd1e16bf1e2a2c1d1e1f202a ,
                        0x9c1e11bb2133b82777872c679a3279b23177b03175ae4081b59cbdd7ffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffeffff8bc5ef4ca6e53c99df ,
                        0x3b97dc3a95da3688c82b619121395321347d1e13b81e16bf1e2a2c1d20242135 ,
                        0x722233942959903176ae337db8337bb6327ab34084b97dabcefdfefeffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffbfdfe8ac5ef52aae6 ,
                        0x3c9adf3b98dd3b96db3a94d8378bcc3175ae2c679b2a5f932a5e8e2c64972f6f ,
                        0xa6337fba3583c03481be3480bb337eb9478dc17dadd1fafcfdffffffffffffff ,
                        0xfffffffffffffffffffffffffffffffffffffffffffffffffffffeffffa8d4f3 ,
                        0x5bb0e845a0e23b98de3b96db3a95d93a93d63991d4388fd1388dcf378bcc3789 ,
                        0xca3687c73686c43584c13d88c25197cb9ec3dffdfefeffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xdfeffb81c2ed58ade74ba4e23b97dc3a95d93a93d73991d4398fd2388dcf378c ,
                        0xcd378acb4593ce519bd079b0d9d9e8f3ffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffdceefa9ccdf06fb8e955aae355a8e154a6df53a5dd53a3db53a2 ,
                        0xd969addc96c3e5d7e8f4ffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffffffffffffffffffffffff1f8fdd8ebf8cae3f5c9e2f4d7e9f6f0f7 ,
                        0xfcffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffff
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Find Next"
                    Picture ="Game-casino-icon_32.bmp"
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="FindNext"
                        End
                    End

                    LayoutCachedLeft =8700
                    LayoutCachedTop =1380
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =1950
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =15527148
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
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
' FORM NAME:    frm_Data_Entry
' Description:  Primary field data entry form
' Data source:  tbl_Locations
' Data access:  edit; allow additions off except for new records
' Pages:        none
' Functions:    Update_Loc_Info, ValidateForm
' References:   fxnSwitchboardIsOpen, fxnGUIDGen
' Source/date:  John R. Boetsch, June 2006
' Revisions:    Simon Kingston, October - January 2006
'                   - extensive updates, adding GUID generation code, new controls
' =================================

Private Sub Form_Open(Cancel As Integer)
    On Error GoTo Err_Handler

    Dim strCaptionSuffix As String
    Dim booEditOn As Boolean
   
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Current()
    'Update fields in header from Locations table
    Update_Loc_Info
    'Enable edit location function if there is an active location
    'Me!cmdEditLocation.Enabled = Not IsNull(Me!txtLocation_ID)
    'Event groups not implemented in this database
    'Me!cmdEditEventGroup.Enabled = Not IsNull(Me!cboEvent_Group_ID)
    RefillImages
End Sub

Private Sub lblLink_To_Plot_Photos_Click()
On Error GoTo Err_Handler

    Dim retVal As Double
    Dim RootFolder As String
    Dim PhotoFolder As String
    
    RootFolder = "T:\I&M"
    PhotoFolder = "T:\I&M\Monitoring\Forest_Vegetation\Photos\"
    If FolderExists(PhotoFolder & Me!txtPlot_Name) Then
        retVal = Shell("explorer /e,/root, " & PhotoFolder & Me!txtPlot_Name, vbNormalFocus)
        GoTo Exit_Procedure
    Else
        If FolderExists(RootFolder) Then
            MsgBox ("Folder for this plot not found....Opening the root of the Photos folder.")
            retVal = Shell("explorer /e,/root, " & PhotoFolder, vbNormalFocus)
            GoTo Exit_Procedure
        Else
            MsgBox ("The network appears to be unavailable. Network access is required to view photos.")
        End If
    End If

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox Err.Description
    Resume Exit_Procedure
End Sub

Private Sub cmdEditLocation_Click()
Dim strOpenargs As String
Dim strCriteria As String

    If Not IsNothing(Me!txtLocation_ID) Then
        strOpenargs = XML_Tag("FormFrom", Me.Name)
        strOpenargs = strOpenargs & XML_Tag("ControlFrom", "txtLocation_ID")
        strCriteria = GetCriteriaString("Location_ID=", "tbl_Locations", "Location_ID", Me.Name, "txtLocation_ID")
        DoCmd.OpenForm "frm_Locations", , , strCriteria, acFormEdit, acWindowNormal, strOpenargs
    End If
End Sub

Public Sub Update_Loc_Info()
' Description:  Updates associated location information when Location_ID is updated
' References:   GetCriteriaString
' Source/date:  Simon Kingston, Sept. 2006
' Revisions:    <name, date, desc - add lines as you go>

Dim strXY As Variant
Dim strCriteria As String

If IsNull(Me!txtLocation_ID) Then
    Me!txtXY = Null
    lblLink_to_Google_Maps.HyperlinkAddress = "http://maps.google.com"
    'lblLink_To_Plot_Photos.Tag = "T:\I&M\Monitoring\Forest_Vegetation\Photos"
Else
    strCriteria = GetCriteriaString("Location_ID=", "tbl_Events", "Location_ID", Me.Name, "txtLocation_ID")
    strXY = "UTM 18N NAD83 E: " & [X_Coord] & "  N: " & [Y_Coord]
    Me!txtXY = strXY
    Me!cboEvent_Date1 = DMin("Event_Date", "tbl_Events", strCriteria)
    Me!cboEvent_Date2 = DMax("Event_Date", "tbl_Events", strCriteria)
    'Me!txtPlot_Name = DLookup("Plot_Name", "tbl_Locations", strCriteria)
    lblLink_to_Google_Maps.HyperlinkAddress = "http://maps.google.com/maps?q=" & Me!txtPlot_Name & "@" & [Lat_WGS84] & "," & [Lon_WGS84] & "&iwloc=A&t=h"
    'lblLink_To_Plot_Photos.Tag = "T:\I&M\Monitoring\Forest_Vegetation\Photos\" & Me!txtPlot_Name
End If
End Sub

Private Sub cmdClose_Click()
    On Error GoTo Err_Handler

    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close , , acSaveNo

Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub Form_Close()
If IsLoaded("frm_Data_Gateway") Then
    Forms("frm_Data_Gateway").Requery
End If
End Sub

Public Sub RefillImages(Optional strName As String = "", Optional strErrorMessage As String = "")
Dim varNames() As Variant
Dim i As Integer
Dim strPictureFileStart As String
Dim strPictureFolder As String
Dim strPictureDateMin As String
Dim strPictureDateMax As String

strPictureFolder = "T:\I&M\Monitoring\Forest_Vegetation\Photos\"  'The first part of the path (same for all images)
'strPictureFolder = "E:\Database\Photos\"  'The first part of the path (same for all images)
strPictureFileStart = strPictureFolder & Me!txtPlot_Name & "\"   'Part two of path, the plot name
strPictureDateMin = "_" & Format(Me!cboEvent_Date1, "yyyymmdd") & "_"  'Part three of path, the date (YYYYMMDD)
strPictureDateMax = "_" & Format(Me!cboEvent_Date2, "yyyymmdd") & "_"  'Part three of path, the date (YYYYMMDD)
'Loop below adds part four (image description) to the path
If Len(strName) = 0 Then
    varNames = Array("plotcenter", "360h", "360v", "120h", "120v", "240h", "240v", "060h", "060v", "180h", "180v", "300h", "300v")
    For i = 0 To UBound(varNames())
        ImageFiller strPictureFileStart + Me!txtPlot_Name + strPictureDateMin + varNames(i) + ".jpg", "img_" & varNames(i), "lbl_" & varNames(i), strErrorMessage
        ImageFiller strPictureFileStart + Me!txtPlot_Name + strPictureDateMax + varNames(i) + ".jpg", "img2_" & varNames(i), "lbl2_" & varNames(i), strErrorMessage
    Next
Else
    ImageFiller strPictureFileStart & Me(strName), "img_" & strName, "lbl_" & strName, strErrorMessage
End If
End Sub

Public Sub ImageFiller(varFileName As Variant, strImageControl As String, strErrorLabel As String, Optional strDefaultErrorMsg As String = "")
Dim strPicture As String
Dim strErrorMessage As String

On Error GoTo Error_Handler

'varFileName is the full path to the image

If IsNull(varFileName) Then
    strErrorMessage = "No image selected"
End If

strErrorMessage = strDefaultErrorMsg
Me(strErrorLabel).Caption = strErrorMessage
Me(strImageControl).Picture = Nz(varFileName, "")
Me(strImageControl).HyperlinkAddress = Nz(varFileName, "")

Exit_Handler:
    Exit Sub

Error_Handler:
    Me(strImageControl).Picture = ""
    Select Case Err.Number
        Case 2114
            'image not supported
            strErrorMessage = "Image type unsupported."
        Case 2220
            If FileExists(varFileName) Then
                strErrorMessage = "Unable to open file."
            Else
                strErrorMessage = "Unable to find file."
            End If
        Case Else
            MsgBox Err.Number & " - " & Err.Description
            strErrorMessage = "Error"
    End Select
    Me(strErrorLabel).Caption = strErrorMessage
    Resume Exit_Handler

End Sub

Private Sub cboEvent_Date1_AfterUpdate()
    RefillImages
End Sub

Private Sub cboEvent_Date2_AfterUpdate()
    RefillImages
End Sub

Public Sub Open_Photo_QA(strError_Detected As Boolean, datEventDate1 As Date, datEventDate2 As Date, strLocationID As String)

DoCmd.OpenForm "frm_Photo_QA", acNormal, , , acFormAdd ', , strOpenargs

Forms("frm_Photo_QA").Error_Detected = strError_Detected
Forms("frm_Photo_QA").Event_Date1 = datEventDate1
Forms("frm_Photo_QA").Event_Date2 = datEventDate2
Forms("frm_Photo_QA").Location_ID = strLocationID
Forms("frm_Photo_QA").AD_Name = NetworkUserName()
Forms("frm_Photo_QA").Error_Date = Now()

If strError_Detected = True Then
    Forms("frm_Photo_QA").txtError_Description.Visible = True
    Forms("frm_Photo_QA").txtError_Description.SetFocus
Else
    Forms("frm_Photo_QA").txtError_Description.Visible = False
    Forms("frm_Photo_QA").cmd_Close_Form.SetFocus
End If

'Forms("frm_Photo_QA").cboTag_ID = strRecordID
'Forms("frm_Photo_QA")!Table_Name = strTableName
'Forms("frm_Photo_QA")!Field_Name = strFieldName
'Forms("frm_Photo_QA")!Record_ID_Field_Name = strRecordIDFieldName
'Forms("frm_Tags_History_Update").txtValue_Old = strOldValue

'Set Forms("frm_Tags_History_Update").ctlToReset = ctlControlToReset
'Set Forms("frm_Tags_History_Update").frmReferrer = frmFormToSave

End Sub

Private Sub cmdNoMatch_Click()
    Open_Photo_QA True, cboEvent_Date1, cboEvent_Date1, txtLocation_ID
End Sub

Private Sub cmdMatch_Click()
    Open_Photo_QA False, cboEvent_Date1, cboEvent_Date1, txtLocation_ID
End Sub
