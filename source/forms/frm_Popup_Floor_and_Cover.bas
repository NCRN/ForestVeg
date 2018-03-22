Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10320
    DatasheetFontHeight =9
    ItemSuffix =87
    Left =6105
    Top =2535
    Right =16710
    Bottom =8970
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2e57b5f08f80e340
    End
    Caption ="Decay Classes"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
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
        Begin FormHeader
            Height =480
            BackColor =11252642
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Top =60
                    Width =8280
                    Height =360
                    FontSize =12
                    FontWeight =700
                    Name ="Label18"
                    Caption ="Floor Conditions and Vegetation Cover"
                    FontName ="Arial"
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9180
                    Top =60
                    Width =930
                    Height =315
                    Name ="cmd_Close_Decay_Popup"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9180
                    LayoutCachedTop =60
                    LayoutCachedWidth =10110
                    LayoutCachedHeight =375
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =5940
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =60
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label30"
                    Caption ="Floor Condition"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =345
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =420
                    Width =10020
                    Height =2280
                    FontSize =10
                    Name ="Label36"
                    Caption ="“Floor conditions” are anything that can occupy part of a quadrat and excludes h"
                        "erbaceous vegetation. The categories are: trees, rocks, CWD (coarse woody debris"
                        ") and “other”. Estimate the % of the quadrat that each of these conditions cover"
                        "s to either the nearest 1% if the cover is less than 10%, or to the nearest 5% i"
                        "f the cover is over 10%. \015\012\015\012The “trees” category only includes the "
                        "area covered by the stem/trunk of woody plants and does not include cover of the"
                        " canopy. This category includes standing dead as well as living trees. CWD inclu"
                        "des any wood laying on the soil with a diameter ≥ 7.5 cm. “Other” can include an"
                        "y other object (such as trash). If the cover of other is greater than 0, note in"
                        " the comments section what object is being measured."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =2700
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =3120
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label42"
                    Caption ="Vegetation Cover"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3120
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =3405
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =3480
                    Width =10020
                    Height =1740
                    FontSize =10
                    Name ="Label48"
                    Caption ="“Vegetation cover” is a measure of several types of vegetation that can occur in"
                        " the quadrat. This data is monitored to help determine if there are trends in un"
                        "derstory cover of these groups over time. These categories are: grasses, sedges,"
                        " herbs, ferns and bryophytes. Estimate the % of the quadrat that each of these c"
                        "onditions covers to either the nearest 1% if the cover is less than 10%, or to t"
                        "he nearest 5% if the cover is over 10%."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3480
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =5220
                End
                Begin Line
                    OverlapFlags =85
                    Top =360
                    Width =10080
                    Name ="Line81"
                    LayoutCachedTop =360
                    LayoutCachedWidth =10080
                    LayoutCachedHeight =360
                End
                Begin Line
                    OverlapFlags =85
                    Left =60
                    Top =3420
                    Width =10080
                    Name ="Line85"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3420
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =3420
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
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

Private Sub cmd_Close_Decay_Popup_Click()
On Error GoTo Err_cmd_Close_Decay_Popup_Click


    DoCmd.Close

Exit_cmd_Close_Decay_Popup_Click:
    Exit Sub

Err_cmd_Close_Decay_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Close_Decay_Popup_Click
    
End Sub
