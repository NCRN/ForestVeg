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
    Width =11880
    DatasheetFontHeight =9
    ItemSuffix =111
    Left =5325
    Top =1785
    Right =17490
    Bottom =10500
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2e57b5f08f80e340
    End
    Caption ="Decay Classes"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
                    Caption ="Conditions and Pests"
                    FontName ="Arial"
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10620
                    Top =60
                    Width =930
                    Height =315
                    Name ="cmd_Close_Decay_Popup"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10620
                    LayoutCachedTop =60
                    LayoutCachedWidth =11550
                    LayoutCachedHeight =375
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =8220
            BackColor =15527148
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
                    Caption ="Advanced Decay"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =345
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =360
                    Width =11700
                    Height =300
                    FontSize =10
                    Name ="Label36"
                    Caption ="Large portions of the tree are undergoing decay (e.g. heart-rot), but the tree i"
                        "s still alive."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =660
                End
                Begin Line
                    OverlapFlags =87
                    Top =360
                    Width =11760
                    Name ="Line81"
                    LayoutCachedTop =360
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =360
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =720
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label87"
                    Caption ="Primary Branch Broken"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =720
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =1005
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =1020
                    Width =11700
                    Height =1020
                    FontSize =10
                    Name ="Label88"
                    Caption ="Main trunk of the tree is broken off. In the case of trees which split, if any m"
                        "ain trunk is broken use this condition and make a note of the circumstances. Thi"
                        "s does not need to be selected if “Alive Broken” was chosen as the tree status ("
                        "see below). In practice this will generally be selected some, but not all, of th"
                        "e main branches on a tree with a large split are broken, or when the very top, b"
                        "ut not the entire crown of a tree is broken."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =1020
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =2040
                End
                Begin Line
                    OverlapFlags =87
                    Top =1020
                    Width =11760
                    Name ="Line89"
                    LayoutCachedTop =1020
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =1020
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =2100
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label90"
                    Caption ="Large Dead Branches"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =2100
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =2385
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =2400
                    Width =11700
                    Height =480
                    FontSize =10
                    Name ="Label91"
                    Caption ="Large branches on the tree are dead. If “Primary Branch Broken” is selected as a"
                        " tree condition or “Alive Broken” is selected as tree status, this should not be"
                        " selected unless you wish to indicate additional damage not already covered by t"
                        "hose choices."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =2400
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =2880
                End
                Begin Line
                    OverlapFlags =87
                    Top =2400
                    Width =11760
                    Name ="Line92"
                    LayoutCachedTop =2400
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =2400
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =2940
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label93"
                    Caption ="Lightning Damage"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =3225
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =3240
                    Width =11700
                    Height =540
                    FontSize =10
                    Name ="Label94"
                    Caption ="Obvious signs of lightning damage, such as large vertical burn marks on tree. Yo"
                        "u should indicate lightning damage only if you are sure this kind of damage has "
                        "occurred. "
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3240
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =3780
                End
                Begin Line
                    OverlapFlags =87
                    Top =3240
                    Width =11760
                    Name ="Line95"
                    LayoutCachedTop =3240
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =3240
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =3840
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label96"
                    Caption ="Wind Damage"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3840
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =4125
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =4140
                    Width =11700
                    Height =780
                    FontSize =10
                    Name ="Label97"
                    Caption ="Allows us to determine the extent of damage to trees due to storm events in the "
                        "NCRN. In some locations, particularly in Catoctin Mountain Park, wind damage fro"
                        "m storms can knock down stands of dozens of trees. Only select this if you are c"
                        "ertain that wind is the cause of the damage you are seeing."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =4140
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =4920
                End
                Begin Line
                    OverlapFlags =87
                    Top =4140
                    Width =11760
                    Name ="Line98"
                    LayoutCachedTop =4140
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =4140
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =4980
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label99"
                    Caption ="Open Wound"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =4980
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =5265
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =5280
                    Width =11700
                    Height =300
                    FontSize =10
                    Name ="Label100"
                    Caption ="Large open wound, such as from where a branch fell off."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =5280
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =5580
                End
                Begin Line
                    OverlapFlags =87
                    Top =5280
                    Width =11760
                    Name ="Line101"
                    LayoutCachedTop =5280
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =5280
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =5640
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label102"
                    Caption ="Vines in the Crown"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =5640
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =5925
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =5940
                    Width =11700
                    Height =540
                    FontSize =10
                    Name ="Label103"
                    Caption ="Record if any vines on the tree grow into the crown. This is not recorded separa"
                        "tely for each vine species as it can be difficult to identify vines that are hig"
                        "h above the ground."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =5940
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =6480
                End
                Begin Line
                    OverlapFlags =87
                    Top =5940
                    Width =11760
                    Name ="Line104"
                    LayoutCachedTop =5940
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =5940
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =6540
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label105"
                    Caption ="Other Visible Damage"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =6540
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =6825
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =6840
                    Width =11700
                    Height =540
                    FontSize =10
                    Name ="Label106"
                    Caption ="Any other damage that you feel could increase the mortality risk of the tree tha"
                        "t is not covered here (note that wind and lightning damage is covered separately"
                        " below). Record in the notes what the damage is."
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =6840
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =7380
                End
                Begin Line
                    OverlapFlags =87
                    Top =6840
                    Width =11760
                    Name ="Line107"
                    LayoutCachedTop =6840
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =6840
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =186
                    Left =60
                    Top =7440
                    Width =7140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="Label108"
                    Caption ="Pests and Diseases"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =7440
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =7725
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =186
                    Left =60
                    Top =7740
                    Width =11700
                    Height =300
                    FontSize =10
                    Name ="Label109"
                    Caption ="Pests and diseases are also recorded in this section"
                    FontName ="Arial"
                    LayoutCachedLeft =60
                    LayoutCachedTop =7740
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =8040
                End
                Begin Line
                    OverlapFlags =87
                    Top =7740
                    Width =11760
                    Name ="Line110"
                    LayoutCachedTop =7740
                    LayoutCachedWidth =11760
                    LayoutCachedHeight =7740
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
