Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =11880
    DatasheetFontHeight =9
    ItemSuffix =119
    Left =4110
    Top =1095
    Right =15990
    Bottom =5505
    DatasheetGridlinesColor =12632256
    Filter ="[Event_ID]='{856EFBAD-929A-49B8-ABB9-7FBC04CDA834}'"
    RecSrcDt = Begin
        0xc74647bc11ade340
    End
    RecordSource ="tbl_Events"
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
                    Caption ="Add or Edit Event Note"
                    FontName ="Arial"
                    LayoutCachedTop =60
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =420
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10620
                    Top =120
                    Width =930
                    Height =315
                    Name ="cmd_Close_Decay_Popup"
                    Caption ="Close Form"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10620
                    LayoutCachedTop =120
                    LayoutCachedWidth =11550
                    LayoutCachedHeight =435
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =3945
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10155
                    Top =75
                    FontSize =10
                    Name ="Text111"
                    ControlSource ="Event_ID"
                    FontName ="Calibri"

                    LayoutCachedLeft =10155
                    LayoutCachedTop =75
                    LayoutCachedWidth =11595
                    LayoutCachedHeight =315
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9075
                            Top =75
                            Width =705
                            Height =240
                            FontSize =10
                            Name ="Label112"
                            Caption ="Text111:"
                            FontName ="Calibri"
                            LayoutCachedLeft =9075
                            LayoutCachedTop =75
                            LayoutCachedWidth =9780
                            LayoutCachedHeight =315
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =705
                    Top =405
                    Width =1575
                    Height =285
                    FontSize =12
                    TabIndex =1
                    Name ="txtEvent_Date"
                    ControlSource ="Event_Date"
                    Format ="Short Date"
                    FontName ="Calibri"

                    LayoutCachedLeft =705
                    LayoutCachedTop =405
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =690
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =45
                            Top =405
                            Width =615
                            Height =285
                            FontSize =12
                            Name ="Label114"
                            Caption ="Date"
                            FontName ="Calibri"
                            LayoutCachedLeft =45
                            LayoutCachedTop =405
                            LayoutCachedWidth =660
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =720
                    Top =780
                    Width =10950
                    Height =2955
                    FontSize =16
                    TabIndex =2
                    Name ="txtEvent_Notes"
                    ControlSource ="Event_Notes"
                    FontName ="Calibri"

                    LayoutCachedLeft =720
                    LayoutCachedTop =780
                    LayoutCachedWidth =11670
                    LayoutCachedHeight =3735
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =75
                            Top =780
                            Width =600
                            Height =315
                            FontSize =12
                            Name ="Label116"
                            Caption ="Notes"
                            FontName ="Calibri"
                            LayoutCachedLeft =75
                            LayoutCachedTop =780
                            LayoutCachedWidth =675
                            LayoutCachedHeight =1095
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =705
                    Top =60
                    Width =1575
                    Height =285
                    FontSize =12
                    TabIndex =3
                    Name ="txtPlot_Name"
                    ControlSource ="=DLookUp(\"[Plot_Name]\",\"tbl_Locations\",\"[Location_ID] =\" & [Location_ID])"
                    FontName ="Calibri"

                    LayoutCachedLeft =705
                    LayoutCachedTop =60
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =345
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =45
                            Top =60
                            Width =615
                            Height =285
                            FontSize =12
                            Name ="Label118"
                            Caption ="Plot"
                            FontName ="Calibri"
                            LayoutCachedLeft =45
                            LayoutCachedTop =60
                            LayoutCachedWidth =660
                            LayoutCachedHeight =345
                        End
                    End
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
    Forms![frm_Events]![fsub_Note_History].Requery

Exit_cmd_Close_Decay_Popup_Click:
    Exit Sub
Err_cmd_Close_Decay_Popup_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Close_Decay_Popup_Click
End Sub
