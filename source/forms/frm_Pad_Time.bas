Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =127
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =3840
    DatasheetFontHeight =10
    ItemSuffix =32
    Left =10830
    Top =3255
    Right =14640
    Bottom =7695
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf8c4ff537de0e240
    End
    Caption ="Number Pad"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
        Begin FormHeader
            Height =600
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =2434
                    Height =480
                    FontSize =18
                    FontWeight =700
                    Name ="txtValue"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2580
                    Top =180
                    Width =1020
                    Height =360
                    FontSize =12
                    Name ="Label30"
                    Caption ="Seconds"
                    LayoutCachedLeft =2580
                    LayoutCachedTop =180
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =540
                End
            End
        End
        Begin Section
            Height =3900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =2880
                    Width =1260
                    Height =959
                    TabIndex =7
                    Name ="cmdAssign"
                    Caption ="Assign"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100040000000000000200000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x80800000c0c0c000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777778 ,
                        0x7777777777777777777777777777777880777777777777777777777777777778 ,
                        0x8007777777777777777777777777777880007777777777777777777777777778 ,
                        0x8000077777777777777788888888888880000077777777777777880000000000 ,
                        0x0000000777777777777788000000000000000000777777777777880000000000 ,
                        0x0000000007777777777788000000000000000000007777777777880000000000 ,
                        0x0000000000077777777788000000000000000000007077777777880000000000 ,
                        0x0000000007077777777788000000000000000000707777777777880000000000 ,
                        0x0000000707777777777788077777777770000070777777777777770000000000 ,
                        0x0000070777777777777777777777777880007077777777777777777777777778 ,
                        0x8007077777777777777777777777777880707777777777777777777777777778 ,
                        0x7007777777777777777777777777777770777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777777777777777777777777777777777777777777777777777 ,
                        0x7777777777777777000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =120
                    Top =1020
                    Width =1260
                    Height =300
                    FontSize =12
                    Name ="Label17"
                    Caption ="Start Time"
                    LayoutCachedLeft =120
                    LayoutCachedTop =1020
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =1320
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =1560
                    Top =660
                    Width =600
                    Height =300
                    FontSize =12
                    Name ="Label18"
                    Caption ="Hour"
                    LayoutCachedLeft =1560
                    LayoutCachedTop =660
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =960
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2220
                    Top =660
                    Width =600
                    Height =300
                    FontSize =12
                    Name ="Label19"
                    Caption ="Min"
                    LayoutCachedLeft =2220
                    LayoutCachedTop =660
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =960
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =2880
                    Top =660
                    Width =600
                    Height =300
                    FontSize =12
                    Name ="Label20"
                    Caption ="Sec"
                    LayoutCachedLeft =2880
                    LayoutCachedTop =660
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =960
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Top =1440
                    Width =1380
                    Height =300
                    FontSize =12
                    Name ="Label21"
                    Caption ="Stop Time"
                    LayoutCachedTop =1440
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =1740
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1020
                    Width =600
                    Height =315
                    FontSize =12
                    Name ="txt_Time_Start_Hour"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1335
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =1020
                    Width =600
                    Height =315
                    FontSize =12
                    TabIndex =1
                    Name ="txt_Time_Start_Min"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =1335
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =1020
                    Width =600
                    Height =315
                    FontSize =12
                    TabIndex =2
                    Name ="txt_Time_Start_Sec"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1020
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1335
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =1440
                    Width =600
                    Height =315
                    FontSize =12
                    TabIndex =3
                    Name ="txt_Time_Stop_Hour"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =1560
                    LayoutCachedTop =1440
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =1755
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2220
                    Top =1440
                    Width =600
                    Height =315
                    FontSize =12
                    TabIndex =4
                    Name ="txt_Time_Stop_Min"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =2220
                    LayoutCachedTop =1440
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =1755
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =1440
                    Width =600
                    Height =315
                    FontSize =12
                    TabIndex =5
                    Name ="txt_Time_Stop_Sec"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =1755
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =180
                    Width =3660
                    Height =360
                    FontSize =11
                    Name ="Label29"
                    Caption ="Enter stopwatch start and stop time"
                    LayoutCachedLeft =60
                    LayoutCachedTop =180
                    LayoutCachedWidth =3720
                    LayoutCachedHeight =540
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =3300
                    Width =1140
                    Height =504
                    FontSize =14
                    TabIndex =6
                    Name ="cmdClear"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =60
                    LayoutCachedTop =3300
                    LayoutCachedWidth =1200
                    LayoutCachedHeight =3804
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =1260
                    Top =3300
                    Width =1140
                    Height =504
                    FontSize =14
                    TabIndex =8
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1260
                    LayoutCachedTop =3300
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =3804
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub cmdAssign_Click()
    If Not IsNull(txtValue) Then CtrlToUpdate = txtValue
    DoCmd.Close acForm, "frm_Pad_Time"
End Sub

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, "frm_Pad_Time"
End Sub

Private Sub cmdClear_Click()
    On Error Resume Next
    txt_Time_Start_Hour = 0
    txt_Time_Start_Min = 0
    txt_Time_Start_Sec = 0
    txt_Time_Stop_Hour = 0
    txt_Time_Stop_Min = 0
    txt_Time_Stop_Sec = 0
    Time_Pad_Recalc
End Sub

Private Sub Time_Pad_Recalc()
    On Error Resume Next
    txtValue = (3600 * (Me!txt_Time_Stop_Hour - Me!txt_Time_Start_Hour)) + (60 * (Me!txt_Time_Stop_Min - Me!txt_Time_Start_Min)) + (Me!txt_Time_Stop_Sec - Me!txt_Time_Start_Sec)
End Sub

Private Sub txt_Time_Start_Hour_AfterUpdate()
    Time_Pad_Recalc
End Sub

Private Sub txt_Time_Start_Min_AfterUpdate()
    Time_Pad_Recalc
End Sub

Private Sub txt_Time_Start_Sec_AfterUpdate()
    Time_Pad_Recalc
End Sub

Private Sub txt_Time_Stop_Hour_AfterUpdate()
    Time_Pad_Recalc
End Sub

Private Sub txt_Time_Stop_Min_AfterUpdate()
    Time_Pad_Recalc
End Sub

Private Sub txt_Time_Stop_Sec_AfterUpdate()
    Time_Pad_Recalc
End Sub
