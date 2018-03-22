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
    Width =3600
    DatasheetFontHeight =10
    ItemSuffix =43
    Left =14535
    Top =3285
    Right =18150
    Bottom =7545
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf8c4ff537de0e240
    End
    Caption ="Keypad"
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
                    Width =2220
                    Height =480
                    FontSize =16
                    FontWeight =700
                    Name ="txtValue"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2340
                    Top =60
                    Width =1253
                    Height =504
                    FontSize =14
                    TabIndex =1
                    Name ="cmdClear"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =3606
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Left =720
                    Top =3000
                    Width =720
                    Height =600
                    FontSize =20
                    TabIndex =1
                    Name ="cmdMinus"
                    Caption ="-"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =1440
                    Top =3000
                    Width =720
                    Height =600
                    FontSize =20
                    TabIndex =2
                    Name ="cmdPeriod"
                    Caption ="."
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =3
                    Name ="cmda"
                    Caption ="a"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =720
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =4
                    Name ="cmdb"
                    Caption ="b"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =1440
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =5
                    Name ="cmdc"
                    Caption ="c"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2160
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =6
                    Name ="cmdd"
                    Caption ="d"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2880
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =7
                    Name ="cmde"
                    Caption ="e"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Top =600
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =8
                    Name ="cmdf"
                    Caption ="f"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =720
                    Top =600
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =9
                    Name ="cmdg"
                    Caption ="g"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =1440
                    Top =600
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =10
                    Name ="cmdh"
                    Caption ="h"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2160
                    Top =600
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =11
                    Name ="cmdi"
                    Caption ="i"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2880
                    Top =600
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =12
                    Name ="cmdj"
                    Caption ="j"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Top =1200
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =13
                    Name ="cmdk"
                    Caption ="k"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Top =1800
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =14
                    Name ="cmdp"
                    Caption ="p"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Top =2400
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    Name ="cmdu"
                    Caption ="u"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Top =3000
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =15
                    Name ="cmdz"
                    Caption ="z"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =720
                    Top =1200
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =16
                    Name ="cmdl"
                    Caption ="l"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =1440
                    Top =1200
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =17
                    Name ="cmdm"
                    Caption ="m"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2160
                    Top =1200
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =18
                    Name ="cmdn"
                    Caption ="n"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2880
                    Top =1200
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =19
                    Name ="cmdo"
                    Caption ="o"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =720
                    Top =1800
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =20
                    Name ="cmdq"
                    Caption ="q"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =1440
                    Top =1800
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =21
                    Name ="cmdr"
                    Caption ="r"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2160
                    Top =1800
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =22
                    Name ="cmds"
                    Caption ="s"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2880
                    Top =1800
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =23
                    Name ="cmdt"
                    Caption ="t"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =720
                    Top =2400
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =24
                    Name ="cmdv"
                    Caption ="v"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =1440
                    Top =2400
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =25
                    Name ="cmdw"
                    Caption ="w"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2160
                    Top =2400
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =26
                    Name ="cmdx"
                    Caption ="x"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2880
                    Top =2400
                    Width =720
                    Height =600
                    FontSize =20
                    FontWeight =700
                    TabIndex =27
                    Name ="cmdy"
                    Caption ="y"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =2160
                    Top =3000
                    Height =606
                    TabIndex =28
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
                        0x7777777777777777
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
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
Dim sPrevious As String

Private Function f_NumberClick(sValue As String)
    txtValue = txtValue & sValue
End Function

Private Sub cmdAssign_Click()
    CtrlToUpdate = txtValue
    DoCmd.Close acForm, "frm_AlphaPad"
End Sub

Private Sub cmdClear_Click()
    On Error Resume Next
    txtValue = ""
End Sub

Private Sub cmdMinus_Click()
On Error Resume Next
f_NumberClick ("-")

End Sub

Private Sub cmdPeriod_Click()
On Error Resume Next
f_NumberClick (".")
End Sub

Private Sub cmda_Click()
On Error Resume Next
f_NumberClick ("a")
End Sub

Private Sub cmdb_Click()
On Error Resume Next
f_NumberClick ("b")
End Sub

Private Sub cmdc_Click()
On Error Resume Next
f_NumberClick ("c")
End Sub

Private Sub cmdd_Click()
On Error Resume Next
f_NumberClick ("d")
End Sub

Private Sub cmde_Click()
On Error Resume Next
f_NumberClick ("e")
End Sub

Private Sub cmdf_Click()
On Error Resume Next
f_NumberClick ("f")
End Sub

Private Sub cmdg_Click()
On Error Resume Next
f_NumberClick ("g")
End Sub

Private Sub cmdh_Click()
On Error Resume Next
f_NumberClick ("h")
End Sub

Private Sub cmdi_Click()
On Error Resume Next
f_NumberClick ("i")
End Sub

Private Sub cmdj_Click()
On Error Resume Next
f_NumberClick ("j")
End Sub

Private Sub cmdk_Click()
On Error Resume Next
f_NumberClick ("k")
End Sub

Private Sub cmdl_Click()
On Error Resume Next
f_NumberClick ("l")
End Sub

Private Sub cmdm_Click()
On Error Resume Next
f_NumberClick ("m")
End Sub

Private Sub cmdn_Click()
On Error Resume Next
f_NumberClick ("n")
End Sub

Private Sub cmdo_Click()
On Error Resume Next
f_NumberClick ("o")
End Sub

Private Sub cmdp_Click()
On Error Resume Next
f_NumberClick ("p")
End Sub

Private Sub cmdq_Click()
On Error Resume Next
f_NumberClick ("q")
End Sub

Private Sub cmdr_Click()
On Error Resume Next
f_NumberClick ("r")
End Sub

Private Sub cmds_Click()
On Error Resume Next
f_NumberClick ("s")
End Sub

Private Sub cmdt_Click()
On Error Resume Next
f_NumberClick ("t")
End Sub

Private Sub cmdu_Click()
On Error Resume Next
f_NumberClick ("u")
End Sub

Private Sub cmdv_Click()
On Error Resume Next
f_NumberClick ("v")
End Sub

Private Sub cmdw_Click()
On Error Resume Next
f_NumberClick ("w")
End Sub

Private Sub cmdx_Click()
On Error Resume Next
f_NumberClick ("x")
End Sub

Private Sub cmdy_Click()
On Error Resume Next
f_NumberClick ("y")
End Sub

Private Sub cmdz_Click()
On Error Resume Next
f_NumberClick ("z")
End Sub
