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
    ItemSuffix =44
    Left =8220
    Top =2430
    Right =11820
    Bottom =6630
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
            Height =599
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Width =2100
                    Height =599
                    ColumnOrder =0
                    FontSize =22
                    FontWeight =700
                    Name ="txtValue"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =599
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =2160
                    Width =720
                    Height =599
                    FontSize =10
                    TabIndex =1
                    Name ="cmdClear"
                    Caption ="Clear"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2160
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =599
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =2880
                    Width =720
                    Height =599
                    FontSize =10
                    TabIndex =2
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =2880
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =599
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
            End
        End
        Begin Section
            Height =3720
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
                    FontWeight =700
                    TabIndex =1
                    Name ="cmdMinus"
                    Caption ="90"
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
                    FontWeight =700
                    TabIndex =2
                    Name ="cmdPeriod"
                    Caption ="95"
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
                    Caption ="0"
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
                    Caption ="1"
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
                    Caption ="2"
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
                    Caption ="3"
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
                    Caption ="4"
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
                    Caption ="5"
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
                    Caption ="6"
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
                    Caption ="7"
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
                    Caption ="8"
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
                    Caption ="9"
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
                    Caption ="10"
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
                    Caption ="35"
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
                    Caption ="60"
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
                    Caption ="85"
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
                    Caption ="15"
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
                    Caption ="20"
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
                    Caption ="25"
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
                    Caption ="30"
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
                    Caption ="40"
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
                    Caption ="45"
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
                    Caption ="50"
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
                    Caption ="55"
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
                    Caption ="65"
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
                    Caption ="70"
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
                    Caption ="75"
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
                    Caption ="80"
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
On Error GoTo Err_Handler
    If Not IsNull(txtValue) Then
        CtrlToUpdate = txtValue
        'The following line is optional, but may be needed if you have calculated field on the subform
        'This line is now remarked because it caused issues with forms that were not ready to be saved, eg, some required fields were not entered yet
        'CtrlToUpdate.Parent.Refresh
    End If
    DoCmd.Close acForm, "frm_Pad_Percent"
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub cmdCancel_Click()
    DoCmd.Close acForm, "frm_Pad_Percent"
End Sub

Private Sub cmdClear_Click()
    On Error Resume Next
    txtValue = ""
End Sub

Private Sub cmdMinus_Click()
On Error Resume Next
f_NumberClick ("90")

End Sub

Private Sub cmdPeriod_Click()
On Error Resume Next
f_NumberClick ("95")
End Sub

Private Sub cmda_Click()
On Error Resume Next
f_NumberClick ("0")
End Sub

Private Sub cmdb_Click()
On Error Resume Next
f_NumberClick ("1")
End Sub

Private Sub cmdc_Click()
On Error Resume Next
f_NumberClick ("2")
End Sub

Private Sub cmdd_Click()
On Error Resume Next
f_NumberClick ("3")
End Sub

Private Sub cmde_Click()
On Error Resume Next
f_NumberClick ("4")
End Sub

Private Sub cmdf_Click()
On Error Resume Next
f_NumberClick ("5")
End Sub

Private Sub cmdg_Click()
On Error Resume Next
f_NumberClick ("6")
End Sub

Private Sub cmdh_Click()
On Error Resume Next
f_NumberClick ("7")
End Sub

Private Sub cmdi_Click()
On Error Resume Next
f_NumberClick ("8")
End Sub

Private Sub cmdj_Click()
On Error Resume Next
f_NumberClick ("9")
End Sub

Private Sub cmdk_Click()
On Error Resume Next
f_NumberClick ("10")
End Sub

Private Sub cmdl_Click()
On Error Resume Next
f_NumberClick ("15")
End Sub

Private Sub cmdm_Click()
On Error Resume Next
f_NumberClick ("20")
End Sub

Private Sub cmdn_Click()
On Error Resume Next
f_NumberClick ("25")
End Sub

Private Sub cmdo_Click()
On Error Resume Next
f_NumberClick ("30")
End Sub

Private Sub cmdp_Click()
On Error Resume Next
f_NumberClick ("35")
End Sub

Private Sub cmdq_Click()
On Error Resume Next
f_NumberClick ("40")
End Sub

Private Sub cmdr_Click()
On Error Resume Next
f_NumberClick ("45")
End Sub

Private Sub cmds_Click()
On Error Resume Next
f_NumberClick ("50")
End Sub

Private Sub cmdt_Click()
On Error Resume Next
f_NumberClick ("55")
End Sub

Private Sub cmdu_Click()
On Error Resume Next
f_NumberClick ("60")
End Sub

Private Sub cmdv_Click()
On Error Resume Next
f_NumberClick ("65")
End Sub

Private Sub cmdw_Click()
On Error Resume Next
f_NumberClick ("70")
End Sub

Private Sub cmdx_Click()
On Error Resume Next
f_NumberClick ("75")
End Sub

Private Sub cmdy_Click()
On Error Resume Next
f_NumberClick ("80")
End Sub

Private Sub cmdz_Click()
On Error Resume Next
f_NumberClick ("85")
End Sub
