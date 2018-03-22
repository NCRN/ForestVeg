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
    ItemSuffix =17
    Left =16185
    Top =3975
    Right =19995
    Bottom =8415
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
                Begin CommandButton
                    OverlapFlags =85
                    Left =2520
                    Top =60
                    Width =1260
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
            Height =3900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    Width =1260
                    Height =960
                    FontSize =36
                    Name ="cmd7"
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
                    Top =960
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =1
                    Name ="cmd4"
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
                    Top =1920
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =2
                    Name ="cmd1"
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
                    Top =2880
                    Width =660
                    Height =960
                    FontSize =36
                    TabIndex =3
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
                    Left =1260
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =4
                    Name ="cmd8"
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
                    Left =1260
                    Top =960
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =5
                    Name ="cmd5"
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
                    Left =1260
                    Top =1920
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =6
                    Name ="cmd2"
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
                    Left =1260
                    Top =2880
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =7
                    Name ="cmd0"
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
                    Left =2520
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =8
                    Name ="cmd9"
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
                    Left =2520
                    Top =960
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =9
                    Name ="cmd6"
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
                    Left =2520
                    Top =1920
                    Width =1260
                    Height =960
                    FontSize =36
                    TabIndex =10
                    Name ="cmd3"
                    Caption ="3"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =660
                    Top =2880
                    Width =600
                    Height =960
                    FontSize =36
                    TabIndex =11
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
                    OverlapFlags =87
                    Left =2520
                    Top =2880
                    Width =1260
                    Height =959
                    TabIndex =12
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
    DoCmd.Close acForm, "frm_Pad_Num"
Exit_Procedure:
    Exit Sub
Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure
End Sub

Private Sub cmd0_Click()
On Error Resume Next
f_NumberClick ("0")
End Sub

Private Sub cmd1_Click()
On Error Resume Next
f_NumberClick ("1")
End Sub

Private Sub cmd2_Click()
On Error Resume Next
f_NumberClick ("2")
End Sub

Private Sub cmd3_Click()
On Error Resume Next
f_NumberClick ("3")
End Sub

Private Sub cmd4_Click()
On Error Resume Next
f_NumberClick ("4")
End Sub

Private Sub cmd5_Click()
On Error Resume Next
f_NumberClick ("5")
End Sub

Private Sub cmd6_Click()
On Error Resume Next
f_NumberClick ("6")
End Sub

Private Sub cmd7_Click()
On Error Resume Next
f_NumberClick ("7")
End Sub

Private Sub cmd8_Click()
On Error Resume Next
f_NumberClick ("8")
End Sub

Private Sub cmd9_Click()
On Error Resume Next
f_NumberClick ("9")
End Sub

Private Sub cmdPeriod_Click()
On Error Resume Next
f_NumberClick (".")
End Sub

Private Sub cmdMinus_Click()
On Error Resume Next
f_NumberClick ("-")
End Sub

Private Sub cmdClear_Click()
On Error Resume Next
txtValue = ""
End Sub
