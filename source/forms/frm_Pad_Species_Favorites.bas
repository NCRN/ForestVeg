Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =5580
    DatasheetFontHeight =11
    ItemSuffix =20
    Left =3135
    Top =3855
    Right =9000
    Bottom =9225
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x240ade1eab79e440
    End
    RecordSource ="SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants_1.Latin_Name A"
        "S Latin_Name_Accepted, tlu_Plants_1.Common, tlu_Plants.Favorite, tlu_Plants.Tree"
        ", tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accept"
        "ed_Found FROM tlu_Plants INNER JOIN tlu_Plants AS tlu_Plants_1 ON tlu_Plants.TSN"
        "_Accepted = tlu_Plants_1.TSN WHERE (((tlu_Plants.Tree)=True)) OR (((tlu_Plants.S"
        "hrub)=True)) ORDER BY tlu_Plants.Latin_Name;"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =255
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
            BorderColor =16777215
            GridlineColor =16777215
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
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
            Height =420
            BackColor =14276557
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =3
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =720
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="Label16"
                    Caption ="Switch to"
                    GridlineColor =10921638
                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    ForeTint =75.0
                End
                Begin Label
                    OverlapFlags =85
                    Left =3480
                    Top =60
                    Width =1260
                    Height =300
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="Label17"
                    Caption ="species list"
                    GridlineColor =10921638
                    LayoutCachedLeft =3480
                    LayoutCachedTop =60
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    ForeTint =75.0
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2100
                    Top =60
                    Width =1260
                    Height =300
                    FontWeight =700
                    ForeColor =4210752
                    Name ="cmdList_Select"
                    Caption ="All"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =60
                    LayoutCachedWidth =3360
                    LayoutCachedHeight =360
                    BackColor =11710639
                    BorderColor =11710639
                    ThemeFontIndex =-1
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =405
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =5400
                    Height =345
                    FontSize =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLatin_Name_Favorite"
                    ControlSource ="Latin_Name"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000084000000010000000100000000000000000000001100000001000000 ,
                        0x8c8c8c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00410063006300650070007400650064005f0046006f0075006e0064005d00 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =405
                    ThemeFontIndex =-1
                    ConditionalFormat14 = Begin
                        0x0100010000000100000000000000010000008c8c8c00ffffff00100000005b00 ,
                        0x410063006300650070007400650064005f0046006f0075006e0064005d000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =780
            BackColor =14276557
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =3
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =4980
                    Top =120
                    Width =576
                    Height =576
                    TabIndex =6
                    ForeColor =4210752
                    Name ="cmdLast"
                    Caption ="Last"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Last Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =4980
                    LayoutCachedTop =120
                    LayoutCachedWidth =5556
                    LayoutCachedHeight =696
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =120
                    Width =606
                    Height =576
                    TabIndex =2
                    ForeColor =4210752
                    Name ="cmdPrevious_50"
                    Caption ="-50"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =1800
                    LayoutCachedTop =120
                    LayoutCachedWidth =2406
                    LayoutCachedHeight =696
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3120
                    Top =120
                    Width =606
                    Height =576
                    TabIndex =4
                    ForeColor =4210752
                    Name ="cmdNext_50"
                    Caption ="+50"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =3120
                    LayoutCachedTop =120
                    LayoutCachedWidth =3726
                    LayoutCachedHeight =696
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =60
                    Top =120
                    Width =576
                    Height =576
                    ForeColor =4210752
                    Name ="cmdFirst"
                    Caption ="First"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="First Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =636
                    LayoutCachedHeight =696
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3780
                    Top =120
                    Width =1146
                    Height =576
                    TabIndex =5
                    ForeColor =4210752
                    Name ="cmdNext_Page"
                    Caption ="Page Down"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Next Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =3780
                    LayoutCachedTop =120
                    LayoutCachedWidth =4926
                    LayoutCachedHeight =696
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =660
                    Top =120
                    Width =1086
                    Height =576
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdPrevious_Page"
                    Caption ="Page Up"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =660
                    LayoutCachedTop =120
                    LayoutCachedWidth =1746
                    LayoutCachedHeight =696
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2460
                    Top =120
                    Width =606
                    Height =576
                    TabIndex =3
                    ForeColor =4210752
                    Name ="cmdNext_100"
                    Caption ="+100"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Previous Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =2460
                    LayoutCachedTop =120
                    LayoutCachedWidth =3066
                    LayoutCachedHeight =696
                    BackColor =11710639
                    BorderColor =11710639
                    HoverColor =13355721
                    PressedColor =6249563
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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

Private Sub cmdList_Select_Click()
On Error GoTo cmdList_Select_Click_Err

    If (cmdList_Select.Caption = "All") Then
        cmdList_Select.Caption = "Favorite"
        Me.FilterOn = False
    Else
        cmdList_Select.Caption = "All"
        Me.Filter = "Favorite = TRUE"
        Me.FilterOn = True
    End If

cmdList_Select_Click_Exit:
    Exit Sub
cmdList_Select_Click_Err:
    MsgBox Error$
    Resume cmdList_Select_Click_Exit
End Sub

Private Sub cmdNext_100_Click()
On Error GoTo cmdNext_100_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext, 100
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If

cmdNext_100_Click_Exit:
    Exit Sub
cmdNext_100_Click_Err:
    MsgBox Error$
    Resume cmdNext_100_Click_Exit
End Sub

Private Sub cmdNext_50_Click()
On Error GoTo cmdNext_50_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext, 50
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If

cmdNext_50_Click_Exit:
    Exit Sub
cmdNext_50_Click_Err:
    MsgBox Error$
    Resume cmdNext_50_Click_Exit
End Sub

Private Sub cmdNext_Page_Click()
On Error GoTo cmdNext_Page_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext, 11
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If

cmdNext_Page_Click_Exit:
    Exit Sub
cmdNext_Page_Click_Err:
    MsgBox Error$
    Resume cmdNext_Page_Click_Exit
End Sub

Private Sub cmdPrevious_50_Click()
On Error GoTo cmdPrevious_50_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious, 50
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

cmdPrevious_50_Click_Exit:
    Exit Sub
cmdPrevious_50_Click_Err:
    MsgBox Error$
    Resume cmdPrevious_50_Click_Exit
End Sub

Private Sub cmdPrevious_Page_Click()
On Error GoTo cmdPrevious_Page_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious, 11
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
    End If

cmdPrevious_Page_Click_Exit:
    Exit Sub
cmdPrevious_Page_Click_Err:
    MsgBox Error$
    Resume cmdPrevious_Page_Click_Exit
End Sub

Private Sub Form_Open(Cancel As Integer)

    Select Case Parent.OpenArgs
        Case "TREE"
            Me.RecordSource = "SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants_1.Latin_Name AS Latin_Name_Accepted, tlu_Plants_1.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found FROM tlu_Plants INNER JOIN tlu_Plants AS tlu_Plants_1 ON tlu_Plants.TSN_Accepted = tlu_Plants_1.TSN WHERE (((tlu_Plants.Tree)=True)) ORDER BY tlu_Plants.Latin_Name;"
        Case "SAPLING"
            Me.RecordSource = "SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants_1.Latin_Name AS Latin_Name_Accepted, tlu_Plants_1.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found FROM tlu_Plants INNER JOIN tlu_Plants AS tlu_Plants_1 ON tlu_Plants.TSN_Accepted = tlu_Plants_1.TSN WHERE (((tlu_Plants.Tree)=True)) OR (((tlu_Plants.Shrub)=True)) ORDER BY tlu_Plants.Latin_Name;"
        Case "SEEDLING"
            Me.RecordSource = "SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants_1.Latin_Name AS Latin_Name_Accepted, tlu_Plants_1.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found FROM tlu_Plants INNER JOIN tlu_Plants AS tlu_Plants_1 ON tlu_Plants.TSN_Accepted = tlu_Plants_1.TSN WHERE (((tlu_Plants.Tree)=True)) OR (((tlu_Plants.Shrub)=True)) ORDER BY tlu_Plants.Latin_Name;"
        Case "CWD"
            Me.RecordSource = "SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants_1.Latin_Name AS Latin_Name_Accepted, tlu_Plants_1.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found FROM tlu_Plants INNER JOIN tlu_Plants AS tlu_Plants_1 ON tlu_Plants.TSN_Accepted = tlu_Plants_1.TSN WHERE (((tlu_Plants.Tree)=True)) OR (((tlu_Plants.Vine)=True)) OR (((tlu_Plants.Shrub)=True)) ORDER BY tlu_Plants.Latin_Name;"
        Case "VINE"
            Me.RecordSource = "SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants_1.Latin_Name AS Latin_Name_Accepted, tlu_Plants_1.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found FROM tlu_Plants INNER JOIN tlu_Plants AS tlu_Plants_1 ON tlu_Plants.TSN_Accepted = tlu_Plants_1.TSN WHERE (((tlu_Plants.Vine)=True)) ORDER BY tlu_Plants.Latin_Name;"
        Case "TARGETED HERB"
            Me.RecordSource = "SELECT tlu_Plants.TSN_Accepted, tlu_Plants.Latin_Name, tlu_Plants_1.Latin_Name AS Latin_Name_Accepted, tlu_Plants_1.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found FROM tlu_Plants INNER JOIN tlu_Plants AS tlu_Plants_1 ON tlu_Plants.TSN_Accepted = tlu_Plants_1.TSN WHERE (((tlu_Plants.Targeted_Herb)=True)) ORDER BY tlu_Plants.Latin_Name;"
        Case Else
    End Select
    
    cmdList_Select.Caption = "All"
    Me.Filter = "Favorite = True"
    Me.FilterOn = True
End Sub

Private Sub cmdLast_Click()
On Error GoTo cmdLast_Click_Err

    DoCmd.GoToRecord , "", acLast

cmdLast_Click_Exit:
    Exit Sub
cmdLast_Click_Err:
    MsgBox Error$
    Resume cmdLast_Click_Exit
End Sub

Private Sub cmdNext_Click()
On Error GoTo cmdNext_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acNext
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If

cmdNext_Click_Exit:
    Exit Sub
cmdNext_Click_Err:
    MsgBox Error$
    Resume cmdNext_Click_Exit
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo cmdPrevious_Click_Err

    On Error Resume Next
    DoCmd.GoToRecord , "", acPrevious
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.Description, vbOKOnly, ""
            End If

cmdPrevious_Click_Exit:
    Exit Sub
cmdPrevious_Click_Err:
    MsgBox Error$
    Resume cmdPrevious_Click_Exit
End Sub

Private Sub cmdFirst_Click()
On Error GoTo cmdFirst_Click_Err

    DoCmd.GoToRecord , "", acFirst

cmdFirst_Click_Exit:
    Exit Sub
cmdFirst_Click_Err:
    MsgBox Error$
    Resume cmdFirst_Click_Exit
End Sub

Private Sub txtLatin_Name_Favorite_Click()
    On Error Resume Next
    Parent.txtValue = [TSN_Accepted]
    Parent.txtLatin = [Latin_Name_Accepted]
    Parent.txtCommon = [Common]
End Sub
