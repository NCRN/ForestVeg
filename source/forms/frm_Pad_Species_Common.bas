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
    ItemSuffix =17
    Left =9615
    Top =3855
    Right =15480
    Bottom =9225
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x2d57dc1eab79e440
    End
    RecordSource ="SELECT tlu_Plants.TSN, Count(tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID) AS"
        " Obs_Count, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_P"
        "lants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Pla"
        "nts.Accepted_Found, tbl_Events.Event_Date, tbl_Quadrat_Data.Event_ID FROM ((tbl_"
        "Events INNER JOIN tbl_Quadrat_Data ON tbl_Events.Event_ID = tbl_Quadrat_Data.Eve"
        "nt_ID) INNER JOIN tbl_Quadrat_Seedlings_Data ON tbl_Quadrat_Data.Quadrat_Data_ID"
        " = tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID) INNER JOIN tlu_Plants ON tbl_Quad"
        "rat_Seedlings_Data.TSN = tlu_Plants.TSN GROUP BY tlu_Plants.TSN, tlu_Plants.Lati"
        "n_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine"
        ", tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Eve"
        "nts.Event_Date, tbl_Quadrat_Data.Event_ID HAVING (((tlu_Plants.Tree) = True) And"
        " ((tbl_Quadrat_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) Or (((tlu_Plan"
        "ts.Shrub) = True) And ((tbl_Quadrat_Data.Event_ID) = [Forms]![frm_Events]![Event"
        "_ID])) ORDER BY Count(tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID) DESC , tl"
        "u_Plants.Latin_Name;"
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
    FilterOnLoad =0
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
                    Left =1920
                    Top =60
                    Width =2265
                    Height =300
                    FontSize =12
                    FontWeight =700
                    BorderColor =8355711
                    ForeColor =4210752
                    Name ="Label17"
                    Caption ="Recently Used Species"
                    GridlineColor =10921638
                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =4185
                    LayoutCachedHeight =360
                    ThemeFontIndex =-1
                    ForeTint =75.0
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
                    Left =480
                    Top =60
                    Width =4980
                    Height =345
                    FontSize =14
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtLatin_Name_Recent"
                    ControlSource ="Latin_Name"
                    OnEnter ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000084000000010000000100000000000000000000001100000001000000 ,
                        0x8c8c8c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00410063006300650070007400650064005f0046006f0075006e0064005d00 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =480
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
                Begin TextBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Top =60
                    Width =420
                    Height =345
                    FontSize =14
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text16"
                    ControlSource ="Obs_Count"
                    ConditionalFormat = Begin
                        0x0100000084000000010000000100000000000000000000001100000001000000 ,
                        0x8c8c8c00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00410063006300650070007400650064005f0046006f0075006e0064005d00 ,
                        0x00000000
                    End
                    GridlineColor =10921638

                    LayoutCachedTop =60
                    LayoutCachedWidth =420
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
                    Left =4920
                    Top =120
                    Width =576
                    Height =576
                    ForeColor =4210752
                    Name ="cmdLast"
                    Caption ="Last"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Last Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =4920
                    LayoutCachedTop =120
                    LayoutCachedWidth =5496
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
                    Top =120
                    Width =576
                    Height =576
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdFirst"
                    Caption ="First"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="First Record"
                    GridlineColor =10921638

                    LayoutCachedTop =120
                    LayoutCachedWidth =576
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
                    Left =3720
                    Top =120
                    Width =1146
                    Height =576
                    TabIndex =2
                    ForeColor =4210752
                    Name ="cmdNext_Page"
                    Caption ="Page Down"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Next Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =3720
                    LayoutCachedTop =120
                    LayoutCachedWidth =4866
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
                    Left =600
                    Top =120
                    Width =1086
                    Height =576
                    TabIndex =3
                    ForeColor =4210752
                    Name ="cmdPrevious_Page"
                    Caption ="Page Up"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Previous Record"
                    GridlineColor =10921638

                    LayoutCachedLeft =600
                    LayoutCachedTop =120
                    LayoutCachedWidth =1686
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

Private Sub Form_Open(Cancel As Integer)
    Select Case Parent.OpenArgs
        Case "TREE"
            Me.RecordSource = "SELECT tlu_Plants.TSN, Count(tbl_Tags.Tag_ID) AS Obs_Count, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Tree_Data.Event_ID " _
            & "FROM ((tbl_Events INNER JOIN tbl_Tree_Data ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) INNER JOIN tbl_Tags ON tbl_Tree_Data.Tag_ID = tbl_Tags.Tag_ID) INNER JOIN tlu_Plants ON tbl_Tags.TSN = tlu_Plants.TSN " _
            & "GROUP BY tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Tree_Data.Event_ID " _
            & "HAVING (((tlu_Plants.Tree) = True) And ((tbl_Tree_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) " _
            & "ORDER BY Count(tbl_Tags.Tag_ID) DESC , tlu_Plants.Latin_Name;"
        Case "SAPLING"
            Me.RecordSource = "SELECT tlu_Plants.TSN, Count(tbl_Tags.Tag_ID) AS Obs_Count, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Sapling_Data.Event_ID " _
            & "FROM ((tbl_Events INNER JOIN tbl_Sapling_Data ON tbl_Events.Event_ID = tbl_Sapling_Data.Event_ID) INNER JOIN tbl_Tags ON tbl_Sapling_Data.Tag_ID = tbl_Tags.Tag_ID) INNER JOIN tlu_Plants ON tbl_Tags.TSN = tlu_Plants.TSN " _
            & "GROUP BY tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Sapling_Data.Event_ID " _
            & "HAVING (((tlu_Plants.Tree) = True) And ((tbl_Sapling_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) Or (((tlu_Plants.Shrub) = True) And ((tbl_Sapling_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) " _
            & "ORDER BY Count(tbl_Tags.Tag_ID) DESC , tlu_Plants.Latin_Name;"
        Case "SEEDLING"
            Me.RecordSource = "SELECT tlu_Plants.TSN, Count(tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID) AS Obs_Count, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Quadrat_Data.Event_ID " _
            & "FROM ((tbl_Events INNER JOIN tbl_Quadrat_Data ON tbl_Events.Event_ID = tbl_Quadrat_Data.Event_ID) INNER JOIN tbl_Quadrat_Seedlings_Data ON tbl_Quadrat_Data.Quadrat_Data_ID = tbl_Quadrat_Seedlings_Data.Quadrat_Data_ID) INNER JOIN tlu_Plants ON tbl_Quadrat_Seedlings_Data.TSN = tlu_Plants.TSN " _
            & "GROUP BY tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Quadrat_Data.Event_ID " _
            & "HAVING (((tlu_Plants.Tree) = True) And ((tbl_Quadrat_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) Or (((tlu_Plants.Shrub) = True) And ((tbl_Quadrat_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) " _
            & "ORDER BY Count(tbl_Quadrat_Seedlings_Data.Quadrat_Seedlings_ID) DESC , tlu_Plants.Latin_Name;"
        Case "CWD"
            Me.RecordSource = "SELECT tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_CWD_Data.Event_ID, Count(tbl_CWD_Data.CWD_Data_ID) AS Obs_Count " _
            & "FROM tbl_Events INNER JOIN (tbl_CWD_Data INNER JOIN tlu_Plants ON tbl_CWD_Data.TSN = tlu_Plants.TSN) ON tbl_Events.Event_ID = tbl_CWD_Data.Event_ID " _
            & "GROUP BY tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_CWD_Data.Event_ID " _
            & "HAVING (((tlu_Plants.Tree) = True) And ((tbl_CWD_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) Or (((tlu_Plants.Vine) = True) And ((tbl_CWD_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) Or (((tlu_Plants.Shrub) = True) And ((tbl_CWD_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) " _
            & "ORDER BY Count(tbl_CWD_Data.CWD_Data_ID) DESC , tlu_Plants.Latin_Name;"
        Case "VINE"
            Me.RecordSource = "SELECT tlu_Plants.TSN, Count(tbl_Tree_Vines.Tree_Vine_ID) AS Obs_Count, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Tree_Data.Event_ID " _
            & "FROM ((tbl_Events INNER JOIN tbl_Tree_Data ON tbl_Events.Event_ID = tbl_Tree_Data.Event_ID) INNER JOIN tbl_Tree_Vines ON tbl_Tree_Data.Tree_Data_ID = tbl_Tree_Vines.Tree_Data_ID) INNER JOIN tlu_Plants ON tbl_Tree_Vines.TSN = tlu_Plants.TSN " _
            & "GROUP BY tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Tree_Data.Event_ID " _
            & "HAVING (((tlu_Plants.Vine) = True) And ((tbl_Tree_Data.Event_ID) = [Forms]![frm_Events]![Event_ID])) " _
            & "ORDER BY Count(tbl_Tree_Vines.Tree_Vine_ID) DESC , tlu_Plants.Latin_Name;"
        Case "TARGETED HERB"
            Me.RecordSource = "SELECT tlu_Plants.TSN, Count(tbl_Quadrat_Herbaceous_Data.Quadrat_Herbaceous_ID) AS Obs_Count, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Quadrat_Data.Event_ID " _
            & "FROM ((tbl_Events INNER JOIN tbl_Quadrat_Data ON tbl_Events.Event_ID = tbl_Quadrat_Data.Event_ID) INNER JOIN tbl_Quadrat_Herbaceous_Data ON tbl_Quadrat_Data.Quadrat_Data_ID = tbl_Quadrat_Herbaceous_Data.Quadrat_Data_ID) INNER JOIN tlu_Plants ON tbl_Quadrat_Herbaceous_Data.TSN = tlu_Plants.TSN " _
            & "GROUP BY tlu_Plants.TSN, tlu_Plants.Latin_Name, tlu_Plants.Common, tlu_Plants.Favorite, tlu_Plants.Tree, tlu_Plants.Vine, tlu_Plants.Shrub, tlu_Plants.Targeted_Herb, tlu_Plants.Accepted_Found, tbl_Events.Event_Date, tbl_Quadrat_Data.Event_ID " _
            & "HAVING (((tlu_Plants.Targeted_Herb)=True) AND ((tbl_Quadrat_Data.Event_ID)=[Forms]![frm_Events]![Event_ID])) " _
            & "ORDER BY Count(tbl_Quadrat_Herbaceous_Data.Quadrat_Herbaceous_ID) DESC , tlu_Plants.Latin_Name;"
        Case Else
    End Select
End Sub

Private Sub txtLatin_Name_Recent_Click()
    On Error Resume Next
    Parent.txtValue = [TSN]
    Parent.txtLatin = [Latin_Name]
    Parent.txtCommon = [Common]
End Sub

Private Sub txtLatin_Name_Recent_Enter()
    On Error Resume Next
    'Parent.txtValue = [TSN]
    'Parent.txtLatin = [Latin_Name]
    'Parent.txtCommon = [Common]
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

Private Sub cmdFirst_Click()
On Error GoTo cmdFirst_Click_Err

    DoCmd.GoToRecord , "", acFirst

cmdFirst_Click_Exit:
    Exit Sub
cmdFirst_Click_Err:
    MsgBox Error$
    Resume cmdFirst_Click_Exit
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
