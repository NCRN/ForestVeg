Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =204
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =9
    ItemSuffix =17
    Left =9975
    Top =3090
    Right =14295
    Bottom =7155
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x5de4b299aba7e340
    End
    RecordSource ="tbl_Events"
    Caption ="Create New Event"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin CommandButton
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
        End
        Begin ComboBox
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
        End
        Begin Section
            Height =4080
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    TextFontCharSet =204
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2160
                    Left =1485
                    Top =1440
                    Width =2475
                    Height =510
                    FontSize =18
                    FontWeight =700
                    TabIndex =1
                    Name ="cboLocation_ID"
                    ControlSource ="Location_ID"
                    RowSourceType ="Table/Query"
                    ColumnWidths ="0;2160"

                    LayoutCachedLeft =1485
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1950
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =1440
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblPick_Plot"
                            Caption ="Plot"
                            LayoutCachedLeft =540
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =1955
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1500
                    Top =2040
                    Width =2460
                    Height =510
                    FontSize =18
                    FontWeight =700
                    TabIndex =2
                    Name ="txtEvent_Date"
                    ControlSource ="Event_Date"
                    Format ="Short Date"
                    DefaultValue ="=Date()"
                    InputMask ="99.99.0000;"

                    LayoutCachedLeft =1500
                    LayoutCachedTop =2040
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =2550
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =540
                            Top =2040
                            Width =885
                            Height =510
                            FontSize =18
                            FontWeight =700
                            Name ="lblEvent_Date"
                            Caption ="Date"
                            LayoutCachedLeft =540
                            LayoutCachedTop =2040
                            LayoutCachedWidth =1425
                            LayoutCachedHeight =2550
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Width =4320
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =275078
                    ForeColor =16777215
                    Name ="lblEvent_Add"
                    Caption ="Create New Event"
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =540
                    BackThemeColorIndex =5
                    BackShade =50.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =1275
                    Top =600
                    Width =2595
                    Height =210
                    ColumnWidth =1320
                    FontSize =8
                    TabIndex =5
                    Name ="txtEvent_ID"
                    ControlSource ="Event_ID"
                    StatusBarText ="M. Event identifier (Event_ID)"

                    LayoutCachedLeft =1275
                    LayoutCachedTop =600
                    LayoutCachedWidth =3870
                    LayoutCachedHeight =810
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =240
                            Top =600
                            Width =975
                            Height =210
                            FontSize =8
                            Name ="lblEvent_ID"
                            Caption ="Event ID:"
                            LayoutCachedLeft =240
                            LayoutCachedTop =600
                            LayoutCachedWidth =1215
                            LayoutCachedHeight =810
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    TextFontCharSet =204
                    Left =540
                    Top =2700
                    Width =2325
                    Height =1080
                    FontSize =14
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdEvent_Create"
                    Caption ="Create Event"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =540
                    LayoutCachedTop =2700
                    LayoutCachedWidth =2865
                    LayoutCachedHeight =3780
                    ForeThemeColorIndex =0
                    UseTheme =1
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
                Begin CommandButton
                    OverlapFlags =87
                    TextFontCharSet =204
                    Left =2940
                    Top =2700
                    Width =1020
                    Height =1080
                    FontSize =14
                    TabIndex =4
                    ForeColor =0
                    Name ="cmdEvent_Cancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =2940
                    LayoutCachedTop =2700
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =3780
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
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =2100
                    Top =600
                    Height =315
                    TabIndex =6
                    Name ="txtProtocol_Name"
                    ControlSource ="Protocol_Name"
                    DefaultValue ="=[Forms]![frm_Switchboard]![Protocol_Name]"

                    LayoutCachedLeft =2100
                    LayoutCachedTop =600
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =960
                            Top =600
                            Width =1080
                            Height =315
                            Name ="lblProtocol_Name"
                            Caption ="Protocol:"
                            LayoutCachedLeft =960
                            LayoutCachedTop =600
                            LayoutCachedWidth =2040
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    ListWidth =2160
                    Left =1485
                    Top =840
                    Width =2475
                    Height =510
                    FontSize =18
                    FontWeight =700
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cboPark_Code"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Unit Code\")) ORDER BY tlu_Enumerations.Enum_Code;"
                    ColumnWidths ="2160"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"\""

                    LayoutCachedLeft =1485
                    LayoutCachedTop =840
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =1350
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =540
                            Top =840
                            Width =870
                            Height =515
                            FontSize =18
                            FontWeight =700
                            Name ="lblPick_Park"
                            Caption ="Park"
                            LayoutCachedLeft =540
                            LayoutCachedTop =840
                            LayoutCachedWidth =1410
                            LayoutCachedHeight =1355
                        End
                    End
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

Private Sub cboPark_Code_AfterUpdate()
    Me.cboLocation_ID.RowSource = "SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, tbl_Locations.Panel, tbl_Locations.Frame, tbl_Locations.Unit_Code FROM tbl_Locations WHERE (((tbl_Locations.Panel) = [Forms]![frm_Switchboard]![Panel]) And ((tbl_Locations.Unit_Code) = '" & Me.cboPark_Code & "')) ORDER BY tbl_Locations.Plot_Name;"
    Me.cboLocation_ID = Me.cboLocation_ID.ItemData(0)
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
'Generate string GUID for Event_ID
    If Me.NewRecord Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me!Event_ID = fxnGUIDGen
        End If
    End If
End Sub

Private Sub cmdEvent_Create_Click()
'Save the new event if all of the needed information is provided, and open the Event form
On Error GoTo Err_cmdEvent_Create_Click

    Dim stDocName As String
    Dim stLinkCriteria As String
    
    If IsNull(Me!cboLocation_ID) Then
        MsgBox "You must select a location before you can enter record details!", vbExclamation, "Enter Location First"
        Me!cboLocation_ID.SetFocus
    Else
        If IsNull(Me!txtEvent_Date) Then
            MsgBox "You must enter a date before you can enter record details!", vbExclamation, "Enter Start Date"
            Me!txtEvent_Date.SetFocus
        Else
            DoCmd.RunCommand acCmdSaveRecord
            stDocName = "frm_Events"
            stLinkCriteria = "[Event_ID]=" & "'" & Me![txtEvent_ID] & "'"
            DoCmd.OpenForm stDocName, , , stLinkCriteria, , , "(Creating)"
            DoCmd.Close acForm, "frm_Event_Add"
        End If
    End If

Exit_cmdEvent_Create_Click:
    Exit Sub
Err_cmdEvent_Create_Click:
    MsgBox Err.Description
    Resume Exit_cmdEvent_Create_Click
End Sub

Private Sub cmdEvent_Cancel_Click()
'Close the Create Event form without creating a record
On Error GoTo Err_cmdEvent_Cancel_Click

    If Me.Dirty Then Me.Undo
    If Not Me.NewRecord Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    
    DoCmd.Close
    
Exit_cmdEvent_Cancel_Click:
    Exit Sub
Err_cmdEvent_Cancel_Click:
    MsgBox Err.Description
    Resume Exit_cmdEvent_Cancel_Click
End Sub
