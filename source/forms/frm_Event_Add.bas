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
                    Name ="cbxLocationID"
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
                            Name ="lblPlot"
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
                    Name ="tbxEventDate"
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
                            Name ="lblEventDate"
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
                    Name ="lblTitle"
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
                    Name ="tbxEventID"
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
                            Name ="lblEventID"
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
                    Name ="btnCreate"
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
                    Name ="btnCancel"
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
                    Name ="tbxProtocolName"
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
                            Name ="lblProtocolName"
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
                    Name ="cbxParkCode"
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
                            Name ="lblPark"
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
Option Explicit

' =================================
' MODULE:       frm_Event_Add
' Level:        Application module
' Version:      1.01
'
' Description:  add event related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 4/5/2018 - 1.01 - added documentation, error handling
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ----------------
'  Events
' ----------------

' ---------------------------------
' SUB:          Form_BeforeUpdate
' Description:  form before update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    'Generate string GUID for Event_ID
    If Me.NewRecord Then
        If GetDataType("tbl_Events", "Event_ID") = dbText Then
            Me!Event_ID = fxnGUIDGen
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[frm_Event_Add])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCreate_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cmdEvent_Create > btnCreate
' ---------------------------------
Private Sub btnCreate_Click()
On Error GoTo Err_Handler

    'Save the new event if all of the needed information is provided, and open the Event form

    Dim strDocName As String
    Dim strLinkCriteria As String
    
    If IsNull(Me!cbxLocationID) Then
        MsgBox "You must select a location before you can enter record details!", _
            vbExclamation, "Enter Location First"
        Me!cbxLocationID.SetFocus
    Else
        If IsNull(Me!tbxEventDate) Then
            MsgBox "You must enter a date before you can enter record details!", _
                vbExclamation, "Enter Start Date"
            Me!tbxEventDate.SetFocus
        Else
            DoCmd.RunCommand acCmdSaveRecord
            strDocName = "frm_Events"
            strLinkCriteria = "[Event_ID]=" & "'" & Me![tbxEventID] & "'"
            DoCmd.OpenForm strDocName, , , strLinkCriteria, , , "(Creating)"
            DoCmd.Close acForm, "frm_Event_Add"
        End If
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCreate_Click[frm_Event_Add])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnCancel_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cmdEvent_Cancel > btnCancel
' ---------------------------------
Private Sub btnCancel_Click()
On Error GoTo Err_Handler

    'Close the Create Event form without creating a record

    If Me.Dirty Then Me.Undo
    If Not Me.NewRecord Then
        DoCmd.RunCommand acCmdDeleteRecord
    End If
    
    DoCmd.Close
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[frm_Event_Add])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxParkCode_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 4/5/2018 - added documentation, error handling
'                                  renamed cboPark_Code > cbxParkCode
' ---------------------------------
Private Sub cbxParkCode_AfterUpdate()
On Error GoTo Err_Handler

    Me.cbxLocationID.RowSource = "SELECT tbl_Locations.Location_ID, tbl_Locations.Plot_Name, " _
            & "tbl_Locations.Panel, tbl_Locations.Frame, tbl_Locations.Unit_Code " _
            & "FROM tbl_Locations " _
            & "WHERE (((tbl_Locations.Panel) = [Forms]![frm_Switchboard]![Panel]) " _
            & "AND ((tbl_Locations.Unit_Code) = '" & Me.cbxParkCode & "')) " _
            & "ORDER BY tbl_Locations.Plot_Name;"

    Me.cbxLocationID = Me.cbxLocationID.ItemData(0)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxParkCode_AfterUpdate[frm_Event_Add])"
    End Select
    Resume Exit_Handler
End Sub
