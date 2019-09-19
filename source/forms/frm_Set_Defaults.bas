Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4560
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =12135
    Top =4350
    Right =17085
    Bottom =8085
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa3c57b9aedcee240
    End
    RecordSource ="tsys_App_Defaults"
    Caption =" Set application default values"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =3114
            BackColor =11056034
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =5760
                    Left =1182
                    Top =840
                    Width =1395
                    Height =252
                    FontSize =9
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboPanel"
                    ControlSource ="Panel"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code, tlu_Enumerations.Enum_Description FROM tlu_En"
                        "umerations WHERE (((tlu_Enumerations.[Enum_Group])=\"Sampling Panel\")) ORDER BY"
                        " tlu_Enumerations.Enum_Code; "
                    ColumnWidths ="720;5040"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =1182
                    LayoutCachedTop =840
                    LayoutCachedWidth =2577
                    LayoutCachedHeight =1092
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =435
                            Top =840
                            Width =660
                            Height =255
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblPanel"
                            Caption ="Panel"
                            FontName ="Calibri"
                            LayoutCachedLeft =435
                            LayoutCachedTop =840
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1095
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1167
                    Top =120
                    Width =3180
                    Height =252
                    FontSize =9
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"0\""
                    Name ="cboUser"
                    ControlSource ="User_name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Last_Name & \"_\" & First_Name FROM tlu_Contacts ORDER BY Last_Name, Firs"
                        "t_Name; "
                    FontName ="Calibri"
                    OnNotInList ="[Event Procedure]"

                    LayoutCachedLeft =1167
                    LayoutCachedTop =120
                    LayoutCachedWidth =4347
                    LayoutCachedHeight =372
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =432
                            Top =120
                            Width =663
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblUser"
                            Caption ="User"
                            FontName ="Calibri"
                            LayoutCachedLeft =432
                            LayoutCachedTop =120
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =372
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1182
                    Top =2340
                    Width =3168
                    Height =252
                    FontSize =9
                    TabIndex =6
                    Name ="cboProject"
                    ControlSource ="Project"
                    FontName ="Calibri"

                    LayoutCachedLeft =1182
                    LayoutCachedTop =2340
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =2592
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =423
                            Top =2340
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblProject"
                            Caption ="Project"
                            FontName ="Calibri"
                            LayoutCachedLeft =423
                            LayoutCachedTop =2340
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =2592
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3690
                    Top =2760
                    Width =720
                    Height =354
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="cmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =3690
                    LayoutCachedTop =2760
                    LayoutCachedWidth =4410
                    LayoutCachedHeight =3114
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3330
                    Top =480
                    Width =1035
                    FontSize =9
                    FontWeight =700
                    TabIndex =7
                    ForeColor =0
                    Name ="cmdNewUser"
                    Caption ="New user"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"
                    ControlTipText ="Add a new user"

                    LayoutCachedLeft =3330
                    LayoutCachedTop =480
                    LayoutCachedWidth =4365
                    LayoutCachedHeight =840
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =3600
                    Left =1182
                    Top =1620
                    Width =1395
                    FontSize =9
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboDatum"
                    ControlSource ="Datum"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"Datu"
                        "m\" ORDER BY Sort_Order; "
                    ColumnWidths ="720;2880"
                    FontName ="Calibri"

                    LayoutCachedLeft =1182
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2577
                    LayoutCachedHeight =1860
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =423
                            Top =1620
                            Width =672
                            Height =252
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblDatum"
                            Caption ="Datum"
                            FontName ="Calibri"
                            LayoutCachedLeft =423
                            LayoutCachedTop =1620
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1872
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2664
                    Left =1170
                    Top =1980
                    Width =1395
                    FontSize =9
                    TabIndex =4
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboUTM_Zone"
                    ControlSource ="UTM_Zone"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"UTM "
                        "Zone\" ORDER BY Sort_Order; "
                    ColumnWidths ="504;2160"
                    FontName ="Calibri"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =1980
                    LayoutCachedWidth =2565
                    LayoutCachedHeight =2220
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =390
                            Top =1980
                            Width =705
                            Height =270
                            FontSize =9
                            FontWeight =700
                            BackColor =11056034
                            Name ="lblDeclination"
                            Caption ="Zone"
                            FontName ="Calibri"
                            LayoutCachedLeft =390
                            LayoutCachedTop =1980
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =2250
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =7200
                    Left =1170
                    Top =480
                    Width =1395
                    FontSize =9
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"100\""
                    Name ="cboProtocol_Name"
                    ControlSource ="Protocol_Name"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code, Enum_Description FROM tlu_Enumerations WHERE Enum_Group=\"prot"
                        "ocol\" ORDER BY Sort_Order; "
                    ColumnWidths ="2160;5040"
                    StatusBarText ="M. The name or code of the protocol governing the event (Protcl_Nam)"
                    FontName ="Calibri"

                    LayoutCachedLeft =1170
                    LayoutCachedTop =480
                    LayoutCachedWidth =2565
                    LayoutCachedHeight =720
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =300
                            Top =480
                            Width =795
                            Height =240
                            FontSize =9
                            FontWeight =700
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label54"
                            Caption ="Protocol"
                            FontName ="Calibri"
                            LayoutCachedLeft =300
                            LayoutCachedTop =480
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =720
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1200
                    Top =1185
                    Width =1395
                    FontSize =9
                    TabIndex =8
                    Name ="Tablet_Role"
                    ControlSource ="Entry_Role"
                    RowSourceType ="Value List"
                    RowSource ="PRIMARY;SECONDARY;SINGLE;OFFICE"
                    StatusBarText ="Data Entry Role of this Computer (Primary, Secondary, Single)"
                    FontName ="Calibri"
                    AllowValueListEdits =1

                    LayoutCachedLeft =1200
                    LayoutCachedTop =1185
                    LayoutCachedWidth =2595
                    LayoutCachedHeight =1425
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =15
                            Top =1185
                            Width =1080
                            Height =240
                            FontSize =9
                            FontWeight =700
                            Name ="Label7"
                            Caption ="Entry_Role:"
                            FontName ="Calibri"
                            LayoutCachedLeft =15
                            LayoutCachedTop =1185
                            LayoutCachedWidth =1095
                            LayoutCachedHeight =1425
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3360
                    Top =1200
                    Width =960
                    Height =255
                    TabIndex =9
                    Name ="Text8"
                    ControlSource ="Timeframe"

                    LayoutCachedLeft =3360
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1455
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2760
                            Top =1200
                            Width =540
                            Height =240
                            Name ="lblYear_Default"
                            Caption ="Year"
                            LayoutCachedLeft =2760
                            LayoutCachedTop =1200
                            LayoutCachedWidth =3300
                            LayoutCachedHeight =1440
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
' FORM NAME:    frm_Set_Defaults
' Description:  Standard module for setting application defaults
' Data source:  tsys_App_Defaults
' Data access:  edit only, no deletions
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, May 16, 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cboUser_NotInList(NewData As String, response As Integer)
    On Error GoTo Err_Handler

    MsgBox "User not found.  To add this user, click the New user button.", vbOKOnly, "User Not In List"
    Me.ActiveControl.Undo
    response = acDataErrContinue
    Me!cmdNewUser.SetFocus

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub

Private Sub cmdNewUser_Click()
    On Error GoTo Err_Handler
    
    ' Open the contacts form
    DoCmd.OpenForm "frm_Contacts", , , , , , "new"

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub


Private Sub cmdOK_Click()
    On Error GoTo Err_Handler

    Dim varOpenArgs As Variant
    
    varOpenArgs = Me.OpenArgs
    
    ' Make sure the information is valid before updating the record
    If varOpenArgs <> 0 Then
        '  Verify that the critical data elements have been completed before saving
        If IsNull(Me!User_name) Then
            MsgBox "Please indicate the user name", vbOKOnly, "Validation error"
            Me!cboUser.SetFocus
            GoTo Exit_Procedure
       ' ElseIf IsNull(Me!Park) Then
       '    MsgBox "Please indicate the park", vbOKOnly, "Validation error"
       '    Me!cboPark.SetFocus
       '    GoTo Exit_Procedure
        End If
    End If

    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.OpenForm "frm_Switchboard"
    Select Case varOpenArgs
        Case 1
            DoCmd.OpenForm "frm_Data_Gateway", , , , , , varOpenArgs
        Case 2
            DoCmd.OpenForm "frm_Browser", , , , , , varOpenArgs
        Case 3
            DoCmd.OpenForm "frm_QA_Tool", , , , , , varOpenArgs
        Case 4
            ' opened by switchboard only ... do nothing
        Case Else
            MsgBox "Error: OpenArgs property out of range", vbCritical
    End Select

Exit_Procedure:
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
