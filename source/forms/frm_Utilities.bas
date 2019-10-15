Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6060
    DatasheetFontHeight =11
    ItemSuffix =9
    Left =1125
    Top =2700
    Right =7185
    Bottom =6165
    DatasheetGridlinesColor =14276557
    RecSrcDt = Begin
        0x2e1f8472d703e440
    End
    Caption ="Utilities"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0xd0020000d0020000d0020000d002000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
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
        Begin Section
            Height =3480
            BackColor =15921906
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Width =6060
                    Height =480
                    FontSize =18
                    FontWeight =700
                    BackColor =0
                    BorderColor =8355711
                    ForeColor =16777215
                    Name ="lblUtilities_Header"
                    Caption ="Utilities and Configuration Tools"
                    GridlineColor =10921638
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =480
                    ThemeFontIndex =-1
                    BackThemeColorIndex =0
                    ForeThemeColorIndex =1
                    ForeTint =100.0
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1680
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =3
                    Name ="btnDataQA"
                    Caption ="QA/QC"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the QA/QC Summary Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Data_QA"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =1680
                    LayoutCachedTop =2040
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =240
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    Name ="btnAppend"
                    Caption =" Append Data"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the Append Data Switchboard ti Import Field Data"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Append_Select_Import_File"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"btnAppend\" Event=\"OnClick\" xmlns=\"http://schemas.microso"
                                "ft.com/office/accessservices/2009/11/application\"><Statements><Action Name=\"Op"
                                "enForm\"><Argument Name=\"FormName"
                        End
                        Begin
                            Comment ="_AXL:\">frm_Append_Select_Import_File</Argument></Action></Statements></UserInte"
                                "rfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =2040
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =240
                    Top =660
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    Name ="btnRelinkTables"
                    Caption =" Relink Tables"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Reset the link to the backend database"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Connect_Tables"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =240
                    LayoutCachedTop =660
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =1920
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =4560
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =4
                    Name ="btnClose"
                    Caption ="Close"
                    FontName ="Franklin Gothic Book"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =8
                        Begin
                            Action ="Close"
                            Argument ="-1"
                            Argument =""
                            Argument ="0"
                        End
                    End

                    LayoutCachedLeft =4560
                    LayoutCachedTop =2040
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =10798077
                    HoverThemeColorIndex =5
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =87
                    Left =1680
                    Top =660
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    Name ="btnBackupBE"
                    Caption ="Create Backup"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Create a Backup of the Backend Database"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638

                    LayoutCachedLeft =1680
                    LayoutCachedTop =660
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1920
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =3120
                    Top =2040
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =5
                    Name ="btnLookups"
                    Caption ="Lookups"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open the QA/QC Summary Form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    GridlineColor =10921638
                    OnClickEmMacro = Begin
                        Version =196611
                        ColumnsShown =0
                        Begin
                            Action ="OpenForm"
                            Argument ="frm_Lookups"
                            Argument ="0"
                            Argument =""
                            Argument =""
                            Argument ="-1"
                            Argument ="0"
                        End
                        Begin
                            Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
                                "nterfaceMacro For=\"cmdLookups\" Event=\"OnClick\" xmlns=\"http://schemas.micros"
                                "oft.com/office/accessservices/2009/11/application\" xmlns:a=\"http://schemas.mic"
                                "rosoft.com/office/accessservices"
                        End
                        Begin
                            Comment ="_AXL:/2009/11/forms\"><Statements><Action Name=\"OpenForm\"><Argument Name=\"For"
                                "mName\">frm_Lookups</Argument></Action></Statements></UserInterfaceMacro>"
                        End
                    End

                    LayoutCachedLeft =3120
                    LayoutCachedTop =2040
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =3300
                    ForeTint =100.0
                    BackColor =8289145
                    BackTint =100.0
                    BorderColor =8289145
                    BorderTint =100.0
                    ThemeFontIndex =-1
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =23
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =87
                    AccessKey =82
                    Left =3120
                    Top =660
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =6
                    Name ="btnPreSeasonPrep"
                    Caption ="P&re-Season Prep"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Create BE backup and purge annual field data from tables"
                    UnicodeAccessKey =114
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3120
                    LayoutCachedTop =660
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =1920
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    BackColor =0
                    BackThemeColorIndex =0
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =87
                    AccessKey =79
                    Left =4560
                    Top =660
                    Width =1260
                    Height =1260
                    FontSize =14
                    FontWeight =700
                    TabIndex =7
                    Name ="btnPostSeasonChecks"
                    Caption ="P&ost-Season Checks"
                    OnClick ="[Event Procedure]"
                    FontName ="Franklin Gothic Book"
                    ControlTipText ="Open a form to check that RIO tags are actually in the office"
                    UnicodeAccessKey =111
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =4560
                    LayoutCachedTop =660
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =1920
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                    Shape =0
                    BackColor =0
                    BackThemeColorIndex =0
                    BackTint =100.0
                    BorderColor =0
                    BorderThemeColorIndex =0
                    BorderTint =100.0
                    HoverColor =65280
                    HoverThemeColorIndex =-1
                    HoverTint =100.0
                    PressedColor =6644321
                    PressedShade =80.0
                    HoverForeColor =0
                    HoverForeTint =100.0
                    PressedForeColor =0
                    PressedForeTint =100.0
                    Shadow =-1
                    QuickStyle =22
                    QuickStyleMask =-1
                    WebImagePaddingTop =1
                    Overlaps =1
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
' MODULE:       frm_Utilities
' Level:        Application form module
' Version:      1.02
' Description:  Standard module - main form for various database functions
' Data source:  -
' Data access:  -
' Pages:        -
' Functions:    none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      Bonnie Campbell, August 16, 2019
' Revisions:
'               ML/GS - unknown   - 1.00 - initial version
'               BLC   - 8/16/2019 - 1.01 - documentation, error handling,
'                                          added Pre-Season Prep for purging db field data tables
'                                          renamed cmdXX to btnXX
'               BLC   - 9/26/2019 - 1.02 - added Post-Season Prep for RIO tag check
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Methods
' ---------------------------------
' ---------------------------------
' SUB:          Form_Open
' Description:  form opening actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 16, 2019
' Adapted:      -
' Revisions:
'   BLC   - 8/16/2019 - initial version
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler
    
    'check for DbAdmin functionality (app level DB_ADMIN set in the db_Module)
    Me.btnPreSeasonPrep.Visible = IIf(Nz(DB_ADMIN, False), True, False)

    Dim strCaption As String

    ' Set the application font to more closely match the forms.
    ' Useful in cases where the subforms use tables directly
    Application.SetOption "Default Font Name", "Arial"
    Application.SetOption "Default Font Size", 9

'    ' Set the table-driven caption of the switchboard
'    strCaption = Nz(DLookup("[Database_title]", "tsys_App_Releases", "[Release_ID] = '" _
'        & Me!Release_ID & "'"), "")
'    Me.Caption = strCaption

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_App_Releases) (#" & Err.Number & " - Form_Open[frm_Utilities])"
        Case 2001   ' Field name in DLookup improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_App_Releases) (#" & Err.Number & " - Form_Open[frm_Utilities])"
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_App_Releases) (#" & Err.Number & " - Form_Open[frm_Utilities])"
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - Form_Open[frm_Utilities])"
    End Select
    Resume Exit_Handler

End Sub

' ---------------------------------
' SUB:          Form_Load
' Description:  form loading actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 16, 2019
' Adapted:      -
' Revisions:
'   BLC   - 8/16/2019 - initial version
' ---------------------------------
Private Sub Form_Load()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - Form_Load[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Form_Current
' Description:  form current actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 16, 2019
' Adapted:      -
' Revisions:
'   BLC   - 8/16/2019 - initial version
' ---------------------------------
Private Sub Form_Current()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - Form_Current[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnBackupBE_Click
' Description:  make a backup database backend file
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      Bonnie Campbell, August 16, 2019
' Revisions:
'   ML/GS - unknown   - initial version
'   BLC   - 8/16/2019 - added documentation & error handling
' ---------------------------------
Private Sub btnBackupBE_Click()
On Error GoTo Err_Handler

    fxnMakeBackup

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnBackupBE_Click[frm_Utilities])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPreSeasonPrep_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 16, 2019
' Revisions:
'   BLC   - 7/16/2019 - initial version
' ---------------------------------
Private Sub btnPreSeasonPrep_Click()
On Error GoTo Err_Handler

    'copy BE db
    BackupDbBE
    
    'copy & purge tables
    PurgeAnnualData
    
' shift msg to PurgeAnnualData to display only when purging is selected
'    'update
'    MsgBox "Pre-season backup & annual data purge is complete." & vbCrLf _
'           & "Review APBU_* data tables before deleting them.", _
'           vbOKOnly + vbInformation, "Pre-Season Backup & Annual Db Prep Complete"

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnPreSeasonPrep_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnPostSeasonChecks_Click
' Description:  button click actions
' Assumptions:  -
' Notes:        -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, September 26, 2019
' Revisions:
'   BLC   - 9/26/2019 - initial version
' ---------------------------------
Private Sub btnPostSeasonChecks_Click()
On Error GoTo Err_Handler

    DoCmd.OpenForm "RIOCheck", acNormal, , , acFormAdd, acDialog

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
        Case Else
          MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
              "Error encountered (#" & Err.Number & " - btnPostSeasonChecks_Click[frm_Switchboard])"
    End Select
    Resume Exit_Handler
End Sub
