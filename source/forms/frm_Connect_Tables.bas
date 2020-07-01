Version =21
VersionRequired =20
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10254
    DatasheetFontHeight =10
    ItemSuffix =96
    Left =4860
    Top =1410
    Right =15375
    Bottom =7515
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb46db5e5f0f8e240
    End
    RecordSource ="tsys_Link_Files"
    Caption =" Update Data Table Connections"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin FormHeader
            Height =486
            BackColor =0
            Name ="FormHeader"
            BackThemeColorIndex =0
            Begin
                Begin Label
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =2
                    Width =10254
                    Height =486
                    FontSize =20
                    FontWeight =700
                    ForeColor =16777215
                    Name ="lblTitle"
                    Caption ="Relink Backend Database Tables"
                    FontName ="Calibri"
                    LayoutCachedWidth =10254
                    LayoutCachedHeight =486
                    ForeThemeColorIndex =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2340
            BackColor =15921906
            Name ="Detail"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =9300
                    Top =1380
                    Width =842
                    Height =842
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="btnBrowse"
                    Caption ="Browse"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9300
                    LayoutCachedTop =1380
                    LayoutCachedWidth =10142
                    LayoutCachedHeight =2222
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =2
                    BackStyle =0
                    Left =180
                    Top =120
                    Width =9897
                    Height =300
                    ColumnWidth =6630
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="tbxLinkDescription"
                    ControlSource ="Link_description"
                    StatusBarText ="Describes the types of data tables included in the link"
                    FontName ="Calibri"

                    LayoutCachedLeft =180
                    LayoutCachedTop =120
                    LayoutCachedWidth =10077
                    LayoutCachedHeight =420
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    Left =3000
                    Top =540
                    Width =6183
                    Height =300
                    ColumnWidth =2520
                    FontSize =12
                    TabIndex =2
                    Name ="tbxCurrentName"
                    ControlSource ="Link_file_name"
                    FontName ="Calibri"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =540
                    LayoutCachedWidth =9183
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =540
                            Width =2706
                            Height =324
                            FontSize =12
                            FontWeight =700
                            Name ="lblCurrentName"
                            Caption ="CURRENT Name and Path:"
                            FontName ="Calibri"
                            LayoutCachedLeft =180
                            LayoutCachedTop =540
                            LayoutCachedWidth =2886
                            LayoutCachedHeight =864
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    BackStyle =0
                    Left =180
                    Top =900
                    Width =8997
                    Height =300
                    ColumnWidth =2205
                    FontSize =9
                    TabIndex =3
                    Name ="tbxCurrentPath"
                    ControlSource ="Link_file_path"
                    StatusBarText ="Current linked file path"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =900
                    LayoutCachedWidth =9177
                    LayoutCachedHeight =1200
                End
                Begin TextBox
                    Locked = NotDefault
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =1800
                    Width =8997
                    Height =312
                    FontSize =9
                    TabIndex =5
                    Name ="tbxNewPath"
                    ControlSource ="New_file_path"
                    StatusBarText ="New linked file path"
                    FontName ="Calibri"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =180
                    LayoutCachedTop =1800
                    LayoutCachedWidth =9177
                    LayoutCachedHeight =2112
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    CanShrink = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3000
                    Top =1440
                    Width =6183
                    Height =312
                    FontSize =12
                    TabIndex =4
                    Name ="tbxNewName"
                    ControlSource ="New_file_name"
                    FontName ="Calibri"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =1440
                    LayoutCachedWidth =9183
                    LayoutCachedHeight =1752
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =180
                            Top =1440
                            Width =2700
                            Height =312
                            FontSize =12
                            FontWeight =700
                            Name ="lblNewName"
                            Caption ="REVISED Name and Path:"
                            FontName ="Calibri"
                            LayoutCachedLeft =180
                            LayoutCachedTop =1440
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =1752
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =960
            BackColor =15921906
            Name ="FormFooter"
            BackThemeColorIndex =1
            BackShade =95.0
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =6600
                    Width =2583
                    Height =842
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
                    Name ="btnUpdateLinks"
                    Caption ="Update links"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Update links to the file(s) indicated"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =6600
                    LayoutCachedWidth =9183
                    LayoutCachedHeight =842
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =6731160
                    HoverThemeColorIndex =7
                    HoverTint =80.0
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
                    Left =9300
                    Width =842
                    Height =842
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="btnClose"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ControlTipText ="Close the form"
                    LeftPadding =60
                    RightPadding =75
                    BottomPadding =120

                    LayoutCachedLeft =9300
                    LayoutCachedWidth =10142
                    LayoutCachedHeight =842
                    ForeThemeColorIndex =0
                    UseTheme =1
                    Shape =1
                    Gradient =12
                    BackColor =8289145
                    BackThemeColorIndex =4
                    BorderColor =8289145
                    BorderThemeColorIndex =4
                    HoverColor =10798077
                    HoverThemeColorIndex =5
                    HoverTint =40.0
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
' FORM:         frm_Connect_Tables form
' Level:        Application form module
' Version:      1.04
'
' Description:  updates back-end db connections, related functions & procedures
'
' Data source:  tsys_Link_Files
' Data access:  edit only, no additions, moving between records or deletions
' Pages:        none
' Functions:    fxnGetLinkFile, fxnRefreshLinks, fxnSwitchboardIsOpen
' References:   -
' Source/date:  Susan Huse, MonitoringSM.mdb v July 28, 2004
' Revisions:    SH  - 5/xx/2005 - 1.00 - initial version
'               JRB - 5/xx/2005 - 1.01 - minor edits
'               JRB - 5/24/2006 - 1.02 - documentation, added error trapping,
'                                        fixed specification of initial directory
'                                        to current directory, simplified a little
'               BLC - 1/30/2019 - 1.03 - added documenation, error handling
'               BLC - 4/2/2020  - 1.04 - added trailing slash for opening directory
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

' ----------------
'  Events
' ----------------

' ----------------
'  Form
' ----------------
' ---------------------------------
' SUB:          Form_Open
' Description:  form open actions
' Assumptions:  -
' Parameters:   Cancel - whether open action(s) should be cancelled (boolean)
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 5/23/2018 - update documentation, error handling
' ---------------------------------
Private Sub Form_Open(Cancel As Integer)
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_Open[frm_Connect_Tables form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxCurrentPath_Click
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 1/30/2019 - update documentation, error handling
' ---------------------------------
Private Sub tbxCurrentPath_Click()
On Error GoTo Err_Handler

    SendKeys "+{F2}"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxCurrentPath_Click[frm_Connect_Tables form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          tbxNewPath
' Description:  textbox click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 1/30/2019 - update documentation, error handling
' ---------------------------------
Private Sub tbxNewPath_Click()
On Error GoTo Err_Handler

    SendKeys "+{F2}"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - tbxNewPath_Click[frm_Connect_Tables form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnBrowse_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 1/30/2019 - update documentation, error handling
'   BLC - 4/2/2020  - update current directory path to add trailing slash
' ---------------------------------
Private Sub btnBrowse_Click()
On Error GoTo Err_Handler

    Dim strCurrentFile As String
    Dim strCurrentDir As String
    Dim varFilePath As Variant
    Dim arrFile() As String

    strCurrentFile = Me!tbxCurrentName
    strCurrentDir = Me!tbxCurrentPath & "\"

    ' Clip to indicate just the folder of the current back-end
    strCurrentDir = Left(strCurrentDir, Len(strCurrentDir) - Len(strCurrentFile) - 1)

    ' Select the file, and start the search in the current back-end folder
    varFilePath = fxnGetLinkFile(strCurrentDir)

    ' Exit if the user didn't specify a file
    If IsNull(varFilePath) Then GoTo Exit_Handler

    ' Update the new path and file name controls
    Me!tbxNewPath = varFilePath
    ' Update the new file name after first storing the path components in an array
    arrFile = Split(varFilePath, "\")
    Me!tbxNewName = arrFile(UBound(arrFile))
    Me!btnUpdateLinks.Enabled = True

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Files)"
            Resume Exit_Handler
        Case 2001   ' Field name in DLookup improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_Link_Files)"
            Resume Exit_Handler
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Files)"
            Resume Exit_Handler
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnBrowse_Click[frm_Connect_Tables form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnUpdateLinks_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 1/30/2019 - update documentation, error handling,
'                     update txtLinkPath to tbxLinkPath
' ---------------------------------
Private Sub btnUpdateLinks_Click()
On Error GoTo Err_Handler

    Dim rst As DAO.Recordset
    Dim strSysTable As String       ' Name of the system table listing linked tables
    Dim strLinkName As String
    Dim strFilePath As String       ' Path of the new database
    Dim strSQL As String
    Dim bHasError As Boolean

    'Save record commanded was required for correct execution of code in Access 2013 (32 bit).
    DoCmd.RunCommand acCmdSaveRecord
    strSysTable = "[tsys_Link_Tables]"  ' Set the name of the system table

    ' Set a loop in case of multiple back-ends.  If errors are encountered on one,
    '   go to the next loop rather than exit
    Set rst = Me.RecordsetClone
    rst.MoveFirst

    bHasError = False   ' Default until an error is encountered

    Do Until rst.EOF
        strLinkName = rst.Fields("Link_type")
        ' If the user didn't specify a different database,
        '   refresh the links to the current linked file
        If IsNull(rst.Fields("New_file_path")) Then
            strFilePath = rst.Fields("Link_file_path")
        Else
            strFilePath = rst.Fields("New_file_path")
        End If

        ' Build a query statement identifying the tables that should be in the file
        strSQL = "SELECT * FROM " & strSysTable & " WHERE " & _
            strSysTable & "![Link_type] = '" & strLinkName & "'"

        ' Verify the file and update the links to the selected file
        If fxnRefreshLinks(strSQL, strFilePath) = False Then
            ' An error was encountered
            MsgBox "Links to this file were not updated or only partially updated", _
                vbExclamation, strLinkName
            bHasError = True
            GoTo NextBackEnd
        ' If no linking error on this back end then update the current path and file
        ElseIf IsNull(rst.Fields("New_file_path")) = False Then
            With rst
                .Edit
                !Link_file_name = rst.Fields("New_file_name").value
                !Link_file_path = rst.Fields("New_file_path").value
                !New_file_name = Null
                !New_file_path = Null
                .Update
                .Bookmark = .LastModified
            End With
        End If

        ' If the switchboard is open and the current file is the primary back-end, then update
        '   the switchboard control for the current file link
        If fxnSwitchboardIsOpen And strLinkName = "Back-end data" And bHasError = False Then
            Forms![frm_Switchboard]![tbxLinkPath] = strFilePath
            Forms!frm_Switchboard.Refresh
            
            '10/23/2018 update
            SetTempVar "BEfilepath", strFilePath

        End If

NextBackEnd:
        On Error Resume Next
        If Err > 0 Then
            MsgBox "Error #" & Err.Number & ": " & Err.Description, _
                vbCritical, "Error encountered while updating database links"
            bHasError = True
        End If
        Err = 0
        rst.MoveNext
    Loop
    ' End the loop accommodating multiple back-end files here

    ' If no connection errors, then notify the user and close
    If bHasError = False Then
        MsgBox "Update complete!", vbExclamation, "Update Back-end Data Connections"
        DoCmd.Close , , acSaveNo
    End If

Exit_Handler:
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 3078   ' Can't find the system table
            MsgBox "Error #" & Err.Number & ":  Missing a system table. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Tables)"
        Case 3265   ' Field name in the tsys_Link_Files improperly specified
            MsgBox "Error #" & Err.Number & ":  System table field not found." & _
                vbCrLf & "Please notify the database administrator before using " & _
                "this application.", vbCritical, "System table error (tsys_Link_Tables)"
        Case 94    ' Missing information in the systems table
            MsgBox "Error #" & Err.Number & ":  Missing system table info. Please notify" & _
                vbCrLf & "the database administrator before using this application.", _
                vbCritical, "System table error (tsys_Link_Tables)"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered when updating database links (#" & Err.Number & " - btnUpdateLinks_Click[frm_Connect_Tables form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnClose_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoff Sanders, unknown
' Adapted:      -
' Revisions:
'   MEL/GS - unknown - initial version
'   BLC - 1/30/2019 - update documentation, error handling
' ---------------------------------
Private Sub btnClose_Click()
On Error GoTo Err_Handler

    DoCmd.Close acForm, Me.Name, acSaveNo
    'clear new file name, new file path
    CurrentDb.Execute "UPDATE tsys_Link_Files SET New_file_name=null, New_file_path=null;"

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnClose_Click[frm_Connect_Tables form])"
    End Select
    Resume Exit_Handler
End Sub
