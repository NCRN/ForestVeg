Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DataEntry = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularCharSet =204
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridX =24
    GridY =24
    Width =12660
    DatasheetFontHeight =9
    ItemSuffix =23
    Left =-31471
    Top =9855
    Right =-18556
    Bottom =14340
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0xde59bba555ace340
    End
    RecordSource ="tbl_Tags_History"
    Caption ="Change Log"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            SizeMode =3
            PictureAlignment =2
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
        Begin FormHeader
            Height =0
            BackColor =3751056
            Name ="FormHeader"
        End
        Begin Section
            Height =4500
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3524
                    Top =4050
                    Width =3030
                    Height =359
                    ColumnWidth =4200
                    FontSize =12
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxTagHistoryID"
                    ControlSource ="Tags_History_ID"
                    StatusBarText ="MA. Field data table row identifier (Data_ID)"

                    LayoutCachedLeft =3524
                    LayoutCachedTop =4050
                    LayoutCachedWidth =6554
                    LayoutCachedHeight =4409
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =1095
                            Top =4050
                            Width =2369
                            Height =359
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblTagHistory"
                            Caption ="Tag_History_ID:"
                            LayoutCachedLeft =1095
                            LayoutCachedTop =4050
                            LayoutCachedWidth =3464
                            LayoutCachedHeight =4409
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2294
                    Top =1920
                    Width =10185
                    Height =1186
                    ColumnWidth =2370
                    FontSize =12
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxHistoryNotes"
                    ControlSource ="Value_History_Notes"
                    StatusBarText ="Comments about this identification change"
                    ConditionalFormat = Begin
                        0x0100000098000000010000000100000000000000000000001b00000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b0074007800740048006900730074006f00 ,
                        0x720079005f004e006f007400650073005d00290000000000
                    End

                    LayoutCachedLeft =2294
                    LayoutCachedTop =1920
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =3106
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a5001a0000004900 ,
                        0x73004e0075006c006c0028005b0074007800740048006900730074006f007200 ,
                        0x79005f004e006f007400650073005d0029000000000000000000000000000000 ,
                        0x00000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Top =1920
                            Width =2219
                            Height =661
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblNotesDescription"
                            Caption ="Please describe why you made this change"
                            LayoutCachedTop =1920
                            LayoutCachedWidth =2219
                            LayoutCachedHeight =2581
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9869
                    Top =3210
                    Width =2610
                    Height =359
                    FontSize =12
                    TabIndex =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="tbxNetworkUserName"
                    ControlSource ="Network_User_Name"
                    StatusBarText ="The network user name of the person making the change"

                    LayoutCachedLeft =9869
                    LayoutCachedTop =3210
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =3569
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =7440
                            Top =3210
                            Width =2369
                            Height =359
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblNetworkUserName"
                            Caption ="Network User Name"
                            LayoutCachedLeft =7440
                            LayoutCachedTop =3210
                            LayoutCachedWidth =9809
                            LayoutCachedHeight =3569
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2279
                    Top =3630
                    Width =4260
                    Height =359
                    FontSize =12
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxChangeDate"
                    ControlSource ="Change_Date"
                    Format ="Short Date"
                    StatusBarText ="Date that species identification was changed for this specimen"
                    InputMask ="99/99/0000;0;_"

                    LayoutCachedLeft =2279
                    LayoutCachedTop =3630
                    LayoutCachedWidth =6539
                    LayoutCachedHeight =3989
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =75
                            Top =3630
                            Width =2129
                            Height =359
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblChangeDate"
                            Caption ="Date of Change"
                            LayoutCachedLeft =75
                            LayoutCachedTop =3630
                            LayoutCachedWidth =2204
                            LayoutCachedHeight =3989
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2279
                    Top =3195
                    Width =4275
                    Height =359
                    FontSize =12
                    TabIndex =3
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    ConditionalFormat = Begin
                        0x0100000092000000010000000100000000000000000000001800000001000000 ,
                        0x00000000dfa7a500000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x490073004e0075006c006c0028005b00630062006f0043006f006e0074006100 ,
                        0x630074005f00490044005d00290000000000
                    End
                    Name ="cbxContactID"
                    ControlSource ="Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                        "Last_Name, tlu_Contacts.First_Name;"
                    ColumnWidths ="0;2880"
                    StatusBarText ="M. Contact identifier (Contact_ID)"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =2279
                    LayoutCachedTop =3195
                    LayoutCachedWidth =6554
                    LayoutCachedHeight =3554
                    ConditionalFormat14 = Begin
                        0x01000100000001000000000000000100000000000000dfa7a500170000004900 ,
                        0x73004e0075006c006c0028005b00630062006f0043006f006e00740061006300 ,
                        0x74005f00490044005d0029000000000000000000000000000000000000000000 ,
                        0x00
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =3195
                            Width =2129
                            Height =359
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblContact"
                            Caption ="Changed By"
                            LayoutCachedLeft =60
                            LayoutCachedTop =3195
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =3554
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10860
                    Top =3885
                    Width =1605
                    Height =450
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    ForeColor =4754549
                    Name ="btnAccept"
                    Caption ="Accept Change"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10860
                    LayoutCachedTop =3885
                    LayoutCachedWidth =12465
                    LayoutCachedHeight =4335
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9240
                    Top =3885
                    Width =1605
                    Height =450
                    FontSize =12
                    TabIndex =5
                    ForeColor =3751056
                    Name ="btnCancel"
                    Caption ="Cancel Change"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =9240
                    LayoutCachedTop =3885
                    LayoutCachedWidth =10845
                    LayoutCachedHeight =4335
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2294
                    Top =975
                    Width =4200
                    Height =359
                    FontSize =14
                    FontWeight =700
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxValueNew"
                    ControlSource ="Value_New"
                    StatusBarText ="New TSN of Specimen"

                    LayoutCachedLeft =2294
                    LayoutCachedTop =975
                    LayoutCachedWidth =6494
                    LayoutCachedHeight =1334
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =600
                            Top =975
                            Width =1589
                            Height =359
                            FontSize =14
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblValueNew"
                            Caption ="New Value"
                            LayoutCachedLeft =600
                            LayoutCachedTop =975
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =1334
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8249
                    Top =960
                    Width =4230
                    Height =359
                    FontSize =14
                    FontWeight =700
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="tbxValueOld"
                    ControlSource ="Value_Old"
                    StatusBarText ="Previous TSN of Specimen"

                    LayoutCachedLeft =8249
                    LayoutCachedTop =960
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =1319
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6705
                            Top =960
                            Width =1424
                            Height =359
                            FontSize =14
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblValueOld"
                            Caption ="Old Value"
                            LayoutCachedLeft =6705
                            LayoutCachedTop =960
                            LayoutCachedWidth =8129
                            LayoutCachedHeight =1319
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2294
                    Top =1425
                    Width =4200
                    Height =374
                    FontSize =12
                    FontWeight =700
                    TabIndex =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="tbxDescriptionNew"
                    StatusBarText ="New TSN of Specimen"

                    LayoutCachedLeft =2294
                    LayoutCachedTop =1425
                    LayoutCachedWidth =6494
                    LayoutCachedHeight =1799
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =600
                            Top =1425
                            Width =1589
                            Height =374
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblDescriptionNew"
                            Caption ="Description"
                            LayoutCachedLeft =600
                            LayoutCachedTop =1425
                            LayoutCachedWidth =2189
                            LayoutCachedHeight =1799
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    FontItalic = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8249
                    Top =1410
                    Width =4230
                    Height =374
                    FontSize =12
                    FontWeight =700
                    TabIndex =11
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    BackColor =11056034
                    Name ="tbxDescriptionOld"
                    StatusBarText ="Previous TSN of Specimen"

                    LayoutCachedLeft =8249
                    LayoutCachedTop =1410
                    LayoutCachedWidth =12479
                    LayoutCachedHeight =1784
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6705
                            Top =1410
                            Width =1424
                            Height =374
                            FontSize =12
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            Name ="lblDescriptionOld"
                            Caption ="Description"
                            LayoutCachedLeft =6705
                            LayoutCachedTop =1410
                            LayoutCachedWidth =8129
                            LayoutCachedHeight =1784
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =120
                    Width =1974
                    Height =479
                    FontSize =18
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =3751056
                    Name ="lblTitle"
                    Caption ="Change Log"
                    GridlineColor =-2147483616
                    HorizontalAnchor =2
                    LayoutCachedLeft =120
                    LayoutCachedWidth =2094
                    LayoutCachedHeight =479
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =525
                    Width =12360
                    Height =315
                    FontSize =13
                    ForeColor =3751056
                    Name ="lblDescription"
                    Caption ="Please confirm the revised SPECIES ID below"
                    LayoutCachedLeft =120
                    LayoutCachedTop =525
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =840
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    BorderWidth =3
                    OverlapFlags =85
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11159
                    Top =60
                    Width =1275
                    Height =299
                    FontWeight =700
                    TabIndex =12
                    BackColor =11056034
                    ForeColor =8355711
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"4\";\"4\""
                    Name ="cbxTag_ID"
                    ControlSource ="Record_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tbl_Tags.Tag_ID, tbl_Tags.Tag FROM tbl_Tags ORDER BY tbl_Tags.Tag; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="MA. Field data table row identifier (Data_ID)"
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22

                    LayoutCachedLeft =11159
                    LayoutCachedTop =60
                    LayoutCachedWidth =12434
                    LayoutCachedHeight =359
                    ForeThemeColorIndex =1
                    ForeShade =50.0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =10380
                            Top =75
                            Width =690
                            Height =285
                            FontWeight =700
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =8355711
                            Name ="lblTag_ID"
                            Caption ="Tag"
                            LayoutCachedLeft =10380
                            LayoutCachedTop =75
                            LayoutCachedWidth =11070
                            LayoutCachedHeight =360
                            ForeThemeColorIndex =1
                            ForeShade =50.0
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =5760
                    Left =1980
                    Top =2640
                    Width =240
                    Height =315
                    FontSize =12
                    TabIndex =2
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cbxQuickComment"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Enumerations.Enum_Code FROM tlu_Enumerations WHERE (((tlu_Enumeration"
                        "s.Enum_Group)=\"Quick Comments\")) ORDER BY tlu_Enumerations.Sort_Order;"
                    ColumnWidths ="5760"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =2640
                    LayoutCachedWidth =2220
                    LayoutCachedHeight =2955
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =60
                            Top =2640
                            Width =1860
                            Height =320
                            Name ="lblQuickComment"
                            Caption ="Quick Comment ->"
                            LayoutCachedLeft =60
                            LayoutCachedTop =2640
                            LayoutCachedWidth =1920
                            LayoutCachedHeight =2960
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
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
' MODULE:       frm_Tags_History_Confirm
' Level:        Form module
' Version:      1.01
'
' Description:  tag history confirmation related functions & procedures
'
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, April 5, 2018
' Revisions:    ML/GS - unknown  - 1.00 - initial version
'               BLC   - 11/5/2018 - 1.01 - added documentation, error handling
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ----------------
'  Properties
' ----------------
Public ctlToReset As Control
Public frmReferrer As Form

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
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
' ---------------------------------
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo Err_Handler

    'Generate string GUID for Tag_History_ID
    If Me.NewRecord Then
        If GetDataType("tbl_Tags_History", "Tag_History_ID") = dbText Then
            Me!Tag_History_ID = fxnGUIDGen
        End If
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Form_BeforeUpdate[frm_Tags_History_Confirm])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnNewValueKeypad_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnNewValueKeypad_Click()
On Error GoTo Err_Handler

  'This routine requires the presence of the Keypad_Utils module.
  Dim strKeypadFormName As String
  Dim strControlToUpdate As String
  Dim frmFormToUpdate As Form
    
  'The two lines below should be changed to reflect the name of the keypad to open
  '    and the name of the control to be updated.
  strKeypadFormName = "frm_Pad_Num"
  strControlToUpdate = "txtValue_New"
  'The lines below should not usually be edited.
  Set frmFormToUpdate = Me
  Call OpenKeypad(strKeypadFormName, frmFormToUpdate, strControlToUpdate)

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxQuickComment_AfterUpdate[frm_Tags_History_Confirm])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          cbxQuickComment_AfterUpdate
' Description:  combobox after update actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
' ---------------------------------
Private Sub cbxQuickComment_AfterUpdate()
On Error GoTo Err_Handler

    Me.tbxHistoryNotes = LTrim(Me.tbxHistoryNotes & " " & Me.cbxQuickComment)
    Me.tbxHistoryNotes.Requery
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cbxQuickComment_AfterUpdate[frm_Tags_History_Confirm])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          btnAccept_Click
' Description:  button click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Mark Lehman/Geoffrey Sanders, unknown
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
'                                   conditionally refresh form events tag history (if open only)
' ---------------------------------
Private Sub btnAccept_Click()
On Error GoTo Err_Handler

    If Me.Dirty Then Me.Dirty = False
    ctlToReset.Value = Me!Value_New
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    frmReferrer.SaveRecord
    
    'refresh tag history on event form (if open)
    If FormIsOpen("frm_Events") Then _
        Forms![frm_Events]![fsub_Tree_Data]![fsub_Tags_History_Summary].Requery

    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnAccept_Click[frm_Tags_History_Confirm])"
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
' Adapted:      Bonnie Campbell, November 5, 2018
' Revisions:    ML/GS - unknown  - initial version
'               BLC   - 11/5/2018 - added documentation, error handling
' ---------------------------------
Private Sub btnCancel_Click()
On Error GoTo Err_Handler
        
    'Command below is not needed when for is called from button instead of BeforeUpdate
    ctlToReset.Value = ctlToReset.OldValue
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdDeleteRecord
    DoCmd.SetWarnings True
    MsgBox "Value was NOT changed", vbInformation, "Change cancelled"
    DoCmd.Close , , acSaveNo

Exit_Handler:
    Exit Sub
    
Err_Handler:
    DoCmd.SetWarnings True
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - btnCancel_Click[frm_Tags_History_Confirm])"
    End Select
    Resume Exit_Handler
End Sub
