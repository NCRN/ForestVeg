Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4863
    DatasheetFontHeight =9
    ItemSuffix =6
    Left =3300
    Top =2175
    Right =8565
    Bottom =4080
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xde0128929108e340
    End
    RecordSource ="xref_Event_Contacts"
    Caption =" Observers"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    DatasheetForeColor12 =33554432
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
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
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =15527148
            Name ="FormHeader"
        End
        Begin Section
            Height =420
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2880
                    Left =60
                    Top =60
                    Width =3003
                    Height =357
                    ColumnWidth =2268
                    FontSize =14
                    ColumnInfo ="\"\";\"\";\"\";\"\";\"10\";\"0\""
                    Name ="cboContact_ID"
                    ControlSource ="Contact_ID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT tlu_Contacts.Contact_ID, [Last_Name] & (\", \"+[First_Name]) AS FullName "
                        "FROM tlu_Contacts ORDER BY tlu_Contacts.Crew, tlu_Contacts.Active, tlu_Contacts."
                        "Last_Name, tlu_Contacts.First_Name; "
                    ColumnWidths ="0;2880"
                    StatusBarText ="Observer identifier"
                    FontName ="Calibri"
                    OnNotInList ="[Event Procedure]"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3063
                    LayoutCachedHeight =417
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ListWidth =2160
                    Left =3120
                    Top =60
                    Width =1743
                    Height =360
                    ColumnWidth =2376
                    FontSize =14
                    TabIndex =1
                    ColumnInfo ="\"\";\"\";\"10\";\"100\""
                    Name ="cboObserver_Role"
                    ControlSource ="Contact_Role"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Enum_Code FROM tlu_Enumerations WHERE Enum_Group=\"Contact Role\" ORDER B"
                        "Y Sort_Order; "
                    ColumnWidths ="2160"
                    StatusBarText ="Comments about the observer specific to this sampling event"
                    FontName ="Calibri"

                    LayoutCachedLeft =3120
                    LayoutCachedTop =60
                    LayoutCachedWidth =4863
                    LayoutCachedHeight =420
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
Option Explicit

' =================================
' FORM NAME:    fsub_Observers
' Description:  Data entry form for observers associated with sampling events
' Data source:  tbl_Observers
' Data access:  edit, add and delete
' Pages:        none
' Functions:    none
' References:   none
' Source/date:  John R. Boetsch, June 2006
' Revisions:    <name, date, desc - add lines as you go>
' =================================

Private Sub cboContact_ID_NotInList(NewData As String, Response As Integer)
    On Error GoTo Err_Handler

    Dim ctl As Control

    Set ctl = Me!cboContact_ID
    ' Prompt user to verify they wish to add new value
    If MsgBox("The contact is not in list. Would you like to add this name?", vbYesNo) = vbYes Then
        Response = acDataErrContinue
        ctl.Undo
        DoCmd.OpenForm "frm_Contacts", , , , , , "new"
    Else
    ' Suppress error message and undo changes
        Response = acDataErrContinue
        ctl.Undo
    End If

Exit_Procedure:
    On Error Resume Next
    Set ctl = Nothing
    Exit Sub

Err_Handler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
    Resume Exit_Procedure

End Sub
