Version =21
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1800
    DatasheetFontHeight =11
    ItemSuffix =6
    Left =660
    Top =1395
    DatasheetGridlinesColor =14276557
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xb1555bbfbe81e540
    End
    RecordSource ="SELECT d.Tag_ID, e.Event_Date, dbh.Tree_Data_ID, dbh.DBH FROM ((tbl_Events e  IN"
        "NER JOIN tbl_Tree_Data d ON d.Event_ID = e.Event_ID)  INNER JOIN tbl_Tree_DBH db"
        "h ON dbh.Tree_Data_ID = d.Tree_Data_ID)  GROUP BY d.Tag_ID, dbh.Tree_Data_ID, e."
        "Event_Date, dbh.DBH ORDER BY e.Event_Date DESC;"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Franklin Gothic Book"
    PrtMip = Begin
        0x6a010000a8000000660100001e01000000000000080700002c01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =255
    FitToPage =1
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
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =225
            Name ="ReportHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    Left =60
                    Width =1620
                    Height =225
                    FontSize =7
                    FontWeight =600
                    Name ="lblPriorDBH"
                    Caption ="Prior DBH Stems (cm)"
                    FontName ="Franklin Gothic Book"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =225
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =300
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =1
                    OldBorderStyle =0
                    TextAlign =2
                    IMESentenceMode =3
                    Width =720
                    FontSize =7
                    ForeColor =16711680
                    Name ="tbxPriorDBH"
                    ControlSource ="=TruncateNumber([DBH],1)"
                    Format ="General Number"
                    ConditionalFormat = Begin
                        0x01000000b0000000010000000100000000000000000000002700000001000000 ,
                        0xffffff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004500760065006e00740044006100740065005d003e003d00 ,
                        0x5b00740062007800530061006d0070006c0069006e0067004500760065006e00 ,
                        0x740044006100740065005d0000000000
                    End

                    LayoutCachedWidth =720
                    LayoutCachedHeight =240
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff00ffffff00260000005b00 ,
                        0x7400620078004500760065006e00740044006100740065005d003e003d005b00 ,
                        0x740062007800530061006d0070006c0069006e0067004500760065006e007400 ,
                        0x44006100740065005d00000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1320
                    Width =300
                    Height =270
                    FontSize =8
                    TabIndex =1
                    Name ="tbxTagID"
                    ControlSource ="Tag_ID"
                    StatusBarText ="Number of physical tag attached to tree"
                    FontName ="Franklin Gothic Book"

                    LayoutCachedLeft =1320
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =270
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    BackStyle =0
                    IMESentenceMode =3
                    Left =1200
                    Width =300
                    Height =270
                    FontSize =8
                    TabIndex =2
                    Name ="tbxDataID"
                    ControlSource ="Tree_Data_ID"
                    FontName ="Franklin Gothic Book"

                    LayoutCachedLeft =1200
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =270
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =720
                    Width =1020
                    Height =255
                    FontSize =8
                    TabIndex =3
                    Name ="tbxEventDate"
                    ControlSource ="Event_Date"
                    Format ="Short Date"
                    StatusBarText ="Number of physical tag attached to tree"
                    FontName ="Franklin Gothic Book"
                    ConditionalFormat = Begin
                        0x01000000b0000000010000000100000000000000000000002700000001000000 ,
                        0xffffff00ffffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b007400620078004500760065006e00740044006100740065005d003e003d00 ,
                        0x5b00740062007800530061006d0070006c0069006e0067004500760065006e00 ,
                        0x740044006100740065005d0000000000
                    End

                    LayoutCachedLeft =720
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =255
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    ConditionalFormat14 = Begin
                        0x010001000000010000000000000001000000ffffff00ffffff00260000005b00 ,
                        0x7400620078004500760065006e00740044006100740065005d003e003d005b00 ,
                        0x740062007800530061006d0070006c0069006e0067004500760065006e007400 ,
                        0x44006100740065005d00000000000000000000000000000000000000000000
                    End
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1320
                    Width =360
                    Height =255
                    FontSize =8
                    TabIndex =4
                    Name ="tbxSamplingEventDate"
                    ControlSource ="=[Parent].[Report].[tbxSamplingEventDate]"
                    FontName ="Franklin Gothic Book"

                    LayoutCachedLeft =1320
                    LayoutCachedWidth =1680
                    LayoutCachedHeight =255
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderShade =100.0
                    ThemeFontIndex =-1
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =29
            Name ="ReportFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    Visible = NotDefault
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Width =788
                    Height =29
                    FontSize =7
                    ForeColor =16711680
                    Name ="lblNoData"
                    Caption ="No Stem Data"
                    LayoutCachedLeft =60
                    LayoutCachedWidth =848
                    LayoutCachedHeight =29
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
                End
                Begin Label
                    Visible = NotDefault
                    FontItalic = NotDefault
                    BackStyle =1
                    TextAlign =2
                    Left =960
                    Width =788
                    Height =29
                    FontSize =7
                    ForeColor =16711680
                    Name ="lblHasData"
                    Caption ="Stem Data"
                    LayoutCachedLeft =960
                    LayoutCachedWidth =1748
                    LayoutCachedHeight =29
                    ThemeFontIndex =-1
                    BackThemeColorIndex =-1
                    BorderThemeColorIndex =-1
                    BorderTint =100.0
                    ForeThemeColorIndex =-1
                    ForeTint =100.0
                    GridlineThemeColorIndex =-1
                    GridlineShade =100.0
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
' Report:       Tree PriorDBHStems
' Level:        Application report
' Version:      1.00
'
' Description:  Tree PriorDBHStems report object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, August 2, 2020
' References:   -
' Revisions:    BLC - 8/2/2020 - 1.00 - initial version
' =================================

'---------------------
' Simulated Inheritance
'---------------------

'---------------------
' Declarations
'---------------------
Private m_Title As String
Private m_Directions As String
Private m_CallingForm As String

'---------------------
' Event Declarations
'---------------------
Public Event InvalidCallingForm(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let CallingForm(Value As String)
    If Len(Value) > 0 Then
        m_CallingForm = Value
    Else
        RaiseEvent InvalidCallingForm(Value)
    End If
End Property

Public Property Get CallingForm() As String
    CallingForm = m_CallingForm
End Property

'---------------------
' Events
'---------------------
' ---------------------------------
' Sub:          Report_Open
' Description:  Report opening event actions
' Assumptions:  Event date is present on rSub_Event_Summary_Unfiltered
'               Form hierarchy is
'                  rpt_Event_Summary_Unfiltered > rSub_Event_Trees > rSub_Event_rSub_Tree_PriorDBH_Stems
'               NOTE:  using Me.Filter results in
'                      Error #2101 - The setting you entered isn't valid for this property
'                      Possibly due to load sequence, however recordsource is more rapidly changed
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 8/2/2020 - initial version
' ---------------------------------
Private Sub Report_Open(Cancel As Integer)
On Error GoTo Err_Handler

'Dim SubFilter As String
'
'    If Not IsNull(Me.Parent.Parent.Event_Date) Then
'        SubFilter = "e.Event_Date > #" & CDate(Me.Parent.Parent.Event_Date) & "#"
'    Debug.Print SubFilter
'        'insert the WHERE *before* the GROUP BY clause
'        Me.RecordSource = Replace(Me.RecordSource, "GROUP BY", " WHERE " & SubFilter & " GROUP BY")
'    '        Me.FilterOn = True
'    End If
'
'    Debug.Print Me.RecordSource

    '----------------------------------------
    'set values for the stems DBH subreport
    '----------------------------------------
    Dim SubFilter As String

    If Not IsNull(Parent.Report.tbxSamplingEventDate) Then 'skip record
    
        Debug.Print Me.RecordSource
        
    End If


'    If Not IsNull(Me.Parent.Parent.Event_Date) And Me.Parent.SetOnce = False Then
'        SubFilter = "e.Event_Date > #" & CDate(Me.Parent.Parent.Event_Date) & "#"
'    Debug.Print SubFilter
'        'insert the WHERE *before* the GROUP BY clause
''        Me.rsub_Event_Tree_PriorDBH_Stems.Report.RecordSource = Replace(Me.rsub_Event_Tree_PriorDBH_Stems.Report.RecordSource, "GROUP BY", " WHERE " & SubFilter & " GROUP BY")
'    '        Me.FilterOn = True
'        Me.Parent.SetOnce = True
'    End If

'    Debug.Print Me.rsub_Event_Tree_PriorDBH_Stems.Report.RecordSource
    '----------------------------------------

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Open[Tree PriorDBHStems Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Report_Close
' Description:  Closing event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 8/2/2020 - initial version
' ---------------------------------
Private Sub Report_Close()
On Error GoTo Err_Handler

    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_Close[Tree PriorDBHStems Report])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Detail_Format
' Description:  detail format actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 5, 2020
' Adapted:      -
' Revisions:
'   BLC - 8/5/2020 - initial version
' ---------------------------------
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
On Error GoTo Err_Handler

    '----------------------------------------
    'set values for the stems DBH subreport
    '----------------------------------------
    Dim SubFilter As String
    
'    If Parent.Report.tbxSamplingEventDate Then 'skip record

'    If Not IsNull(Me.Parent.Parent.Event_Date) And Me.Parent.SetOnce = False Then
'        SubFilter = "e.Event_Date > #" & CDate(Me.Parent.Parent.Event_Date) & "#"
'    Debug.Print SubFilter
'        'insert the WHERE *before* the GROUP BY clause
''        Me.rsub_Event_Tree_PriorDBH_Stems.Report.RecordSource = Replace(Me.rsub_Event_Tree_PriorDBH_Stems.Report.RecordSource, "GROUP BY", " WHERE " & SubFilter & " GROUP BY")
'    '        Me.FilterOn = True
'        Me.Parent.SetOnce = True
'    End If

'    Debug.Print Me.rsub_Event_Tree_PriorDBH_Stems.Report.RecordSource
    '----------------------------------------
'End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Detail_Format[Tree PriorDBHStems Report])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Report_NoData
' Description:  Closing event actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, August 2, 2020
' Adapted:      -
' Revisions:
'   BLC - 8/2/2020 - initial version
' ---------------------------------
Private Sub Report_NoData(Cancel As Integer)
On Error GoTo Err_Handler

'    Me.lblNoData.visible = True
    ' hide subreport if no data
    Cancel = True
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Report_NoData[Tree PriorDBHStems Report])"
    End Select
    Resume Exit_Handler
End Sub
