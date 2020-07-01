Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =3240
    Top =390
    Right =10695
    Bottom =9375
    DatasheetGridlinesColor =12632256
    Filter ="ID='6'"
    RecSrcDt = Begin
        0x724af1170fb2e340
    End
    RecordSource ="tsys_Import_Tables"
    Caption ="Select Import Tables"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =2940
            BackColor =15527148
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =120
                    Width =6420
                    Height =480
                    FontSize =18
                    Name ="Label2"
                    Caption ="Select the tables you would like to import:"
                    FontName ="Calibri"
                End
                Begin Label
                    FontItalic = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =660
                    Width =6360
                    Height =240
                    FontSize =10
                    Name ="Label4"
                    Caption ="Note that the default is not to import 'tbl_Locations'"
                    FontName ="Calibri"
                End
                Begin OptionGroup
                    OverlapFlags =93
                    Left =120
                    Top =1200
                    Width =6597
                    Height =605
                    ColumnOrder =0
                    Name ="optframe_Step1Import"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =215
                            Left =240
                            Top =1080
                            Width =1140
                            Height =300
                            FontSize =12
                            FontWeight =700
                            BackColor =15527148
                            ForeColor =10485760
                            Name ="Label7"
                            Caption ="Step 1"
                            FontName ="Calibri"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3960
                            Top =1440
                            Width =1260
                            Height =300
                            FontWeight =700
                            OptionValue =1
                            Name ="Toggle9"
                            Caption ="One Tablet"
                            FontName ="Calibri"

                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =5280
                            Top =1440
                            Width =1320
                            Height =300
                            FontWeight =700
                            TabIndex =1
                            OptionValue =2
                            Name ="Toggle10"
                            Caption ="Two Tablets"
                            FontName ="Calibri"

                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =240
                    Top =1500
                    Width =3570
                    Height =210
                    Name ="Label11"
                    Caption ="How many tablets were used to collect the data:"
                    FontName ="Calibri"
                End
                Begin OptionGroup
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =120
                    Top =2040
                    Width =6597
                    Height =725
                    ColumnOrder =1
                    TabIndex =1
                    Name ="optframe_Step2Import"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            BackStyle =1
                            OverlapFlags =215
                            Left =240
                            Top =1920
                            Width =1140
                            Height =300
                            FontSize =12
                            FontWeight =700
                            BackColor =15527148
                            ForeColor =10485760
                            Name ="Label13"
                            Caption ="Step 2"
                            FontName ="Calibri"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3720
                            Top =2340
                            Width =1140
                            Height =330
                            FontWeight =700
                            OptionValue =1
                            Name ="Toggle15"
                            Caption ="Primary Tablet"
                            FontName ="Calibri"

                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =4920
                            Top =2340
                            Width =1695
                            Height =330
                            FontWeight =700
                            TabIndex =1
                            OptionValue =2
                            Name ="Toggle16"
                            Caption ="Secondary Tablet"
                            FontName ="Calibri"

                            WebImagePaddingLeft =2
                            WebImagePaddingTop =2
                            WebImagePaddingRight =1
                            WebImagePaddingBottom =1
                        End
                    End
                End
                Begin Label
                    OverlapFlags =215
                    Left =240
                    Top =2280
                    Width =3360
                    Height =420
                    Name ="Label17"
                    Caption ="Are you importing the main tablet or secondary tablet:"
                    FontName ="Calibri"
                End
            End
        End
        Begin Section
            Height =345
            BackColor =15527148
            Name ="Detail"
            OnClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1140
                    Top =30
                    Width =3240
                    ColumnWidth =2475
                    Name ="txt_Table_Name"
                    ControlSource ="Table_Name"
                    FontName ="Calibri"

                    LayoutCachedLeft =1140
                    LayoutCachedTop =30
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =270
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =30
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="Table Name:"
                            FontName ="Calibri"
                            LayoutCachedLeft =60
                            LayoutCachedTop =30
                            LayoutCachedWidth =1080
                            LayoutCachedHeight =270
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6120
                    Top =45
                    Height =210
                    ColumnWidth =1680
                    TabIndex =1
                    Name ="chk_Import"
                    ControlSource ="Import"

                    LayoutCachedLeft =6120
                    LayoutCachedTop =45
                    LayoutCachedWidth =6380
                    LayoutCachedHeight =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4605
                            Top =45
                            Width =1380
                            Height =240
                            Name ="Label1"
                            Caption ="Import this table?"
                            FontName ="Calibri"
                            LayoutCachedLeft =4605
                            LayoutCachedTop =45
                            LayoutCachedWidth =5985
                            LayoutCachedHeight =285
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =540
            BackColor =15527148
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =60
                    Top =60
                    Width =1680
                    Height =405
                    FontWeight =700
                    Name ="cmd_Open_Import_Form"
                    Caption ="Select Import File"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1920
                    Top =60
                    Width =720
                    Height =405
                    FontWeight =700
                    TabIndex =1
                    Name ="cmd_Cancel"
                    Caption ="Cancel"
                    OnClick ="[Event Procedure]"
                    FontName ="Calibri"

                    LayoutCachedLeft =1920
                    LayoutCachedTop =60
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =465
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
Option Explicit


Private Sub Form_Load()

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rsImportTables As DAO.Recordset
Set db = CurrentDb

DoCmd.RunCommand acCmdSaveRecord

Set rsImportTables = db.OpenRecordset("tsys_Import_Tables")
'
Set rs = Me.RecordsetClone

If intImport2 = 2 Then

    Me!optframe_Step1Import = 2
    Me.Requery
    Me!optframe_Step2Import.Enabled = True
    Me!optframe_Step2Import = 2
    Me.Requery
    Me!optframe_Step2Import.Enabled = True
    Me.Requery
    Me!cmd_Open_Import_Form.Enabled = True

    rsImportTables.MoveFirst

Me.RecordSource = "qry_Append_Secondary_Tablet_Import"
'Set rs = Me.Recordset
Me.Requery

        Do Until rsImportTables.EOF
            rsImportTables.Edit
            If rsImportTables![Table_Name] = "tbl_Locations" Then
                rsImportTables![Import] = False

            ElseIf rsImportTables![Table_Name] = "tbl_Events" Then
                rsImportTables![Import] = True

            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Data" Then
                rsImportTables![Import] = True

            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Seedlings_Data" Then
                rsImportTables![Import] = True

            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Herbaceous_Data" Then
                rsImportTables![Import] = True

            ElseIf rsImportTables![Table_Name] = "tbl_CWD_Data" Then
                rsImportTables![Import] = True

            Else
                rsImportTables![Import] = False

            End If
            rsImportTables.Update
        rsImportTables.MoveNext
        Loop


' Do Until rs.EOF
''    'Not clear why this was needed
'    Me.chk_Import.Value = 1
'    rs.MoveNext
' Loop

    intImport2 = 0
'
Else
'
'Set rs = Me.Recordset
rsImportTables.MoveFirst
'
 Do Until rsImportTables.EOF
    If Me.txt_Table_Name = "tbl_Locations" Then
        Me.chk_Import.value = 0
    Else
        Me.chk_Import.value = 1

    End If

 rsImportTables.MoveNext
 Loop

End If
'
Set db = Nothing
Set rs = Nothing
Set rsImportTables = Nothing

End Sub

Private Sub cmd_Open_Import_Form_Click()
On Error GoTo Err_cmd_Open_Import_Form_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "frm_Append_Select_Import_File"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "frm_Append_Select_Import_Tables"
    
Exit_cmd_Open_Import_Form_Click:
    Exit Sub
Err_cmd_Open_Import_Form_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Open_Import_Form_Click
End Sub

Private Sub Detail_Click()
If Me!optframe_Step1Import = "" Or IsNull(Me!optframe_Step1Import) Then
    MsgBox "Please answer the import questions above prior to selecting the tables " _
        & "to import.", , "Import Tables"
    Me!optframe_Step1Import.SetFocus
ElseIf Me!optframe_Step1Import = 2 Then
    MsgBox "Please select whether you are importing data from the main tablet " _
        & "or the secondary tablet.", , "Import Tables"
    Me!optframe_Step2Import.SetFocus
End If
End Sub

Private Sub cmd_Cancel_Click()
On Error GoTo Err_cmd_Cancel_Click

    DoCmd.Close

Exit_cmd_Cancel_Click:
    Exit Sub
Err_cmd_Cancel_Click:
    MsgBox Err.Description
    Resume Exit_cmd_Cancel_Click
End Sub


Private Sub optframe_Step1Import_AfterUpdate()

 'Set the public variable
 intImport2 = Me!optframe_Step1Import.value

Select Case optframe_Step1Import.value
Case 1
     optframe_Step2Import.Enabled = False
     Me!optframe_Step2Import.value = 0
     Me.RecordSource = "tsys_Import_Tables"
     Me!cmd_Open_Import_Form.Enabled = True
     
Case 2
    optframe_Step2Import.Enabled = True
    Me!optframe_Step2Import.value = 0
    Me!optframe_Step2Import.SetFocus
    
Case Else
    optframe_Step2Import.Enabled = False
  
End Select
Me.Refresh

End Sub

Private Sub optframe_Step2Import_AfterUpdate()

Dim db As DAO.Database
    Set db = CurrentDb
    Dim rsImportTables As DAO.Recordset
    Set rsImportTables = db.OpenRecordset("tsys_Import_Tables")
    
Me!cmd_Open_Import_Form.Enabled = True
   
Select Case optframe_Step2Import.value

Case 1
Me.RecordSource = "qry_Append_Primary_Tablet_Import"
Me.Requery
        rsImportTables.MoveFirst
        Do Until rsImportTables.EOF
            
            rsImportTables.Edit
            'Until all plots are updated with Slope and Aspect data we will import the Locations Table in case the updates come from the field database.
'            If rsImportTables![Table_Name] = "tbl_Locations" Then
'                rsImportTables![Import] = False
            
'            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Data" Then
'                rsImportTables![Import] = False
            If rsImportTables![Table_Name] = "tbl_Quadrat_Data" Then
                 rsImportTables![Import] = False
            
            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Seedlings_Data" Then
                rsImportTables![Import] = False
            
            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Herbaceous_Data" Then
                rsImportTables![Import] = False
            
            ElseIf rsImportTables![Table_Name] = "tbl_CWD_Data" Then
                rsImportTables![Import] = False
            
            Else
                rsImportTables![Import] = True
            
            End If
            rsImportTables.Update
            
        rsImportTables.MoveNext
        Loop
Case 2

Me.RecordSource = "qry_Append_Secondary_Tablet_Import"
Me.Requery
   
        rsImportTables.MoveFirst
        Do Until rsImportTables.EOF
            rsImportTables.Edit
            If rsImportTables![Table_Name] = "tbl_Locations" Then
                rsImportTables![Import] = False
                
            ElseIf rsImportTables![Table_Name] = "tbl_Events" Then
                rsImportTables![Import] = True
                
            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Data" Then
                rsImportTables![Import] = True
                
            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Seedlings_Data" Then
                rsImportTables![Import] = True
                
            ElseIf rsImportTables![Table_Name] = "tbl_Quadrat_Herbaceous_Data" Then
                rsImportTables![Import] = True
            
            ElseIf rsImportTables![Table_Name] = "tbl_CWD_Data" Then
                rsImportTables![Import] = True
                
            Else
                rsImportTables![Import] = False
                
            End If
            rsImportTables.Update
        rsImportTables.MoveNext
        Loop

End Select
Me.Refresh

End Sub
