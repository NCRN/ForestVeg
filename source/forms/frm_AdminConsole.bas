Version =21
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =127
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6360
    DatasheetFontHeight =9
    ItemSuffix =14
    Left =5895
    Top =3525
    Right =12000
    Bottom =5475
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xeaa3a1cb300be340
    End
    Caption ="Administrative Console"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =2220
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =6240
                    Height =390
                    FontSize =14
                    Name ="Label0"
                    Caption ="Administrative Console"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =600
                    Width =2430
                    Height =300
                    Name ="cmdManageLinks"
                    Caption ="Manage Linked Databases"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =180
                    Top =960
                    Width =2430
                    Height =300
                    TabIndex =1
                    Name ="cmdReleaseHistory"
                    Caption ="Add/Edit Release History/Bugs"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4320
                    Top =660
                    Width =1920
                    TabIndex =2
                    Name ="txtProjectName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2880
                            Top =660
                            Width =1260
                            Height =240
                            Name ="Label7"
                            Caption ="Project Name:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =1500
                    Top =1740
                    FontWeight =700
                    TabIndex =3
                    Name ="cmdSave"
                    Caption ="&Save && Close"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =83

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =3180
                    Top =1740
                    FontWeight =700
                    TabIndex =4
                    Name ="cmdCancel"
                    Caption ="&Cancel"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =67

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

Private Sub cmdCancel_Click()
DoCmd.Close acForm, Me.Name
End Sub

Private Sub cmdManageLinks_Click()
DoCmd.OpenForm "frm_ManageLinks"
End Sub

Private Sub cmdReleaseHistory_Click()
DoCmd.OpenForm "frm_App_Releases"
End Sub

Private Sub cmdSave_Click()
Dim strSQL As String

strSQL = "UPDATE tsys_App_Defaults SET Project="
If IsNull(Me!txtProjectName) Then
    strSQL = strSQL & "NULL;"
Else
    strSQL = strSQL & CorrectText(Me!txtProjectName) & ";"
End If
CurrentDb.Execute strSQL
DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Close()
DoCmd.OpenForm "frm_Switchboard"
End Sub

Private Sub Form_Load()
Me!txtProjectName = DLookup("[Project]", "tsys_App_Defaults")
End Sub
