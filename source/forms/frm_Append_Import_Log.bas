Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11100
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =2940
    Top =4230
    Right =12705
    Bottom =11805
    DatasheetGridlinesColor =12632256
    OrderBy ="Import_Date DESC"
    RecSrcDt = Begin
        0x2f53581f0fb2e340
    End
    RecordSource ="tsys_Import_Log"
    Caption ="Import Log"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
        End
        Begin Section
            Height =2880
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =300
                    ColumnWidth =4815
                    Name ="Table_Name"
                    ControlSource ="Table_Name"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =300
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="Table_Name:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =600
                    ColumnWidth =1500
                    TabIndex =1
                    Name ="Import_Date"
                    ControlSource ="Import_Date"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =600
                            Width =1050
                            Height =240
                            Name ="Label1"
                            Caption ="Import_Date:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1560
                    Top =900
                    ColumnWidth =1935
                    TabIndex =2
                    Name ="Import_Records_Count"
                    ControlSource ="Import_Records"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =900
                            Width =1440
                            Height =240
                            Name ="Label2"
                            Caption ="Imported Records:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1260
                    Top =1200
                    ColumnWidth =1350
                    TabIndex =3
                    Name ="Delete_Table"
                    ControlSource ="Delete_Table"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1200
                            Width =1140
                            Height =240
                            Name ="Label3"
                            Caption ="Table Deleted?"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =1560
                    ColumnWidth =1230
                    TabIndex =4
                    Name ="Delete_Date"
                    ControlSource ="Delete_Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1560
                            Width =1125
                            Height =240
                            Name ="Label4"
                            Caption =" Date Deleted:"
                        End
                    End
                End
            End
        End
    End
End
