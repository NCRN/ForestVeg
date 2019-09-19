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
    Top =2550
    Right =9675
    Bottom =7710
    DatasheetGridlinesColor =12632256
    OrderBy ="[tsys_Import_Log].[Import_Date] DESC"
    RecSrcDt = Begin
        0x2f53581f0fb2e340
    End
    RecordSource ="tsys_Import_Log"
    Caption ="Import Log"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
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
            Height =2880
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1200
                    Top =300
                    ColumnWidth =5670
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
