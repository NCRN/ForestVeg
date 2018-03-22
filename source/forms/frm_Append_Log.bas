Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ScrollBars =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =1050
    Top =3390
    Right =14445
    Bottom =5835
    DatasheetGridlinesColor =12632256
    OrderBy ="[tsys_Append_Log].[Append_Date] DESC, [tsys_Append_Log].[Table_Name], [tsys_Appe"
        "nd_Log].[Append_Table_Name] DESC"
    RecSrcDt = Begin
        0x3319a2e210b2e340
    End
    RecordSource ="tsys_Append_Log"
    Caption ="Append Log"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =780
                    Top =180
                    ColumnWidth =3870
                    Name ="Table_Name"
                    ControlSource ="Table_Name"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =180
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="Table_Name:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =480
                    Top =720
                    ColumnWidth =1530
                    TabIndex =1
                    Name ="Append_Date"
                    ControlSource ="Append_Date"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =720
                            Width =1125
                            Height =240
                            Name ="Label1"
                            Caption ="Append_Date:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =840
                    Top =1260
                    ColumnWidth =6885
                    TabIndex =2
                    Name ="Append_Table"
                    ControlSource ="Append_Table_Name"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =1260
                            Width =1170
                            Height =240
                            Name ="Label2"
                            Caption ="Append_Table:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =840
                    Top =1680
                    ColumnWidth =2565
                    TabIndex =3
                    Name ="Append_Records_Count"
                    ControlSource ="Append_Records"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Top =1680
                            Width =1890
                            Height =240
                            Name ="Label3"
                            Caption ="Append_Records_Count:"
                        End
                    End
                End
            End
        End
    End
End
