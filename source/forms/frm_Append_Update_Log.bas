Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9300
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =735
    Top =1080
    Right =10320
    Bottom =6315
    DatasheetGridlinesColor =12632256
    OrderBy ="[tsys_Update_Log].[Update_Date] DESC"
    RecSrcDt = Begin
        0xd7107ed611b2e340
    End
    RecordSource ="tsys_Update_Log"
    Caption ="Update Log"
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
            BackColor =15527148
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =1320
                    ColumnWidth =1575
                    Name ="Table_Name"
                    ControlSource ="Table_Name"
                    StatusBarText ="Name of the table being updated"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =1320
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="Table Name:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =1620
                    ColumnWidth =4230
                    TabIndex =1
                    Name ="Update_Table"
                    ControlSource ="Update_Table"
                    StatusBarText ="Name of the table with the new records"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =1620
                            Width =1140
                            Height =240
                            Name ="Label1"
                            Caption ="Update Table:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =1920
                    ColumnWidth =1395
                    TabIndex =2
                    Name ="Update_Date"
                    ControlSource ="Update_Date"
                    Format ="Short Date"
                    StatusBarText ="Date of the update"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =1920
                            Width =1095
                            Height =240
                            Name ="Label2"
                            Caption ="Update Date:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3960
                    Top =2220
                    ColumnWidth =2685
                    TabIndex =3
                    Name ="Update_Records"
                    ControlSource ="Update_Records"
                    StatusBarText ="Number of records updated"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =2220
                            Width =2175
                            Height =240
                            Name ="Label3"
                            Caption ="Number of Records Updated:"
                        End
                    End
                End
            End
        End
    End
End
