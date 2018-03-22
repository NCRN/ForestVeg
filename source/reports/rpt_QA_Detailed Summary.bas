Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =178
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13920
    DatasheetFontHeight =9
    ItemSuffix =36
    Left =2760
    DatasheetGridlinesColor =15062992
    RecSrcDt = Begin
        0x33ff37304fc0e340
    End
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x550100006801000055010000680100000000000060360000c07b000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =0
            FontSize =11
            FontName ="Calibri"
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =0
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            ShowDatePicker =0
        End
        Begin Subform
            BorderLineStyle =0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =705
            Name ="ReportHeader"
            AutoHeight =1
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =34
                    Top =60
                    Width =8475
                    Height =645
                    FontSize =22
                    FontWeight =700
                    ForeColor =5054976
                    Name ="Auto_Title0"
                    Caption ="NCRN Forest Vegetation QA Summary"
                    FontName ="Segoe UI"
                    LayoutCachedTop =60
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =705
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8550
                    Top =45
                    Width =2625
                    Height =315
                    ColumnOrder =0
                    Name ="Text0"
                    ControlSource ="=Date()"
                    Format ="Short Date"

                    LayoutCachedLeft =8550
                    LayoutCachedTop =45
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =8550
                    Top =360
                    Width =2625
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text1"
                    ControlSource ="=Time()"
                    Format ="Long Time"

                    LayoutCachedLeft =8550
                    LayoutCachedTop =360
                    LayoutCachedWidth =11175
                    LayoutCachedHeight =675
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =31680
            Name ="Detail"
            Begin
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =540
                    Width =11445
                    Height =1039
                    Name ="rSub_QA_qQA_C_Plot_Floor_Incomplete"
                    SourceObject ="Report.rSub_QA_qQA_C_Plot_Floor_Incomplete"

                    LayoutCachedTop =540
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =1579
                End
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =1560
                    Width =11385
                    Height =1290
                    TabIndex =1
                    Name ="rSub_QA_qQA_C_CWD_Incomplete"
                    SourceObject ="Report.rSub_QA_qQA_C_CWD_Incomplete"

                    LayoutCachedTop =1560
                    LayoutCachedWidth =11385
                    LayoutCachedHeight =2850
                End
                Begin Subform
                    Locked = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =2880
                    Width =11400
                    Height =1260
                    TabIndex =2
                    Name ="rSub_QA_qQA_C_Herbaceous_Incomplete"
                    SourceObject ="Report.rSub_QA_qQA_C_Herbaceous_Incomplete"

                    LayoutCachedTop =2880
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =4140
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =60
                    Top =7020
                    Width =11275
                    Height =1399
                    TabIndex =3
                    Name ="rSub_QA_qQA_C_Sapling_Incomplete"
                    SourceObject ="Report.rSub_QA_qQA_C_Live_Sapling_Incomplete"

                    LayoutCachedLeft =60
                    LayoutCachedTop =7020
                    LayoutCachedWidth =11335
                    LayoutCachedHeight =8419
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =9780
                    Width =10980
                    Height =1557
                    TabIndex =4
                    Name ="rSub_QA_qQA_C_Sapling_GT_10cm_DBH"
                    SourceObject ="Report.rSub_QA_qQA_C_Sapling_GT_10cm_DBH"

                    LayoutCachedLeft =120
                    LayoutCachedTop =9780
                    LayoutCachedWidth =11100
                    LayoutCachedHeight =11337
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =4140
                    Width =11400
                    Height =1482
                    TabIndex =5
                    Name ="rSub_QA_qQA_C_Seedlings_Height_Incomplete"
                    SourceObject ="Report.rSub_QA_qQA_C_Seedlings_Height_Incomplete"

                    LayoutCachedTop =4140
                    LayoutCachedWidth =11400
                    LayoutCachedHeight =5622
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Top =12840
                    Width =11380
                    Height =1445
                    TabIndex =6
                    Name ="rSub_QA_qQA_C_Tree_Crown_Class_or_Status_or_Checked_Missing"
                    SourceObject ="Report.rSub_QA_qQA_C_Live_Tree_Incomplete"

                    LayoutCachedTop =12840
                    LayoutCachedWidth =11380
                    LayoutCachedHeight =14285
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =15480
                    Width =11115
                    Height =1505
                    TabIndex =7
                    Name ="rSub_QA_qQA_C_Tree_LT_10cm_DBH"
                    SourceObject ="Report.rSub_QA_qQA_C_Tree_LT_10cm_DBH"

                    LayoutCachedLeft =120
                    LayoutCachedTop =15480
                    LayoutCachedWidth =11235
                    LayoutCachedHeight =16985
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =60
                    Top =28680
                    Width =11385
                    Height =1535
                    TabIndex =8
                    Name ="rSub_QA_qQA_I_Find duplicates for tbl_Quadrat_Data"
                    SourceObject ="Report.rSub_QA_qQA_I_Find duplicates for tbl_Quadrat_Data"
                    EventProcPrefix ="rSub_QA_qQA_I_Find_duplicates_for_tbl_Quadrat_Data"

                    LayoutCachedLeft =60
                    LayoutCachedTop =28680
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =30215
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =60
                    Top =30190
                    Width =11355
                    Height =1490
                    TabIndex =9
                    Name ="rSub_QA_qQA_I_Photos_Not_Taken"
                    SourceObject ="Report.rSub_QA_qQA_I_Photos_Not_Taken"

                    LayoutCachedLeft =60
                    LayoutCachedTop =30190
                    LayoutCachedWidth =11415
                    LayoutCachedHeight =31680
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =19200
                    Width =11310
                    Height =1482
                    TabIndex =10
                    Name ="rSub_QA_qQA_W_Missing_or_Extra_Quadrat_Records"
                    SourceObject ="Report.rSub_QA_qQA_W_Missing_or_Extra_Quadrat_Records"

                    LayoutCachedLeft =120
                    LayoutCachedTop =19200
                    LayoutCachedWidth =11430
                    LayoutCachedHeight =20682
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =20760
                    Width =11456
                    Height =1482
                    TabIndex =11
                    Name ="rSub_QA_qQA_W_Sapling_AND_Tree_Records_Exist"
                    SourceObject ="Report.rSub_QA_qQA_W_Sapling_AND_Tree_Records_Exist"

                    LayoutCachedLeft =120
                    LayoutCachedTop =20760
                    LayoutCachedWidth =11576
                    LayoutCachedHeight =22242
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =22290
                    Width =11325
                    Height =1430
                    TabIndex =12
                    Name ="rSub_QA_qQA_W_Seedling_Missing_or_Non-accepted_TSN"
                    SourceObject ="Report.rSub_QA_qQA_W_Seedling_Missing_or_Non-accepted_TSN"
                    EventProcPrefix ="rSub_QA_qQA_W_Seedling_Missing_or_Non_accepted_TSN"

                    LayoutCachedLeft =120
                    LayoutCachedTop =22290
                    LayoutCachedWidth =11445
                    LayoutCachedHeight =23720
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =23715
                    Width =11460
                    Height =1482
                    TabIndex =13
                    Name ="rSub_QA_qQA_W_Tag_Record_Incomplete_or_TSN_Not_Accepted"
                    SourceObject ="Report.rSub_QA_qQA_W_Tag_Record_Incomplete_or_TSN_Not_Accepted"

                    LayoutCachedLeft =120
                    LayoutCachedTop =23715
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =25197
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =17220
                    Width =11370
                    Height =1445
                    TabIndex =14
                    Name ="rSub_QA_qQA_C_Quadrat_Conditions_Incomplete"
                    SourceObject ="Report.rSub_QA_qQA_C_Quadrat_Conditions_Incomplete"

                    LayoutCachedLeft =120
                    LayoutCachedTop =17220
                    LayoutCachedWidth =11490
                    LayoutCachedHeight =18665
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =26655
                    Width =11460
                    Height =1482
                    TabIndex =15
                    Name ="Child21"
                    SourceObject ="Report.rSub_QA_qQA_W_Current_Year_Measured_As_Sapl_Status_Not_Sapl"

                    LayoutCachedLeft =120
                    LayoutCachedTop =26655
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =28137
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =25155
                    Width =11460
                    Height =1482
                    TabIndex =16
                    Name ="Child22"
                    SourceObject ="Report.rSub_QA_qQA_W_Current_Year_Measured_As_Tree_Status_Not_Tree"

                    LayoutCachedLeft =120
                    LayoutCachedTop =25155
                    LayoutCachedWidth =11580
                    LayoutCachedHeight =26637
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    Left =120
                    Top =14280
                    Width =11380
                    Height =1205
                    TabIndex =17
                    Name ="Child24"
                    SourceObject ="Report.rSub_QA_qQA_C_Dead_Tree_Incomplete"

                    LayoutCachedLeft =120
                    LayoutCachedTop =14280
                    LayoutCachedWidth =11500
                    LayoutCachedHeight =15485
                End
                Begin Subform
                    OldBorderStyle =0
                    Left =120
                    Top =11340
                    Width =10673
                    Height =1446
                    TabIndex =18
                    Name ="rSub_QA_qQA_C_All_Trees_Required"
                    SourceObject ="Report.rSub_QA_qQA_C_All_Trees_Required"

                    LayoutCachedLeft =120
                    LayoutCachedTop =11340
                    LayoutCachedWidth =10793
                    LayoutCachedHeight =12786
                End
                Begin Subform
                    Left =60
                    Top =8400
                    Width =11220
                    Height =1377
                    TabIndex =19
                    Name ="rSub_QA_qQA_C_Dead_Sapling_Incomplete"
                    SourceObject ="Report.rSub_QA_qQA_C_Dead_Sapling_Incomplete"

                    LayoutCachedLeft =60
                    LayoutCachedTop =8400
                    LayoutCachedWidth =11280
                    LayoutCachedHeight =9777
                End
                Begin Subform
                    Top =5580
                    Width =11220
                    Height =1445
                    TabIndex =20
                    Name ="rSub_QA_qQA_C_All_Saplings_Required"
                    SourceObject ="Report.rSub_QA_qQA_C_All_Saplings_Required"

                    LayoutCachedTop =5580
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =7025
                End
                Begin Label
                    TextFontFamily =34
                    Left =180
                    Top =60
                    Width =1785
                    Height =420
                    FontSize =16
                    Name ="Label32"
                    Caption ="Critical Errors"
                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =1965
                    LayoutCachedHeight =480
                End
                Begin Label
                    TextFontFamily =34
                    Left =300
                    Top =18720
                    Width =1995
                    Height =420
                    FontSize =16
                    Name ="Label34"
                    Caption ="Warning Errors"
                    LayoutCachedLeft =300
                    LayoutCachedTop =18720
                    LayoutCachedWidth =2295
                    LayoutCachedHeight =19140
                End
                Begin Label
                    TextFontFamily =34
                    Left =240
                    Top =28200
                    Width =2640
                    Height =420
                    FontSize =16
                    Name ="Label35"
                    Caption ="Informational Errors"
                    LayoutCachedLeft =240
                    LayoutCachedTop =28200
                    LayoutCachedWidth =2880
                    LayoutCachedHeight =28620
                End
            End
        End
        Begin PageFooter
            Height =840
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3960
                    Top =480
                    Width =3600
                    Height =315
                    Name ="Text2"
                    ControlSource ="=\"Page \" & [Page]"

                    LayoutCachedLeft =3960
                    LayoutCachedTop =480
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =795
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
