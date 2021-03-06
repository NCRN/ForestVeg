Option Compare Database
Option Explicit

' =================================
' MODULE:       fw_mod_App_Enum
' Level:        Application module
' Version:      1.03
' Description:  enum functions & procedures specific to this application
'
' Note:  This module is re-generated by the application from
'        the Enum table when the application is initialized.
'        The framework module mod_Enums & Enum table are used
'        when the CreateEnums() call is made during app initialization.
'        This allows enums to be changed via table vs. hardcoded.
'
'        [_First] & [_Last] enum values are hidden values
'        that allow for iteration w/in the enum
'        (any bracketed string begun with an underscore is hidden)
'
' *************************************************************
' *  IMPORTANT!                                               *
' *                                                           *
' *        DO NOT EDIT this module!                           *
' *        ALL changes WILL be LOST when it is regenerated!   *
' *                                                           *
' *************************************************************
'
' Source/date:  Bonnie Campbell, 11/5/2015
'
' References:  Chip Pearson, November 6, 2013
'              http://www.cpearson.com/excel/Enums.aspx
'
' Revisions:    BLC - 11/5/2015  - 1.00 - initial version
' Revisions:    BLC - 4/12/2015  - 1.01 - revised rs to use SQL to retrieve
'                                         sorted results, .Sort doesn't apply to table recordsets
'                                         added hidden _First & _Last values for @ enum
'               app - 8/4/2017  - 1.02 - latest enum update from db
'                                         last updated: 8/4/2017 6:47:48 AM
'               BLC - 5/16/2019 - 1.03 - added fw_ module prefix
' =================================

'-----------------------------
'  Actions
'-----------------------------
Public Enum Actions
    [_First] = 74
    Observe = 74
    DataEntry = 75
    Verify = 76
    Certify = 77
    Download = 78
    Upload = 79
    Change = 80
    Record = 81
    [_Last] = 81
End Enum

'-----------------------------
'  DirectionFacings
'-----------------------------
Public Enum DirectionFacing
    [_First] = 13
    US = 13
    DS = 14
    RR = 15
    RL = 16
    [_Last] = 16
End Enum

'-----------------------------
'  ModWentworthClassSizes
'-----------------------------
Public Enum ModWentworthClassSize
    [_First] = 64
    f = 64
    CL = 65
    LC = 66
    LO = 67
    SA = 68
    GR = 69
    PE = 70
    CO = 71
    BL = 72
    BR = 73
    [_Last] = 73
End Enum

'-----------------------------
'  PhotoTypes
'-----------------------------
Public Enum PhotoType
    [_First] = 1
    Feature = 1
    Transect = 2
    Overview = 3
    Reference = 4
    Animals = 5
    Plants = 6
    Cultural = 7
    Scenic = 8
    Disturbance = 9
    Weather = 10
    Fieldwork = 11
    other = 12
    [_Last] = 12
End Enum

'-----------------------------
'  PlantTypes
'-----------------------------
Public Enum PlantType
    [_First] = 90
    herb = 90
    Shrub = 91
    Tree = 92
    grass = 93
    sedge = 94
    other = 95
    [_Last] = 95
End Enum

'-----------------------------
'  Priority
'-----------------------------
Public Enum Priority
    [_First] = 51
    Critical = 51
    high = 52
    Medium = 53
    low = 54
    [_Last] = 54
End Enum

'-----------------------------
'  RegExs
'-----------------------------
Public Enum RegEx
    [_First] = 96
    PhotoNumRegex = 96
    [_Last] = 96
End Enum

'-----------------------------
'  Rivers
'-----------------------------
Public Enum River
    [_First] = 20
    CAC = 20
    CBC = 21
    Green = 22
    GAC = 23
    GBC = 24
    Gunnison = 25
    Yampa = 26
    [_Last] = 26
End Enum

'-----------------------------
'  SlopeChangeCauses
'-----------------------------
Public Enum SlopeChangeCause
    [_First] = 55
    Debris = 55
    Ground = 56
    Rock = 57
    Veg = 58
    Water = 59
    [_Last] = 59
End Enum

'-----------------------------
'  Status
'-----------------------------
Public Enum Status
    [_First] = 47
    Opened = 47
    InProgress = 48
    Completed = 49
    Deferred = 50
    [_Last] = 50
End Enum

'-----------------------------
'  SyntaxTypes
'-----------------------------
Public Enum SyntaxType
    [_First] = 97
    T_SQL = 97
    Jet_SQL = 98
    text = 99
    [_Last] = 99
End Enum

'-----------------------------
'  TaglineTypes
'-----------------------------
Public Enum TaglineType
    [_First] = 82
    h = 82
    WRS = 83
    rs = 84
    v = 85
    G = 86
    w = 87
    r = 88
    d = 89
    [_Last] = 89
End Enum

'-----------------------------
'  TaskTypes
'-----------------------------
Public Enum TaskType
    [_First] = 37
    Site = 37
    Feature = 38
    Photo = 39
    Transect = 40
    Plot = 41
    [_Last] = 41
End Enum

'-----------------------------
'  TransducerTypes
'-----------------------------
Public Enum TransducerType
    [_First] = 17
    US = 17
    DS = 18
    Air = 19
    [_Last] = 19
End Enum

'-----------------------------
'  WentworthClassSizes
'-----------------------------
Public Enum WentworthClassSize
    [_First] = 27
    s = 27
    FG = 28
    MG = 29
    CG = 30
    sp = 31
    LP = 32
    sc = 33
    LC = 34
    b = 35
    BED = 36
    [_Last] = 36
End Enum